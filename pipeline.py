"""
Background processing job.

Runs in its own thread so the Streamlit UI stays responsive (live progress,
working Stop button). All Google calls happen here, never in the UI thread.

Phases
  1. connect      verify Drive access, ensure Input/ and Output/ folders
  2. parse        read every file locally (legacy .doc converted through Drive)
  3. serials      reserve a contiguous QBGDOC range (Output/_serial_counter.json
                  + process-wide lock)
  4. originals    upload each source file to Input/<Category>/<File Name>.docx
  5. questions    build each question .docx and upload it, unconverted, to
                  Output/<Category>/<File Name>/QBGDOC000001.docx (thread pool)
  6. record       when a file finishes, its question-level Excel is uploaded to
                  Output/<Category>/<File Name>.xlsx
"""
from __future__ import annotations

import csv
import io
import os
import threading
import time
import traceback
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional

try:  # package layout (qbg/...) or flat layout (all files next to app.py)
    from .builder import QuestionDocBuilder
    from .google_services import DOC_MIME, DOCX_MIME, XLSX_MIME, GoogleDrive
    from .parser import Question, parse_docx
except ImportError:
    from builder import QuestionDocBuilder
    from google_services import DOC_MIME, DOCX_MIME, XLSX_MIME, GoogleDrive
    from parser import Question, parse_docx

IST = timezone(timedelta(hours=5, minutes=30))
SERIAL_PREFIX = "QBGDOC"
SERIAL_WIDTH = 6

# One Streamlit process serves every visitor, so this lock makes serial
# allocation collision-free across simultaneous users of the deployment.
_SERIAL_LOCK = threading.Lock()
_SERIAL_HIGH_WATER = {"n": 0}


def now_ist() -> str:
    return datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")


def fmt_serial(n: int) -> str:
    return f"{SERIAL_PREFIX}{n:0{SERIAL_WIDTH}d}"


@dataclass
class FileTask:
    index: int
    uploaded_name: str
    data: bytes
    category: str
    file_name: str

    @property
    def ext(self) -> str:
        return os.path.splitext(self.uploaded_name)[1].lower()


@dataclass
class ResultRow:
    category: str
    file_name: str
    serial: str
    original_link: str
    doc_link: str
    status: str
    remarks: str
    file_index: int = 0
    order: int = 0

    def as_list(self) -> List[str]:
        return [self.category, self.file_name, self.serial, self.original_link,
                self.doc_link, self.status, self.remarks]


@dataclass
class FileState:
    task: FileTask
    questions: List[Question] = field(default_factory=list)
    answer_keys: int = 0
    serial_start: int = 0
    original_link: str = ""
    output_folder: str = ""
    category_out_folder: str = ""
    excel_link: str = ""
    error: str = ""
    pending: int = 0
    rows: Dict[int, ResultRow] = field(default_factory=dict)
    flushed: bool = False


class Job:
    def __init__(self, tasks: List[FileTask], workspace_factory, *, language: str = "English",
                 qtype: str = "SCQ", max_workers: int = 12):
        self.id = datetime.now(IST).strftime("RUN-%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:4].upper()
        self.tasks = tasks
        self.workspace_factory = workspace_factory
        self.language = language
        self.qtype = qtype
        self.max_workers = max_workers
        self.stop_event = threading.Event()
        self.lock = threading.Lock()

        self.status = "queued"          # queued | running | stopping | stopped | done | error
        self.phase = "Waiting to start"
        self.error = ""
        self.started_at: Optional[float] = None
        self.finished_at: Optional[float] = None
        self.files: List[FileState] = [FileState(task=t) for t in tasks]
        self.total_detected = 0
        self.total_answer_keys = 0
        self.formatted = 0
        self.failed = 0
        self.processed = 0
        self.log: List[str] = []
        self.thread: Optional[threading.Thread] = None
        self._out_root: Optional[str] = None

    # ------------------------------------------------------------------ #
    def start(self):
        self.status = "running"
        self.started_at = time.time()
        self.thread = threading.Thread(target=self._run, name=f"job-{self.id}", daemon=True)
        self.thread.start()

    def stop(self):
        if self.status in ("running", "queued"):
            self.stop_event.set()
            with self.lock:
                self.status = "stopping"
                self.phase = "Stopping: finishing the uploads already in flight"
            self._log("Stop requested by user")

    @property
    def active(self) -> bool:
        return self.status in ("queued", "running", "stopping")

    @property
    def elapsed(self) -> float:
        if not self.started_at:
            return 0.0
        return (self.finished_at or time.time()) - self.started_at

    def _log(self, msg: str):
        with self.lock:
            self.log.append(f"{datetime.now(IST).strftime('%H:%M:%S')}  {msg}")
            self.log = self.log[-300:]

    def _set_phase(self, msg: str):
        with self.lock:
            self.phase = msg
        self._log(msg)

    # ------------------------------------------------------------------ #
    def results(self) -> List[ResultRow]:
        out: List[ResultRow] = []
        with self.lock:
            for fs in self.files:
                out.extend(fs.rows[k] for k in sorted(fs.rows))
        return out

    # ------------------------------------------------------------------ #
    def _run(self):
        ws: Optional[GoogleDrive] = None
        try:
            self._set_phase("Connecting to Google Drive")
            ws = self.workspace_factory()
            ws.check_access()
            in_root = ws.ensure_folder("Input", ws.root)
            out_root = ws.ensure_folder("Output", ws.root)
            self._out_root = out_root

            # ---- parse ------------------------------------------------- #
            self._set_phase(f"Reading {len(self.files)} file(s)")
            for fs in self.files:
                if self.stop_event.is_set():
                    break
                t = fs.task
                try:
                    data = t.data
                    if t.ext == ".doc":
                        data = ws.doc_to_docx(t.data, t.uploaded_name, in_root)
                    fs.questions, fs.answer_keys = parse_docx(data)
                    if not fs.questions:
                        fs.error = ("No questions detected. Check that questions are numbered "
                                    "(1., 2., ...) and each has an 'Ans.' / 'Answer:' line.")
                except Exception as e:  # corrupt / password-protected / not a Word file
                    fs.error = f"File could not be read as a Word document ({type(e).__name__}: {e})"
                with self.lock:
                    self.total_detected += len(fs.questions)
                    self.total_answer_keys += fs.answer_keys
                msg = f"{t.file_name}: {len(fs.questions)} question(s) detected"
                if fs.answer_keys and fs.answer_keys != len(fs.questions):
                    msg += f" (answer keys in file: {fs.answer_keys} — check remarks)"
                self._log(msg if not fs.error else f"{t.file_name}: {fs.error}")

            # ---- serials ----------------------------------------------- #
            self._set_phase("Reserving question serial numbers")
            need = sum(len(fs.questions) for fs in self.files)
            with _SERIAL_LOCK:
                first = ws.reserve_serials(out_root, SERIAL_PREFIX, need,
                                           floor=_SERIAL_HIGH_WATER["n"]) if need else 0
                last = first - 1
                cursor = first
                for fs in self.files:
                    fs.serial_start = cursor
                    cursor += len(fs.questions)
                _SERIAL_HIGH_WATER["n"] = max(_SERIAL_HIGH_WATER["n"], last + need)
            if need:
                self._log(f"Serials {fmt_serial(last + 1)} to {fmt_serial(last + need)} reserved")

            # ---- originals + folders ----------------------------------- #
            self._set_phase("Uploading original files to Drive")

            def upload_original(fs: FileState):
                t = fs.task
                cat_in = ws.ensure_folder(t.category, in_root)
                mime = DOC_MIME if t.ext == ".doc" else DOCX_MIME
                _, link = ws.upload_file(t.data, f"{t.file_name}{t.ext}", cat_in, mime)
                fs.original_link = link
                fs.category_out_folder = ws.ensure_folder(t.category, out_root)
                if fs.questions:
                    fs.output_folder = ws.ensure_folder(t.file_name, fs.category_out_folder)

            with ThreadPoolExecutor(max_workers=min(4, len(self.files)) or 1) as pool:
                futs = {pool.submit(upload_original, fs): fs for fs in self.files
                        if not self.stop_event.is_set()}
                for f in as_completed(futs):
                    fs = futs[f]
                    try:
                        f.result()
                    except Exception as e:
                        fs.error = fs.error or f"Original upload failed: {e}"
                        self._log(f"{fs.task.file_name}: original upload failed: {e}")

            # ---- questions --------------------------------------------- #
            for fs in self.files:
                fs.pending = len(fs.questions)
                if fs.error or not fs.questions:
                    self._file_level_rows(fs)
                    self._flush_file(ws, fs)

            work = [(fs, k) for fs in self.files if not fs.error for k in range(len(fs.questions))]
            self._set_phase(f"Formatting and uploading {len(work)} question doc(s)")

            builders = {}

            def process(fs: FileState, k: int) -> ResultRow:
                q = fs.questions[k]
                t = fs.task
                serial = fmt_serial(fs.serial_start + k)
                base = dict(category=t.category, file_name=t.file_name, serial=serial,
                            original_link=fs.original_link, file_index=t.index, order=k)
                if self.stop_event.is_set():
                    return ResultRow(doc_link="", status="Failure",
                                     remarks="Not processed: run stopped by user", **base)
                if not q.ok:
                    return ResultRow(doc_link="", status="Failure", remarks=q.remarks(), **base)
                b = builders.get(t.category)
                if b is None:
                    b = builders[t.category] = QuestionDocBuilder(t.category, self.language, self.qtype)
                try:
                    data = b.build(q)
                    _, link = ws.upload_file(data, f"{serial}.docx", fs.output_folder, DOCX_MIME)
                    return ResultRow(doc_link=link, status="Success", remarks=q.remarks(), **base)
                except Exception as e:
                    return ResultRow(doc_link="", status="Failure",
                                     remarks=f"Source {q.label}: doc upload failed ({e})", **base)

            with ThreadPoolExecutor(max_workers=self.max_workers) as pool:
                futs = {pool.submit(process, fs, k): (fs, k) for fs, k in work}
                for f in as_completed(futs):
                    fs, k = futs[f]
                    try:
                        row = f.result()
                    except Exception as e:
                        row = ResultRow(fs.task.category, fs.task.file_name,
                                        fmt_serial(fs.serial_start + k), fs.original_link, "",
                                        "Failure", f"Unexpected error: {e}", fs.task.index, k)
                    with self.lock:
                        fs.rows[k] = row
                        fs.pending -= 1
                        self.processed += 1
                        if row.status == "Success":
                            self.formatted += 1
                        else:
                            self.failed += 1
                        done = fs.pending == 0
                    if done:
                        self._flush_file(ws, fs)

            for fs in self.files:          # anything left (stop mid-way)
                if not fs.flushed:
                    self._flush_file(ws, fs)

            with self.lock:
                self.finished_at = time.time()
                if self.stop_event.is_set():
                    self.status, self.phase = "stopped", "Stopped. Finished questions and their Excel record are saved in Drive."
                else:
                    self.status, self.phase = "done", "All files processed"
            self._log(f"Finished: {self.formatted} formatted, {self.failed} failed, "
                      f"in {self.elapsed:.1f}s")
        except Exception as e:
            with self.lock:
                self.status = "error"
                self.error = f"{type(e).__name__}: {e}"
                self.phase = "Stopped by an error"
                self.finished_at = time.time()
            self._log("ERROR " + self.error)
            self._log(traceback.format_exc(limit=3))
            if ws is not None:
                for fs in self.files:
                    try:
                        if not fs.flushed and fs.rows:
                            self._flush_file(ws, fs)
                    except Exception:
                        pass

    # ------------------------------------------------------------------ #
    def _file_level_rows(self, fs: FileState):
        """A file with no usable questions still gets an audit row."""
        t = fs.task
        if fs.error and fs.questions:
            for k, q in enumerate(fs.questions):
                fs.rows[k] = ResultRow(t.category, t.file_name, fmt_serial(fs.serial_start + k),
                                       fs.original_link, "", "Failure",
                                       f"Source {q.label}: {fs.error}", t.index, k)
                self.processed += 1
                self.failed += 1
            fs.pending = 0
        elif not fs.questions:
            fs.rows[0] = ResultRow(t.category, t.file_name, "", fs.original_link, "", "Failure",
                                   fs.error or "No questions detected", t.index, 0)

    def _flush_file(self, ws: GoogleDrive, fs: FileState):
        with self.lock:
            if fs.flushed:
                return
            fs.flushed = True
            t = fs.task
            # rows for questions never reached (stop) so serials stay accounted for
            for k in range(len(fs.questions)):
                if k not in fs.rows:
                    fs.rows[k] = ResultRow(t.category, t.file_name, fmt_serial(fs.serial_start + k),
                                           fs.original_link, "", "Failure",
                                           "Not processed: run stopped by user", t.index, k)
                    self.processed += 1
                    self.failed += 1
            rows = [fs.rows[k] for k in sorted(fs.rows)]
            ok = sum(1 for r in rows if r.status == "Success")
        n_q = len(fs.questions)
        status = ("Failure" if (fs.error or ok == 0) else
                  "Success" if ok == n_q else "Partial")
        summary = {
            "Run ID": self.id, "Processed at (IST)": now_ist(), "QBG Category": t.category,
            "File Name": t.file_name, "Uploaded file": t.uploaded_name,
            "Original doc link": fs.original_link, "Total questions given": n_q,
            "Total questions formatted": ok, "Failed / not processed": n_q - ok,
            "Answer keys found in file": fs.answer_keys, "File status": status,
        }
        if fs.error:
            summary["Remarks"] = fs.error
        if not fs.category_out_folder:
            if not self._out_root:
                return
            try:
                fs.category_out_folder = ws.ensure_folder(t.category, self._out_root)
            except Exception as e:
                self._log(f"{t.file_name}: could not create the Output folder ({e})")
                return
        try:
            data = to_excel(rows, summary)
            name = f"{t.file_name}.xlsx"
            if ws.find_file(name, fs.category_out_folder):
                name = f"{t.file_name} ({self.id}).xlsx"     # never overwrite an earlier record
            _, link = ws.upload_file(data, name, fs.category_out_folder, XLSX_MIME)
            with self.lock:
                fs.excel_link = link
            self._log(f"{t.file_name}: Excel record saved to Drive ({ok}/{n_q} formatted)")
        except Exception as e:
            self._log(f"{t.file_name}: could not save the Excel record to Drive ({e}); "
                      f"use the download button to keep these results")


# ---------------------------------------------------------------------- #
# Exports
# ---------------------------------------------------------------------- #
EXPORT_HEADERS = ["QBG Category", "File Name", "Question Serial No.", "Original Doc Link",
                  "Question Doc Link", "Status", "Remarks"]


def to_csv(rows: List[ResultRow]) -> bytes:
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(EXPORT_HEADERS)
    for r in rows:
        w.writerow(r.as_list())
    return ("\ufeff" + buf.getvalue()).encode("utf-8")   # BOM so Excel opens UTF-8 cleanly


def to_excel(rows: List[ResultRow], summary: Dict[str, object]) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Output"
    ws.append(EXPORT_HEADERS)
    head_fill = PatternFill("solid", fgColor="1F2A5C")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
        c.alignment = Alignment(vertical="center")
    ok_font, bad_font = Font(color="1E7B4F", bold=True), Font(color="B42318", bold=True)
    link_font = Font(color="2348C8", underline="single")
    for r in rows:
        ws.append(r.as_list())
        row = ws.max_row
        for col in (4, 5):
            cell = ws.cell(row=row, column=col)
            if cell.value:
                cell.hyperlink = cell.value
                cell.font = link_font
        ws.cell(row=row, column=6).font = ok_font if r.status == "Success" else bad_font
    widths = [18, 28, 20, 48, 48, 11, 70]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{max(ws.max_row, 1)}"

    s = wb.create_sheet("Summary")
    for k, v in summary.items():
        s.append([k, v])
    s.column_dimensions["A"].width = 32
    s.column_dimensions["B"].width = 40
    for c in s["A"]:
        c.font = Font(bold=True)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
