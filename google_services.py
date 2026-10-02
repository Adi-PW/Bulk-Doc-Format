"""
Google Drive access (Drive is the only Google dependency).

Auth (configured in Streamlit secrets, see README):

* ``[google_oauth]``  client_id / client_secret / refresh_token   (recommended)
  Files are owned by the real Google account that owns the Drive folder.
* ``[gcp_service_account]``  service-account JSON fields
  Works ONLY if the root folder lives in a *Shared Drive* the service account
  is a member of: service accounts have no My Drive storage quota, so uploads
  into a normal My Drive folder fail with ``storageQuotaExceeded``.
  Optional ``impersonate = "user@pw.live"`` uses domain-wide delegation.

Question docs are uploaded as plain .docx (no Google Docs conversion), which is
the fastest upload Drive offers. googleapiclient's HTTP transport is not
thread-safe, so every worker thread gets its own Drive client.
"""
from __future__ import annotations

import io
import json
import random
import re
import threading
import time
from typing import Dict, Optional, Tuple

from google.oauth2 import service_account
from google.oauth2.credentials import Credentials as UserCredentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]
FOLDER_MIME = "application/vnd.google-apps.folder"
GDOC_MIME = "application/vnd.google-apps.document"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
DOC_MIME = "application/msword"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
COUNTER_NAME = "_serial_counter.json"


def build_credentials(secrets) -> object:
    if "google_oauth" in secrets:
        s = secrets["google_oauth"]
        return UserCredentials(
            token=None, refresh_token=s["refresh_token"], client_id=s["client_id"],
            client_secret=s["client_secret"], token_uri="https://oauth2.googleapis.com/token",
            scopes=SCOPES)
    if "gcp_service_account" in secrets:
        info = dict(secrets["gcp_service_account"])
        subject = info.pop("impersonate", None)
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        return creds.with_subject(subject) if subject else creds
    raise RuntimeError("No Google credentials found in secrets "
                       "(add [google_oauth] or [gcp_service_account]).")


def _retry(fn, *, tries: int = 7, base: float = 0.8):
    """Exponential backoff with jitter for rate limits and transient errors."""
    for attempt in range(tries):
        try:
            return fn()
        except HttpError as e:
            status = getattr(e, "status_code", None) or (e.resp.status if e.resp else 0)
            reason = str(e)
            transient = status in (429, 500, 502, 503, 504) or (
                status == 403 and re.search(r"rate ?limit|userRateLimitExceeded|quota", reason, re.I)
                and "storageQuotaExceeded" not in reason)
            if not transient or attempt == tries - 1:
                raise
        except (ConnectionError, TimeoutError, OSError):
            if attempt == tries - 1:
                raise
        time.sleep(min(base * (2 ** attempt), 30) + random.random() * 0.5)


def _q(s: str) -> str:
    return s.replace("\\", "\\\\").replace("'", "\\'")


class GoogleDrive:
    def __init__(self, credentials, root_folder_id: str):
        self.creds = credentials
        self.root = root_folder_id
        self._local = threading.local()
        self._folder_lock = threading.Lock()
        self._folder_cache: Dict[Tuple[str, str], str] = {}

    @property
    def drive(self):
        d = getattr(self._local, "drive", None)
        if d is None:
            d = build("drive", "v3", credentials=self.creds, cache_discovery=False)
            self._local.drive = d
        return d

    # ---- lookups ---------------------------------------------------------- #
    def check_access(self) -> str:
        meta = _retry(lambda: self.drive.files().get(
            fileId=self.root, fields="id,name", supportsAllDrives=True).execute())
        return meta.get("name", "")

    def find_file(self, name: str, parent: str, folder: bool = False) -> Optional[str]:
        mime = f"mimeType {'=' if folder else '!='} '{FOLDER_MIME}'"
        q = f"name = '{_q(name)}' and '{parent}' in parents and {mime} and trashed = false"
        res = _retry(lambda: self.drive.files().list(
            q=q, fields="files(id)", pageSize=5, supportsAllDrives=True,
            includeItemsFromAllDrives=True).execute())
        files = res.get("files", [])
        return files[0]["id"] if files else None

    def ensure_folder(self, name: str, parent: str) -> str:
        name = name.strip() or "Untitled"
        key = (parent, name.lower())
        with self._folder_lock:
            if key in self._folder_cache:
                return self._folder_cache[key]
            fid = self.find_file(name, parent, folder=True)
            if not fid:
                meta = {"name": name, "mimeType": FOLDER_MIME, "parents": [parent]}
                fid = _retry(lambda: self.drive.files().create(
                    body=meta, fields="id", supportsAllDrives=True).execute())["id"]
            self._folder_cache[key] = fid
            return fid

    # ---- files ------------------------------------------------------------ #
    def upload_file(self, data: bytes, name: str, parent: str, mime: str,
                    convert_to_gdoc: bool = False) -> Tuple[str, str]:
        meta = {"name": name, "parents": [parent]}
        if convert_to_gdoc:
            meta["mimeType"] = GDOC_MIME

        def go():
            media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=False)
            return self.drive.files().create(body=meta, media_body=media, fields="id",
                                             supportsAllDrives=True).execute()
        fid = _retry(go)["id"]
        if convert_to_gdoc:
            return fid, f"https://docs.google.com/document/d/{fid}/edit"
        return fid, f"https://drive.google.com/file/d/{fid}/view"

    def update_file(self, file_id: str, data: bytes, mime: str):
        def go():
            media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=False)
            return self.drive.files().update(fileId=file_id, media_body=media, fields="id",
                                             supportsAllDrives=True).execute()
        _retry(go)

    def download(self, file_id: str) -> bytes:
        return _retry(lambda: self.drive.files().get_media(
            fileId=file_id, supportsAllDrives=True).execute())

    def doc_to_docx(self, data: bytes, name: str, parent: str) -> bytes:
        """Legacy .doc -> Google Doc -> .docx export (temp file is deleted)."""
        fid, _ = self.upload_file(data, f"__tmp__{name}", parent, DOC_MIME, convert_to_gdoc=True)
        try:
            return _retry(lambda: self.drive.files().export(fileId=fid, mimeType=DOCX_MIME).execute())
        finally:
            try:
                self.drive.files().delete(fileId=fid, supportsAllDrives=True).execute()
            except Exception:
                pass

    # ---- serial counter --------------------------------------------------- #
    def _max_serial_in_folder_tree(self, prefix: str) -> int:
        """Safety net if the counter file is missing: highest QBGDOC name in Drive."""
        try:
            res = _retry(lambda: self.drive.files().list(
                q=f"name contains '{_q(prefix)}' and trashed = false",
                orderBy="name desc", pageSize=20, fields="files(name)",
                supportsAllDrives=True, includeItemsFromAllDrives=True,
                corpora="allDrives").execute())
        except Exception:
            return 0
        best = 0
        for f in res.get("files", []):
            m = re.match(rf"^{re.escape(prefix)}(\d+)", f.get("name", ""))
            if m:
                best = max(best, int(m.group(1)))
        return best

    def reserve_serials(self, output_root: str, prefix: str, count: int, floor: int = 0) -> int:
        """Reserve ``count`` serials; returns the first number. The counter file is
        written *before* any upload, so a crash or stop can never reuse a serial."""
        fid = self.find_file(COUNTER_NAME, output_root)
        last = 0
        if fid:
            try:
                last = int(json.loads(self.download(fid).decode("utf-8")).get("last_serial", 0))
            except Exception:
                last = 0
        if not fid or last == 0:
            last = max(last, self._max_serial_in_folder_tree(prefix))
        last = max(last, floor)
        payload = json.dumps({"prefix": prefix, "last_serial": last + count,
                              "updated_at": time.strftime("%Y-%m-%d %H:%M:%S")}, indent=2).encode()
        if fid:
            self.update_file(fid, payload, "application/json")
        else:
            self.upload_file(payload, COUNTER_NAME, output_root, "application/json")
        return last + 1
