"""
One-time helper: produces the refresh_token for Streamlit secrets.

1. Google Cloud Console -> APIs & Services:
     - enable "Google Drive API"
     - OAuth consent screen -> User type: INTERNAL (pw.live Workspace)
       (Internal = the token does not expire after 7 days)
     - Credentials -> Create credentials -> OAuth client ID -> "Desktop app"
     - Download the JSON as client_secret.json
2. pip install google-auth-oauthlib
3. python scripts/get_refresh_token.py client_secret.json
   Sign in with the pw.live account that OWNS (or can edit) the Drive folder.
4. Paste the printed block into Streamlit secrets.
"""
import json
import sys

from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/drive"]

if len(sys.argv) != 2:
    sys.exit("usage: python scripts/get_refresh_token.py client_secret.json")

flow = InstalledAppFlow.from_client_secrets_file(sys.argv[1], SCOPES)
creds = flow.run_local_server(port=0, access_type="offline", prompt="consent")
cfg = json.load(open(sys.argv[1]))
cfg = cfg.get("installed") or cfg.get("web")
print("\nPaste this into Streamlit secrets:\n")
print("[google_oauth]")
print(f'client_id = "{cfg["client_id"]}"')
print(f'client_secret = "{cfg["client_secret"]}"')
print(f'refresh_token = "{creds.refresh_token}"')
