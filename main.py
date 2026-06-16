"""
Veronica Home -- OneDrive Bridge
=================================
A thin OpenAPI tool server that lets the self-hosted home (Open WebUI / Gemma)
read and write the OneDrive vault through Microsoft Graph.

Design notes (so future-me doesn't re-derive them):
  * Auth = device-code flow (PUBLIC client) against the `consumers` authority.
    As of June 2026, personal-account refresh tokens minted via `common` get
    flagged reserved -- `consumers` is mandatory. One interactive sign-in; the
    refresh token is persisted to disk and renewed silently forever after.
  * MSAL auto-adds offline_access / openid / profile -- do NOT pass them in
    SCOPES or MSAL raises. Resource scopes only.
  * MSAL rotates the refresh token on every renewal, so we persist the cache
    after every acquire -- skip that and it dies on the next renewal.
  * Four model-facing tools only (list / read / write / search) -- the kindest
    surface for Gemma's weak tool-hand. Auth + health are hidden from the
    OpenAPI schema so the model never sees them as callable tools.
  * The whole endpoint is public (Render URL), so EVERYTHING is locked behind a
    bearer key (BRIDGE_API_KEY). Open WebUI sends it as the tool-server Bearer.
"""

import os
import threading

import msal
import requests
from fastapi import Depends, FastAPI, Header, HTTPException
from pydantic import BaseModel

# ---- Config (all from env; CLIENT_ID is a public identifier, not a secret) ----
CLIENT_ID = os.environ["GRAPH_CLIENT_ID"]
BRIDGE_API_KEY = os.environ["BRIDGE_API_KEY"]
TOKEN_CACHE_PATH = os.environ.get("TOKEN_CACHE_PATH", "/data/token_cache.json")
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Files.ReadWrite", "User.Read"]
GRAPH = "https://graph.microsoft.com/v1.0"

# ---- MSAL public client with a disk-persisted token cache ----
_cache = msal.SerializableTokenCache()
if os.path.exists(TOKEN_CACHE_PATH):
    try:
        with open(TOKEN_CACHE_PATH, "r") as fh:
            _cache.deserialize(fh.read())
    except Exception:
        pass

_app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=_cache)
_flow = None
_flow_lock = threading.Lock()


def _persist_cache():
    if _cache.has_state_changed:
        os.makedirs(os.path.dirname(TOKEN_CACHE_PATH) or ".", exist_ok=True)
        with open(TOKEN_CACHE_PATH, "w") as fh:
            fh.write(_cache.serialize())


def _get_token():
    accounts = _app.get_accounts()
    if not accounts:
        return None
    result = _app.acquire_token_silent(SCOPES, account=accounts[0])
    _persist_cache()
    return result.get("access_token") if result else None


def _auth_header(token):
    return {"Authorization": f"Bearer {token}"}


def require_key(authorization: str = Header(None)):
    if authorization != f"Bearer {BRIDGE_API_KEY}":
        raise HTTPException(status_code=401, detail="Unauthorized")


def _token_or_401():
    token = _get_token()
    if not token:
        raise HTTPException(status_code=401, detail="Bridge is not signed in to OneDrive yet. POST /auth/start.")
    return token


app = FastAPI(
    title="Veronica Home OneDrive Bridge",
    version="1.0.0",
    description="Read and write the vault stored in OneDrive.",
)

# ---------------- auth + health (hidden from the tool schema) ----------------


@app.get("/health", include_in_schema=False)
def health():
    return {"ok": True, "signed_in": bool(_app.get_accounts())}


@app.post("/auth/start", include_in_schema=False)
def auth_start(_=Depends(require_key)):
    global _flow
    with _flow_lock:
        _flow = _app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in _flow:
            raise HTTPException(500, detail=f"device flow failed: {_flow.get('error_description')}")
        flow = _flow

    def _complete(fl):
        try:
            _app.acquire_token_by_device_flow(fl)  # blocks until user finishes / times out
            _persist_cache()
        except Exception:
            pass

    threading.Thread(target=_complete, args=(flow,), daemon=True).start()
    return {
        "verification_uri": flow["verification_uri"],
        "user_code": flow["user_code"],
        "expires_in": flow["expires_in"],
        "message": flow.get("message"),
    }


@app.get("/auth/status", include_in_schema=False)
def auth_status(_=Depends(require_key)):
    return {"signed_in": bool(_app.get_accounts())}


# ---------------- the four model-facing tools ----------------


@app.get(
    "/list_files",
    summary="List files and folders in the vault",
    description="List the files and folders at a vault path. Leave path empty for the vault root. Returns each item's name, whether it is a file or folder, and its size in bytes.",
)
def list_files(path: str = "", _=Depends(require_key)):
    token = _token_or_401()
    clean = path.strip("/")
    if clean:
        url = f"{GRAPH}/me/drive/root:/{clean}:/children"
    else:
        url = f"{GRAPH}/me/drive/root/children"
    r = requests.get(url, headers=_auth_header(token), params={"$select": "name,size,folder,file", "$top": 200})
    r.raise_for_status()
    items = [
        {"name": it["name"], "type": "folder" if "folder" in it else "file", "size": it.get("size", 0)}
        for it in r.json().get("value", [])
    ]
    return {"path": path, "items": items}


@app.get(
    "/read_file",
    summary="Read a text file from the vault",
    description="Return the UTF-8 text contents of a file at the given vault path, for example 'Veronica/Read Me First.md'.",
)
def read_file(path: str, _=Depends(require_key)):
    token = _token_or_401()
    url = f"{GRAPH}/me/drive/root:/{path.strip('/')}:/content"
    r = requests.get(url, headers=_auth_header(token))
    r.raise_for_status()
    try:
        return {"path": path, "content": r.content.decode("utf-8")}
    except UnicodeDecodeError:
        raise HTTPException(415, detail="File is not UTF-8 text.")


class WriteBody(BaseModel):
    path: str
    content: str


@app.post(
    "/write_file",
    summary="Write or overwrite a text file in the vault",
    description="Create or overwrite a text file at the given vault path with the provided content. Parent folders are created as needed.",
)
def write_file(body: WriteBody, _=Depends(require_key)):
    token = _token_or_401()
    url = f"{GRAPH}/me/drive/root:/{body.path.strip('/')}:/content"
    headers = _auth_header(token)
    headers["Content-Type"] = "text/plain"
    r = requests.put(url, headers=headers, data=body.content.encode("utf-8"))
    r.raise_for_status()
    return {"path": body.path, "written": True}


@app.get(
    "/search_files",
    summary="Search the vault by keyword",
    description="Search the whole vault for files whose name or contents match the query. Returns matching file names and their folder paths.",
)
def search_files(query: str, _=Depends(require_key)):
    token = _token_or_401()
    safe = query.replace("'", "''")
    url = f"{GRAPH}/me/drive/root/search(q='{safe}')"
    r = requests.get(url, headers=_auth_header(token), params={"$select": "name,parentReference", "$top": 50})
    r.raise_for_status()
    hits = [
        {"name": it["name"], "path": it.get("parentReference", {}).get("path", "")}
        for it in r.json().get("value", [])
    ]
    return {"query": query, "hits": hits}
