import os, io, time, base64, json
from flask import Flask, request, jsonify
import requests
from pdfminer.high_level import extract_text as pdf_extract_text
from docx import Document as DocxDocument

app = Flask(__name__)

# =========================
# Config / ENV
# =========================
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
API_KEY = os.getenv("API_KEY")
ONEDRIVE_USER = os.getenv("ONEDRIVE_USER")  # upn o id utente (es. r_nuti@...onmicrosoft.com)

GRAPH_SCOPE = "https://graph.microsoft.com/.default"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# =========================
# Token cache
# =========================
_token_cache = {"access_token": None, "exp": 0}

def get_graph_token():
    now = time.time()
    if _token_cache["access_token"] and now < _token_cache["exp"] - 60:
        return _token_cache["access_token"]
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": GRAPH_SCOPE,
        "grant_type": "client_credentials",
    }
    r = requests.post(url, data=data, timeout=30)
    r.raise_for_status()
    tok = r.json()
    _token_cache["access_token"] = tok["access_token"]
    _token_cache["exp"] = now + tok.get("expires_in", 3600)
    return _token_cache["access_token"]

def gheaders():
    return {"Authorization": f"Bearer {get_graph_token()}"}

# =========================
# Helpers Graph (retry)
# =========================
def gget(url, params=None, stream=False):
    for i in range(5):
        r = requests.get(url, headers=gheaders(), params=params, timeout=60, stream=stream)
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(1.5 * (i + 1))
            continue
        r.raise_for_status()
        return r
    r.raise_for_status()

def gpost(url, json_payload=None):
    for i in range(5):
        r = requests.post(url, headers={**gheaders(), "Content-Type": "application/json"}, json=json_payload, timeout=60)
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(1.5 * (i + 1))
            continue
        r.raise_for_status()
        return r
    r.raise_for_status()

# =========================
# Security
# =========================
def require_api_key(req):
    return req.headers.get("x-api-key") == API_KEY

# =========================
# Drive cache (user drive id)
# =========================
_drive_cache = {"id": None}

def get_user_drive_id():
    if _drive_cache["id"]:
        return _drive_cache["id"]
    if not ONEDRIVE_USER:
        raise RuntimeError("ONEDRIVE_USER non configurato")
    url = f"{GRAPH_BASE}/users/{ONEDRIVE_USER}/drive"
    data = gget(url).json()
    _drive_cache["id"] = data.get("id")
    return _drive_cache["id"]

# =========================
# Text extraction
# =========================
def guess_ext(name):
    if not name:
        return ""
    name = name.lower()
    for ext in (".pdf", ".docx", ".txt", ".md", ".csv"):
        if name.endswith(ext):
            return ext
    return ""

def extract_text_from_bytes(content: bytes, filename: str):
    ext = guess_ext(filename)
    if ext == ".pdf":
        return pdf_extract_text(io.BytesIO(content))
    if ext == ".docx":
        f = io.BytesIO(content)
        doc = DocxDocument(f)
        return "\n".join([p.text for p in doc.paragraphs])
    if ext in (".txt", ".md", ".csv"):
        try:
            return content.decode("utf-8", errors="ignore")
        except:
            return content.decode("latin-1", errors="ignore")
    return ""  # binari o formati non gestiti

# =========================
# Health
# =========================
@app.get("/health")
def health():
    return jsonify({"ok": True})

# =========================
# DEBUG endpoints
# =========================
@app.get("/debug/drive")
def debug_drive():
    if not require_api_key(request):
        return jsonify({"error": "unauthorized"}), 401
    try:
        did = get_user_drive_id()
        return jsonify({"driveId": did})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.get("/debug/list")
def debug_list():
    """
    Lista i children di un path su OneDrive utente
    GET /debug/list?path=Documents/Struttura_Normativa_GPT
    """
    if not require_api_key(request):
        return jsonify({"error": "unauthorized"}), 401
    path = request.args.get("path", "").strip().strip("/")
    if not path:
        path = "Documents"
    try:
        did = get_user_drive_id()
        url = f"{GRAPH_BASE}/drives/{did}/root:/{path}:/children"
        items = []
        next_url = url
        while next_url:
            resp = gget(next_url).json()
            items.extend(resp.get("value", []))
            next_url = resp.get("@odata.nextLink")
        return jsonify([
            {
                "id": it.get("id"),
                "name": it.get("name"),
                "type": "folder" if "folder" in it else "file",
                "webUrl": it.get("webUrl"),
            } for it in items
        ])
    except requests.HTTPError as e:
        return jsonify({"error": "graph_error", "status": e.response.status_code, "detail": e.response.text}), 502
    except Exception as e:
        return jsonify({"error": "internal_error", "detail": str(e)}), 500

# =========================
# SEARCH (per nome, con path opzionale)
# =========================
@app.get("/search")
def search():
    """
    GET /search?q=DPR%20380&path=Documents/Struttura_Normativa_GPT
    - se path presente: lista la cartella e filtra per nome contiene q
    - altrimenti: usa Graph search sul drive utente
    """
    if not require_api_key(request):
        return jsonify({"error": "unauthorized"}), 401

    q = (request.args.get("q") or "").strip()
    path = (request.args.get("path") or "").strip().strip("/")

    if not q:
        return jsonify({"error": "missing q"}), 400

    try:
        did = get_user_drive_id()

        # preferisci path per evitare ambiguità di "Documents/Documenti"
        if path:
            url = f"{GRAPH_BASE}/drives/{did}/root:/{path}:/children"
            items, next_url = [], url
            while next_url:
                resp = gget(next_url).json()
                items.extend(resp.get("value", []))
                next_url = resp.get("@odata.nextLink")
            qlow = q.lower()
            results = [
                {
                    "id": it.get("id"),
                    "driveId": it.get("parentReference", {}).get("driveId"),
                    "name": it.get("name"),
                    "type": "folder" if "folder" in it else "file",
                    "webUrl": it.get("webUrl"),
                }
                for it in items
                if it.get("name","").lower().find(qlow) != -1
            ]
            return jsonify({"scope": "path", "path": path, "results": results})

        # fallback: Graph search
        # NB: search lavora bene ma può includere percorsi fuori dalla cartella attesa
        url = f"{GRAPH_BASE}/drives/{did}/root/search(q='{q}')"
        resp = gget(url).json()
        results = [
            {
                "id": it.get("id"),
                "driveId": it.get("parentReference", {}).get("driveId"),
                "name": it.get("name"),
                "type": "folder" if "folder" in it else "file",
                "webUrl": it.get("webUrl"),
            } for it in resp.get("value", [])
        ]
        return jsonify({"scope": "drive_search", "results": results})

    except requests.HTTPError as e:
        return jsonify({"error": "graph_error", "status": e.response.status_code, "detail": e.response.text}), 502
    except Exception as e:
        return jsonify({"error": "internal_error", "detail": str(e)}), 500

# =========================
# READ by path (legacy)
# =========================
@app.post("/read")
def read_by_path():
    """
    POST /read { "path": "Documents/Struttura_Normativa_GPT/DPR_380_2001.docx" }
    """
    if not require_api_key(request):
        return jsonify({"error": "unauthorized"}), 401
    try:
        payload = request.get_json(force=True)
        path = (payload.get("path") or "").strip().strip("/")
        if not path:
            return jsonify({"error": "missing path"}), 400

        did = get_user_drive_id()
        meta_url = f"{GRAPH_BASE}/drives/{did}/root:/{path}"
        meta = gget(meta_url).json()
        name = meta.get("name", "")
        content_url = f"{meta_url}:/content"

        r = gget(content_url, stream=True)
        content = r.content
        text = extract_text_from_bytes(content, name)
        return jsonify({"path": path, "name": name, "text": text, "bytes": len(content)})
    except requests.HTTPError as e:
        return jsonify({"error": "graph_error", "status": e.response.status_code, "detail": e.response.text}), 502
    except Exception as e:
        return jsonify({"error": "internal_error", "detail": str(e)}), 500

# =========================
# RESOLVE link (nuovo, robusto)
# =========================
@app.get("/resolve")
def resolve_link():
    """
    GET /resolve?url=<link OneDrive/SharePoint>
    Ritorna id/driveId e children se cartella.
    """
    if not require_api_key(request):
        return jsonify({"error": "unauthorized"}), 401

    url = request.args.get("url", "").strip()
    if not url:
        return jsonify({"error": "missing url param"}), 400

    try:
        b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
        item_url = f"{GRAPH_BASE}/shares/{b64}/driveItem"
        item = gget(item_url).json()

        drive_id = item.get("parentReference", {}).get("driveId")
        item_id  = item.get("id")
        name     = item.get("name")
        is_folder = "folder" in item

        result = {
            "name": name,
            "id": item_id,
            "driveId": drive_id,
            "webUrl": item.get("webUrl"),
            "type": "folder" if is_folder else "file",
        }

        if is_folder:
            children = []
            next_url = f"{GRAPH_BASE}/shares/{b64}/driveItem/children"
            while next_url:
                resp = gget(next_url).json()
                children.extend(resp.get("value", []))
                next_url = resp.get("@odata.nextLink")
            result["children"] = [
                {
                    "id": c["id"],
                    "name": c.get("name"),
                    "driveId": c.get("parentReference", {}).get("driveId"),
                    "type": "folder" if "folder" in c else "file",
                    "webUrl": c.get("webUrl"),
                }
                for c in children
            ]

        return jsonify(result)
    except requests.HTTPError as e:
        return jsonify({"error": "graph_error", "status": e.response.status_code, "detail": e.response.text}), 502
    except Exception as e:
        return jsonify({"error": "internal_error", "detail": str(e)}), 500

# =========================
# READ by id (nuovo, robusto)
# =========================
@app.post("/read_by_id")
def read_by_id():
    """
    POST /read_by_id { "id": "...", "driveId": "b!...optional...", "name": "file.ext" }
    """
    if not require_api_key(request):
        return jsonify({"error": "unauthorized"}), 401
    try:
        payload = request.get_json(force=True)
        item_id = payload.get("id")
        drive_id = payload.get("driveId")
        name = payload.get("name", "")

        if not item_id:
            return jsonify({"error": "missing id"}), 400

        # prova con driveId se presente
        if drive_id:
            meta_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
        else:
            meta_url = f"{GRAPH_BASE}/drive/items/{item_id}"

        meta = gget(meta_url).json()
        if not name:
            name = meta.get("name", "")
        content_url = f"{meta_url}/content"

        r = gget(content_url, stream=True)
        content = r.content
        text = extract_text_from_bytes(content, name)
        return jsonify({"id": item_id, "driveId": drive_id, "name": name, "text": text, "bytes": len(content)})
    except requests.HTTPError as e:
        return jsonify({"error": "graph_error", "status": e.response.status_code, "detail": e.response.text}), 502
    except Exception as e:
        return jsonify({"error": "internal_error", "detail": str(e)}), 500

# =========================
# Entrypoint
# =========================
if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port)
