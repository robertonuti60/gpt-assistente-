import os, io, json, requests
from flask import Flask, request, jsonify
import msal

# === Config via variabili d'ambiente ===
CLIENT_ID     = os.environ.get("CLIENT_ID")       # Entra → App registrations
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")   # Entra → client secret
TENANT_ID     = os.environ.get("TENANT_ID")       # Entra → Directory (tenant) ID
ONEDRIVE_USER = os.environ.get("ONEDRIVE_USER")   # UPN o id utente (es. r_nuti@...onmicrosoft.com)
API_KEY       = os.environ.get("API_KEY")         # chiave condivisa per l’Azione GPT
MAX_CHARS     = int(os.environ.get("MAX_CHARS", "500000"))

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE  = "https://graph.microsoft.com/v1.0"

app = Flask(__name__)

# --- simple API-key guard ---
def require_api_key(f):
    def wrapper(*args, **kwargs):
        if API_KEY and request.headers.get("x-api-key") != API_KEY:
            return jsonify({"error": "Forbidden"}), 403
        return f(*args, **kwargs)
    wrapper.__name__ = f.__name__
    return wrapper

def get_token():
    appc = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = appc.acquire_token_silent(GRAPH_SCOPE, account=None) or appc.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if "access_token" not in result:
        raise RuntimeError(f"Auth error: {result}")
    return result["access_token"]

def gget(url, params=None, stream=False):
    token = get_token()
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, params=params, stream=stream, timeout=60)
    r.raise_for_status()
    return r

def path_url(path: str) -> str:
    p = path.lstrip("/")
    return f"{GRAPH_BASE}/users/{ONEDRIVE_USER}/drive/root:/{p}"

@app.get("/health")
def health():
    return {"ok": True}

@app.get("/search")
@require_api_key
def search():
    q    = request.args.get("q", "")
    path = request.args.get("path", "").strip()
    if path:
        url = path_url(path) + f":/search(q='{q}')"
    else:
        url = f"{GRAPH_BASE}/users/{ONEDRIVE_USER}/drive/root/search(q='{q}')"
    data = gget(url).json()
    out = []
    for it in data.get("value", []):
        parent = it.get("parentReference", {}).get("path", "")
        full   = (parent + "/" + it.get("name")).replace("/drive/root:", "")
        out.append({
            "name": it.get("name"),
            "path": full,                  # <— questo ‘path’ servirà al /read
            "size": it.get("size"),
            "webUrl": it.get("webUrl"),
            "id": it.get("id"),
        })
    return {"results": out}

@app.post("/read")
@require_api_key
def read():
    body = request.get_json(force=True)
    path = body.get("path")
    if not path:
        return {"error": "Missing 'path'"}, 400

    # scarica il file
    content = gget(path_url(path) + ":/content", stream=True).content
    lower   = path.lower()

    try:
        if lower.endswith(".pdf"):
            text = extract_pdf(content)
        elif lower.endswith(".docx"):
            text = extract_docx(content)
        elif lower.endswith(".txt"):
            text = content.decode("utf-8", errors="ignore")
        else:
            return {"error": "Unsupported file type"}, 415
    except Exception as e:
        return {"error": f"Extraction failed: {e}"}, 500

    if len(text) > MAX_CHARS:
        text = text[:MAX_CHARS] + "\n\n[...troncato...]"
    return {"path": path, "text": text}

# --- estrattori ---
def extract_pdf(binary: bytes) -> str:
    from pdfminer.high_level import extract_text
    return extract_text(io.BytesIO(binary))

def extract_docx(binary: bytes) -> str:
    from docx import Document
    doc = Document(io.BytesIO(binary))
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            parts.append("\t".join(c.text for c in row.cells))
    return "\n".join(parts)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
