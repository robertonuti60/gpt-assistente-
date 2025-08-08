import os
import io
import json
import requests
from flask import Flask, request, jsonify
import msal

# ====== Config da variabili d'ambiente ======
CLIENT_ID     = os.environ.get("CLIENT_ID")        # Azure AD app client id
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")    # Azure AD client secret
TENANT_ID     = os.environ.get("TENANT_ID")        # Azure AD tenant id
ONEDRIVE_USER = os.environ.get("ONEDRIVE_USER")    # utente/UPN (es. name@tenant.onmicrosoft.com)
API_KEY       = os.environ.get("API_KEY")          # chiave condivisa per /search e /read
MAX_CHARS     = int(os.environ.get("MAX_CHARS", "500000"))

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE  = "https://graph.microsoft.com/v1.0"

app = Flask(__name__)

# ====== Helpers ======
def require_api_key(f):
    """Semplice guard che verifica l'header x-api-key se impostato in env."""
    def wrapper(*args, **kwargs):
        if API_KEY:
            if request.headers.get("x-api-key") != API_KEY:
                return jsonify({"error": "Forbidden"}), 403
        return f(*args, **kwargs)
    wrapper.__name__ = f.__name__
    return wrapper

def get_token():
    """Ottiene un Access Token per Microsoft Graph via MSAL (client_credential)."""
    if not (CLIENT_ID and CLIENT_SECRET and TENANT_ID):
        raise RuntimeError("CLIENT_ID / CLIENT_SECRET / TENANT_ID non configurati.")
    appc = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = appc.acquire_token_silent(GRAPH_SCOPE, account=None)
    if not result:
        result = appc.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if "access_token" not in result:
        raise RuntimeError(f"Errore auth MSAL: {result}")
    return result["access_token"]

def gget(url, params=None, stream=False):
    """GET su Graph con Bearer token."""
    token = get_token()
    r = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        params=params,
        stream=stream,
        timeout=60
    )
    r.raise_for_status()
    return r

def path_url(path: str) -> str:
    """Converte un path OneDrive in URL Graph (drive/root:/path)."""
    p = path.lstrip("/")
    if not ONEDRIVE_USER:
        raise RuntimeError("ONEDRIVE_USER non configurato.")
    return f"{GRAPH_BASE}/users/{ONEDRIVE_USER}/drive/root:/{p}"

# ====== Estrattori testo ======
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

# ====== ROUTES ======

# Home con documentazione rapida
@app.route("/", methods=["GET"], strict_slashes=False)
def index():
    return jsonify({
        "name": "gpt-assistente-api",
        "ok": True,
        "note": "Usa x-api-key: <API_KEY> per /search e /read.",
        "endpoints": {
            "health": "/health",
            "ping": "/ping",
            "search": "/search?q=term&path=<facoltativo>",
            "read": "POST /read  { path: '/cartella/file.pdf' }"
        }
    })

# Health (robusto: accetta /health e /health/)
@app.route("/health", methods=["GET"], strict_slashes=False)
def health():
    return jsonify({"ok": True}), 200

# Ping test
@app.route("/ping", methods=["GET"], strict_slashes=False)
def ping():
    return "pong", 200

# Cerca file su OneDrive
@app.route("/search", methods=["GET"], strict_slashes=False)
@require_api_key
def search():
    q    = request.args.get("q", "").strip()
    path = request.args.get("path", "").strip()
    if not q:
        return jsonify({"error": "Parametro q mancante"}), 400

    try:
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
                "path": full,   # questo path lo userai sul /read
                "size": it.get("size"),
                "webUrl": it.get("webUrl"),
                "id": it.get("id"),
            })
        return jsonify({"results": out}), 200
    except Exception as e:
        return jsonify({"error": f"Search failed: {e}"}), 500

# Legge un file (PDF/DOCX/TXT)
@app.route("/read", methods=["POST"], strict_slashes=False)
@require_api_key
def read():
    try:
        body = request.get_json(force=True, silent=True) or {}
        path = body.get("path", "").strip()
        if not path:
            return jsonify({"error": "Missing 'path'"}), 400

        # scarica il file
        content = gget(path_url(path) + ":/content", stream=True).content
        lower   = path.lower()

        if lower.endswith(".pdf"):
            text = extract_pdf(content)
        elif lower.endswith(".docx"):
            text = extract_docx(content)
        elif lower.endswith(".txt"):
            text = content.decode("utf-8", errors="ignore")
        else:
            return jsonify({"error": "Unsupported file type"}), 415

        if len(text) > MAX_CHARS:
            text = text[:MAX_CHARS] + "\n\n[...troncato...]"

        return jsonify({"path": path, "text": text}), 200
    except Exception as e:
        return jsonify({"error": f"Read failed: {e}"}), 500


# ====== avvio locale ======
if __name__ == "__main__":
    # Render usa gunicorn, ma questo serve per test locale:
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
