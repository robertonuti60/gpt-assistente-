import os
import io
import json
import requests
from flask import Flask, request, jsonify
import msal

# ==========
# Config
# ==========
CLIENT_ID     = os.environ.get("CLIENT_ID")          # Azure AD → App registrations → Application (client) ID
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")      # Azure AD → App registrations → Client secret (value)
TENANT_ID     = os.environ.get("TENANT_ID")          # Azure AD → Tenant ID
ONEDRIVE_USER = os.environ.get("ONEDRIVE_USER")      # UPN o id utente es. r_nuti@xxx.onmicrosoft.com
API_KEY       = os.environ.get("API_KEY")            # chiave condivisa per proteggere gli endpoint
MAX_CHARS     = int(os.environ.get("MAX_CHARS", "500000"))

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE  = "https://graph.microsoft.com/v1.0"

app = Flask(__name__)


# ==========
# Helpers
# ==========
def require_api_key(f):
    """Decoratore: richiede header x-api-key che matcha API_KEY."""
    def wrapper(*args, **kwargs):
        if API_KEY and request.headers.get("x-api-key") != API_KEY:
            return jsonify({"error": "Forbidden"}), 403
        return f(*args, **kwargs)
    wrapper.__name__ = f.__name__
    return wrapper


def get_token():
    """Ottiene un access token per Microsoft Graph via MSAL (client credentials)."""
    appc = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = appc.acquire_token_silent(GRAPH_SCOPE, account=None) or \
             appc.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if "access_token" not in result:
        raise RuntimeError(f"Auth error: {result}")
    return result["access_token"]


def gget(url, params=None, stream=False):
    """GET helper verso Graph con bearer token, timeout e raise_for_status()."""
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
    """Costruisce l’URL Graph per un path OneDrive/SharePoint dell’utente indicato."""
    p = (path or "").lstrip("/")
    return f"{GRAPH_BASE}/users/{ONEDRIVE_USER}/drive/root:/{p}"


# ==========
# Estrattori di testo
# ==========
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


# ==========
# Routes
# ==========
@app.get("/")
def root():
    return jsonify({
        "name": "gpt-assistente-api",
        "ok": True,
        "endpoints": {
            "health": "/health",
            "search": "/search?q=term&path=facoltativo",
            "read": "POST /read { path: '/cartella/file.pdf' }"
        },
        "note": "Usa header 'x-api-key: <API_KEY>' per /search e /read."
    })


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/search")
@require_api_key
def search():
    """
    Cerca file in OneDrive.
    query:
      - q (string)   : termine di ricerca
      - path (string): opzionale, path relativo (es. 'Normativa/NTC') per cercare dentro quella cartella
    """
    q    = (request.args.get("q") or "").strip()
    path = (request.args.get("path") or "").strip()

    if not q:
        return {"error": "Missing 'q' (query term)"}, 400

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
                "path": full,                  # questo path lo rimandi a /read
                "size": it.get("size"),
                "webUrl": it.get("webUrl"),
                "id": it.get("id"),
            })
        return {"results": out}
    except requests.HTTPError as http_err:
        return {"error": "Graph error", "details": str(http_err), "body": http_err.response.text}, 502
    except Exception as e:
        return {"error": "Internal error", "details": str(e)}, 500


@app.post("/read")
@require_api_key
def read():
    """
    Legge un file e ne estrae il testo.
    body JSON: { "path": "/cartella/nomefile.pdf" }
    """
    body = request.get_json(silent=True) or {}
    path = (body.get("path") or "").strip()
    if not path:
        return {"error": "Missing 'path'"}, 400

    try:
        content = gget(path_url(path) + ":/content", stream=True).content
        lower   = path.lower()

        if lower.endswith(".pdf"):
            text = extract_pdf(content)
        elif lower.endswith(".docx"):
            text = extract_docx(content)
        elif lower.endswith(".txt"):
            text = content.decode("utf-8", errors="ignore")
        else:
            return {"error": "Unsupported file type. Use .pdf, .docx, .txt"}, 415

        if len(text) > MAX_CHARS:
            text = text[:MAX_CHARS] + "\n\n[...troncato per MAX_CHARS...]"

        return {"path": path, "text": text}
    except requests.HTTPError as http_err:
        return {"error": "Graph error", "details": str(http_err), "body": http_err.response.text}, 502
    except Exception as e:
        return {"error": "Extraction failed", "details": str(e)}, 500


# ==========
# Entrypoint locale
# ==========
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
