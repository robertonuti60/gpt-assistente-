import os
import io
import json
import requests
from flask import Flask, request, jsonify
import msal

# ============== Config tramite variabili d'ambiente ==============
CLIENT_ID     = os.environ.get("CLIENT_ID")          # Entra → App registrations → Application (client) ID
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")      # Entra → App registrations → Certificates & secrets
TENANT_ID     = os.environ.get("TENANT_ID")          # Entra → Overview → Directory (tenant) ID
ONEDRIVE_USER = os.environ.get("ONEDRIVE_USER")      # UPN dell'utente es. user@tenant.onmicrosoft.com
API_KEY       = os.environ.get("API_KEY")            # chiave condivisa per /search e /read
MAX_CHARS     = int(os.environ.get("MAX_CHARS", "500000"))

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE  = "https://graph.microsoft.com/v1.0"

app = Flask(__name__)

# ====================== Sicurezza con API Key =====================
def require_api_key(f):
    def _wrap(*args, **kwargs):
        if API_KEY:
            given = request.headers.get("x-api-key")
            if not given or given != API_KEY:
                return jsonify({"error": "Forbidden"}), 403
        return f(*args, **kwargs)
    _wrap.__name__ = f.__name__
    return _wrap

# ====================== Auth MS Graph (MSAL) ======================
def get_token():
    appc = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = appc.acquire_token_silent(GRAPH_SCOPE, account=None)
    if not result:
        result = appc.acquire_token_for_client(scopes=GRAPH_SCOPE)

    if "access_token" not in result:
        raise RuntimeError(f"Auth error: {result}")
    return result["access_token"]

def gget(url, params=None, stream=False):
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
    """Costruisce l'URL Graph per un path nel drive dell'utente."""
    p = path.lstrip("/")
    return f"{GRAPH_BASE}/users/{ONEDRIVE_USER}/drive/root:/{p}"

# ============================= Endpoints =============================
@app.get("/")
def info():
    return {
        "name": "gpt-assistente-api",
        "ok": True,
        "note": "Usa l'header 'x-api-key: <API_KEY>' per /search e /read.",
        "endpoints": {
            "health": "/health",
            "search": "/search?q=term&path=<facoltativo>",
            "read": "POST /read  body: { 'path': '/cartella/file.pdf' }"
        }
    }

@app.get("/health")
def health():
    return {"ok": True}

@app.get("/search")
@require_api_key
def search():
    """
    Cerca file nel OneDrive.
    Query:
      - q    (obbligatorio): stringa di ricerca
      - path (facoltativo) : cartella di partenza (es. 'Documenti/Norme')
    """
    q = (request.args.get("q") or "").strip()
    path = (request.args.get("path") or "").strip()

    if not q:
        return {"error": "Parametro 'q' mancante"}, 400

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
                "path": full,
                "size": it.get("size"),
                "webUrl": it.get("webUrl"),
                "id": it.get("id"),
                "lastModified": it.get("lastModifiedDateTime"),
            })
        return {"count": len(out), "results": out}

    except requests.HTTPError as e:
        return {"error": "Graph error", "detail": e.response.text}, 502
    except Exception as e:
        return {"error": str(e)}, 500

@app.post("/read")
@require_api_key
def read():
    """
    Scarica un file e ne estrae il testo.
    Body JSON: { "path": "/cartella/file.pdf" }
    Supporta: .pdf, .docx, .txt
    """
    try:
        body = request.get_json(force=True)
    except Exception:
        body = None

    if not body or not isinstance(body, dict):
        return {"error": "JSON non valido"}, 400

    path = (body.get("path") or "").strip()
    if not path:
        return {"error": "Campo 'path' mancante"}, 400

    try:
        # Scarica contenuto
        url_content = path_url(path) + ":/content"
        content = gget(url_content, stream=True).content

        lower = path.lower()
        if lower.endswith(".pdf"):
            text = extract_pdf(content)
        elif lower.endswith(".docx"):
            text = extract_docx(content)
        elif lower.endswith(".txt"):
            text = content.decode("utf-8", errors="ignore")
        else:
            return {"error": "Tipo file non supportato. Usa pdf/docx/txt"}, 415

        if len(text) > MAX_CHARS:
            text = text[:MAX_CHARS] + "\n\n[...troncato...]"

        return {"path": path, "characters": len(text), "text": text}

    except requests.HTTPError as e:
        return {"error": "Graph error", "detail": e.response.text}, 502
    except Exception as e:
        return {"error": str(e)}, 500

# ============================ Estrattori =============================
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

# ============================= Run locale ============================
if __name__ == "__main__":
    # In produzione Render usa gunicorn/WSGI. Questo è per il run locale.
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
