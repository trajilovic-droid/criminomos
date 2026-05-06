# Iusprudentia Poenalis - Serveur MCP v8
# Streamable HTTP avec gestion des sessions (Mcp-Session-Id)
import json
import re
import unicodedata
import os
import uuid
import logging
from pathlib import Path

import httpx
from bs4 import BeautifulSoup
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import JSONResponse, Response
from starlette.routing import Route
from starlette.middleware import Middleware
from starlette.middleware.cors import CORSMiddleware
import uvicorn

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Chargement des données
# ---------------------------------------------------------------------------
DATA_PATH = Path(__file__).parent / "arrets.json"
with open(DATA_PATH, encoding="utf-8") as f:
    ARRETS = json.load(f)
ARRETS_BY_ID = {r["arret"]: r for r in ARRETS if r.get("arret")}

logger.info(f"=== Iusprudentia Poenalis MCP v8 : {len(ARRETS)} arrets charges ===")

# Stockage des sessions actives
SESSIONS = {}

# ---------------------------------------------------------------------------
# Définition des outils
# ---------------------------------------------------------------------------
TOOLS = [
    {
        "name": "search_arrets",
        "description": "Recherche des arrets du Tribunal federal suisse en droit penal par mots-cles, infraction ou article de loi.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "Mots-cles de recherche"},
                "infraction": {"type": "string", "description": "Type d infraction ex: Expulsion"},
                "article": {"type": "string", "description": "Article de loi ex: 66a CP"},
                "annee": {"type": "string", "description": "Annee ex: 2024"},
                "limite": {"type": "integer", "description": "Nombre de resultats max 30"}
            },
            "required": ["query"]
        }
    },
    {
        "name": "get_fulltext",
        "description": "Charge le texte integral d un arret depuis bger.ch.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "arret_id": {"type": "string", "description": "Numero d arret ex: 6B_409/2024"}
            },
            "required": ["arret_id"]
        }
    },
    {
        "name": "get_references",
        "description": "Extrait les ATF et arrets TF cites dans un arret.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "arret_id": {"type": "string", "description": "Numero d arret ex: 6B_409/2024"}
            },
            "required": ["arret_id"]
        }
    },
    {
        "name": "get_arret_by_reference",
        "description": "Charge le texte d un ATF ou arret TF cite en reference.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "reference": {"type": "string", "description": "ex: ATF 148 IV 409 ou 6B_123/2021"}
            },
            "required": ["reference"]
        }
    }
]

# ---------------------------------------------------------------------------
# Logique métier
# ---------------------------------------------------------------------------
def normalize(text):
    text = text.lower()
    text = unicodedata.normalize("NFD", text)
    return "".join(c for c in text if unicodedata.category(c) != "Mn")


def score_arret(arret, query_words):
    score = 0
    for word in query_words:
        w = normalize(word)
        if w in normalize(arret.get("objet", "")): score += 4
        if w in normalize(arret.get("resume", "")): score += 2
        if w in normalize(arret.get("articles", "")): score += 3
        if w in normalize(arret.get("arret", "")): score += 5
    if arret.get("interet") == "oui":
        score += 1
    return score


def search_arrets(query="", infraction="", article="", annee="", limite=10):
    limite = min(int(limite), 30)
    query_words = query.strip().split() if query.strip() else []
    results = []
    for arret in ARRETS:
        if infraction and not normalize(arret.get("objet", "")).startswith(normalize(infraction)):
            continue
        if article and normalize(article) not in normalize(arret.get("articles", "")):
            continue
        if annee and not (arret.get("decision", "").startswith(annee) or arret.get("parution", "").startswith(annee)):
            continue
        score = score_arret(arret, query_words) if query_words else 1
        if score > 0 or not query_words:
            results.append((score, arret))
    results.sort(key=lambda x: x[0], reverse=True)
    results = results[:limite]
    if not results:
        return "Aucun arret trouve."
    lines = [str(len(results)) + " arret(s) trouve(s)\n"]
    for _, r in results:
        lines.append("### " + r["arret"])
        lines.append("- Objet : " + r.get("objet", "-"))
        lines.append("- Date : " + r.get("decision", "-"))
        lines.append("- Articles : " + r.get("articles", "-"))
        if r.get("admis"): lines.append("- Resultat : " + r["admis"])
        if r.get("interet") == "oui": lines.append("- Arret d interet")
        if r.get("resume"): lines.append("- Resume : " + r["resume"])
        if r.get("url"): lines.append("- URL : " + r["url"])
        lines.append("")
    return "\n".join(lines)


def get_fulltext(arret_id):
    arret = ARRETS_BY_ID.get(arret_id)
    url = arret.get("url") if arret else None
    if not url:
        clean_id = re.sub(r"[^A-Za-z0-9_/.-]", "", arret_id)[:40]
        url = "https://www.bger.ch/ext/eurospider/live/fr/php/aza/http/index.php?lang=fr&type=show_document&highlight_docid=aza://" + clean_id
    try:
        resp = httpx.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"}, follow_redirects=True)
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "nav", "header", "footer"]):
            tag.decompose()
        lines = [l.rstrip() for l in soup.get_text(separator="\n").splitlines() if l.strip()]
        return "Arret " + arret_id + "\nURL : " + url + "\n" + "-"*60 + "\n\n" + "\n".join(lines)[:15000]
    except Exception as e:
        return "Erreur : " + str(e)


def get_references(arret_id):
    arret = ARRETS_BY_ID.get(arret_id)
    url = arret.get("url") if arret else None
    if not url:
        clean_id = re.sub(r"[^A-Za-z0-9_/.-]", "", arret_id)[:40]
        url = "https://www.bger.ch/ext/eurospider/live/fr/php/aza/http/index.php?lang=fr&type=show_document&highlight_docid=aza://" + clean_id
    try:
        resp = httpx.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"}, follow_redirects=True)
        text = BeautifulSoup(resp.text, "html.parser").get_text(separator=" ")
    except Exception as e:
        return "Erreur : " + str(e)
    atf_refs = sorted(set(re.findall(r"ATF\s+\d{2,3}\s+[IVX]+\s+\d+", text)))
    tf_refs = sorted(set(r for r in re.findall(r"\b[0-9][A-Z]{1,2}_\d{1,4}/20\d{2}\b", text) if r != arret_id))
    lines = ["References de l arret " + arret_id + "\n"]
    if atf_refs:
        lines.append("ATF cites (" + str(len(atf_refs)) + ") :")
        for r in atf_refs: lines.append("- " + r)
    if tf_refs:
        lines.append("\nArrets TF cites (" + str(len(tf_refs)) + ") :")
        for r in tf_refs:
            if r in ARRETS_BY_ID:
                a = ARRETS_BY_ID[r]
                lines.append("- " + r + " - " + a.get("objet", "-") + " (" + a.get("decision", "-") + ") [base]")
            else:
                lines.append("- " + r)
    if not atf_refs and not tf_refs:
        lines.append("Aucune reference trouvee.")
    return "\n".join(lines)


def get_arret_by_reference(reference):
    reference = reference.strip()
    if re.match(r"^[0-9][A-Z]{1,2}_\d{1,4}/20\d{2}$", reference):
        return get_fulltext(reference)
    atf_match = re.match(r"ATF\s+(\d{2,3})\s+([IVX]+)\s+(\d+)", reference, re.IGNORECASE)
    if atf_match:
        vol, part, page = atf_match.groups()
        url = "https://www.bger.ch/ext/eurospider/live/fr/php/clir/http/index.php?lang=fr&type=show_document&highlight_docid=atf:///" + vol + "/" + part + "/" + page
        try:
            resp = httpx.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"}, follow_redirects=True)
            soup = BeautifulSoup(resp.text, "html.parser")
            for tag in soup(["script", "style", "nav", "header", "footer"]):
                tag.decompose()
            lines = [l.rstrip() for l in soup.get_text(separator="\n").splitlines() if l.strip()]
            return reference + "\nURL : " + url + "\n" + "-"*60 + "\n\n" + "\n".join(lines)[:15000]
        except Exception as e:
            return "Erreur : " + str(e)
    return "Format non reconnu : " + reference


def call_tool(name, args):
    if name == "search_arrets": return search_arrets(**args)
    elif name == "get_fulltext": return get_fulltext(**args)
    elif name == "get_references": return get_references(**args)
    elif name == "get_arret_by_reference": return get_arret_by_reference(**args)
    return "Outil inconnu : " + name

# ---------------------------------------------------------------------------
# Handlers MCP Streamable HTTP
# ---------------------------------------------------------------------------
def make_response(req_id, result):
    return JSONResponse(
        {"jsonrpc": "2.0", "id": req_id, "result": result},
        headers={"Content-Type": "application/json"}
    )

def make_error(req_id, code, message):
    return JSONResponse(
        {"jsonrpc": "2.0", "id": req_id, "error": {"code": code, "message": message}},
        headers={"Content-Type": "application/json"}
    )


async def handle_mcp(request: Request):
    # GET /mcp — retourne info serveur
    if request.method == "GET":
        return JSONResponse({
            "name": "iusprudentia-poenalis",
            "version": "1.0.0",
            "protocolVersion": "2025-11-25"
        })

    # DELETE /mcp — ferme une session
    if request.method == "DELETE":
        session_id = request.headers.get("mcp-session-id")
        if session_id and session_id in SESSIONS:
            del SESSIONS[session_id]
        return Response(status_code=200)

    # POST /mcp — traite les requêtes JSON-RPC
    try:
        data = await request.json()
    except Exception:
        return make_error(None, -32700, "Parse error")

    logger.info(f"POST /mcp method={data.get('method')} session={request.headers.get('mcp-session-id', 'none')}")

    method = data.get("method", "")
    params = data.get("params", {})
    req_id = data.get("id", 1)

    # Gestion des sessions
    session_id = request.headers.get("mcp-session-id")

    if method == "initialize":
        # Créer nouvelle session
        new_session_id = str(uuid.uuid4())
        SESSIONS[new_session_id] = {"initialized": True}
        result = {
            "protocolVersion": "2025-11-25",
            "capabilities": {"tools": {"listChanged": False}},
            "serverInfo": {"name": "iusprudentia-poenalis", "version": "1.0.0"}
        }
        response = make_response(req_id, result)
        response.headers["mcp-session-id"] = new_session_id
        return response

    if method == "notifications/initialized":
        return Response(status_code=202)

    if method == "ping":
        return make_response(req_id, {})

    if method == "tools/list":
        return make_response(req_id, {"tools": TOOLS})

    if method == "tools/call":
        tool = params.get("name", "")
        args = params.get("arguments", {})
        try:
            result = call_tool(tool, args)
        except Exception as e:
            result = "Erreur : " + str(e)
        return make_response(req_id, {"content": [{"type": "text", "text": result}]})

    return make_error(req_id, -32601, "Method not found")


async def handle_health(request: Request):
    return JSONResponse({
        "status": "ok",
        "name": "iusprudentia-poenalis",
        "arrets": len(ARRETS),
        "version": "8.0"
    })


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------
middleware = [
    Middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_methods=["GET", "POST", "DELETE", "OPTIONS"],
        allow_headers=["*"],
        expose_headers=["mcp-session-id"]
    )
]

app = Starlette(
    routes=[
        Route("/", handle_health, methods=["GET", "HEAD"]),
        Route("/mcp", handle_mcp, methods=["GET", "POST", "DELETE"]),
    ],
    middleware=middleware
)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
