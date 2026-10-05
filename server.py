# Criminomos - Serveur MCP v18
# Sans authentification + fix SSL bger.ch + recherche profonde niveau 3
import json
import re
import unicodedata
import os
import uuid
import logging
import io
from pathlib import Path
from collections import defaultdict
from datetime import datetime, timezone

import httpx
import pandas as pd
from openpyxl import load_workbook
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
# Configuration
# ---------------------------------------------------------------------------
GDRIVE_FILE_ID  = os.environ.get("GDRIVE_FILE_ID", "18ylKTce78zSdEpeJ8tBbPchPIGs-kG4w")
RELOAD_KEY      = os.environ.get("RELOAD_KEY", "iuris2026!")
GDRIVE_URL      = f"https://docs.google.com/spreadsheets/d/{GDRIVE_FILE_ID}/export?format=xlsx"
BASE_URL        = os.environ.get("BASE_URL", "https://mcp.criminomos.ch")
MAX_LOG_ENTRIES = 500

# ---------------------------------------------------------------------------
# Journal d'accès
# ---------------------------------------------------------------------------
ACCESS_LOG = []

def log_access(client_id, ip, action):
    entry = {
        "client_id": client_id,
        "ip":        ip,
        "action":    action,
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    }
    ACCESS_LOG.append(entry)
    if len(ACCESS_LOG) > MAX_LOG_ENTRIES:
        ACCESS_LOG.pop(0)
    logger.info(f"ACCESS | client={client_id} | ip={ip} | action={action}")

# ---------------------------------------------------------------------------
# Dictionnaire multilingue
# ---------------------------------------------------------------------------
MULTILANG = {
    "expulsion": ["landesverweisung", "espulsione"],
    "viol": ["vergewaltigung", "violenza carnale"],
    "meurtre": ["mord", "omicidio"],
    "homicide": ["tötung", "omicidio"],
    "lésions corporelles": ["körperverletzung", "lesioni corporali"],
    "escroquerie": ["betrug", "truffa"],
    "abus de confiance": ["veruntreuung", "appropriazione indebita"],
    "contrainte": ["nötigung", "coazione"],
    "menaces": ["drohung", "minaccia"],
    "brigandage": ["raub", "rapina"],
    "vol": ["diebstahl", "furto"],
    "faux": ["urkundenfälschung", "falsità in documenti"],
    "diffamation": ["verleumdung", "diffamazione"],
    "injure": ["beschimpfung", "ingiuria"],
    "incendie": ["brandstiftung", "incendio"],
    "tentative": ["versuch", "tentativo"],
    "complicité": ["gehilfenschaft", "complicità"],
    "instigation": ["anstiftung", "istigazione"],
    "récidive": ["rückfall", "recidiva"],
    "concours": ["konkurrenz", "concorso"],
    "détention provisoire": ["untersuchungshaft", "carcerazione preventiva"],
    "détention": ["haft", "detenzione"],
    "arrestation": ["verhaftung", "arresto"],
    "ordonnance pénale": ["strafbefehl", "decreto d'accusa"],
    "classement": ["einstellung", "abbandono"],
    "non-entrée en matière": ["nichtanhandnahme", "non luogo a procedere"],
    "acquittement": ["freispruch", "assoluzione"],
    "condamnation": ["verurteilung", "condanna"],
    "appel": ["berufung", "appello"],
    "recours": ["beschwerde", "ricorso"],
    "révision": ["revision", "revisione"],
    "récusation": ["ausstand", "ricusazione"],
    "scellés": ["siegelung", "sigillazione"],
    "séquestre": ["beschlagnahme", "sequestro"],
    "perquisition": ["hausdurchsuchung", "perquisizione"],
    "surveillance": ["überwachung", "sorveglianza"],
    "expertise": ["gutachten", "perizia"],
    "peine privative de liberté": ["freiheitsstrafe", "pena detentiva"],
    "peine pécuniaire": ["geldstrafe", "pena pecuniaria"],
    "sursis": ["aufschub", "sospensione"],
    "libération conditionnelle": ["bedingte entlassung", "liberazione condizionale"],
    "internement": ["verwahrung", "internamento"],
    "mesure": ["massnahme", "misura"],
    "culpabilité": ["schuld", "colpevolezza"],
    "négligence": ["fahrlässigkeit", "negligenza"],
    "causalité": ["kausalität", "causalità"],
    "prescription": ["verjährung", "prescrizione"],
    "légitime défense": ["notwehr", "legittima difesa"],
    "état de nécessité": ["notstand", "stato di necessità"],
    "victime": ["opfer", "vittima"],
    "prévenu": ["beschuldigte", "imputato"],
    "ministère public": ["staatsanwaltschaft", "ministero pubblico"],
}

def expand_query_multilang(query_words):
    expanded    = list(query_words)
    query_lower = " ".join(query_words).lower()
    for fr_term, translations in MULTILANG.items():
        if fr_term in query_lower:
            expanded.extend(translations)
    return expanded

# ---------------------------------------------------------------------------
# Chargement des données
# ---------------------------------------------------------------------------
def parse_excel(content: bytes) -> list:
    sheets  = pd.read_excel(io.BytesIO(content), sheet_name=None)
    wb      = load_workbook(io.BytesIO(content))
    url_map = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.hyperlink and cell.value:
                    url_map[str(cell.value).strip()] = cell.hyperlink.target
    all_rows = []
    for sheet_name, df in sheets.items():
        if sheet_name == "2021-2024":
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)
        for _, row in df.iterrows():
            r = {}
            for col in df.columns:
                val = row.get(col, None)
                if pd.notna(val):
                    r[str(col).strip()] = str(val)
            all_rows.append(r)
    trimmed = []
    for r in all_rows:
        arret    = r.get("Arrêt", "").strip()
        parution = r.get("Date de parution", "")
        decision = r.get("Date de la décision", "")
        trimmed.append({
            "arret":    arret,
            "parution": parution.split(" ")[0] if "00:00:00" in parution else parution,
            "decision": decision.split(" ")[0] if "00:00:00" in decision else decision,
            "objet":    r.get("Objet", ""),
            "articles": r.get("Articles", ""),
            "resume":   r.get("Résumé", ""),
            "langue":   r.get("Langue", "").strip() if r.get("Langue") else "",
            "interet":  r.get("Arrêt d'intérêt", "").strip() if r.get("Arrêt d'intérêt") else r.get("Arrêt d intérêt", ""),
            "admis":    r.get("Admis/rejeté", ""),
            "peine":    r.get("Peine prononcée", ""),
            "url":      url_map.get(arret, ""),
        })
    return [r for r in trimmed if r["arret"]]


def load_from_gdrive():
    logger.info("Téléchargement depuis Google Drive...")
    resp = httpx.get(GDRIVE_URL, timeout=60, follow_redirects=True)
    resp.raise_for_status()
    data  = parse_excel(resp.content)
    by_id = {r["arret"]: r for r in data}
    logger.info(f"=== {len(data)} arrêts chargés ===")
    return data, by_id


def load_from_local():
    local = Path(__file__).parent / "arrets.json"
    if local.exists():
        with open(local, encoding="utf-8") as f:
            data = json.load(f)
        by_id = {r["arret"]: r for r in data if r.get("arret")}
        logger.info(f"=== {len(data)} arrêts chargés (local) ===")
        return data, by_id
    return [], {}


try:
    ARRETS, ARRETS_BY_ID = load_from_gdrive()
except Exception as e:
    logger.warning(f"Google Drive inaccessible ({e}), chargement local.")
    ARRETS, ARRETS_BY_ID = load_from_local()

# ---------------------------------------------------------------------------
# Outils MCP
# ---------------------------------------------------------------------------
TOOLS = [
    {
        "name": "search_arrets",
        "description": "Recherche des arrets du Tribunal federal suisse en droit penal. Recherche automatiquement dans les trois langues (FR/DE/IT).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "query":      {"type": "string",  "description": "Mots-cles de recherche"},
                "infraction": {"type": "string",  "description": "Type d infraction ex: Expulsion"},
                "article":    {"type": "string",  "description": "Article de loi ex: 66a CP"},
                "annee":      {"type": "string",  "description": "Annee ex: 2024"},
                "langue":     {"type": "string",  "description": "Langue : F, D ou I"},
                "limite":     {"type": "integer", "description": "Nombre de resultats max 30"}
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
        "description": "Extrait les ATF et arrets TF cites dans un arret (niveau 1).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "arret_id": {"type": "string", "description": "Numero d arret ex: 6B_409/2024"}
            },
            "required": ["arret_id"]
        }
    },
    {
        "name": "get_references_deep",
        "description": """Remonte les references sur plusieurs niveaux de profondeur.
Par defaut explore 2 niveaux. Pour une recherche exhaustive, utilisez profondeur=3.""",
        "inputSchema": {
            "type": "object",
            "properties": {
                "arret_id":   {"type": "string",  "description": "Numero d arret de depart"},
                "max_refs":   {"type": "integer", "description": "Nombre max de refs niveau 1 (defaut 10, max 15)"},
                "profondeur": {"type": "integer", "description": "Profondeur : 2 (defaut) ou 3 (exhaustif)"}
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
        if w in normalize(arret.get("objet",    "")): score += 4
        if w in normalize(arret.get("resume",   "")): score += 2
        if w in normalize(arret.get("articles", "")): score += 3
        if w in normalize(arret.get("arret",    "")): score += 5
    if arret.get("interet") == "oui":
        score += 1
    return score


def search_arrets(query="", infraction="", article="", annee="", langue="", limite=10):
    limite      = min(int(limite), 30)
    query_words = query.strip().split() if query.strip() else []
    expanded    = expand_query_multilang(query_words)
    results     = []
    for arret in ARRETS:
        if infraction and not normalize(arret.get("objet", "")).startswith(normalize(infraction)):
            continue
        if article and normalize(article) not in normalize(arret.get("articles", "")):
            continue
        if annee and not (arret.get("decision", "").startswith(annee) or
                          arret.get("parution",  "").startswith(annee)):
            continue
        if langue and arret.get("langue", "").upper() != langue.upper():
            continue
        score = score_arret(arret, expanded) if expanded else 1
        if score > 0 or not query_words:
            results.append((score, arret))
    results.sort(key=lambda x: x[0], reverse=True)
    results = results[:limite]
    if not results:
        return "Aucun arret trouve."
    extra  = [w for w in expanded if w not in query_words]
    header = ("Recherche etendue aux equivalents : " + ", ".join(extra) + "\n\n") if extra else ""
    lines  = [header + str(len(results)) + " arret(s) trouve(s)\n"]
    for _, r in results:
        lines.append("### " + r["arret"])
        lines.append("- Objet : "    + r.get("objet",    "-"))
        lines.append("- Date : "     + r.get("decision", "-"))
        lines.append("- Langue : "   + r.get("langue",   "-"))
        lines.append("- Articles : " + r.get("articles", "-"))
        if r.get("admis"):             lines.append("- Resultat : " + r["admis"])
        if r.get("interet") == "oui": lines.append("- Arret d interet")
        if r.get("resume"):            lines.append("- Resume : " + r["resume"])
        if r.get("url"):               lines.append("- URL : " + r["url"])
        lines.append("")
    return "\n".join(lines)


def _fetch_text(url):
    # verify=False nécessaire : bger.ch a une chaîne de certificats non reconnue sur Render
    resp = httpx.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"},
                     follow_redirects=True, verify=False)
    soup = BeautifulSoup(resp.text, "html.parser")
    for tag in soup(["script", "style", "nav", "header", "footer"]):
        tag.decompose()
    lines = [l.rstrip() for l in soup.get_text(separator="\n").splitlines() if l.strip()]
    return "\n".join(lines)


def _extract_refs(text, arret_id=""):
    atf_refs = sorted(set(re.findall(r"ATF\s+\d{2,3}\s+[IVX]+\s+\d+", text)))
    tf_refs  = sorted(set(
        r for r in re.findall(r"\b[0-9][A-Z]{1,2}_\d{1,4}/20\d{2}\b", text)
        if r != arret_id
    ))
    return atf_refs, tf_refs


def _arret_url(arret_id):
    arret = ARRETS_BY_ID.get(arret_id)
    if arret and arret.get("url"):
        return arret["url"]
    cid = re.sub(r"[^A-Za-z0-9_/.-]", "", arret_id)[:40]
    return ("https://www.bger.ch/ext/eurospider/live/fr/php/aza/http/index.php"
            f"?lang=fr&type=show_document&highlight_docid=aza://{cid}")


def get_fulltext(arret_id):
    url = _arret_url(arret_id)
    try:
        text = _fetch_text(url)
        return "Arret " + arret_id + "\nURL : " + url + "\n" + "-"*60 + "\n\n" + text[:30000]
    except Exception as e:
        return "Erreur : " + str(e)


def get_references(arret_id):
    url = _arret_url(arret_id)
    try:
        text = _fetch_text(url)
    except Exception as e:
        return "Erreur : " + str(e)
    atf_refs, tf_refs = _extract_refs(text, arret_id)
    lines = ["References de l arret " + arret_id + "\n"]
    if atf_refs:
        lines.append("ATF cites (" + str(len(atf_refs)) + ") :")
        for r in atf_refs: lines.append("- " + r)
    if tf_refs:
        lines.append("\nArrets TF cites (" + str(len(tf_refs)) + ") :")
        for r in tf_refs:
            if r in ARRETS_BY_ID:
                a = ARRETS_BY_ID[r]
                lines.append("- " + r + " - " + a.get("objet", "-") +
                              " (" + a.get("decision", "-") + ") [base]")
            else:
                lines.append("- " + r)
    if not atf_refs and not tf_refs:
        lines.append("Aucune reference trouvee.")
    return "\n".join(lines)


def get_references_deep(arret_id, max_refs=10, profondeur=2):
    max_refs   = min(int(max_refs), 15)
    profondeur = min(max(int(profondeur), 2), 3)
    url        = _arret_url(arret_id)
    lines      = ["=== REFERENCES PROFONDES (niveau " + str(profondeur) + ") : " + arret_id + " ===\n"]
    try:
        text1 = _fetch_text(url)
    except Exception as e:
        return "Erreur : " + str(e)
    atf1, tf1 = _extract_refs(text1, arret_id)
    lines.append("NIVEAU 1 — References directes")
    lines.append("ATF cites : "       + (", ".join(atf1)     if atf1 else "aucun"))
    lines.append("Arrets TF cites : " + (", ".join(tf1[:30]) if tf1  else "aucun"))
    lines.append("")
    explored_tf  = {arret_id}
    explored_atf = set()
    niveau2_tf_refs  = {}
    niveau2_atf_refs = {}
    for ref_id in tf1[:max_refs]:
        if ref_id in explored_tf: continue
        explored_tf.add(ref_id)
        try:
            text2     = _fetch_text(_arret_url(ref_id))
            atf2, tf2 = _extract_refs(text2, ref_id)
            niveau2_tf_refs[ref_id] = (atf2, tf2)
            meta  = ARRETS_BY_ID.get(ref_id)
            objet = meta.get("objet", "—") if meta else "hors base"
            lines.append("\n  [N2] " + ref_id + " (" + objet + ")")
            if atf2: lines.append("       ATF cites : "       + ", ".join(atf2[:10]))
            if tf2:  lines.append("       Arrets TF cites : " + ", ".join(tf2[:10]))
            if not atf2 and not tf2: lines.append("       (aucune reference)")
        except Exception as e:
            lines.append("\n  [N2] " + ref_id + " — Erreur : " + str(e))
    for atf_ref in atf1[:8]:
        if atf_ref in explored_atf: continue
        explored_atf.add(atf_ref)
        m = re.match(r"ATF\s+(\d{2,3})\s+([IVX]+)\s+(\d+)", atf_ref)
        if m:
            vol, part, page = m.groups()
            atf_url = ("https://www.bger.ch/ext/eurospider/live/fr/php/clir/http/index.php"
                       f"?lang=fr&type=show_document&highlight_docid=atf:///{vol}/{part}/{page}")
            try:
                text_atf  = _fetch_text(atf_url)
                atf2, tf2 = _extract_refs(text_atf, "")
                niveau2_atf_refs[atf_ref] = (atf2, tf2)
                lines.append("\n  [N2] " + atf_ref)
                if atf2: lines.append("       ATF cites : "       + ", ".join(atf2[:10]))
                if tf2:  lines.append("       Arrets TF cites : " + ", ".join(tf2[:10]))
            except Exception as e:
                lines.append("\n  [N2] " + atf_ref + " — Erreur : " + str(e))
    if profondeur >= 3:
        lines.append("\n\nNIVEAU 3 — Exploration exhaustive")
        niveau3_candidates = []
        for ref_id, (atf2, tf2) in niveau2_tf_refs.items():
            for r in tf2[:5]:
                if r not in explored_tf:
                    niveau3_candidates.append(r)
        for atf_ref, (atf2, tf2) in niveau2_atf_refs.items():
            for r in tf2[:5]:
                if r not in explored_tf:
                    niveau3_candidates.append(r)
        niveau3_candidates = list(dict.fromkeys(niveau3_candidates))[:15]
        if not niveau3_candidates:
            lines.append("  Aucune nouvelle reference au niveau 3.")
        else:
            for ref_id in niveau3_candidates:
                explored_tf.add(ref_id)
                try:
                    text3     = _fetch_text(_arret_url(ref_id))
                    atf3, tf3 = _extract_refs(text3, ref_id)
                    meta      = ARRETS_BY_ID.get(ref_id)
                    objet     = meta.get("objet", "—") if meta else "hors base"
                    lines.append("\n    [N3] " + ref_id + " (" + objet + ")")
                    if atf3: lines.append("         ATF cites : "       + ", ".join(atf3[:8]))
                    if tf3:  lines.append("         Arrets TF cites : " + ", ".join(tf3[:8]))
                    if not atf3 and not tf3: lines.append("         (aucune reference)")
                except Exception as e:
                    lines.append("\n    [N3] " + ref_id + " — Erreur : " + str(e))
        lines.append(f"\n\nTotal explores : {len(explored_tf)} arrets | {len(explored_atf)} ATF")
    return "\n".join(lines)


def get_arret_by_reference(reference):
    reference = reference.strip()
    if re.match(r"^[0-9][A-Z]{1,2}_\d{1,4}/20\d{2}$", reference):
        return get_fulltext(reference)
    m = re.match(r"ATF\s+(\d{2,3})\s+([IVX]+)\s+(\d+)", reference, re.IGNORECASE)
    if m:
        vol, part, page = m.groups()
        url = ("https://www.bger.ch/ext/eurospider/live/fr/php/clir/http/index.php"
               f"?lang=fr&type=show_document&highlight_docid=atf:///{vol}/{part}/{page}")
        try:
            text = _fetch_text(url)
            return reference + "\nURL : " + url + "\n" + "-"*60 + "\n\n" + text[:30000]
        except Exception as e:
            return "Erreur : " + str(e)
    return "Format non reconnu : " + reference


def call_tool(name, args):
    if name == "search_arrets":            return search_arrets(**args)
    elif name == "get_fulltext":           return get_fulltext(**args)
    elif name == "get_references":         return get_references(**args)
    elif name == "get_references_deep":    return get_references_deep(**args)
    elif name == "get_arret_by_reference": return get_arret_by_reference(**args)
    return "Outil inconnu : " + name

# ---------------------------------------------------------------------------
# Helpers JSON-RPC
# ---------------------------------------------------------------------------
def ok(req_id, result):
    return JSONResponse({"jsonrpc": "2.0", "id": req_id, "result": result},
                        headers={"Content-Type": "application/json"})

def err(req_id, code, msg):
    return JSONResponse({"jsonrpc": "2.0", "id": req_id, "error": {"code": code, "message": msg}},
                        headers={"Content-Type": "application/json"})

# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------
async def handle_health(request: Request):
    return JSONResponse({
        "status":  "ok",
        "name":    "criminomos",
        "arrets":  len(ARRETS),
        "version": "18.0",
        "auth":    "disabled",
    })


async def handle_reload(request: Request):
    global ARRETS, ARRETS_BY_ID
    key = request.query_params.get("key", "")
    if key != RELOAD_KEY:
        return JSONResponse({"error": "Clé invalide"}, status_code=403)
    try:
        ARRETS, ARRETS_BY_ID = load_from_gdrive()
        return JSONResponse({"status": "ok", "arrets": len(ARRETS)})
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)


async def handle_log(request: Request):
    key = request.query_params.get("key", "")
    if key != RELOAD_KEY:
        return JSONResponse({"error": "Clé invalide"}, status_code=403)
    return JSONResponse({"total": len(ACCESS_LOG), "entries": list(reversed(ACCESS_LOG[-100:]))})


async def handle_mcp(request: Request):
    if request.method == "GET":
        return JSONResponse(
            {"name": "criminomos", "version": "18.0", "protocolVersion": "2025-11-25"},
            headers={"MCP-Protocol-Version": "2025-11-25"}
        )
    if request.method == "HEAD":
        return Response(status_code=200, headers={"MCP-Protocol-Version": "2025-11-25"})
    if request.method == "DELETE":
        return Response(status_code=200)

    ip = request.client.host if request.client else "unknown"

    try:
        data = await request.json()
    except Exception:
        return err(None, -32700, "Parse error")

    method = data.get("method", "")
    params = data.get("params", {})
    req_id = data.get("id", 1)

    if method == "initialize":
        sid = str(uuid.uuid4())
        log_access("anonymous", ip, "connect")
        resp = ok(req_id, {
            "protocolVersion": "2025-11-25",
            "capabilities":    {"tools": {"listChanged": False}},
            "serverInfo":      {"name": "criminomos", "version": "18.0"}
        })
        resp.headers["mcp-session-id"] = sid
        return resp

    if method == "notifications/initialized":
        return Response(status_code=202)

    if method == "ping":
        return ok(req_id, {})

    if method == "tools/list":
        log_access("anonymous", ip, "tools/list")
        return ok(req_id, {"tools": TOOLS})

    if method == "tools/call":
        tool = params.get("name", "")
        args = params.get("arguments", {})
        log_access("anonymous", ip, "call:" + tool)
        try:
            result = call_tool(tool, args)
        except Exception as e:
            result = "Erreur : " + str(e)
        return ok(req_id, {"content": [{"type": "text", "text": result}]})

    return err(req_id, -32601, "Method not found")


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------
middleware = [
    Middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_methods=["GET", "POST", "DELETE", "HEAD", "OPTIONS"],
        allow_headers=["*"],
        expose_headers=["mcp-session-id", "MCP-Protocol-Version"]
    )
]

app = Starlette(
    routes=[
        Route("/",       handle_health, methods=["GET", "HEAD"]),
        Route("/reload", handle_reload, methods=["GET"]),
        Route("/log",    handle_log,    methods=["GET"]),
        Route("/mcp",    handle_mcp,    methods=["GET", "POST", "DELETE", "HEAD"]),
    ],
    middleware=middleware
)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
