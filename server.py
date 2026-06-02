# Criminomos - Serveur MCP v10
# Recherche multilingue + références profondes
import json
import re
import unicodedata
import os
import uuid
import logging
import io
from pathlib import Path

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
GDRIVE_FILE_ID = os.environ.get("GDRIVE_FILE_ID", "18ylKTce78zSdEpeJ8tBbPchPIGs-kG4w")
RELOAD_KEY     = os.environ.get("RELOAD_KEY", "iuris2026!")
GDRIVE_URL     = f"https://docs.google.com/spreadsheets/d/{GDRIVE_FILE_ID}/export?format=xlsx"

# ---------------------------------------------------------------------------
# Dictionnaire multilingue juridique pénal suisse (FR → DE/IT)
# ---------------------------------------------------------------------------
MULTILANG = {
    # Infractions
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
    # Procédure
    "détention provisoire": ["untersuchungshaft", "carcerazione preventiva"],
    "détention": ["haft", "detenzione"],
    "arrestation": ["verhaftung", "arresto"],
    "mise en accusation": ["anklage", "accusa"],
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
    # Peines
    "peine privative de liberté": ["freiheitsstrafe", "pena detentiva"],
    "peine pécuniaire": ["geldstrafe", "pena pecuniaria"],
    "travail d'intérêt général": ["gemeinnützige arbeit", "lavoro di pubblica utilità"],
    "sursis": ["aufschub", "sospensione"],
    "libération conditionnelle": ["bedingte entlassung", "liberazione condizionale"],
    "internement": ["verwahrung", "internamento"],
    "mesure": ["massnahme", "misura"],
    "expulsion": ["landesverweisung", "espulsione"],
    # Notions générales
    "culpabilité": ["schuld", "colpevolezza"],
    "dol": ["vorsatz", "dolo"],
    "négligence": ["fahrlässigkeit", "negligenza"],
    "causalité": ["kausalität", "causalità"],
    "prescription": ["verjährung", "prescrizione"],
    "légitime défense": ["notwehr", "legittima difesa"],
    "état de nécessité": ["notstand", "stato di necessità"],
    "victime": ["opfer", "vittima"],
    "plaignant": ["geschädigte", "accusatore privato"],
    "prévenu": ["beschuldigte", "imputato"],
    "ministère public": ["staatsanwaltschaft", "ministero pubblico"],
}

def expand_query_multilang(query_words):
    """Étend une liste de mots-clés avec leurs équivalents multilingues."""
    expanded = list(query_words)
    query_lower = " ".join(query_words).lower()
    for fr_term, translations in MULTILANG.items():
        if fr_term in query_lower:
            expanded.extend(translations)
    return expanded

# ---------------------------------------------------------------------------
# Chargement des données
# ---------------------------------------------------------------------------
def parse_excel(content: bytes) -> list:
    sheets = pd.read_excel(io.BytesIO(content), sheet_name=None)
    wb     = load_workbook(io.BytesIO(content))
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
    logger.info(f"Téléchargement depuis Google Drive...")
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

SESSIONS = {}

# ---------------------------------------------------------------------------
# Outils MCP
# ---------------------------------------------------------------------------
TOOLS = [
    {
        "name": "search_arrets",
        "description": """Recherche des arrets du Tribunal federal suisse en droit penal.
Recherche automatiquement dans les trois langues (FR/DE/IT) — si vous tapez 'expulsion',
l'outil cherche aussi 'Landesverweisung' et 'espulsione' automatiquement.
Utilisez cet outil en premier pour identifier les arrets pertinents.""",
        "inputSchema": {
            "type": "object",
            "properties": {
                "query":      {"type": "string",  "description": "Mots-cles en francais, allemand ou italien"},
                "infraction": {"type": "string",  "description": "Type d infraction ex: Expulsion, Viol"},
                "article":    {"type": "string",  "description": "Article de loi ex: 66a CP, 190 CP"},
                "annee":      {"type": "string",  "description": "Annee ex: 2024"},
                "langue":     {"type": "string",  "description": "Langue de l arret : F, D ou I"},
                "limite":     {"type": "integer", "description": "Nombre de resultats max 30"}
            },
            "required": ["query"]
        }
    },
    {
        "name": "get_fulltext",
        "description": "Charge le texte integral d un arret depuis bger.ch. Utilise pour lire le raisonnement juridique complet.",
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
        "description": """Remonte les references sur 2 niveaux de profondeur.
Charge l arret, extrait ses references (niveau 1), puis charge chaque reference
et extrait ses propres references (niveau 2). Retourne un arbre complet des citations.
Utilise pour une analyse approfondie de la jurisprudence de reference sur un sujet.""",
        "inputSchema": {
            "type": "object",
            "properties": {
                "arret_id": {"type": "string", "description": "Numero d arret de depart ex: 6B_409/2024"},
                "max_refs":  {"type": "integer", "description": "Nombre max de references niveau 1 a explorer (defaut 5, max 10)"}
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

    # Expansion multilingue automatique
    expanded_words = expand_query_multilang(query_words)

    results = []
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
        score = score_arret(arret, expanded_words) if expanded_words else 1
        if score > 0 or not query_words:
            results.append((score, arret))

    results.sort(key=lambda x: x[0], reverse=True)
    results = results[:limite]
    if not results:
        return "Aucun arret trouve."

    # Indiquer les termes multilingues utilisés
    extra_terms = [w for w in expanded_words if w not in query_words]
    header = ""
    if extra_terms:
        header = "Recherche etendue aux equivalents : " + ", ".join(extra_terms) + "\n\n"

    lines = [header + str(len(results)) + " arret(s) trouve(s)\n"]
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
    resp = httpx.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"}, follow_redirects=True)
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
        return "Arret " + arret_id + "\nURL : " + url + "\n" + "-"*60 + "\n\n" + text[:15000]
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


def get_references_deep(arret_id, max_refs=5):
    """Remonte les références sur 2 niveaux."""
    max_refs = min(int(max_refs), 10)
    url = _arret_url(arret_id)
    lines = ["=== REFERENCES PROFONDES : " + arret_id + " ===\n"]

    # Niveau 1
    try:
        text1 = _fetch_text(url)
    except Exception as e:
        return "Erreur chargement arrêt principal : " + str(e)

    atf1, tf1 = _extract_refs(text1, arret_id)
    lines.append("NIVEAU 1 — References directes")
    lines.append("ATF cites : " + (", ".join(atf1) if atf1 else "aucun"))
    lines.append("Arrets TF cites : " + (", ".join(tf1[:20]) if tf1 else "aucun"))
    lines.append("")

    # Niveau 2 — explorer les premières références TF
    refs_to_explore = tf1[:max_refs]
    if refs_to_explore:
        lines.append("NIVEAU 2 — References des references (top " + str(len(refs_to_explore)) + ")")
        for ref_id in refs_to_explore:
            ref_url = _arret_url(ref_id)
            try:
                text2 = _fetch_text(ref_url)
                atf2, tf2 = _extract_refs(text2, ref_id)
                meta = ARRETS_BY_ID.get(ref_id)
                objet = meta.get("objet", "—") if meta else "hors base"
                lines.append("\n  " + ref_id + " (" + objet + ")")
                if atf2:
                    lines.append("  ATF cites : " + ", ".join(atf2[:5]))
                if tf2:
                    lines.append("  Arrets TF cites : " + ", ".join(tf2[:5]))
                if not atf2 and not tf2:
                    lines.append("  (aucune reference trouvee)")
            except Exception as e:
                lines.append("\n  " + ref_id + " — Erreur : " + str(e))

    # Niveau 2 — explorer les premiers ATF
    atf_to_explore = atf1[:3]
    if atf_to_explore:
        lines.append("\nNIVEAU 2 — References des ATF cites (top 3)")
        for atf_ref in atf_to_explore:
            m = re.match(r"ATF\s+(\d{2,3})\s+([IVX]+)\s+(\d+)", atf_ref)
            if m:
                vol, part, page = m.groups()
                atf_url = ("https://www.bger.ch/ext/eurospider/live/fr/php/clir/http/index.php"
                           f"?lang=fr&type=show_document&highlight_docid=atf:///{vol}/{part}/{page}")
                try:
                    text_atf = _fetch_text(atf_url)
                    atf2, tf2 = _extract_refs(text_atf, "")
                    lines.append("\n  " + atf_ref)
                    if atf2:
                        lines.append("  ATF cites : " + ", ".join(atf2[:5]))
                    if tf2:
                        lines.append("  Arrets TF cites : " + ", ".join(tf2[:5]))
                except Exception as e:
                    lines.append("\n  " + atf_ref + " — Erreur : " + str(e))

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
            return reference + "\nURL : " + url + "\n" + "-"*60 + "\n\n" + text[:15000]
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
# Handlers HTTP
# ---------------------------------------------------------------------------
def ok(req_id, result):
    return JSONResponse({"jsonrpc": "2.0", "id": req_id, "result": result},
                        headers={"Content-Type": "application/json"})

def err(req_id, code, msg):
    return JSONResponse({"jsonrpc": "2.0", "id": req_id, "error": {"code": code, "message": msg}},
                        headers={"Content-Type": "application/json"})


async def handle_health(request: Request):
    return JSONResponse({
        "status":  "ok",
        "name":    "criminomos",
        "arrets":  len(ARRETS),
        "version": "10.0"
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


async def handle_mcp(request: Request):
    if request.method == "GET":
        return JSONResponse({
            "name": "criminomos",
            "version": "10.0",
            "protocolVersion": "2025-11-25"
        })

    if request.method == "DELETE":
        sid = request.headers.get("mcp-session-id")
        if sid and sid in SESSIONS:
            del SESSIONS[sid]
        return Response(status_code=200)

    try:
        data = await request.json()
    except Exception:
        return err(None, -32700, "Parse error")

    method = data.get("method", "")
    params = data.get("params", {})
    req_id = data.get("id", 1)

    logger.info(f"MCP {method}")

    if method == "initialize":
        sid = str(uuid.uuid4())
        SESSIONS[sid] = True
        resp = ok(req_id, {
            "protocolVersion": "2025-11-25",
            "capabilities":    {"tools": {"listChanged": False}},
            "serverInfo":      {"name": "criminomos", "version": "10.0"}
        })
        resp.headers["mcp-session-id"] = sid
        return resp

    if method == "notifications/initialized":
        return Response(status_code=202)

    if method == "ping":
        return ok(req_id, {})

    if method == "tools/list":
        return ok(req_id, {"tools": TOOLS})

    if method == "tools/call":
        tool = params.get("name", "")
        args = params.get("arguments", {})
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
        allow_methods=["GET", "POST", "DELETE", "OPTIONS"],
        allow_headers=["*"],
        expose_headers=["mcp-session-id"]
    )
]

app = Starlette(
    routes=[
        Route("/",       handle_health, methods=["GET", "HEAD"]),
        Route("/reload", handle_reload, methods=["GET"]),
        Route("/mcp",    handle_mcp,    methods=["GET", "POST", "DELETE"]),
    ],
    middleware=middleware
)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
