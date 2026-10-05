import os
import re
import json
import time
import uuid
import base64
import logging
import secrets
import asyncio
import httpx
import msal
from datetime import datetime
from urllib.parse import parse_qs, unquote, urlparse
from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware
from authlib.integrations.starlette_client import OAuth
from authlib.common.urls import add_params_to_qs
from authlib.oauth2.rfc7523 import private_key_jwt_sign
from joserfc.jwk import RSAKey
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.serialization import pkcs12
from dotenv import load_dotenv

load_dotenv()

# ==========================================
# 1. SETUP & OIDC (MICROSOFT ENTRA ID)
# ==========================================
app = FastAPI(title="SharePoint Sync Portal")
# Session Secret aus der .env laden
SESSION_SECRET = os.getenv("SESSION_SECRET", "fallback-nur-fuer-dev")
# Die Bearbeiten-Berechtigung wird beim Login ermittelt - nach 8 Stunden neu anmelden (per SSO unbemerkt),
# damit Änderungen an der Entra-Gruppe zeitnah wirken
SESSION_SECONDS = 8 * 60 * 60
# Öffentliche Adresse, wenn die App hinter einem Reverse-Proxy (z.B. Traefik) läuft, etwa https://sync.example.com
# Bestimmt die Rücksprung-Adresse nach dem Microsoft-Login; leer = Adresse aus der Anfrage (lokaler Betrieb).
PUBLIC_URL = os.getenv("PUBLIC_URL", "").strip().rstrip("/")
# Optional: E-Mail-Domain der Benutzer (z.B. contoso.com) für die Anmeldung ohne Kontoauswahl
LOGIN_DOMAIN_HINT = os.getenv("LOGIN_DOMAIN_HINT", "").strip()
app.add_middleware(
    SessionMiddleware,
    secret_key=SESSION_SECRET,
    max_age=SESSION_SECONDS,
    # Hinter HTTPS das Sitzungs-Cookie nie über unverschlüsselte Verbindungen senden
    https_only=PUBLIC_URL.startswith("https://"),
)

oauth = OAuth()
TENANT_ID = os.getenv("SP_TENANT_ID", "common")
CLIENT_ID = os.getenv("SP_CLIENT_ID")
# Dasselbe Zertifikat wie in sync.py - ein Client-Secret wird nicht mehr benötigt
CERT_PATH = os.getenv("SP_CERT_PATH", "sp_sync_cert.pfx")

class EntraCertificateAuth:
    """Weist die App am Token-Endpunkt mit dem Zertifikat aus (private_key_jwt) statt mit einem Client-Secret"""
    name = "private_key_jwt"

    def __init__(self, pfx_path):
        if not os.path.exists(pfx_path):
            raise SystemExit(f"Zertifikat {pfx_path} nicht gefunden! Abbruch.")
        with open(pfx_path, 'rb') as f:
            private_key, certificate, _ = pkcs12.load_key_and_certificates(f.read(), None)
        self.private_key = RSAKey.import_key(private_key.private_bytes(
            serialization.Encoding.PEM,
            serialization.PrivateFormat.PKCS8,
            serialization.NoEncryption()
        ))
        # Entra erkennt das Zertifikat am SHA-1-Fingerabdruck im JWT-Header
        self.thumbprint = base64.urlsafe_b64encode(certificate.fingerprint(hashes.SHA1())).rstrip(b"=").decode()

    def __call__(self, auth, method, uri, headers, body):
        client_assertion = private_key_jwt_sign(
            self.private_key,
            client_id=auth.client_id,
            token_endpoint=uri,
            header={"x5t": self.thumbprint},
            expires_in=600
        )
        body = add_params_to_qs(
            body or "",
            [
                ("client_id", auth.client_id),
                ("client_assertion_type", "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"),
                ("client_assertion", client_assertion),
            ],
        )
        if "Content-Length" in headers:
            headers["Content-Length"] = str(len(body))
        return uri, headers, body

# SharePoint-Adresse des Tenants, z.B. https://contoso.sharepoint.com (dieselbe Variable wie in sync.py)
TENANT_URL = os.getenv("SP_TENANT_URL", "").strip().rstrip("/")
if not re.fullmatch(r"https://[a-z0-9-]+\.sharepoint\.com", TENANT_URL, re.IGNORECASE):
    raise SystemExit("SP_TENANT_URL fehlt oder ist ungültig (erwartet z.B. https://contoso.sharepoint.com). Abbruch.")
SITE_PATTERN = re.compile(re.escape(TENANT_URL) + r"/(sites|teams)/[^/]+", re.IGNORECASE)

# Für die Site-/Ordner-Auswahl fragt die Web-App SharePoint im Namen des angemeldeten Benutzers ab,
# damit jeder nur sieht, worauf er selbst Zugriff hat (delegierte Berechtigungen der App-Registrierung).
# offline_access liefert ein Refresh-Token, weil das Zugriffstoken nur ca. 1 Stunde gilt.
SHAREPOINT_SCOPES = f"offline_access {TENANT_URL}/AllSites.Read {TENANT_URL}/Sites.Search.All"

oauth.register(
    name='microsoft',
    client_id=CLIENT_ID,
    server_metadata_url=f'https://login.microsoftonline.com/{TENANT_ID}/v2.0/.well-known/openid-configuration',
    client_kwargs={'scope': f'openid profile email {SHAREPOINT_SCOPES}', 'token_endpoint_auth_method': 'private_key_jwt'},
    client_auth_methods=[EntraCertificateAuth(CERT_PATH)],
)

# Die Vorlagen liegen neben dieser Datei; die Daten (config.json, runs, Log) im Arbeitsverzeichnis.
# Im Docker-Container sind das zwei verschiedene Ordner (/app und /data).
templates = Jinja2Templates(directory=os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates"))

CONFIG_FILE = 'config.json'
LOG_FILE = 'sharepoint_sync.log'
RUNS_DIR = 'runs'

# Wer darf die Web-App benutzen? Alle anderen werden nach dem Microsoft-Login abgewiesen.
# SYNC_ADMIN_GROUP_ID: Objekt-ID einer Entra-Gruppe (benötigt den "groups"-Claim im ID-Token)
# SYNC_ADMINS:         zusätzlich/alternativ eine kommagetrennte Liste von Anmeldenamen
# Sind beide leer, darf jeder Benutzer des Tenants die Web-App benutzen.
SYNC_ADMIN_GROUP_ID = os.getenv("SYNC_ADMIN_GROUP_ID", "").strip().lower()
SYNC_ADMINS = [a.strip().lower() for a in os.getenv("SYNC_ADMINS", "").split(",") if a.strip()]

# Ein Lauf ohne Lebenszeichen seit so vielen Sekunden gilt als abgebrochen
STALE_RUN_SECONDS = 15 * 60
# Maximal angezeigte Ereignisse pro Lauf (Fehler werden immer angezeigt)
MAX_EVENTS_SHOWN = 2000

STATUS_LABELS = {
    "ok": ("Erfolgreich", "bg-green-100 text-green-800"),
    "errors": ("Mit Fehlern", "bg-yellow-100 text-yellow-800"),
    "failed": ("Fehlgeschlagen", "bg-red-100 text-red-800"),
    "running": ("Läuft", "bg-blue-100 text-blue-800"),
    "aborted": ("Abgebrochen", "bg-gray-200 text-gray-700"),
}
ACTION_LABELS = {
    "copied": ("Kopiert", "text-green-700"),
    "deleted": ("Gelöscht", "text-orange-700"),
    "folder": ("Ordner erstellt", "text-blue-700"),
    "error": ("Fehler", "text-red-700"),
}

def format_datetime(value):
    if not value:
        return "–"
    try:
        return datetime.fromisoformat(value).strftime("%d.%m.%Y %H:%M:%S")
    except ValueError:
        return value

def format_duration(seconds):
    if seconds is None:
        return "–"
    minutes, secs = divmod(int(seconds), 60)
    hours, minutes = divmod(minutes, 60)
    if hours:
        return f"{hours} h {minutes:02d} min"
    if minutes:
        return f"{minutes} min {secs:02d} s"
    return f"{secs} s"

templates.env.filters["dt"] = format_datetime
def format_initials(name):
    """Initialen für den Platzhalter, wenn es kein Profilbild gibt: 'Erika Mustermann | Contoso' -> 'EM'"""
    words = [word for word in re.split(r"[^\w]+", (name or "").split("|")[0]) if word]
    if not words:
        return "?"
    return (words[0][0] + (words[-1][0] if len(words) > 1 else "")).upper()

templates.env.filters["duration"] = format_duration
templates.env.filters["initials"] = format_initials
# Anzeigename der Anwendung (Kopfzeile, Seitentitel)
templates.env.globals["APP_TITLE"] = os.getenv("APP_TITLE", "").strip() or "SharePoint Sync"
templates.env.globals["TENANT_URL"] = TENANT_URL
templates.env.globals["STATUS_LABELS"] = STATUS_LABELS
templates.env.globals["ACTION_LABELS"] = ACTION_LABELS

# ==========================================
# 2. AUTHENTIFIZIERUNG
# ==========================================
@app.get("/login")
async def login(request: Request, switch: str = ""):
    redirect_uri = f"{PUBLIC_URL}/auth" if PUBLIC_URL else request.url_for('auth')
    params = {}
    # Der domain_hint (z.B. contoso.com) erzwingt echtes Seamless SSO ohne "Konto auswählen"-Bildschirm
    if LOGIN_DOMAIN_HINT:
        params["domain_hint"] = LOGIN_DOMAIN_HINT
    if switch:
        # Nach "Kein Zugriff": Kontoauswahl anzeigen, sonst meldet SSO sofort wieder dasselbe Konto an
        params["prompt"] = "select_account"
    return await oauth.microsoft.authorize_redirect(request, redirect_uri, **params)

@app.get("/auth")
async def auth(request: Request):
    token = await oauth.microsoft.authorize_access_token(request)
    user = token.get('userinfo')
    if user:
        allowed = not access_is_restricted() or is_sync_admin(user)
        # Hilft bei der Fehlersuche, wenn jemand trotz Gruppenmitgliedschaft abgewiesen wird
        logging.getLogger("uvicorn.error").info(
            "Login %s: groups-Claim %s, Zugriff %s",
            user.get("preferred_username"),
            f"mit {len(user.get('groups') or [])} Gruppen" if "groups" in user else "fehlt",
            "erlaubt" if allowed else "verweigert",
        )
        if not allowed:
            # Keine Sitzung anlegen - wer nicht berechtigt ist, sieht nur diese Seite
            request.session.pop('user', None)
            return templates.TemplateResponse(
                request=request,
                name="denied.html",
                status_code=403,
                context={
                    "user": None,
                    "account": user.get("preferred_username") or user.get("name"),
                    "groups_missing": bool(SYNC_ADMIN_GROUP_ID) and "groups" not in user,
                    "groups_count": len(user.get("groups") or []),
                },
            )

        # Die SharePoint-Tokens des Benutzers bleiben auf dem Server, im Cookie steht nur ein Schlüssel dazu
        sid = secrets.token_urlsafe(16)
        store_sharepoint_token(sid, token)
        # Nur das Nötigste in die Session legen: Das Session-Cookie fasst ca. 4 KB,
        # die Gruppenliste aus dem Token würde es bei vielen Gruppen sprengen.
        request.session['user'] = {
            "sid": sid,
            "name": user.get("name"),
            "preferred_username": user.get("preferred_username"),
            "email": user.get("email"),
            "is_sync_admin": is_sync_admin(user),
        }
    return RedirectResponse(url='/')

@app.get("/logout")
async def logout(request: Request):
    user = request.session.pop('user', None)
    if user:
        USER_TOKENS.pop(user.get("sid"), None)
    return RedirectResponse(url='/login')

# SharePoint-Tokens der angemeldeten Benutzer: Sitzungs-Schlüssel -> Token.
# Nur im Arbeitsspeicher; nach einem Neustart der Web-App ist für die Ordner-Auswahl ein neuer Login nötig.
USER_TOKENS = {}

def store_sharepoint_token(sid, token):
    now = time.time()
    # Einträge abgelaufener Sitzungen entfernen
    for old_sid in [key for key, entry in USER_TOKENS.items() if entry["session_ends"] < now]:
        del USER_TOKENS[old_sid]
    USER_TOKENS[sid] = {
        "access_token": token.get("access_token"),
        "refresh_token": token.get("refresh_token"),
        "expires_at": token.get("expires_at") or now + int(token.get("expires_in") or 0),
        "session_ends": now + SESSION_SECONDS,
    }

async def get_sharepoint_token(user):
    """Liefert ein gültiges SharePoint-Zugriffstoken des Benutzers oder None (dann ist ein neuer Login nötig)"""
    entry = USER_TOKENS.get(user.get("sid"))
    if not entry or not entry["access_token"]:
        return None
    if entry["expires_at"] - 60 > time.time():
        return entry["access_token"]

    # Zugriffstoken abgelaufen - mit dem Refresh-Token ein neues holen
    if not entry["refresh_token"]:
        return None
    try:
        token = await oauth.microsoft.fetch_access_token(grant_type="refresh_token", refresh_token=entry["refresh_token"])
    except Exception as e:
        logging.getLogger("uvicorn.error").warning("SharePoint-Token konnte nicht erneuert werden: %s", e)
        USER_TOKENS.pop(user.get("sid"), None)
        return None
    entry["access_token"] = token.get("access_token")
    entry["refresh_token"] = token.get("refresh_token") or entry["refresh_token"]
    entry["expires_at"] = token.get("expires_at") or time.time() + int(token.get("expires_in") or 0)
    return entry["access_token"]

# Dependency zum Prüfen, ob User eingeloggt ist
def get_current_user(request: Request):
    user = request.session.get('user')
    if not user:
        return None
    # Sitzungen aus älteren Versionen fehlen Berechtigung bzw. SharePoint-Token - neu anmelden lassen
    if "is_sync_admin" not in user or "sid" not in user:
        return None
    # Sitzungen von Benutzern, die (inzwischen) nicht mehr berechtigt sind, gelten nicht
    if not can_edit(user):
        return None
    return user

def access_is_restricted():
    return bool(SYNC_ADMIN_GROUP_ID or SYNC_ADMINS)

def is_sync_admin(claims):
    """Prüft Gruppe und Namensliste anhand der Claims aus dem ID-Token"""
    login_names = {str(claims.get(key) or "").lower() for key in ("preferred_username", "email", "upn")}
    if any(admin in login_names for admin in SYNC_ADMINS):
        return True
    groups = {str(group).lower() for group in claims.get("groups") or []}
    return bool(SYNC_ADMIN_GROUP_ID) and SYNC_ADMIN_GROUP_ID in groups

def can_edit(user):
    """Prüft, ob der Benutzer die Web-App benutzen (und damit Sync-Projekte ändern) darf"""
    if not access_is_restricted():
        return True
    # Die Gruppenzugehörigkeit steht nur beim Login zur Verfügung und liegt als Flag in der Session
    return bool(user.get("is_sync_admin")) or is_sync_admin(user)

# --- Profilbild des angemeldeten Benutzers (Kopfzeile) ---
GRAPH_PHOTO_URL = "https://graph.microsoft.com/v1.0/me/photos/48x48/$value"

async def fetch_profile_photo(refresh_token):
    """Holt das Profilbild über Microsoft Graph (delegierte Berechtigung User.Read).

    Liefert (neues Refresh-Token oder None, (Content-Type, Bilddaten) oder None).
    """
    # Das Login-Token gilt nur für SharePoint; für Graph mit dem Refresh-Token ein eigenes Token holen
    token = await oauth.microsoft.fetch_access_token(
        grant_type="refresh_token",
        refresh_token=refresh_token,
        scope="https://graph.microsoft.com/User.Read",
    )
    async with httpx.AsyncClient(timeout=15) as client:
        response = await client.get(GRAPH_PHOTO_URL, headers={"Authorization": f"Bearer {token['access_token']}"})
    photo = None
    if response.status_code == 200:
        photo = (response.headers.get("Content-Type", "image/jpeg"), response.content)
    elif response.status_code != 404:  # 404 = Benutzer hat kein Profilbild
        logging.getLogger("uvicorn.error").warning("Profilbild-Abruf: HTTP %s %s", response.status_code, response.text[:200])
    return token.get("refresh_token"), photo

# --- Favicon (ohne Anmeldung abrufbar, Browser fragen es auch auf der Login-/Fehlerseite an) ---
STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")

@app.get("/favicon.ico", include_in_schema=False)
async def favicon_ico():
    return FileResponse(os.path.join(STATIC_DIR, "favicon.ico"), media_type="image/x-icon",
                        headers={"Cache-Control": "public, max-age=86400"})

@app.get("/favicon.svg", include_in_schema=False)
async def favicon_svg():
    return FileResponse(os.path.join(STATIC_DIR, "favicon.svg"), media_type="image/svg+xml",
                        headers={"Cache-Control": "public, max-age=86400"})

@app.get("/avatar")
async def avatar(request: Request):
    """Profilbild des angemeldeten Benutzers; 404, wenn es keins gibt (die Seite zeigt dann die Initialen)"""
    # Nicht im Browser zwischenspeichern: Die Adresse ist für alle Benutzer gleich, nach einem
    # Kontowechsel am selben Browser würde sonst das Bild des vorigen Benutzers erscheinen
    no_store = {"Cache-Control": "private, no-store"}
    user = get_current_user(request)
    entry = USER_TOKENS.get(user.get("sid")) if user else None
    if not entry:
        return Response(status_code=404, headers=no_store)

    # Nur einmal je Sitzung abrufen; danach liegt das Bild (oder "keins") im Arbeitsspeicher
    if "photo" not in entry:
        entry["photo"] = None
        if entry["refresh_token"]:
            try:
                new_refresh_token, entry["photo"] = await fetch_profile_photo(entry["refresh_token"])
                entry["refresh_token"] = new_refresh_token or entry["refresh_token"]
            except Exception as e:
                # z.B. fehlende Zustimmung für User.Read - die Seite funktioniert auch ohne Bild
                logging.getLogger("uvicorn.error").warning("Profilbild konnte nicht geladen werden: %s", e)

    if not entry["photo"]:
        return Response(status_code=404, headers=no_store)
    content_type, data = entry["photo"]
    return Response(content=data, media_type=content_type, headers=no_store)

# ==========================================
# 3. KONFIGURATION (SYNC-PROJEKTE)
# ==========================================
def save_config(config_data):
    tmp_file = CONFIG_FILE + ".tmp"
    with open(tmp_file, 'w', encoding='utf-8') as f:
        json.dump(config_data, f, ensure_ascii=False, indent=2)
    os.replace(tmp_file, CONFIG_FILE)

def load_config():
    # Frische Installation: noch keine Sync-Projekte angelegt
    if not os.path.exists(CONFIG_FILE):
        return []
    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
        config_data = json.load(f)

    # Ältere Einträge haben noch keine ID - einmalig nachtragen
    missing_ids = [item for item in config_data if not item.get("id")]
    if missing_ids:
        for item in missing_ids:
            item["id"] = uuid.uuid4().hex[:12]
        save_config(config_data)
    return config_data

def find_project(config_data, project_id):
    for item in config_data:
        if item.get("id") == project_id:
            return item
    return None

async def read_form(request: Request):
    """Liest ein klassisches HTML-Formular (ohne zusätzliche Abhängigkeit python-multipart)"""
    body = (await request.body()).decode('utf-8')
    return {key: values[0].strip() for key, values in parse_qs(body, keep_blank_values=True).items()}

def normalize_folder(value):
    """Akzeptiert auch eine komplette URL und macht daraus den server-relativen Pfad"""
    if value.lower().startswith("https://"):
        value = unquote(urlparse(value).path)
    return value.rstrip("/")

def project_from_form(form):
    return {
        "name": form.get("name", ""),
        "source_site": form.get("source_site", "").rstrip("/"),
        "source_folder": normalize_folder(form.get("source_folder", "")),
        "target_site": form.get("target_site", "").rstrip("/"),
        "target_folder": normalize_folder(form.get("target_folder", "")),
        "enabled": form.get("enabled") == "on",
    }

def validate_project(project, config_data, project_id=None):
    errors = []

    if not project["name"]:
        errors.append("Bitte einen Namen angeben.")
    elif any(item["name"].lower() == project["name"].lower() and item.get("id") != project_id for item in config_data):
        errors.append("Es gibt bereits ein Sync-Projekt mit diesem Namen.")

    for side, label in (("source", "Quelle"), ("target", "Ziel")):
        site = project[f"{side}_site"]
        folder = project[f"{side}_folder"]
        if not SITE_PATTERN.fullmatch(site):
            errors.append(f"{label}: Die Site-URL muss die Form {TENANT_URL}/sites/<Name> haben.")
            continue
        site_path = urlparse(site).path
        if not folder.lower().startswith(site_path.lower() + "/"):
            errors.append(f"{label}: Der Ordner muss mit {site_path}/ beginnen (server-relativer Pfad).")

    source = project["source_folder"].lower()
    target = project["target_folder"].lower()
    if source and target:
        if source == target:
            errors.append("Quelle und Ziel dürfen nicht derselbe Ordner sein.")
        elif target.startswith(source + "/") or source.startswith(target + "/"):
            errors.append("Quelle und Ziel dürfen nicht ineinander liegen.")
    return errors

# ==========================================
# 4. LAUF-HISTORIE (wird von sync.py geschrieben)
# ==========================================
def read_json_file(path):
    # sync.py ersetzt die Datei während eines Laufs - kurz erneut versuchen
    for attempt in range(3):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (OSError, ValueError):
            time.sleep(0.1)
    return None

def prepare_run(run):
    """Ergänzt Summen, Dauer und erkennt hängengebliebene Läufe"""
    summary_path = os.path.join(RUNS_DIR, f"{run['id']}.json")
    events_path = os.path.join(RUNS_DIR, f"{run['id']}.events.jsonl")

    if run["status"] == "running":
        last_sign = max(os.path.getmtime(p) for p in (summary_path, events_path) if os.path.exists(p))
        if time.time() - last_sign > STALE_RUN_SECONDS:
            run["status"] = "aborted"
            for pair in run["pairs"]:
                if pair["status"] == "running":
                    pair["status"] = "aborted"

    for key in ("copied", "deleted", "errors"):
        run[key] = sum(pair.get(key, 0) for pair in run["pairs"])
    run["failed_pairs"] = sum(1 for pair in run["pairs"] if pair["status"] == "failed")

    for item in [run] + run["pairs"]:
        item["seconds"] = None
        try:
            start = datetime.fromisoformat(item["started"])
            if item.get("finished"):
                item["seconds"] = (datetime.fromisoformat(item["finished"]) - start).total_seconds()
            elif item["status"] == "running":
                item["seconds"] = (datetime.now() - start).total_seconds()
        except (KeyError, TypeError, ValueError):
            pass
    return run

def load_runs(limit=50):
    if not os.path.isdir(RUNS_DIR):
        return []
    names = sorted((n for n in os.listdir(RUNS_DIR) if re.fullmatch(r"\d{8}-\d{6}\.json", n)), reverse=True)
    runs = []
    for name in names:
        if len(runs) >= limit:
            break
        run = read_json_file(os.path.join(RUNS_DIR, name))
        # Läufe ohne Sync-Projekt (ältere Versionen haben sie auch ohne Projekte angelegt) nicht anzeigen
        if run and (run["pairs"] or run["status"] == "running"):
            runs.append(prepare_run(run))
    return runs

def load_run(run_id):
    # Die ID landet in einem Dateipfad - nur das erwartete Format zulassen
    if not re.fullmatch(r"\d{8}-\d{6}", run_id):
        return None
    path = os.path.join(RUNS_DIR, f"{run_id}.json")
    if not os.path.exists(path):
        return None
    run = read_json_file(path)
    return prepare_run(run) if run else None

def load_events(run_id, pair_count):
    """Liefert je Sync-Paar die Ereignisse; Fehler immer, alles andere bis MAX_EVENTS_SHOWN"""
    events = [{"entries": [], "hidden": 0} for _ in range(pair_count)]
    path = os.path.join(RUNS_DIR, f"{run_id}.events.jsonl")
    if not os.path.exists(path):
        return events

    with open(path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                entry = json.loads(line)
                bucket = events[entry["pair"]]
            except (ValueError, KeyError, IndexError, TypeError):
                continue
            if entry.get("action") == "error" or len(bucket["entries"]) < MAX_EVENTS_SHOWN:
                bucket["entries"].append(entry)
            else:
                bucket["hidden"] += 1
    return events

def last_result_per_project(runs):
    """Ordnet jedem Projekt (ID oder Name) das Ergebnis seines letzten Laufs zu"""
    results = {}
    for run in runs:  # neueste zuerst
        for pair in run["pairs"]:
            for key in (pair.get("id"), pair.get("name")):
                if key and key not in results:
                    results[key] = {"run_id": run["id"], "started": run["started"], "pair": pair}
    return results

# ==========================================
# 5. WEB-FRONTEND
# ==========================================
def render(request: Request, name: str, user, **context):
    context.update({
        "user": user,
        "can_edit": can_edit(user),
        "access_restricted": access_is_restricted(),
        "flash": request.session.pop('flash', None),
    })
    return templates.TemplateResponse(request=request, name=name, context=context)

@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login")

    config_data = load_config()
    runs = load_runs()
    last_results = last_result_per_project(runs)
    for item in config_data:
        item["last"] = last_results.get(item["id"]) or last_results.get(item["name"])

    # Letzte 50 Zeilen vom Log lesen
    log_lines = []
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, 'r', encoding='utf-8') as lf:
            log_lines = lf.readlines()[-50:]

    return render(
        request, "dashboard.html", user,
        config=config_data,
        runs=runs,
        logs=log_lines,
        is_running=any(run["status"] == "running" for run in runs),
    )

@app.get("/projects/new", response_class=HTMLResponse)
async def new_project(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login")
    if not can_edit(user):
        return RedirectResponse(url="/")

    project = {"name": "", "source_site": "", "source_folder": "", "target_site": "", "target_folder": "", "enabled": True}
    return render(request, "project_form.html", user, project=project, errors=[], is_new=True)

@app.post("/projects/new", response_class=HTMLResponse)
async def create_project(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    if not can_edit(user):
        return RedirectResponse(url="/", status_code=303)

    config_data = load_config()
    project = project_from_form(await read_form(request))
    errors = validate_project(project, config_data)
    if errors:
        return render(request, "project_form.html", user, project=project, errors=errors, is_new=True)

    project = {"id": uuid.uuid4().hex[:12], **project}
    config_data.append(project)
    save_config(config_data)
    request.session['flash'] = f"Sync-Projekt „{project['name']}“ wurde angelegt."
    return RedirectResponse(url="/", status_code=303)

@app.get("/projects/{project_id}/edit", response_class=HTMLResponse)
async def edit_project(request: Request, project_id: str):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login")
    project = find_project(load_config(), project_id)
    if not project or not can_edit(user):
        return RedirectResponse(url="/")

    project.setdefault("enabled", True)
    return render(request, "project_form.html", user, project=project, errors=[], is_new=False)

@app.post("/projects/{project_id}/edit", response_class=HTMLResponse)
async def update_project(request: Request, project_id: str):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    config_data = load_config()
    existing = find_project(config_data, project_id)
    if not existing or not can_edit(user):
        return RedirectResponse(url="/", status_code=303)

    project = project_from_form(await read_form(request))
    errors = validate_project(project, config_data, project_id)
    if errors:
        project["id"] = project_id
        return render(request, "project_form.html", user, project=project, errors=errors, is_new=False)

    existing.update(project)
    save_config(config_data)
    request.session['flash'] = f"Sync-Projekt „{project['name']}“ wurde gespeichert."
    return RedirectResponse(url="/", status_code=303)

@app.post("/projects/{project_id}/toggle")
async def toggle_project(request: Request, project_id: str):
    """Pausiert ein Sync-Projekt bzw. nimmt es wieder auf, ohne es zu löschen"""
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    config_data = load_config()
    project = find_project(config_data, project_id)
    if project and can_edit(user):
        project["enabled"] = not project.get("enabled", True)
        save_config(config_data)
        state = "wird beim nächsten Lauf wieder synchronisiert" if project["enabled"] else "ist pausiert und wird nicht mehr synchronisiert"
        request.session['flash'] = f"Sync-Projekt „{project['name']}“ {state}."
    return RedirectResponse(url="/", status_code=303)

@app.post("/projects/{project_id}/delete")
async def delete_project(request: Request, project_id: str):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    config_data = load_config()
    project = find_project(config_data, project_id)
    if project and can_edit(user):
        config_data.remove(project)
        save_config(config_data)
        request.session['flash'] = f"Sync-Projekt „{project['name']}“ wurde entfernt. Bereits kopierte Dateien bleiben im Ziel erhalten."
    return RedirectResponse(url="/", status_code=303)

@app.get("/runs/{run_id}", response_class=HTMLResponse)
async def run_detail(request: Request, run_id: str):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login")
    run = load_run(run_id)
    if not run:
        request.session['flash'] = "Dieser Lauf wurde nicht gefunden."
        return RedirectResponse(url="/")

    events = load_events(run_id, len(run["pairs"]))
    return render(request, "run_detail.html", user, run=run, events=events)

# ==========================================
# 6. SITE- UND ORDNER-AUSWAHL (für das Projekt-Formular)
# ==========================================
# Bibliotheken, die es in jeder Site gibt und die als Sync-Ordner nicht in Frage kommen
SYSTEM_LIBRARIES = {"style library", "formservertemplates", "siteassets", "sitepages"}

def api_error(message, status_code=400, relogin=False):
    return JSONResponse({"error": message, "relogin": relogin}, status_code=status_code)

async def get_browse_token(request: Request):
    """Liefert (Token, None) oder (None, Fehler-Antwort) für die Auswahl-Abfragen"""
    user = get_current_user(request)
    if not user:
        return None, api_error("Sie sind nicht mehr angemeldet.", 401, relogin=True)
    if not can_edit(user):
        return None, api_error("Sie dürfen keine Sync-Projekte bearbeiten.", 403)
    token = await get_sharepoint_token(user)
    if not token:
        return None, api_error("Für die Auswahl wird eine neue SharePoint-Anmeldung benötigt.", 401, relogin=True)
    return token, None

async def sharepoint_get(token, url, params):
    """GET im Namen des Benutzers; bei Drosselung (429) oder Aussetzern (5xx) bis zu zweimal wiederholen.

    Die SharePoint-Suche antwortet gelegentlich mit 500 "Unbekannter Fehler" und klappt beim nächsten Versuch.
    """
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json;odata=nometadata"}
    async with httpx.AsyncClient(timeout=30) as client:
        for attempt in range(3):
            response = await client.get(url, params=params, headers=headers)
            if response.status_code != 429 and response.status_code < 500:
                return response
            if attempt < 2:
                retry_after = response.headers.get("Retry-After", "")
                await asyncio.sleep(min(int(retry_after), 5) if retry_after.isdigit() else 1 + attempt)
        return response

def sharepoint_error(response):
    logging.getLogger("uvicorn.error").warning("SharePoint-Abfrage %s: %s %s", response.url, response.status_code, response.text[:300])
    if response.status_code == 401:
        return api_error("SharePoint hat die Anmeldung abgelehnt.", 401, relogin=True)
    if response.status_code == 403:
        return api_error("Kein Zugriff. Sie haben dort keine Berechtigung, oder der App fehlt die delegierte SharePoint-Berechtigung.", 403)
    return api_error(f"SharePoint-Abfrage fehlgeschlagen ({response.status_code}).", 502)

async def search_sites(token, querytext):
    """Sites aus der SharePoint-Suche, auf die der Benutzer Zugriff hat. Liefert (Liste, None) oder (None, Fehler-Antwort)"""
    sites = {}
    page_size = 500
    for start_row in range(0, 5000, page_size):
        response = await sharepoint_get(token, f"{TENANT_URL}/_api/search/query", {
            "querytext": f"'{querytext}'",
            "selectproperties": "'Title,Path,WebTemplate'",
            "rowlimit": page_size,
            "startrow": start_row,
            "trimduplicates": "false",
        })
        if response.status_code != 200:
            return None, sharepoint_error(response)
        rows = response.json()["PrimaryQueryResult"]["RelevantResults"]["Table"]["Rows"]
        for row in rows:
            cells = {cell["Key"]: cell["Value"] for cell in row["Cells"]}
            url = (cells.get("Path") or "").rstrip("/")
            # Die Suche liefert auch OneDrive- und interne Sites - nur sync-fähige Sites anbieten
            if SITE_PATTERN.fullmatch(url):
                sites[url.lower()] = {
                    "title": cells.get("Title") or url.rsplit("/", 1)[1],
                    "url": url,
                    # Private und freigegebene Teams-Kanäle sind eigene Sites mit dieser Vorlage
                    "channel": (cells.get("WebTemplate") or "").upper() == "TEAMCHANNEL",
                }
        if len(rows) < page_size:
            break
    return sorted(sites.values(), key=lambda site: site["title"].lower()), None

# App-Zugang (Zertifikat) nur zum Auffinden von Sites. Ob der Benutzer Zugriff hat, prüft immer sein eigenes Token.
APP_MSAL = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    client_credential={"private_key_pfx_path": CERT_PATH},
)

async def get_app_token():
    # msal hält das Token zwischengespeichert; der Abruf blockiert, daher in einem Thread
    result = await asyncio.to_thread(APP_MSAL.acquire_token_for_client, scopes=[f"{TENANT_URL}/.default"])
    if "access_token" not in result:
        logging.getLogger("uvicorn.error").warning("App-Token nicht erhalten: %s", result.get("error_description"))
    return result.get("access_token")

async def find_accessible_sites(user_token, querytext):
    """Sites, auf die der Benutzer tatsächlich Zugriff hat. Liefert (Liste, None) oder (None, Fehler-Antwort).

    Die Suche im Namen des Benutzers verlässt sich auf den Suchindex, der neue Berechtigungen
    teils erst nach Stunden oder Tagen kennt - dann fehlen Sites, die der Benutzer längst öffnen kann.
    Deshalb findet die App die Kandidaten mit ihrem eigenen Zugang, und jede Site wird live mit
    dem Benutzer-Token geprüft. Dabei fällt die Gruppen-ID für die Kanal-Zuordnung mit ab.
    """
    app_token = await get_app_token()
    candidates, error = await search_sites(app_token or user_token, querytext)
    if error:
        return None, error

    limit = asyncio.Semaphore(20)  # SharePoint nicht mit allen Anfragen auf einmal fluten

    async def check(site):
        async with limit:
            try:
                response = await sharepoint_get(user_token, f"{site['url']}/_api/site", {"$select": "RelatedGroupId"})
            except httpx.HTTPError:
                return None
        if response.status_code != 200:
            return None  # kein Zugriff
        return {**site, "related_group": (response.json().get("RelatedGroupId") or "").lower()}

    results = await asyncio.gather(*(check(site) for site in candidates))
    return [site for site in results if site], None

async def get_team_channels(token, site):
    """Private/freigegebene Kanäle des Teams hinter site, auf die der Benutzer Zugriff hat.

    Liefert (Kanäle, Fehler ja/nein). Jede Kanal-Site trägt die Gruppen-ID ihres Teams als
    RelatedGroupId; die URL taugt nicht zur Zuordnung, weil sie nach einer Umbenennung des
    Teams nicht mitwandert.
    """
    response = await sharepoint_get(token, f"{site}/_api/site", {"$select": "GroupId"})
    if response.status_code != 200:
        return [], True
    group_id = (response.json().get("GroupId") or "").lower()
    if not group_id or group_id == "00000000-0000-0000-0000-000000000000":
        return [], False  # keine Teams-Site (oder selbst ein Kanal)

    channel_sites, error = await find_accessible_sites(token, "contentclass:STS_Site WebTemplate:TEAMCHANNEL")
    if error:
        return [], True

    team_title = ""
    title_response = await sharepoint_get(token, f"{site}/_api/web", {"$select": "Title"})
    if title_response.status_code == 200:
        team_title = title_response.json().get("Title") or ""

    channels = []
    for channel in channel_sites:
        if channel["related_group"] == group_id:
            # Der Site-Titel lautet "<Team>-<Kanal>" - für die Anzeige nur den Kanalnamen
            name = channel["title"]
            if team_title and name.lower().startswith(team_title.lower() + "-"):
                name = name[len(team_title) + 1:]
            channels.append({"name": name, "url": channel["url"]})
    return sorted(channels, key=lambda channel: channel["name"].lower()), False

@app.get("/api/sites")
async def api_sites(request: Request):
    """Alle Sites, auf die der angemeldete Benutzer Zugriff hat (über die SharePoint-Suche)"""
    token, error = await get_browse_token(request)
    if error:
        return error
    sites, error = await find_accessible_sites(token, "contentclass:STS_Site")
    if error:
        return error
    return {"sites": [{"title": site["title"], "url": site["url"], "channel": site["channel"]} for site in sites]}

@app.get("/api/folders")
async def api_folders(request: Request, site: str = "", path: str = ""):
    """Unterordner von path; ohne path die Dokumentbibliotheken der Site"""
    token, error = await get_browse_token(request)
    if error:
        return error

    # Das Benutzer-Token darf nur an Sites des eigenen Tenants gehen
    site = site.strip().rstrip("/")
    if not SITE_PATTERN.fullmatch(site):
        return api_error(f"Bitte zuerst eine gültige Site wählen ({TENANT_URL}/sites/<Name>).")
    site_path = urlparse(site).path
    path = path.rstrip("/")

    if not path:
        response = await sharepoint_get(token, f"{site}/_api/web/lists", {
            "$filter": "BaseTemplate eq 101 and Hidden eq false",
            "$select": "Title,RootFolder/ServerRelativeUrl",
            "$expand": "RootFolder",
        })
        if response.status_code != 200:
            return sharepoint_error(response)
        folders = []
        for library in response.json()["value"]:
            library_path = library["RootFolder"]["ServerRelativeUrl"]
            name = library_path.rsplit("/", 1)[1]
            if name.lower() not in SYSTEM_LIBRARIES:
                folders.append({"name": name, "title": library["Title"], "path": library_path})
        # Wie die Ansicht "In Kanälen": Kanäle liegen nicht als Ordner im Team, sondern in eigenen Sites
        channels, channels_failed = await get_team_channels(token, site)
        return {
            "path": "",
            "parent": None,
            "folders": sorted(folders, key=lambda folder: folder["name"].lower()),
            "channels": channels,
            "channels_failed": channels_failed,
        }

    if not path.lower().startswith(site_path.lower() + "/"):
        return api_error("Der Ordner gehört nicht zur gewählten Site.")
    response = await sharepoint_get(token, f"{site}/_api/web/GetFolderByServerRelativePath(decodedurl=@p)/Folders", {
        "@p": "'" + path.replace("'", "''") + "'",
        "$select": "Name,ServerRelativeUrl",
        "$orderby": "Name",
        "$top": 5000,
    })
    if response.status_code != 200:
        return sharepoint_error(response)
    folders = [
        {"name": folder["Name"], "path": folder["ServerRelativeUrl"]}
        for folder in response.json()["value"]
        if folder["Name"] != "Forms"
    ]
    parent = path.rsplit("/", 1)[0]
    # Oberhalb einer Bibliothek geht es zurück zur Liste der Bibliotheken
    if parent.lower() == site_path.lower():
        parent = ""
    return {"path": path, "parent": parent, "folders": folders}

if __name__ == "__main__":
    import uvicorn
    # Startet den Server auf Port 8000
    uvicorn.run(app, host="0.0.0.0", port=8000)
