import json
import os
import logging
import time
import msal
import requests
from datetime import datetime
from urllib.parse import quote, unquote, urlparse
from logging.handlers import RotatingFileHandler
from dotenv import load_dotenv

from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.utilities.move_copy_util import MoveCopyUtil
from office365.sharepoint.utilities.move_copy_options import MoveCopyOptions

# ==========================================
# 1. ENTERPRISE LOGGING SETUP
# ==========================================
# Rotiert Log-Dateien: Max 5 MB pro Datei, behält die letzten 5 Dateien
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
log_file = 'sharepoint_sync.log'

file_handler = RotatingFileHandler(log_file, mode='a', maxBytes=5*1024*1024, backupCount=5, encoding='utf-8')
file_handler.setFormatter(log_formatter)

console_handler = logging.StreamHandler()
console_handler.setFormatter(log_formatter)

logger = logging.getLogger('SPSync')
logger.setLevel(logging.INFO)
logger.addHandler(file_handler)
logger.addHandler(console_handler)

# ==========================================
# 2. KONFIGURATION & SECRETS
# ==========================================
# Lädt die Variablen aus der .env Datei
load_dotenv()
TENANT_ID = os.getenv("SP_TENANT_ID")
CLIENT_ID = os.getenv("SP_CLIENT_ID")
# SharePoint akzeptiert App-Only-Tokens nur mit Zertifikat (kein Client-Secret!)
CERT_PATH = os.getenv("SP_CERT_PATH", "sp_sync_cert.pfx")

# SharePoint-Adresse des Tenants, z.B. https://contoso.sharepoint.com
TENANT_URL = os.getenv("SP_TENANT_URL", "").strip().rstrip("/")

if not TENANT_ID or not CLIENT_ID or not TENANT_URL:
    logger.critical("SP_TENANT_ID, SP_CLIENT_ID oder SP_TENANT_URL fehlen in der .env Datei! Abbruch.")
    exit(1)

if not os.path.exists(CERT_PATH):
    logger.critical(f"Zertifikat {CERT_PATH} nicht gefunden! Abbruch.")
    exit(1)

# Frische Installation: Die Datei entsteht erst, wenn im Web-Dashboard ein Sync-Projekt angelegt wird
SYNC_PAIRS = []
if os.path.exists('config.json'):
    with open('config.json', 'r', encoding='utf-8') as f:
        SYNC_PAIRS = json.load(f)

# Lauf-Historie für das Web-Dashboard (wird von web_app.py gelesen)
RUNS_DIR = 'runs'

# Dateitypen, die SharePoint beim Kopieren selbst verändert (es schreibt Metadaten in die Datei).
# Ihr Hash weicht im Ziel immer ab, ohne dass sich der Inhalt geändert hat - für sie zählt nur das Änderungsdatum.
# Würde hier ein betroffener Typ fehlen, würden dessen Dateien bei jedem Lauf erneut kopiert.
REWRITTEN_ON_COPY = {
    ".docx", ".docm", ".dotx", ".doc", ".dot",
    ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xls",
    ".pptx", ".pptm", ".ppsx", ".potx", ".ppt",
    ".vsdx", ".msg"
}

# Maximale Wartezeit auf einen SharePoint-Kopierauftrag (große Dateien), in Sekunden
COPY_JOB_TIMEOUT = 30 * 60

# ==========================================
# 3. LAUF-HISTORIE
# ==========================================
class RunRecorder:
    """Protokolliert einen Sync-Lauf für das Web-Dashboard.

    Pro Lauf entstehen zwei Dateien in RUNS_DIR:
    <id>.json          Zusammenfassung (Status, Zähler je Sync-Paar)
    <id>.events.jsonl  ein Ereignis pro Zeile (kopiert / gelöscht / Ordner / Fehler)
    """
    def __init__(self):
        os.makedirs(RUNS_DIR, exist_ok=True)
        self.id = datetime.now().strftime("%Y%m%d-%H%M%S")
        self.summary_path = os.path.join(RUNS_DIR, f"{self.id}.json")
        self.events_path = os.path.join(RUNS_DIR, f"{self.id}.events.jsonl")
        self.summary = {
            "id": self.id,
            "started": self._now(),
            "finished": None,
            "status": "running",
            "pairs": []
        }
        self.save()

    @staticmethod
    def _now():
        return datetime.now().isoformat(timespec='seconds')

    def start_pair(self, pair):
        """Legt den Eintrag für ein Sync-Paar an; das Dict dient zugleich als Zähler"""
        stats = {
            "id": pair.get("id"),
            "name": pair["name"],
            "source_folder": pair["source_folder"],
            "target_folder": pair["target_folder"],
            "status": "running",
            "started": self._now(),
            "finished": None,
            "source_count": None,
            "target_count": None,
            "copied": 0,
            "deleted": 0,
            "folders": 0,
            "errors": 0,
            "message": ""
        }
        self.summary["pairs"].append(stats)
        self.save()
        return stats

    def finish_pair(self, stats, status, message=""):
        stats["status"] = status
        stats["message"] = message
        stats["finished"] = self._now()
        self.save()

    def event(self, action, path, message=""):
        entry = {
            "time": self._now(),
            "pair": len(self.summary["pairs"]) - 1,
            "action": action,
            "path": path,
            "message": message
        }
        with open(self.events_path, 'a', encoding='utf-8') as ef:
            ef.write(json.dumps(entry, ensure_ascii=False) + "\n")
        # Zähler regelmäßig sichern, damit das Dashboard den Fortschritt live zeigt
        if time.time() - self.last_save > 5:
            self.save()

    def finish(self):
        # Paare, die noch "running" sind, wurden unterbrochen (z.B. Strg+C)
        for p in self.summary["pairs"]:
            if p["status"] == "running":
                p["status"] = "aborted"
                p["finished"] = self._now()

        statuses = [p["status"] for p in self.summary["pairs"]]
        if "failed" in statuses:
            self.summary["status"] = "failed"
        elif "aborted" in statuses:
            self.summary["status"] = "aborted"
        elif "errors" in statuses:
            self.summary["status"] = "errors"
        else:
            self.summary["status"] = "ok"
        self.summary["finished"] = self._now()
        self.save()

    def save(self):
        self.last_save = time.time()
        tmp_path = self.summary_path + ".tmp"
        with open(tmp_path, 'w', encoding='utf-8') as sf:
            json.dump(self.summary, sf, ensure_ascii=False, indent=2)
        # Unter Windows schlägt das Ersetzen fehl, solange das Dashboard die Datei gerade liest
        for attempt in range(5):
            try:
                os.replace(tmp_path, self.summary_path)
                return
            except PermissionError:
                time.sleep(0.2)
        logger.warning(f"Lauf-Zusammenfassung {self.summary_path} konnte nicht aktualisiert werden.")

# ==========================================
# 4. HILFSFUNKTIONEN
# ==========================================
def sp_get_json(url, headers):
    """GET gegen SharePoint mit Wiederholung, falls SharePoint drosselt (429/503)"""
    for attempt in range(5):
        response = requests.get(url, headers=headers, timeout=60)
        if response.status_code in (429, 503):
            time.sleep(int(response.headers.get("Retry-After", 5)))
            continue
        if response.status_code != 200:
            raise Exception(f"Abruf fehlgeschlagen ({response.status_code}) für {url}: {response.text[:300]}")
        return response.json()
    raise Exception(f"SharePoint drosselt dauerhaft, Abruf abgebrochen: {url}")

def get_files_recursively(site_url, folder_url, file_dict):
    """Holt alle Dateien unter folder_url (inkl. Unterordner).

    Liest über die Drive-Schnittstelle (_api/v2.0), weil nur sie den Inhalts-Hash
    (quickXorHash) jeder Datei mitliefert.

    Fehler werden bewusst NICHT abgefangen: Eine unvollständige Liste würde
    dazu führen, dass der Sync am Ziel Dateien löscht, die es in der Quelle noch gibt.
    """
    headers = {
        "Authorization": f"Bearer {get_sp_token()['access_token']}",
        "Accept": "application/json"
    }

    # Dokumentbibliothek (Drive) finden, in der der Ordner liegt
    drive_id = None
    drive_path = ""
    for drive in sp_get_json(f"{site_url}/_api/v2.0/drives?$select=id,webUrl", headers)["value"]:
        candidate_path = unquote(urlparse(drive["webUrl"]).path)
        if (folder_url + "/").lower().startswith(candidate_path.lower() + "/") and len(candidate_path) > len(drive_path):
            drive_id = drive["id"]
            drive_path = candidate_path
    if not drive_id:
        raise Exception(f"Keine Dokumentbibliothek für {folder_url} in {site_url} gefunden")

    drive_url = f"{site_url}/_api/v2.0/drives/{drive_id}"
    query = "?$top=500&$select=id,name,size,file,folder,lastModifiedDateTime"
    path_in_drive = folder_url[len(drive_path):]
    if path_in_drive:
        start_url = f"{drive_url}/root:{quote(path_in_drive)}:/children{query}"
    else:
        start_url = f"{drive_url}/root/children{query}"

    # Noch einzulesende Ordner: (Abruf-URL, Pfad relativ zu folder_url)
    pending_folders = [(start_url, "")]
    while pending_folders:
        url, relative_dir = pending_folders.pop()
        while url:
            page = sp_get_json(url, headers)
            for item in page["value"]:
                rel_path = f"{relative_dir}/{item['name']}"
                if "folder" in item:
                    if item["name"] not in ["Forms", "hidden"]:
                        pending_folders.append((f"{drive_url}/items/{item['id']}/children{query}", rel_path))
                elif "file" in item:
                    file_dict[rel_path] = {
                        "url": f"{folder_url}{rel_path}",
                        "modified": datetime.fromisoformat(item["lastModifiedDateTime"].replace("Z", "+00:00")),
                        "size": item.get("size", 0),
                        "hash": item["file"].get("hashes", {}).get("quickXorHash")
                    }
            # Große Ordner liefert SharePoint seitenweise
            url = page.get("@odata.nextLink")

def ensure_target_folder(ctx, target_base, relative_dir, known_folders, recorder, stats):
    """Erstellt fehlende Unterordner dynamisch"""
    if not relative_dir or relative_dir == "/":
        return

    current_path = target_base
    parts = [p for p in relative_dir.split('/') if p]

    for part in parts:
        current_path = f"{current_path}/{part}"
        # Bereits geprüfte Ordner nicht für jede Datei erneut abfragen
        if current_path in known_folders:
            continue
        # Die ..._path-Aufrufe kommen (anders als ..._url) mit % und # im Namen zurecht
        try:
            folder = ctx.web.get_folder_by_server_relative_path(current_path)
            ctx.load(folder, ["Exists"])
            ctx.execute_query()
            exists = folder.properties.get("Exists", True)
        except Exception:
            exists = False

        if exists:
            known_folders.add(current_path)
        else:
            try:
                parent_url = current_path.rsplit('/', 1)[0]
                parent_folder = ctx.web.get_folder_by_server_relative_path(parent_url)
                parent_folder.folders.add_using_path(part, False)
                ctx.execute_query()
                known_folders.add(current_path)
                logger.info(f"Neuer Ordner am Ziel erstellt: {part}")
                stats["folders"] += 1
                recorder.event("folder", current_path[len(target_base):])
            except Exception as e:
                logger.error(f"Konnte Ordner {part} nicht erstellen: {str(e)}")
                recorder.event("error", current_path[len(target_base):], f"Ordner konnte nicht erstellt werden: {str(e)}")

def get_copy_reason(rel_path, s_info, t_info):
    """Entscheidet, ob eine Quelldatei kopiert werden muss.

    Liefert den Grund als Text oder None, wenn die Zieldatei aktuell ist.
    """
    if t_info is None:
        return "neu"

    extension = os.path.splitext(rel_path)[1].lower()
    if extension not in REWRITTEN_ON_COPY and s_info["hash"] and t_info["hash"]:
        # Der Inhalts-Hash ist maßgeblich - gleicher Inhalt heißt nichts zu tun, egal welches Datum
        if s_info["hash"] == t_info["hash"]:
            return None
        return "aktualisiert (Inhalt geändert)"

    # Ohne vergleichbaren Hash zählt das Änderungsdatum
    if s_info["modified"] > t_info["modified"]:
        return "aktualisiert"
    return None

def get_sp_token():
    """Holt einen modernen OAuth2 Token vom Microsoft Entra ID"""
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential={"private_key_pfx_path": CERT_PATH}
    )
    scopes = [f"{TENANT_URL}/.default"]
    result = app.acquire_token_for_client(scopes=scopes)

    if "access_token" in result:
        return result
    else:
        logger.critical(f"Auth-Fehler von Entra ID: {result.get('error_description')}")
        raise Exception("Konnte kein Token abrufen!")

def copy_file_with_job(source_site, abs_source_url, abs_target_folder_url):
    """Kopiert eine Datei über einen SharePoint-Kopierauftrag (CreateCopyJobs).

    Die direkte Kopie (CopyFileByPath) hat ein Größenlimit. Der Kopierauftrag läuft
    asynchron in SharePoint, hat dieses Limit nicht und ersetzt eine vorhandene Zieldatei.
    Der Aufruf geht direkt per REST, weil die Bibliothek die Anfrage in einem Format
    sendet, das SharePoint ablehnt.
    """
    headers = {
        "Authorization": f"Bearer {get_sp_token()['access_token']}",
        "Accept": "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata"
    }
    payload = {
        "exportObjectUris": [abs_source_url],
        "destinationUri": abs_target_folder_url,
        "options": {
            "IgnoreVersionHistory": True,
            "IsMoveMode": False,
            "AllowSchemaMismatch": True,
            "NameConflictBehavior": 1  # 1 = vorhandene Datei ersetzen
        }
    }
    response = requests.post(f"{source_site}/_api/site/CreateCopyJobs", headers=headers, json=payload, timeout=60)
    if response.status_code != 200:
        raise Exception(f"Kopierauftrag konnte nicht angelegt werden ({response.status_code}): {response.text[:300]}")
    job = response.json()["value"][0]

    deadline = time.time() + COPY_JOB_TIMEOUT
    job_errors = []
    while True:
        response = requests.post(f"{source_site}/_api/site/GetCopyJobProgress", headers=headers, json={"copyJobInfo": job}, timeout=60)
        if response.status_code != 200:
            raise Exception(f"Status des Kopierauftrags nicht abrufbar ({response.status_code}): {response.text[:300]}")
        progress = response.json()

        # Jede Log-Zeile wird nur einmal geliefert - Fehler deshalb über alle Abfragen sammeln
        for line in progress["Logs"]:
            entry = json.loads(line)
            if entry.get("Event") in ("JobError", "JobFatalError"):
                job_errors.append(entry.get("Message") or line)

        if progress["JobState"] == 0:
            break
        if time.time() > deadline:
            raise Exception(f"Kopierauftrag {job['JobId']} war nach {COPY_JOB_TIMEOUT // 60} Minuten nicht fertig")
        time.sleep(2)

    if job_errors:
        raise Exception("Kopierauftrag fehlgeschlagen: " + "; ".join(job_errors))

# ==========================================
# 5. HAUPT-SYNC-LOGIK
# ==========================================
def run_sync():
    active_pairs = [pair for pair in SYNC_PAIRS if pair.get("enabled", True)]
    if not active_pairs:
        # Nichts zu tun: kein Eintrag in Log und Lauf-Historie, nur ein Lebenszeichen in der Konsole (docker logs)
        print(f"{datetime.now():%Y-%m-%d %H:%M:%S} - Keine aktiven Sync-Projekte, nichts zu tun.", flush=True)
        return

    logger.info("=== STARTE SHAREPOINT MIRROR SYNC ===")
    recorder = RunRecorder()

    try:
        for pair in active_pairs:
            logger.info(f"Verarbeite Sync-Paar: {pair['name']}")
            stats = recorder.start_pair(pair)

            try:
                ctx_source = ClientContext(pair["source_site"]).with_access_token(get_sp_token)
                ctx_target = ClientContext(pair["target_site"]).with_access_token(get_sp_token)

                source_files = {}
                target_files = {}

                logger.info("Lese Quell-Ordner ein...")
                get_files_recursively(pair["source_site"], pair["source_folder"], source_files)
                stats["source_count"] = len(source_files)

                logger.info("Lese Ziel-Ordner ein...")
                get_files_recursively(pair["target_site"], pair["target_folder"], target_files)
                stats["target_count"] = len(target_files)
                recorder.save()

                # --- 1. LÖSCHEN ---
                for rel_path, t_info in target_files.items():
                    if rel_path not in source_files:
                        try:
                            logger.info(f"Lösche Datei (nicht mehr in Quelle): {rel_path}")
                            # ..._path statt ..._url: funktioniert auch mit % und # im Namen
                            file_to_delete = ctx_target.web.get_file_by_server_relative_path(t_info["url"])
                            file_to_delete.delete_object()
                            ctx_target.execute_query()
                            stats["deleted"] += 1
                            recorder.event("deleted", rel_path)
                        except Exception as e:
                            logger.error(f"Fehler beim Löschen von {rel_path}: {str(e)}")
                            stats["errors"] += 1
                            recorder.event("error", rel_path, f"Löschen fehlgeschlagen: {str(e)}")

                # --- 2. KOPIEREN / UPDATEN ---
                known_folders = set()
                for rel_path, s_info in source_files.items():
                    copy_note = get_copy_reason(rel_path, s_info, target_files.get(rel_path))

                    if copy_note:
                        try:
                            relative_dir = os.path.dirname(rel_path)
                            ensure_target_folder(ctx_target, pair["target_folder"], relative_dir, known_folders, recorder, stats)

                            # rel_path beginnt immer mit "/" (auch bei Dateien direkt im Hauptordner)
                            target_rel_url = f"{pair['target_folder']}{rel_path}"

                            abs_source_url = f"{TENANT_URL}{s_info['url']}"
                            abs_target_url = f"{TENANT_URL}{target_rel_url}"

                            logger.info(f"Cloud-Kopie gestartet: {rel_path}")

                            # Die direkte Kopie kann nicht in Pfade mit % oder # schreiben - dort gleich den Kopierauftrag nehmen
                            use_copy_job = "%" in rel_path or "#" in rel_path
                            if not use_copy_job:
                                try:
                                    # KeepBoth=False + overwrite: vorhandene Zieldatei ersetzen statt "Datei1.pdf" anzulegen
                                    options = MoveCopyOptions(KeepBoth=False)
                                    MoveCopyUtil.copy_file_by_path(ctx_source, abs_source_url, abs_target_url, True, options=options)
                                    ctx_source.execute_query()
                                except Exception as direct_error:
                                    # z.B. "Die Dateigröße überschreitet den Grenzwert" - große Dateien per Kopierauftrag
                                    logger.info(f"Direkte Kopie nicht möglich ({str(direct_error)}), starte Kopierauftrag: {rel_path}")
                                    use_copy_job = True

                            if use_copy_job:
                                abs_target_folder_url = abs_target_url.rsplit('/', 1)[0]
                                copy_file_with_job(pair["source_site"], abs_source_url, abs_target_folder_url)
                                copy_note += " (Kopierauftrag)"

                            stats["copied"] += 1
                            recorder.event("copied", rel_path, copy_note)

                        except Exception as e:
                            logger.error(f"Fehler beim Kopieren von {rel_path}: {str(e)}")
                            stats["errors"] += 1
                            recorder.event("error", rel_path, f"Kopieren fehlgeschlagen: {str(e)}")

                logger.info(f"Zusammenfassung '{pair['name']}': {stats['copied']} kopiert/aktualisiert, {stats['deleted']} gelöscht, {stats['errors']} Fehler.")
                recorder.finish_pair(stats, "errors" if stats["errors"] else "ok")

            except Exception as e:
                logger.critical(f"Kritischer Fehler im Sync-Paar '{pair['name']}': {str(e)}")
                recorder.event("error", "", f"Kritischer Fehler: {str(e)}")
                recorder.finish_pair(stats, "failed", str(e))
    finally:
        recorder.finish()

    logger.info("=== SHAREPOINT MIRROR SYNC BEENDET ===")

if __name__ == "__main__":
    run_sync()
