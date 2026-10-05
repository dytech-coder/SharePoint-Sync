# SharePoint Sync

Spiegelt Ordner zwischen SharePoint-Online-Sites eines Microsoft-365-Tenants – mit Web-Oberfläche zur Verwaltung der Zuordnungen.

Ein Sync-Projekt verbindet einen **Quellordner** mit einem **Zielordner**. In festen Abständen (Standard: alle 4 Stunden) wird jedes Ziel exakt an seine Quelle angeglichen: neue und geänderte Dateien werden kopiert, im Ziel überzählige Dateien gelöscht. Typischer Einsatz: Datenblätter, Preislisten oder Zertifikate aus einem internen Bereich automatisch in eine Team-Site oder einen privaten Teams-Kanal übertragen.

> **Achtung – Spiegelung:** Dateien, die nur im Zielordner liegen, werden **endgültig** gelöscht (ohne Papierkorb). Als Ziel nur Ordner verwenden, die ausschließlich vom Sync befüllt werden.

Oberfläche und Dokumentation im Code sind deutschsprachig.

## Funktionen

- **Sync-Projekte verwalten:** anlegen, bearbeiten, pausieren, entfernen
- **Site- und Ordner-Auswahl** wie in Power Automate – inklusive privater und freigegebener Teams-Kanäle; jeder Benutzer sieht nur, worauf er selbst Zugriff hat
- **Lauf-Historie** mit Details je Datei (kopiert, gelöscht, Ordner erstellt, Fehler) und Filter
- **Zugriff nur für eine Entra-Gruppe** (Gruppen-Claim im ID-Token), optional zusätzliche Notzugangs-Konten
- **Änderungserkennung per Inhalts-Hash** (`quickXorHash`); Office-Dateien und `.msg` per Änderungsdatum, weil SharePoint sie beim Kopieren verändert
- **Serverseitiges Kopieren** in SharePoint; große Dateien und Pfade mit `%`/`#` über SharePoint-Kopieraufträge
- **Anmeldung ausschließlich per Zertifikat** – kein Client-Secret
- Betrieb als **Docker-Compose-Stack hinter Traefik**

## Aufbau

| Komponente | Aufgabe |
|---|---|
| `web_app.py` | Web-Oberfläche (FastAPI): Login über Microsoft Entra ID, Projektverwaltung, Lauf-Historie, Site-/Ordner-Auswahl |
| `sync.py` | Sync-Dienst: gleicht alle aktiven Projekte ab und schreibt Protokoll und Lauf-Historie |
| `compose.yml` | Dienste `web` (hinter Traefik) und `sync` (Schleife mit Pause) |
| `data/` | Laufzeitdaten beider Dienste: `config.json` (Projekte), `runs/` (Historie), `sharepoint_sync.log` |

Die Web-App und der Sync nutzen dieselbe App-Registrierung auf drei Wegen:

- **App-Zugang** (Zertifikat): Sync liest, kopiert und löscht; die Web-App findet Sites und Kanäle für die Auswahl.
- **Benutzer-Anmeldung** (OpenID Connect): Login und Gruppenprüfung.
- **Benutzer-Zugriff** (delegiert): prüft live, auf welche Sites der Benutzer selbst zugreifen darf; Profilbild aus Microsoft Graph.

## Einrichtung in Microsoft Entra ID

1. **App-Registrierung** anlegen (Kontotyp: nur dieses Verzeichnis).
2. **Redirect URI** (Plattform „Web“): `https://<SYNC_HOSTNAME>/auth`
3. **Zertifikat** erzeugen (siehe unten) und die `.cer` unter *Zertifikate & Geheimnisse* hochladen. Ein Client-Secret wird nicht benötigt – SharePoint lehnt App-Zugriffe mit Client-Secret ab.
4. **API-Berechtigungen** hinzufügen und **Administratorzustimmung erteilen**:

   | API | Berechtigung | Typ |
   |---|---|---|
   | SharePoint | `Sites.ReadWrite.All` | Anwendung |
   | SharePoint | `AllSites.Read` | Delegiert |
   | SharePoint | `Sites.Search.All` | Delegiert |
   | Microsoft Graph | `User.Read` | Delegiert |

5. **Tokenkonfiguration → Gruppenanspruch hinzufügen:** Sicherheitsgruppen, für das ID-Token mit „Gruppen-ID“.
6. **Sicherheitsgruppe** für die berechtigten Benutzer anlegen und ihre Objekt-ID notieren.

> `Sites.ReadWrite.All` als Anwendungsberechtigung erlaubt Schreiben und Löschen in **allen** Sites des Tenants. Das Zertifikat entsprechend schützen und nur vertrauenswürdige Benutzer in die Gruppe aufnehmen – die Web-Oberfläche erlaubt es, beliebige Sites des Tenants als Quelle oder Ziel einzutragen.

### Zertifikat erzeugen

Die Anwendung erwartet eine PFX-Datei **ohne Passwort**:

```bash
openssl req -x509 -newkey rsa:2048 -nodes -days 730 \
  -subj "/CN=SharePoint Sync" -keyout sp_sync_cert.key -out sp_sync_cert.cer
openssl pkcs12 -export -inkey sp_sync_cert.key -in sp_sync_cert.cer -out sp_sync_cert.pfx \
  -keypbe NONE -certpbe NONE -passout pass:
shred -u sp_sync_cert.key
```

`sp_sync_cert.cer` wird in Entra ID hochgeladen, `sp_sync_cert.pfx` bleibt auf dem Server.

## Installation mit Docker Compose und Traefik

Voraussetzungen: Docker mit Compose-Plugin, ein laufender Traefik v3 mit Docker-Provider, einem externen Docker-Netz (Standard `proxy`), einem HTTPS-Entrypoint (Standard `websecure`) und einem Zertifikats-Resolver (Standard `letsencrypt`); ein DNS-Eintrag für `SYNC_HOSTNAME`.

```bash
git clone https://github.com/dytech-coder/SharePoint-Sync.git
cd SharePoint-Sync
cp .env.example .env              # ausfüllen, siehe unten
cp /pfad/zu/sp_sync_cert.pfx .    # Zertifikat ablegen
mkdir -p data
sudo chown -R 1000:1000 data sp_sync_cert.pfx   # die Container laufen als UID 1000
chmod 600 .env
sudo docker compose up -d --build
```

Prüfen:

```bash
sudo docker compose ps
sudo docker compose logs --tail 50
curl -sI https://<SYNC_HOSTNAME>/login | grep -i location   # Weiterleitung zu login.microsoftonline.com
```

Der erste Sync-Lauf startet sofort nach dem Hochfahren. Solange kein aktives Projekt existiert, tut er nichts.

## Konfiguration (`.env`)

Vorlage mit Erläuterungen: [`.env.example`](.env.example)

| Variable | Pflicht | Bedeutung |
|---|---|---|
| `SP_TENANT_ID` | ja | Verzeichnis-ID (Tenant-ID) |
| `SP_CLIENT_ID` | ja | Anwendungs-ID der App-Registrierung |
| `SP_TENANT_URL` | ja | SharePoint-Adresse, z. B. `https://contoso.sharepoint.com` |
| `SESSION_SECRET` | ja | Schlüssel für Sitzungs-Cookies (`openssl rand -hex 32`) |
| `SYNC_ADMIN_GROUP_ID` | ja* | Objekt-ID der berechtigten Entra-Gruppe |
| `SYNC_ADMINS` | nein | Kommagetrennte Anmeldenamen mit Zugriff unabhängig von der Gruppe |
| `SYNC_HOSTNAME` | ja | Öffentlicher Hostname, z. B. `sync.contoso.com` |
| `APP_TITLE` | nein | Anzeigename in Kopfzeile und Browser-Tab (Standard „SharePoint Sync“) |
| `LOGIN_DOMAIN_HINT` | nein | E-Mail-Domain der Benutzer für die Anmeldung ohne Kontoauswahl |
| `SYNC_INTERVAL_SECONDS` | nein | Pause zwischen zwei Läufen (Standard 14400 = 4 Stunden) |
| `TZ` | nein | Zeitzone (Standard `Europe/Berlin`) |
| `TRAEFIK_NETWORK` / `TRAEFIK_ENTRYPOINT` / `TRAEFIK_CERTRESOLVER` | nein | Traefik-Anbindung (Standard `proxy` / `websecure` / `letsencrypt`) |

\* Sind weder `SYNC_ADMIN_GROUP_ID` noch `SYNC_ADMINS` gesetzt, kann sich **jeder** Benutzer des Tenants anmelden und Projekte ändern.

## Betrieb

| Aufgabe | Befehl |
|---|---|
| Status | `sudo docker compose ps` |
| Protokolle | `sudo docker compose logs -f` |
| Sync sofort starten | `sudo docker compose restart sync` |
| Update einspielen | `git pull && sudo docker compose up -d --build` |

- Nach einem Neustart von `web` müssen sich Benutzer für die Site-/Ordner-Auswahl neu anmelden (die Benutzer-Tokens liegen nur im Arbeitsspeicher).
- Zu sichern sind `.env`, `sp_sync_cert.pfx` und `data/config.json`.
- Das Zertifikat rechtzeitig vor Ablauf erneuern – danach schlagen Anmeldung und Sync gleichzeitig fehl.

### Ohne Docker (Windows-Server)

Unter [`deploy/`](deploy) liegen PowerShell-Skripte für den Betrieb auf einem Windows-Server (Web-App als NSSM-Dienst, Sync als geplante Aufgabe) sowie eine Traefik-Datei-Provider-Konfiguration, um die App von dort aus zu veröffentlichen.

## Verhalten im Detail

- **Einlesen:** Quelle und Ziel werden vollständig über die SharePoint-Drive-Schnittstelle eingelesen (inkl. Hash). Schlägt das Einlesen fehl, wird das Projekt übersprungen und im Ziel nichts gelöscht.
- **Kopieren:** bei abweichendem Inhalts-Hash; für Word-, Excel-, PowerPoint-, Visio- und `.msg`-Dateien bei neuerem Änderungsdatum der Quelle (Liste `REWRITTEN_ON_COPY` in `sync.py`).
- **Löschen:** endgültig, ohne Papierkorb. Leere Ordner bleiben im Ziel stehen.
- **Versionen:** Im Ziel entsteht jeweils eine aktuelle Kopie; der Versionsverlauf wird nicht übertragen.

## Lizenz

Copyright (C) 2026 Dytech IT Solutions GmbH

Dieses Programm ist freie Software: Sie können es unter den Bedingungen der GNU General Public License, Version 3, wie von der Free Software Foundation veröffentlicht, weitergeben und/oder verändern. Es wird ohne jede Gewährleistung bereitgestellt. Den vollständigen Lizenztext enthält die Datei [`LICENSE`](LICENSE).
