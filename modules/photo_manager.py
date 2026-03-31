import os
import re
from datetime import datetime
from PIL import Image
import config


# ──────────────────────────────────────────────
# Datum aus Foto ermitteln
# ──────────────────────────────────────────────

def _datum_aus_exif(pfad: str) -> str | None:
    try:
        exif = Image.open(pfad).getexif()
        if 36867 in exif:
            return exif[36867][:10].replace(":", "-")
    except Exception:
        pass
    return None


def _datum_aus_dateiname(name: str) -> str | None:
    m = re.search(r"(20\d{2})[-_]?(\d{2})[-_]?(\d{2})", name)
    if m:
        return f"{m[1]}-{m[2]}-{m[3]}"
    return None


def datum_des_fotos(pfad: str) -> str:
    """Gibt das Aufnahmedatum als 'YYYY-MM-DD' zurück."""
    d = _datum_aus_exif(pfad)
    if d:
        return d
    d = _datum_aus_dateiname(os.path.basename(pfad))
    if d:
        return d
    ts = os.path.getmtime(pfad)
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d")


# ──────────────────────────────────────────────
# Lokale Fotos
# ──────────────────────────────────────────────

def lade_lokale_fotos(eingang_pfad: str) -> dict[str, list[str]]:
    """
    Liest alle Fotos aus dem Eingang-Ordner.
    Gibt ein Dict zurück: {"YYYY-MM-DD": [pfad1, pfad2, ...]}
    """
    tage: dict[str, list[str]] = {}
    if not os.path.exists(eingang_pfad):
        return tage

    for f in os.listdir(eingang_pfad):
        if f.lower().endswith((".jpg", ".jpeg", ".png")):
            pfad = os.path.join(eingang_pfad, f)
            datum = datum_des_fotos(pfad)
            tage.setdefault(datum, []).append(pfad)

    return tage


# ──────────────────────────────────────────────
# Google Drive Fotos
# ──────────────────────────────────────────────

def _get_drive_service():
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
    creds = None

    if os.path.exists(config.GOOGLE_TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(config.GOOGLE_TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                config.GOOGLE_CREDENTIALS_FILE, SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open(config.GOOGLE_TOKEN_FILE, "w") as token:
            token.write(creds.to_json())

    return build("drive", "v3", credentials=creds)


def lade_drive_fotos(ziel_ordner: str) -> dict[str, list[str]]:
    """
    Lädt Fotos direkt aus dem konfigurierten Google Drive Ordner (per ID).
    Speichert sie lokal und gibt dict zurück: {"YYYY-MM-DD": [pfad1, ...]}
    """
    from googleapiclient.http import MediaIoBaseDownload
    import io

    service = _get_drive_service()
    ordner_id = config.DRIVE_FOLDER_ID

    if not ordner_id:
        raise ValueError("DRIVE_FOLDER_ID fehlt in der .env Datei.")

    result = service.files().list(
        q=f"'{ordner_id}' in parents and trashed=false",
        fields="files(id, name, mimeType, createdTime)",
        pageSize=200
    ).execute()

    os.makedirs(ziel_ordner, exist_ok=True)
    tage: dict[str, list[str]] = {}

    for datei in result.get("files", []):
        if not datei["mimeType"].startswith("image/"):
            continue

        name = datei["name"]
        lok_pfad = os.path.join(ziel_ordner, name)

        if not os.path.exists(lok_pfad):
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, service.files().get_media(fileId=datei["id"]))
            done = False
            while not done:
                _, done = downloader.next_chunk()
            with open(lok_pfad, "wb") as f:
                f.write(fh.getvalue())

        datum = datum_des_fotos(lok_pfad)
        tage.setdefault(datum, []).append(lok_pfad)

    return tage


# ──────────────────────────────────────────────
# Einheitliche Schnittstelle
# ──────────────────────────────────────────────

def lade_fotos(projekt_pfad: str) -> dict[str, list[str]]:
    """
    Lädt Fotos je nach FOTO_QUELLE Einstellung in config.
    Gibt immer {"YYYY-MM-DD": [pfad, ...]} zurück.
    """
    if config.FOTO_QUELLE == "google_drive":
        eingang = os.path.join(projekt_pfad, "Eingang_Fotos")
        return lade_drive_fotos(eingang)
    else:
        eingang = os.path.join(projekt_pfad, "Eingang_Fotos")
        return lade_lokale_fotos(eingang)
