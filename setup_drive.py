"""
Google Drive Einrichtung – einmalig ausführen.

Voraussetzungen:
1. Google Cloud Console: https://console.cloud.google.com
2. Neues Projekt erstellen (z.B. "BauDoku")
3. Google Drive API aktivieren
4. Anmeldedaten → OAuth 2.0-Client-ID erstellen (Typ: Desktop-App)
5. JSON herunterladen → als "credentials.json" in diesen Ordner speichern
6. Dieses Script ausführen: python setup_drive.py
"""

import os
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.json"

def main():
    print("=" * 50)
    print("BauDoku Pro – Google Drive Einrichtung")
    print("=" * 50)

    if not os.path.exists(CREDENTIALS_FILE):
        print(f"\n❌ Datei '{CREDENTIALS_FILE}' nicht gefunden!")
        print("\nSchritte:")
        print("1. Öffne: https://console.cloud.google.com")
        print("2. Neues Projekt erstellen → 'BauDoku'")
        print("3. APIs & Dienste → Google Drive API aktivieren")
        print("4. Anmeldedaten → OAuth 2.0-Client-ID (Desktop-Anwendung)")
        print("5. JSON herunterladen → als 'credentials.json' speichern")
        print("6. Dieses Script erneut starten")
        return

    print("\nStarte Google-Anmeldung im Browser...")
    print("(Falls kein Browser öffnet, den Link manuell kopieren)\n")

    flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
    creds = flow.run_local_server(port=0)

    with open(TOKEN_FILE, "w") as f:
        f.write(creds.to_json())

    print(f"\n✅ Erfolgreich! Token gespeichert in '{TOKEN_FILE}'")
    print("\nNächster Schritt:")
    print("In der .env Datei eintragen:")
    print("  FOTO_QUELLE=google_drive")
    print("  DRIVE_FOLDER_NAME=Baudoku   (Name deines Google Drive Ordners)")


if __name__ == "__main__":
    main()
