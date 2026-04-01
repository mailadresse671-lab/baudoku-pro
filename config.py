import os
from dotenv import load_dotenv

load_dotenv()

# API Keys
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# Google Drive - Ordner-ID direkt aus der URL
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID", "1j8H5CI8ME2erPRUHKaf4y1FELsQQk6pt")
GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")
GOOGLE_TOKEN_FILE = os.getenv("GOOGLE_TOKEN_FILE", "token.json")

# Foto-Quelle: "local" oder "google_drive"
FOTO_QUELLE = os.getenv("FOTO_QUELLE", "local")

# Pfade
BASIS_PFAD = "Bau_Projekte"
MAX_BILDER_PRO_TAG = 10

# Claude Modell
CLAUDE_MODEL = "claude-sonnet-4-6"
