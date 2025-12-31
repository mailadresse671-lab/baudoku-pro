import os
import subprocess
import datetime
import time

def git_push():
    try:
        # Zeitstempel
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        commit_message = f"Update: {now}"

        print("--- Starte Upload zu GitHub ---")
        print("Bitte warten... das kann bei Fotos kurz dauern.")

        # 1. Alles hinzufuegen
        subprocess.run(["git", "add", "."], check=True)
        
        # 2. Verpacken
        subprocess.run(["git", "commit", "-m", commit_message], check=True)
        
        # 3. Hochladen
        subprocess.run(["git", "push"], check=True)
        
        print("\nGRUEN: Alles erfolgreich hochgeladen!")
        print(f"Zeitpunkt: {now}")

    except Exception as e:
        print(f"\nFEHLER oder nichts Neues zum Speichern.\nMeldung: {e}")

    print("\nFenster schliesst sich in 10 Sekunden...")
    time.sleep(10)

if __name__ == "__main__":
    git_push()