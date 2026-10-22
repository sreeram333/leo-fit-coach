import urllib.request
import csv
import io
import sys

SHEET_EXPORT_URL = "https://docs.google.com/spreadsheets/d/1RARxGSBzzrw-4Sy2wR02aQie7NasvvjjG0py0BDM0bo/export?format=csv"

def sync_data():
    try:
        print("[*] Fetching latest entries from Google Sheets...")
        req = urllib.request.Request(SHEET_EXPORT_URL, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as resp:
            content = resp.read().decode('utf-8')
        
        # Save local copy
        with open("daily_log.csv", "w", encoding="utf-8") as f:
            f.write(content)
        
        reader = csv.DictReader(io.StringIO(content))
        rows = list(reader)
        print(f"[+] Successfully synced {len(rows)} logged day(s) to local daily_log.csv")
        
        if rows:
            latest = rows[-1]
            print("\n--- Latest Logged Entry ---")
            for k, v in latest.items():
                print(f"  {k}: {v}")
    except Exception as e:
        print(f"[!] Error fetching sheet: {e}")

if __name__ == "__main__":
    sync_data()
