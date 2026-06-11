import os
from dotenv import load_dotenv

load_dotenv()

CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE")
SPREADSHEET_ID   = os.getenv("SPREADSHEET_ID")
GEMINI_API_KEY   = os.getenv("GEMINI_API_KEY")

print("=" * 50)
print("TEST KONEKSI PIPELINE MAXY DCB")
print("=" * 50)

# ── TEST 1: Google Sheets ──────────────────────────
print("\n[1/3] Menghubungkan ke Google Sheets...")
try:
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    sheet       = spreadsheet.worksheet("Form Responses 1")
    semua_data  = sheet.get_all_values()

    print(f"  BERHASIL! Spreadsheet: '{spreadsheet.title}'")
    print(f"  Tab ditemukan: {[ws.title for ws in spreadsheet.worksheets()]}")
    print(f"  Jumlah baris data (termasuk header): {len(semua_data)}")
    print(f"  Header kolom: {semua_data[0]}")
    if len(semua_data) > 1:
        print(f"  Contoh baris pertama: {semua_data[1][:3]}")
except Exception as e:
    print(f"  GAGAL: {e}")

# ── TEST 2: Google Drive ───────────────────────────
print("\n[2/3] Menghubungkan ke Google Drive API...")
try:
    from googleapiclient.discovery import build

    drive_service = build("drive", "v3", credentials=creds)
    hasil = drive_service.files().list(pageSize=1).execute()
    print("  BERHASIL! Google Drive API terhubung.")
except Exception as e:
    print(f"  GAGAL: {e}")

# ── TEST 3: Gemini API ─────────────────────────────
print("\n[3/3] Menghubungkan ke Gemini API...")
try:
    from google import genai

    client = genai.Client(api_key=GEMINI_API_KEY)
    respon = client.models.generate_content(
        model="gemini-2.5-flash",
        contents="Balas hanya dengan kata: TERHUBUNG"
    )
    print(f"  BERHASIL! Respon Gemini: {respon.text.strip()}")
except Exception as e:
    print(f"  GAGAL: {e}")

print("\n" + "=" * 50)
print("TEST SELESAI")
print("=" * 50)
