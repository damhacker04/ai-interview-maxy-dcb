"""
Pipeline Analisis Video Interview
Digital Career Bootcamp - Maxy Academy
"""

import os
import re
import json
import time
import datetime
from pathlib import Path
from dotenv import load_dotenv

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

from google import genai

# ─────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────
load_dotenv()

CREDENTIALS_FILE  = os.getenv("GOOGLE_CREDENTIALS_FILE")
SPREADSHEET_ID    = os.getenv("SPREADSHEET_ID")
GEMINI_API_KEY    = os.getenv("GEMINI_API_KEY")
DOWNLOAD_FOLDER   = os.getenv("DOWNLOAD_FOLDER", "downloads")
BATAS_PER_HARI    = 20      # Free tier Gemini: 20 request/hari
JEDA_ANTAR_VIDEO  = 15      # Detik jeda antar video (hindari rate limit)

TAB_FORM   = "Form Responses 1"
TAB_HASIL  = "Hasil Analisis"

HEADER_HASIL = [
    "Timestamp", "Nama Lengkap", "Link Video",
    "Perkenalan Diri (1-5)", "Motivasi (1-5)",
    "Tantangan Karier (1-5)", "Tujuan Karier (1-5)",
    "Relevansi Program (1-5)", "Komitmen (1-5)",
    "Komunikasi (1-5)", "Rata-rata",
    "Ringkasan", "Rekomendasi", "Status", "Waktu Diproses"
]

# ─────────────────────────────────────────
# PROMPT UNTUK GEMINI
# ─────────────────────────────────────────
PROMPT_ANALISIS = """
Kamu adalah evaluator rekrutmen profesional untuk program Digital Career Bootcamp (DCB) Maxy Academy.

Tonton video wawancara ini dengan seksama. Nilai kandidat berdasarkan 7 kriteria di bawah.
Perhatikan BUKAN hanya apa yang dikatakan, tetapi juga cara penyampaian, sikap tubuh,
kepercayaan diri, intonasi suara, dan komunikasi nonverbal kandidat.

KRITERIA PENILAIAN (skala 1-5):

1. Perkenalan Diri
   Nilai kejelasan memperkenalkan nama, kesibukan saat ini, dan kesadaran diri tentang skill.
   1=Tidak jelas/tidak lengkap | 2=Ada tapi kurang | 3=Cukup | 4=Baik dan jelas | 5=Sangat detail dan percaya diri

2. Motivasi mengikuti Program
   Nilai kedalaman dan ketulusan motivasi memilih DCB Maxy Academy.
   1=Tidak jelas/sangat umum | 2=Ada tapi lemah | 3=Cukup | 4=Jelas dan spesifik | 5=Sangat kuat dan meyakinkan

3. Tantangan Karier & Cara Mengatasinya
   Nilai kemampuan menceritakan tantangan nyata dan solusi konkret yang diambil.
   1=Tidak disebutkan | 2=Dangkal | 3=Cukup | 4=Konkret dan reflektif | 5=Sangat insightful

4. Tujuan Karier 3-5 Tahun
   Nilai kejelasan dan realisme target karier jangka menengah.
   1=Tidak ada tujuan | 2=Sangat umum | 3=Cukup jelas | 4=Spesifik dan realistis | 5=Sangat detail dan terencana

5. Relevansi Program dengan Target Karier
   Nilai seberapa jelas kandidat menghubungkan program DCB dengan tujuan kariernya.
   1=Tidak ada koneksi | 2=Lemah | 3=Cukup | 4=Jelas dan logis | 5=Sangat kuat dan meyakinkan

6. Komitmen & Manajemen Waktu
   Nilai kesungguhan dan rencana nyata untuk mengikuti program secara penuh.
   1=Tidak berkomitmen | 2=Ragu-ragu | 3=Cukup | 4=Berkomitmen dengan rencana | 5=Sangat berkomitmen dan detail

7. Komunikasi & Kejelasan Penyampaian
   Nilai kualitas komunikasi verbal, intonasi, kepercayaan diri, dan ekspresi saat berbicara.
   1=Sangat sulit dipahami | 2=Kurang jelas | 3=Cukup | 4=Jelas dan percaya diri | 5=Sangat komunikatif dan meyakinkan

Berikan output dalam format JSON berikut SAJA (tanpa teks lain di luar JSON):
{
  "perkenalan_diri": <angka 1-5>,
  "motivasi": <angka 1-5>,
  "tantangan_karier": <angka 1-5>,
  "tujuan_karier": <angka 1-5>,
  "relevansi_program": <angka 1-5>,
  "komitmen": <angka 1-5>,
  "komunikasi": <angka 1-5>,
  "ringkasan": "<2-3 kalimat Bahasa Indonesia: kekuatan utama dan kelemahan kandidat>",
  "rekomendasi": "<Pass atau Review atau Reject>"
}

Aturan rekomendasi otomatis:
- Rata-rata skor >= 4.0  →  Pass
- Rata-rata skor 3.0-3.9 →  Review
- Rata-rata skor < 3.0   →  Reject
"""


# ─────────────────────────────────────────
# FUNGSI SETUP
# ─────────────────────────────────────────
def setup_clients():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds         = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    sheets_client = gspread.authorize(creds)
    drive_service = build("drive", "v3", credentials=creds)
    gemini_client = genai.Client(api_key=GEMINI_API_KEY)
    return sheets_client, drive_service, gemini_client


def get_or_create_hasil_tab(spreadsheet):
    try:
        ws = spreadsheet.worksheet(TAB_HASIL)
        print(f"  Tab '{TAB_HASIL}' sudah ada.")
        return ws
    except gspread.WorksheetNotFound:
        print(f"  Membuat tab '{TAB_HASIL}' baru...")
        ws = spreadsheet.add_worksheet(title=TAB_HASIL, rows=1000, cols=20)
        ws.append_row(HEADER_HASIL)
        return ws


def get_nama_sudah_diproses(hasil_ws):
    data = hasil_ws.get_all_values()
    if len(data) <= 1:
        return set()
    return {baris[1] for baris in data[1:] if len(baris) > 1 and baris[1]}


# ─────────────────────────────────────────
# FUNGSI GOOGLE DRIVE
# ─────────────────────────────────────────
def extract_drive_id(url):
    """Ekstrak ID dari berbagai format URL Google Drive."""
    # Format: /file/d/ID/
    m = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
    if m:
        return m.group(1), 'file'
    # Format: /folders/ID
    m = re.search(r'/folders/([a-zA-Z0-9_-]+)', url)
    if m:
        return m.group(1), 'folder'
    # Format: ?id=ID atau &id=ID
    m = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', url)
    if m:
        return m.group(1), 'file'
    return None, None


def get_file_id_dari_folder(drive_service, folder_id):
    """Cari file video pertama di dalam folder Google Drive."""
    hasil = drive_service.files().list(
        q=f"'{folder_id}' in parents and mimeType contains 'video/'",
        fields="files(id, name, mimeType)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    files = hasil.get('files', [])
    if not files:
        raise Exception("Tidak ada file video di folder Drive ini")
    print(f"    File ditemukan: {files[0]['name']}")
    return files[0]['id']


def download_video(drive_service, file_id, nama_pelamar):
    """Download video dari Google Drive ke folder lokal."""
    nama_file = re.sub(r'[^\w\s-]', '', nama_pelamar).strip().replace(' ', '_')
    path = os.path.join(DOWNLOAD_FOLDER, f"{nama_file}.mp4")

    request  = drive_service.files().get_media(
        fileId=file_id,
        supportsAllDrives=True
    )
    with open(path, 'wb') as f:
        downloader = MediaIoBaseDownload(f, request, chunksize=10 * 1024 * 1024)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                print(f"    Download: {int(status.progress() * 100)}%")
    return path


# ─────────────────────────────────────────
# FUNGSI GEMINI
# ─────────────────────────────────────────
def analisis_video_gemini(gemini_client, video_path):
    """Upload video ke Gemini dan dapatkan hasil analisis."""

    # Upload video
    print("    Mengupload video ke Gemini...")
    video_file = gemini_client.files.upload(file=video_path)

    # Tunggu sampai video selesai diproses Gemini (maks 5 menit)
    print("    Menunggu video diproses Gemini...")
    batas_tunggu = 300
    sudah_tunggu = 0
    while sudah_tunggu < batas_tunggu:
        video_file = gemini_client.files.get(name=video_file.name)
        if video_file.state.name == "ACTIVE":
            break
        elif video_file.state.name == "FAILED":
            raise Exception("Gemini gagal memproses video ini")
        time.sleep(10)
        sudah_tunggu += 10
        print(f"    Masih memproses... ({sudah_tunggu}s)")

    if video_file.state.name != "ACTIVE":
        raise Exception("Video tidak aktif setelah menunggu 5 menit")

    # Analisis
    print("    Menganalisis video dengan Gemini...")
    response = gemini_client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[video_file, PROMPT_ANALISIS]
    )

    # Hapus file dari Gemini setelah selesai (hemat kuota storage)
    try:
        gemini_client.files.delete(name=video_file.name)
    except Exception:
        pass

    return response.text


def parse_hasil_gemini(teks):
    """Ambil data JSON dari response Gemini."""
    m = re.search(r'\{[\s\S]*\}', teks)
    if not m:
        raise Exception(f"Format response tidak dikenali: {teks[:200]}")
    return json.loads(m.group())


def hitung_rata_rata(data):
    skor = [
        data['perkenalan_diri'], data['motivasi'], data['tantangan_karier'],
        data['tujuan_karier'], data['relevansi_program'],
        data['komitmen'], data['komunikasi']
    ]
    return round(sum(skor) / len(skor), 2)


# ─────────────────────────────────────────
# PIPELINE UTAMA
# ─────────────────────────────────────────
def main():
    print("=" * 58)
    print("  PIPELINE ANALISIS VIDEO INTERVIEW - DCB MAXY ACADEMY")
    print("=" * 58)
    print(f"  Waktu mulai : {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Batas hari  : {BATAS_PER_HARI} video")
    print("=" * 58)

    Path(DOWNLOAD_FOLDER).mkdir(exist_ok=True)

    # Setup koneksi
    print("\n[Setup] Menghubungkan ke Google & Gemini...")
    sheets_client, drive_service, gemini_client = setup_clients()
    spreadsheet = sheets_client.open_by_key(SPREADSHEET_ID)
    form_ws     = spreadsheet.worksheet(TAB_FORM)
    hasil_ws    = get_or_create_hasil_tab(spreadsheet)
    print("  Semua koneksi berhasil.")

    # Ambil data pelamar
    semua_data        = form_ws.get_all_values()[1:]  # skip baris header
    sudah_diproses    = get_nama_sudah_diproses(hasil_ws)
    belum_diproses    = [
        baris for baris in semua_data
        if len(baris) >= 3 and baris[1] and baris[1] not in sudah_diproses
    ]

    akan_diproses = belum_diproses[:BATAS_PER_HARI]

    print(f"\n  Total pelamar di Sheet : {len(semua_data)}")
    print(f"  Sudah dianalisis       : {len(sudah_diproses)}")
    print(f"  Akan dianalisis hari ini: {len(akan_diproses)}")
    sisa = len(belum_diproses) - len(akan_diproses)
    if sisa > 0:
        print(f"  Sisa untuk hari berikut: {sisa} (jalankan lagi besok)")

    if not akan_diproses:
        print("\n  Semua pelamar sudah selesai dianalisis!")
        return

    print("\n" + "-" * 58)

    # Loop per pelamar
    for i, baris in enumerate(akan_diproses):
        timestamp   = baris[0]
        nama        = baris[1]
        link_video  = baris[2]
        video_path  = None

        print(f"\n[{i+1}/{len(akan_diproses)}] {nama}")

        try:
            # LANGKAH 1: Download video
            print("  → Download video dari Google Drive...")
            drive_id, tipe = extract_drive_id(link_video)
            if not drive_id:
                raise Exception("Format link Google Drive tidak dikenali")

            if tipe == 'folder':
                file_id = get_file_id_dari_folder(drive_service, drive_id)
            else:
                file_id = drive_id

            video_path = download_video(drive_service, file_id, nama)
            ukuran_mb  = os.path.getsize(video_path) / (1024 * 1024)
            print(f"  → Download selesai ({ukuran_mb:.1f} MB)")

            # LANGKAH 2: Analisis dengan Gemini
            print("  → Mengirim ke Gemini untuk dianalisis...")
            raw_response = analisis_video_gemini(gemini_client, video_path)

            # LANGKAH 3: Parse hasil
            hasil = parse_hasil_gemini(raw_response)
            rata  = hitung_rata_rata(hasil)
            waktu = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

            # LANGKAH 4: Tulis ke Google Sheet
            baris_hasil = [
                timestamp, nama, link_video,
                hasil['perkenalan_diri'],
                hasil['motivasi'],
                hasil['tantangan_karier'],
                hasil['tujuan_karier'],
                hasil['relevansi_program'],
                hasil['komitmen'],
                hasil['komunikasi'],
                rata,
                hasil['ringkasan'],
                hasil['rekomendasi'],
                "Selesai",
                waktu
            ]
            hasil_ws.append_row(baris_hasil)

            print(f"  ✓ Selesai! Rata-rata: {rata} → {hasil['rekomendasi']}")

        except Exception as e:
            print(f"  ✗ ERROR: {e}")
            waktu = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            hasil_ws.append_row([
                timestamp, nama, link_video,
                "", "", "", "", "", "", "", "",
                f"Error: {str(e)}", "Error", "Error", waktu
            ])

        finally:
            # Hapus video dari disk setelah selesai
            if video_path and os.path.exists(video_path):
                os.remove(video_path)
                print("  → File video lokal dihapus.")

        # Jeda antar video
        if i < len(akan_diproses) - 1:
            print(f"  → Jeda {JEDA_ANTAR_VIDEO} detik sebelum video berikutnya...")
            time.sleep(JEDA_ANTAR_VIDEO)

    # Selesai
    print("\n" + "=" * 58)
    print("  PIPELINE SELESAI")
    print(f"  Waktu selesai: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("  Buka Google Sheet untuk melihat hasil analisis.")
    print("=" * 58)


if __name__ == "__main__":
    main()
