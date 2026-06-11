"""
Dashboard Rekrutmen - Digital Career Bootcamp Maxy Academy
Diakses oleh tim HR untuk melihat hasil analisis video interview
"""

import io
import os
import re
import json
import time
from datetime import datetime
from pathlib import Path

import gspread
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google import genai

load_dotenv()

CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")
DOWNLOAD_FOLDER  = "downloads"

TAB_FORM  = "Form Responses 1"
TAB_HASIL = "Hasil Analisis"

KOLOM_SKOR = [
    "Perkenalan Diri (1-5)", "Motivasi (1-5)", "Tantangan Karier (1-5)",
    "Tujuan Karier (1-5)", "Relevansi Program (1-5)", "Komitmen (1-5)", "Komunikasi (1-5)",
]

HEADER_HASIL = [
    "Timestamp", "Nama Lengkap", "Link Video",
    "Perkenalan Diri (1-5)", "Motivasi (1-5)", "Tantangan Karier (1-5)",
    "Tujuan Karier (1-5)", "Relevansi Program (1-5)", "Komitmen (1-5)",
    "Komunikasi (1-5)", "Rata-rata", "Ringkasan", "Rekomendasi", "Status", "Waktu Diproses"
]

PROMPT_ANALISIS = """
Kamu adalah evaluator rekrutmen profesional untuk program Digital Career Bootcamp (DCB) Maxy Academy.

Tonton video wawancara ini dengan seksama. Nilai kandidat berdasarkan 7 kriteria di bawah.
Perhatikan BUKAN hanya apa yang dikatakan, tetapi juga cara penyampaian, sikap tubuh,
kepercayaan diri, intonasi suara, dan komunikasi nonverbal kandidat.

KRITERIA PENILAIAN (skala 1-5):

1. Perkenalan Diri
   1=Tidak jelas/tidak lengkap | 2=Ada tapi kurang | 3=Cukup | 4=Baik dan jelas | 5=Sangat detail dan percaya diri

2. Motivasi mengikuti Program
   1=Tidak jelas/sangat umum | 2=Ada tapi lemah | 3=Cukup | 4=Jelas dan spesifik | 5=Sangat kuat dan meyakinkan

3. Tantangan Karier & Cara Mengatasinya
   1=Tidak disebutkan | 2=Dangkal | 3=Cukup | 4=Konkret dan reflektif | 5=Sangat insightful

4. Tujuan Karier 3-5 Tahun
   1=Tidak ada tujuan | 2=Sangat umum | 3=Cukup jelas | 4=Spesifik dan realistis | 5=Sangat detail dan terencana

5. Relevansi Program dengan Target Karier
   1=Tidak ada koneksi | 2=Lemah | 3=Cukup | 4=Jelas dan logis | 5=Sangat kuat dan meyakinkan

6. Komitmen & Manajemen Waktu
   1=Tidak berkomitmen | 2=Ragu-ragu | 3=Cukup | 4=Berkomitmen dengan rencana | 5=Sangat berkomitmen dan detail

7. Komunikasi & Kejelasan Penyampaian
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
# FUNGSI KONEKSI
# ─────────────────────────────────────────
def get_spreadsheet_id():
    try:
        val = st.secrets["SPREADSHEET_ID"]
        if val: return val
    except (KeyError, FileNotFoundError):
        pass
    val = os.getenv("SPREADSHEET_ID")
    if not val:
        st.error("❌ SPREADSHEET_ID belum dikonfigurasi.")
        st.stop()
    return val


def get_google_creds():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    try:
        sa_info = st.secrets["gcp_service_account"]
        private_key = sa_info["private_key"]
        if "\\n" in private_key:
            private_key = private_key.replace("\\n", "\n")
        creds_dict = {
            "type":                        sa_info["type"],
            "project_id":                  sa_info["project_id"],
            "private_key_id":              sa_info["private_key_id"],
            "private_key":                 private_key,
            "client_email":                sa_info["client_email"],
            "client_id":                   sa_info["client_id"],
            "auth_uri":                    sa_info["auth_uri"],
            "token_uri":                   sa_info["token_uri"],
            "auth_provider_x509_cert_url": sa_info["auth_provider_x509_cert_url"],
            "client_x509_cert_url":        sa_info["client_x509_cert_url"],
        }
        return Credentials.from_service_account_info(creds_dict, scopes=scopes)
    except (KeyError, FileNotFoundError):
        pass
    except Exception:
        st.error(
            "❌ **Streamlit Secrets ada, tapi `private_key` tidak valid.**\n\n"
            "Buka **Manage app → Settings → Secrets**, ganti format `private_key` "
            "dengan triple-quote multi-baris:\n"
            "```toml\nprivate_key = \"\"\"\n-----BEGIN PRIVATE KEY-----\n"
            "(baris-baris key)\n-----END PRIVATE KEY-----\n\"\"\"\n```"
        )
        st.stop()
    if not os.path.exists(CREDENTIALS_FILE):
        st.error(
            "❌ Google credentials tidak ditemukan.\n\n"
            "Di Streamlit Cloud: tambahkan **[gcp_service_account]** di Secrets.\n"
            "Di lokal: pastikan `credentials.json` ada di folder proyek."
        )
        st.stop()
    return Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)


def get_gemini_key():
    try:
        k = st.secrets["GEMINI_API_KEY"]
        if k: return k
    except (KeyError, FileNotFoundError):
        pass
    k = os.getenv("GEMINI_API_KEY")
    if not k:
        st.error("❌ GEMINI_API_KEY tidak ditemukan. Tambahkan ke Streamlit Secrets atau file .env.")
        st.stop()
    return k


# ─────────────────────────────────────────
# FUNGSI ANALISIS (sama dengan pipeline.py)
# ─────────────────────────────────────────
def extract_drive_id(url):
    m = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
    if m: return m.group(1), 'file'
    m = re.search(r'/folders/([a-zA-Z0-9_-]+)', url)
    if m: return m.group(1), 'folder'
    m = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', url)
    if m: return m.group(1), 'file'
    return None, None


def get_file_id_dari_folder(drive_service, folder_id):
    hasil = drive_service.files().list(
        q=f"'{folder_id}' in parents and mimeType contains 'video/'",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = hasil.get('files', [])
    if not files:
        raise Exception("Tidak ada file video di folder Drive ini")
    return files[0]['id']


def download_video(drive_service, file_id, nama_pelamar):
    Path(DOWNLOAD_FOLDER).mkdir(exist_ok=True)
    nama_file = re.sub(r'[^\w\s-]', '', nama_pelamar).strip().replace(' ', '_')
    path = os.path.join(DOWNLOAD_FOLDER, f"{nama_file}.mp4")
    request = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)
    with open(path, 'wb') as f:
        downloader = MediaIoBaseDownload(f, request, chunksize=10 * 1024 * 1024)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return path


def analisis_video_gemini(gemini_client, video_path):
    video_file = gemini_client.files.upload(file=video_path)
    batas, sudah = 300, 0
    while sudah < batas:
        video_file = gemini_client.files.get(name=video_file.name)
        if video_file.state.name == "ACTIVE":
            break
        if video_file.state.name == "FAILED":
            raise Exception("Gemini gagal memproses video")
        time.sleep(10)
        sudah += 10
    if video_file.state.name != "ACTIVE":
        raise Exception("Video tidak aktif setelah 5 menit")
    response = gemini_client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[video_file, PROMPT_ANALISIS],
    )
    try:
        gemini_client.files.delete(name=video_file.name)
    except Exception:
        pass
    return response.text


def parse_hasil_gemini(teks):
    m = re.search(r'\{[\s\S]*\}', teks)
    if not m:
        raise Exception(f"Format response tidak dikenali: {teks[:200]}")
    return json.loads(m.group())


def hitung_rata_rata(data):
    skor = [
        data['perkenalan_diri'], data['motivasi'], data['tantangan_karier'],
        data['tujuan_karier'], data['relevansi_program'], data['komitmen'], data['komunikasi'],
    ]
    return round(sum(skor) / len(skor), 2)


def jalankan_analisis_satu(nama, timestamp, link_video, hasil_ws):
    """Analisis satu kandidat dan tulis hasilnya ke Google Sheet."""
    video_path = None
    try:
        creds         = get_google_creds()
        drive_service = build("drive", "v3", credentials=creds)
        gemini_client = genai.Client(api_key=get_gemini_key())

        drive_id, tipe = extract_drive_id(link_video)
        if not drive_id:
            raise Exception("Format link Google Drive tidak dikenali")
        file_id    = get_file_id_dari_folder(drive_service, drive_id) if tipe == 'folder' else drive_id
        video_path = download_video(drive_service, file_id, nama)

        raw    = analisis_video_gemini(gemini_client, video_path)
        hasil  = parse_hasil_gemini(raw)
        rata   = hitung_rata_rata(hasil)
        waktu  = datetime.now().strftime("%Y-%m-%d %H:%M")

        hasil_ws.append_row([
            timestamp, nama, link_video,
            hasil['perkenalan_diri'], hasil['motivasi'], hasil['tantangan_karier'],
            hasil['tujuan_karier'], hasil['relevansi_program'], hasil['komitmen'],
            hasil['komunikasi'], rata, hasil['ringkasan'], hasil['rekomendasi'],
            "Selesai", waktu,
        ])
        return hasil['rekomendasi'], rata

    except Exception:
        waktu = datetime.now().strftime("%Y-%m-%d %H:%M")
        try:
            hasil_ws.append_row([
                timestamp, nama, link_video,
                "", "", "", "", "", "", "", "",
                f"Error: analisis gagal", "Error", "Error", waktu,
            ])
        except Exception:
            pass
        raise
    finally:
        if video_path and os.path.exists(video_path):
            os.remove(video_path)


# ─────────────────────────────────────────
# KONFIGURASI HALAMAN
# ─────────────────────────────────────────
st.set_page_config(
    page_title="DCB Maxy Academy – Interview Analyzer",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    .header-box {
        background: linear-gradient(135deg, #0f2d5c 0%, #1a6eb5 100%);
        padding: 2rem 2.5rem; border-radius: 12px;
        color: white; margin-bottom: 1.5rem;
    }
    .header-box h1 { margin: 0; font-size: 1.8rem; }
    .header-box p  { margin: 0.3rem 0 0; opacity: 0.85; font-size: 1rem; }
    div[data-testid="stMetric"] {
        background: white; padding: 1rem; border-radius: 10px;
        box-shadow: 0 1px 6px rgba(0,0,0,0.08);
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────
@st.cache_data(ttl=120)
def load_data():
    creds  = get_google_creds()
    client = gspread.authorize(creds)
    sheet  = client.open_by_key(get_spreadsheet_id())

    form_data  = sheet.worksheet(TAB_FORM).get_all_values()
    total_form = max(0, len(form_data) - 1)

    try:
        hasil_data = sheet.worksheet(TAB_HASIL).get_all_values()
    except gspread.WorksheetNotFound:
        return pd.DataFrame(), total_form, form_data

    if len(hasil_data) <= 1:
        return pd.DataFrame(), total_form, form_data

    df = pd.DataFrame(hasil_data[1:], columns=hasil_data[0])
    for col in KOLOM_SKOR + ["Rata-rata"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df, total_form, form_data


def style_baris(row):
    warna = {
        "Pass":   "background-color: #f0fff4",
        "Review": "background-color: #fffdf0",
        "Reject": "background-color: #fff5f5",
        "Error":  "background-color: #f5f5f5",
    }.get(row.get("Rekomendasi", ""), "")
    return [warna] * len(row)


# ─────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────
def main():

    st.markdown("""
    <div class="header-box">
        <h1>🎯 DCB Maxy Academy — Video Interview Analyzer</h1>
        <p>Dashboard rekrutmen Digital Career Bootcamp · Hasil analisis AI</p>
    </div>
    """, unsafe_allow_html=True)

    col_btn, col_waktu = st.columns([1, 5])
    with col_btn:
        if st.button("🔄 Refresh", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    with col_waktu:
        st.caption(f"Data dimuat pada: {datetime.now().strftime('%d %b %Y  %H:%M:%S')}")

    with st.spinner("Memuat data dari Google Sheet..."):
        df, total_form, form_data = load_data()

    tab1, tab2 = st.tabs(["📊 Hasil Analisis", "🔬 Analisis Kandidat Baru"])

    # ══════════════════════════════════════
    # TAB 1 — HASIL ANALISIS
    # ══════════════════════════════════════
    with tab1:
        if df.empty:
            st.info("⏳ Belum ada hasil analisis. Gunakan tab **Analisis Kandidat Baru**.")
            st.metric("Total Pelamar di Form", total_form)
        else:
            df_selesai = df[df["Status"] == "Selesai"]
            df_error   = df[df["Status"] == "Error"]

            n_pass    = (df_selesai["Rekomendasi"] == "Pass").sum()
            n_review  = (df_selesai["Rekomendasi"] == "Review").sum()
            n_reject  = (df_selesai["Rekomendasi"] == "Reject").sum()
            n_selesai = len(df_selesai)

            # ── Ringkasan ──
            st.subheader("📊 Ringkasan")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("👥 Total Pelamar",    total_form)
            c2.metric("✅ Sudah Dianalisis", n_selesai)
            c3.metric("⏳ Belum Dianalisis", max(total_form - len(df), 0))
            c4.metric("⚠️ Gagal Diproses",  len(df_error))

            st.divider()

            # ── Rekomendasi ──
            st.subheader("🏆 Hasil Rekomendasi")
            r1, r2, r3 = st.columns(3)
            with r1:
                st.metric("✅ Pass", n_pass)
                st.caption("Rata-rata ≥ 4.0 · Direkomendasikan lanjut")
            with r2:
                st.metric("🔍 Perlu Ditinjau", n_review)
                st.caption("Rata-rata 3.0–3.9 · Perlu review manual")
            with r3:
                st.metric("❌ Tidak Lolos", n_reject)
                st.caption("Rata-rata < 3.0 · Tidak direkomendasikan")

            st.divider()

            # ── Tabel ──
            st.subheader("📋 Daftar Semua Pelamar")
            fc1, fc2 = st.columns([2, 3])
            with fc1:
                filter_rek = st.selectbox(
                    "Filter rekomendasi:",
                    ["Semua", "Pass", "Perlu Ditinjau (Review)", "Tidak Lolos (Reject)", "Error"],
                )
            with fc2:
                cari = st.text_input("🔍 Cari nama:", placeholder="Ketik nama pelamar...")

            df_gabung = pd.concat([df_selesai, df_error], ignore_index=True)
            peta = {"Pass": "Pass", "Perlu Ditinjau (Review)": "Review",
                    "Tidak Lolos (Reject)": "Reject", "Error": "Error"}
            if filter_rek != "Semua":
                df_gabung = df_gabung[df_gabung["Rekomendasi"] == peta[filter_rek]]
            if cari:
                df_gabung = df_gabung[
                    df_gabung["Nama Lengkap"].str.lower().str.contains(cari.lower(), na=False)
                ]

            st.caption(f"Menampilkan **{len(df_gabung)}** dari **{len(df)}** pelamar")
            kolom_tampil = [c for c in ["Nama Lengkap"] + KOLOM_SKOR + ["Rata-rata", "Rekomendasi"]
                            if c in df_gabung.columns]
            fmt = {"Rata-rata": "{:.2f}"}
            fmt.update({k: "{:.0f}" for k in KOLOM_SKOR if k in df_gabung.columns})
            st.dataframe(
                df_gabung[kolom_tampil].style.apply(style_baris, axis=1).format(fmt, na_rep="–"),
                use_container_width=True, height=420, hide_index=True,
            )

            st.divider()

            # ── Detail ──
            st.subheader("🔎 Detail Pelamar")
            pilihan = st.selectbox(
                "Pilih pelamar:", ["— Pilih nama pelamar —"] + df_gabung["Nama Lengkap"].tolist(),
                label_visibility="collapsed",
            )
            if pilihan != "— Pilih nama pelamar —":
                row = df_gabung[df_gabung["Nama Lengkap"] == pilihan].iloc[0]
                rek   = row.get("Rekomendasi", "–")
                emoji = {"Pass": "✅", "Review": "🔍", "Reject": "❌"}.get(rek, "⚠️")
                kiri, kanan = st.columns(2)
                with kiri:
                    st.markdown(f"**Nama:** {row.get('Nama Lengkap', '–')}")
                    st.markdown(f"**Tanggal Submit:** {row.get('Timestamp', '–')}")
                    st.markdown(f"**Rekomendasi:** {emoji} **{rek}**")
                    st.markdown(f"**Rata-rata Skor:** {row.get('Rata-rata', '–')}")
                    st.markdown(f"**Dianalisis:** {row.get('Waktu Diproses', '–')}")
                    st.markdown("**Ringkasan AI:**")
                    ringkasan = str(row.get("Ringkasan", "–"))
                    if "Error" in ringkasan:
                        st.error(ringkasan)
                    else:
                        st.info(ringkasan)
                with kanan:
                    st.markdown("**Skor per Kriteria:**")
                    def warna_skor(val):
                        try:
                            v = float(val)
                            if v >= 4: return "background-color:#d4edda;color:#155724"
                            if v >= 3: return "background-color:#fff3cd;color:#856404"
                            return "background-color:#f8d7da;color:#721c24"
                        except Exception: return ""
                    skor_df = pd.DataFrame({
                        "Kriteria": ["Perkenalan Diri", "Motivasi", "Tantangan Karier",
                                     "Tujuan Karier", "Relevansi Program", "Komitmen", "Komunikasi"],
                        "Skor": [row.get(k, "–") for k in KOLOM_SKOR],
                    })
                    st.dataframe(skor_df.style.map(warna_skor, subset=["Skor"]),
                                 hide_index=True, use_container_width=True, height=284)

            st.divider()

            # ── Export ──
            st.subheader("📥 Export Data")
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df_gabung.to_excel(writer, index=False, sheet_name="Hasil Analisis")
            st.download_button(
                label="⬇️ Download Excel (data yang sedang ditampilkan)",
                data=buf.getvalue(),
                file_name=f"DCB_Hasil_Analisis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.caption("File Excel berisi data sesuai filter aktif.")

    # ══════════════════════════════════════
    # TAB 2 — ANALISIS KANDIDAT BARU
    # ══════════════════════════════════════
    with tab2:
        st.markdown("### 🔬 Analisis Kandidat Baru")
        st.caption("Klik **Analyze** untuk menganalisis video menggunakan AI · Proses ~2–5 menit per video · Jangan tutup tab selama proses berjalan")

        # Nama yang sudah ada di hasil (selesai maupun error)
        nama_sudah = set(df["Nama Lengkap"].tolist()) if not df.empty else set()
        nama_error = set(df[df["Status"] == "Error"]["Nama Lengkap"].tolist()) if not df.empty else set()

        if len(form_data) <= 1:
            st.info("Belum ada data kandidat di Google Form.")
        else:
            form_rows = form_data[1:]  # skip header

            # Kandidat belum pernah dianalisis
            belum = [r for r in form_rows if len(r) >= 3 and r[1] and r[1] not in nama_sudah]

            # Kandidat error (bisa di-retry) — ambil dari form bukan dari df_error
            retry = [r for r in form_rows if len(r) >= 3 and r[1] and r[1] in nama_error]

            semua = belum + retry

            # Deduplikasi berdasarkan nama (cegah duplicate key jika submit form > 1x)
            seen_names: set = set()
            semua_unik = []
            for r in semua:
                n = r[1] if len(r) > 1 else ""
                if n and n not in seen_names:
                    seen_names.add(n)
                    semua_unik.append(r)
            semua = semua_unik

            if not semua:
                st.success("✅ Semua kandidat sudah berhasil dianalisis!")
            else:
                # Session state
                if "analyzing_set" not in st.session_state:
                    st.session_state.analyzing_set = set()
                if "done_set" not in st.session_state:
                    st.session_state.done_set = set()

                belum_count = len(belum)
                retry_count = len(retry)
                info_parts = []
                if belum_count: info_parts.append(f"**{belum_count}** belum dianalisis")
                if retry_count: info_parts.append(f"**{retry_count}** perlu di-retry")
                st.info(" · ".join(info_parts))

                # ── Header kolom ──
                h0, h1, h2, h3 = st.columns([3, 2, 4, 2])
                h0.markdown("**Nama Kandidat**")
                h1.markdown("**Status**")
                h2.markdown("**Link Video**")
                h3.markdown("**Aksi**")
                st.markdown("---")

                for row in semua:
                    timestamp  = row[0] if len(row) > 0 else ""
                    nama       = row[1] if len(row) > 1 else ""
                    link_video = row[2] if len(row) > 2 else ""
                    if not nama:
                        continue

                    is_retry     = nama in nama_error
                    is_analyzing = nama in st.session_state.analyzing_set
                    is_done      = nama in st.session_state.done_set

                    c0, c1, c2, c3 = st.columns([3, 2, 4, 2])

                    with c0:
                        st.write(nama)
                    with c1:
                        if is_done:
                            st.markdown("✅ Selesai")
                        elif is_analyzing:
                            st.markdown("⏳ Dianalisis...")
                        elif is_retry:
                            st.markdown("⚠️ Error — retry")
                        else:
                            st.markdown("🟡 Belum dianalisis")
                    with c2:
                        tampil = link_video[:50] + "…" if len(link_video) > 50 else link_video
                        st.caption(tampil)
                    with c3:
                        if is_done:
                            st.write("—")
                        else:
                            label = "⏳ Menganalisis..." if is_analyzing else ("🔁 Retry" if is_retry else "▶️ Analyze")
                            # Tombol disabled jika sedang dianalisis atau sudah selesai
                            if st.button(label, key=f"btn_{nama}",
                                         disabled=is_analyzing or is_done,
                                         use_container_width=True):
                                st.session_state.analyzing_set.add(nama)
                                st.rerun()

                st.markdown("---")

                # ── Jalankan analisis untuk yang sedang di-queue (satu per satu) ──
                for row in semua:
                    nama       = row[1] if len(row) > 1 else ""
                    timestamp  = row[0] if len(row) > 0 else ""
                    link_video = row[2] if len(row) > 2 else ""

                    if nama not in st.session_state.analyzing_set:
                        continue
                    if nama in st.session_state.done_set:
                        continue

                    with st.spinner(f"Menganalisis **{nama}** — harap tunggu, jangan tutup tab ini..."):
                        try:
                            creds_ws  = get_google_creds()
                            client_ws = gspread.authorize(creds_ws)
                            sheet_ws  = client_ws.open_by_key(get_spreadsheet_id())

                            # Jika retry: hapus baris error lama dulu
                            if nama in nama_error:
                                try:
                                    hw = sheet_ws.worksheet(TAB_HASIL)
                                    semua_baris = hw.get_all_values()
                                    for idx, baris in enumerate(semua_baris):
                                        if len(baris) > 1 and baris[1] == nama:
                                            hw.delete_rows(idx + 1)
                                            break
                                except Exception:
                                    pass

                            # Pastikan tab Hasil Analisis ada
                            try:
                                hasil_ws = sheet_ws.worksheet(TAB_HASIL)
                            except gspread.WorksheetNotFound:
                                hasil_ws = sheet_ws.add_worksheet(title=TAB_HASIL, rows=1000, cols=20)
                                hasil_ws.append_row(HEADER_HASIL)

                            rek, rata = jalankan_analisis_satu(nama, timestamp, link_video, hasil_ws)
                            st.session_state.done_set.add(nama)
                            st.session_state.analyzing_set.discard(nama)
                            st.cache_data.clear()
                            st.success(f"✅ **{nama}** selesai dianalisis! Rekomendasi: **{rek}** · Rata-rata: **{rata}**")

                        except Exception as e:
                            st.session_state.analyzing_set.discard(nama)
                            st.error(f"❌ Gagal menganalisis **{nama}**: {e}")

                    break  # satu per satu — tunggu selesai baru bisa klik yang lain

                # Tombol refresh setelah semua selesai
                if st.session_state.done_set:
                    st.info("Analisis selesai. Klik tombol di bawah untuk melihat hasil di tab **Hasil Analisis**.")
                    if st.button("🔄 Lihat Hasil Analisis", type="primary"):
                        st.session_state.done_set.clear()
                        st.session_state.analyzing_set.clear()
                        st.cache_data.clear()
                        st.rerun()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        st.warning("⚠️ Terjadi gangguan teknis. Silakan refresh halaman.")
        if st.button("🔄 Refresh Halaman"):
            st.cache_data.clear()
            st.rerun()
