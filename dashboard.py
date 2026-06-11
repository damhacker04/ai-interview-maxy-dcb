"""
Dashboard Rekrutmen - Digital Career Bootcamp Maxy Academy
Diakses oleh tim HR untuk melihat hasil analisis video interview
"""

import io
import os
from datetime import datetime

import gspread
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials

load_dotenv()

CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")

def get_spreadsheet_id():
    """Baca Spreadsheet ID dari Streamlit Secrets (cloud) atau .env (lokal)."""
    # Coba dari Streamlit Secrets dulu
    try:
        val = st.secrets["SPREADSHEET_ID"]
        if val:
            return val
    except (KeyError, FileNotFoundError):
        pass
    # Fallback ke .env (lokal)
    val = os.getenv("SPREADSHEET_ID")
    if not val:
        st.error("❌ SPREADSHEET_ID belum dikonfigurasi. Tambahkan ke Streamlit Secrets atau file .env.")
        st.stop()
    return val

def get_google_creds():
    """Baca Google credentials dari Streamlit Secrets (cloud) atau file lokal."""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    # Coba dari Streamlit Secrets dulu
    try:
        sa_info = st.secrets["gcp_service_account"]
        private_key = sa_info["private_key"]
        # Handle literal \n (TOML single-quoted string tidak konversi \n otomatis)
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
        # Secrets belum dikonfigurasi → coba file lokal
        pass
    except Exception as e:
        # Secrets ada tapi private_key salah format
        st.error(
            "❌ **Streamlit Secrets ditemukan, tapi `private_key` tidak valid.**\n\n"
            "**Solusi:** Buka **Manage app → Settings → Secrets**, hapus baris `private_key` "
            "yang ada, lalu ganti dengan format multi-baris:\n\n"
            "```toml\n"
            'private_key = """\n'
            "-----BEGIN PRIVATE KEY-----\n"
            "(salin setiap baris key di sini)\n"
            "-----END PRIVATE KEY-----\n"
            '"""\n'
            "```\n\n"
            "Jalankan perintah ini di PowerShell (di folder proyek) untuk melihat "
            "isi key yang benar:\n"
            "```\npython -c \"import json; d=json.load(open('credentials.json')); "
            "print(d['private_key'])\"\n```"
        )
        st.stop()

    # Fallback ke file lokal
    if not os.path.exists(CREDENTIALS_FILE):
        st.error(
            "❌ Google credentials tidak ditemukan.\n\n"
            "Jika kamu di Streamlit Cloud: tambahkan **[gcp_service_account]** di bagian Secrets.\n\n"
            "Jika kamu di lokal: pastikan file `credentials.json` ada di folder proyek."
        )
        st.stop()
    return Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
TAB_FORM         = "Form Responses 1"
TAB_HASIL        = "Hasil Analisis"

KOLOM_SKOR = [
    "Perkenalan Diri (1-5)",
    "Motivasi (1-5)",
    "Tantangan Karier (1-5)",
    "Tujuan Karier (1-5)",
    "Relevansi Program (1-5)",
    "Komitmen (1-5)",
    "Komunikasi (1-5)",
]

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
    /* Header utama */
    .header-box {
        background: linear-gradient(135deg, #0f2d5c 0%, #1a6eb5 100%);
        padding: 2rem 2.5rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1.5rem;
    }
    .header-box h1 { margin: 0; font-size: 1.8rem; }
    .header-box p  { margin: 0.3rem 0 0; opacity: 0.85; font-size: 1rem; }

    /* Badge rekomendasi di tabel */
    .badge-pass   { background:#d4edda; color:#155724; padding:2px 10px;
                    border-radius:20px; font-weight:600; font-size:0.85rem; }
    .badge-review { background:#fff3cd; color:#856404; padding:2px 10px;
                    border-radius:20px; font-weight:600; font-size:0.85rem; }
    .badge-reject { background:#f8d7da; color:#721c24; padding:2px 10px;
                    border-radius:20px; font-weight:600; font-size:0.85rem; }
    .badge-error  { background:#e2e3e5; color:#383d41; padding:2px 10px;
                    border-radius:20px; font-size:0.85rem; }

    /* Kartu detail */
    .detail-card {
        background: #f8f9fa;
        border-left: 4px solid #1a6eb5;
        padding: 1rem 1.2rem;
        border-radius: 0 8px 8px 0;
        margin-bottom: 0.5rem;
    }
    div[data-testid="stMetric"] {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 1px 6px rgba(0,0,0,0.08);
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────
# AMBIL DATA DARI GOOGLE SHEET
# ─────────────────────────────────────────
@st.cache_data(ttl=120)
def load_data():
    creds  = get_google_creds()
    client = gspread.authorize(creds)
    sheet  = client.open_by_key(get_spreadsheet_id())

    # Hitung total pelamar dari tab form
    form_data    = sheet.worksheet(TAB_FORM).get_all_values()
    total_form   = max(0, len(form_data) - 1)

    # Ambil hasil analisis
    try:
        hasil_data = sheet.worksheet(TAB_HASIL).get_all_values()
    except gspread.WorksheetNotFound:
        return pd.DataFrame(), total_form

    if len(hasil_data) <= 1:
        return pd.DataFrame(), total_form

    df = pd.DataFrame(hasil_data[1:], columns=hasil_data[0])

    # Konversi kolom angka
    for col in KOLOM_SKOR + ["Rata-rata"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df, total_form


# ─────────────────────────────────────────
# WARNA BARIS TABEL
# ─────────────────────────────────────────
def style_baris(row):
    rek = row.get("Rekomendasi", "")
    warna = {
        "Pass":   "background-color: #f0fff4",
        "Review": "background-color: #fffdf0",
        "Reject": "background-color: #fff5f5",
        "Error":  "background-color: #f5f5f5",
    }.get(rek, "")
    return [warna] * len(row)


# ─────────────────────────────────────────
# HALAMAN UTAMA
# ─────────────────────────────────────────
def main():

    # ── Header ──────────────────────────────
    st.markdown("""
    <div class="header-box">
        <h1>🎯 DCB Maxy Academy — Video Interview Analyzer</h1>
        <p>Dashboard rekrutmen Digital Career Bootcamp · Hasil analisis AI</p>
    </div>
    """, unsafe_allow_html=True)

    # ── Tombol refresh ───────────────────────
    col_btn, col_waktu = st.columns([1, 5])
    with col_btn:
        if st.button("🔄 Refresh", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    with col_waktu:
        st.caption(f"Data dimuat pada: {datetime.now().strftime('%d %b %Y  %H:%M:%S')}")

    # ── Muat data ────────────────────────────
    with st.spinner("Memuat data dari Google Sheet..."):
        df, total_form = load_data()

    if df.empty:
        st.divider()
        st.info("⏳ Belum ada hasil analisis. Jalankan **pipeline.py** terlebih dahulu.")
        st.metric("Total Pelamar di Form", total_form)
        return

    df_selesai = df[df["Status"] == "Selesai"]
    df_error   = df[df["Status"] == "Error"]

    n_pass   = (df_selesai["Rekomendasi"] == "Pass").sum()
    n_review = (df_selesai["Rekomendasi"] == "Review").sum()
    n_reject = (df_selesai["Rekomendasi"] == "Reject").sum()
    n_error  = len(df_error)
    n_selesai = len(df_selesai)
    n_belum   = total_form - len(df)

    # ── Statistik ────────────────────────────
    st.subheader("📊 Ringkasan")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("👥 Total Pelamar",     total_form)
    c2.metric("✅ Sudah Dianalisis",  n_selesai)
    c3.metric("⏳ Belum Dianalisis",  max(n_belum, 0))
    c4.metric("⚠️ Gagal Diproses",   n_error)

    st.divider()

    # ── Rekomendasi ──────────────────────────
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

    # ── Filter & Tabel ───────────────────────
    st.subheader("📋 Daftar Semua Pelamar")

    fc1, fc2 = st.columns([2, 3])
    with fc1:
        filter_rek = st.selectbox(
            "Filter rekomendasi:",
            ["Semua", "Pass", "Perlu Ditinjau (Review)", "Tidak Lolos (Reject)", "Error"],
        )
    with fc2:
        cari = st.text_input("🔍 Cari nama pelamar:", placeholder="Ketik nama...")

    # Gabung selesai + error untuk tabel
    df_gabung = pd.concat([df_selesai, df_error], ignore_index=True)

    # Terapkan filter rekomendasi
    peta_filter = {
        "Pass":                       "Pass",
        "Perlu Ditinjau (Review)":    "Review",
        "Tidak Lolos (Reject)":       "Reject",
        "Error":                      "Error",
    }
    if filter_rek != "Semua":
        df_gabung = df_gabung[df_gabung["Rekomendasi"] == peta_filter[filter_rek]]

    # Terapkan pencarian nama
    if cari:
        df_gabung = df_gabung[
            df_gabung["Nama Lengkap"].str.lower().str.contains(cari.lower(), na=False)
        ]

    st.caption(f"Menampilkan **{len(df_gabung)}** dari **{len(df)}** pelamar")

    # Kolom yang tampil di tabel
    kolom_tampil = (
        ["Nama Lengkap"] + KOLOM_SKOR + ["Rata-rata", "Rekomendasi"]
    )
    kolom_tampil = [c for c in kolom_tampil if c in df_gabung.columns]

    fmt = {"Rata-rata": "{:.2f}"}
    fmt.update({k: "{:.0f}" for k in KOLOM_SKOR if k in df_gabung.columns})

    styled = (
        df_gabung[kolom_tampil]
        .style
        .apply(style_baris, axis=1)
        .format(fmt, na_rep="–")
    )
    st.dataframe(styled, use_container_width=True, height=420, hide_index=True)

    st.divider()

    # ── Detail Pelamar ───────────────────────
    st.subheader("🔎 Detail Pelamar")

    nama_options = ["— Pilih nama pelamar —"] + df_gabung["Nama Lengkap"].tolist()
    pilihan = st.selectbox("Pilih pelamar:", nama_options, label_visibility="collapsed")

    if pilihan != "— Pilih nama pelamar —":
        row = df_gabung[df_gabung["Nama Lengkap"] == pilihan].iloc[0]

        rek   = row.get("Rekomendasi", "–")
        rata  = row.get("Rata-rata", "–")
        emoji = {"Pass": "✅", "Review": "🔍", "Reject": "❌"}.get(rek, "⚠️")

        col_kiri, col_kanan = st.columns([1, 1])

        with col_kiri:
            st.markdown(f"**Nama:** {row.get('Nama Lengkap', '–')}")
            st.markdown(f"**Tanggal Submit:** {row.get('Timestamp', '–')}")
            st.markdown(f"**Rekomendasi:** {emoji} **{rek}**")
            st.markdown(f"**Rata-rata Skor:** {rata}")
            st.markdown(f"**Dianalisis:** {row.get('Waktu Diproses', '–')}")
            st.markdown("**Ringkasan AI:**")
            ringkasan = row.get("Ringkasan", "–")
            if "Error" in str(ringkasan):
                st.error(ringkasan)
            else:
                st.info(ringkasan)

        with col_kanan:
            st.markdown("**Skor per Kriteria:**")
            skor_df = pd.DataFrame({
                "Kriteria": [
                    "Perkenalan Diri", "Motivasi", "Tantangan Karier",
                    "Tujuan Karier", "Relevansi Program", "Komitmen", "Komunikasi"
                ],
                "Skor": [
                    row.get("Perkenalan Diri (1-5)", "–"),
                    row.get("Motivasi (1-5)", "–"),
                    row.get("Tantangan Karier (1-5)", "–"),
                    row.get("Tujuan Karier (1-5)", "–"),
                    row.get("Relevansi Program (1-5)", "–"),
                    row.get("Komitmen (1-5)", "–"),
                    row.get("Komunikasi (1-5)", "–"),
                ],
            })

            def warna_skor(val):
                try:
                    v = float(val)
                    if v >= 4:  return "background-color: #d4edda; color: #155724"
                    if v >= 3:  return "background-color: #fff3cd; color: #856404"
                    return              "background-color: #f8d7da; color: #721c24"
                except Exception:
                    return ""

            st.dataframe(
                skor_df.style.map(warna_skor, subset=["Skor"]),
                hide_index=True,
                use_container_width=True,
                height=284,
            )

    st.divider()

    # ── Export Excel ─────────────────────────
    st.subheader("📥 Export Data")

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_gabung.to_excel(writer, index=False, sheet_name="Hasil Analisis")
    nama_file = f"DCB_Hasil_Analisis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    st.download_button(
        label="⬇️ Download Excel (data yang sedang ditampilkan)",
        data=buf.getvalue(),
        file_name=nama_file,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )
    st.caption("File Excel hanya berisi data yang tampil di tabel (sesuai filter aktif).")


if __name__ == "__main__":
    main()
