# app.py
import streamlit as st
import pandas as pd
from datetime import datetime

# Import logika baru dari core engine
from core.kalkulator_puc import (
    # hitung_puc_karyawan_v19, 
    muat_tabel_mortalita_dinamis, 
    muat_template_uuck,
    muat_tabel_spot_rate,
    hitung_biaya_bunga_dari_template,
    ambil_nama_pt_dari_template,
    proses_puc_seluruh_karyawan
)

st.set_page_config(page_title="Sistem Aktuaria Terintegrasi PSAK 219", layout="wide")

st.title("🛡️ Sistem Perhitungan Imbalan Kerja Terintegrasi (PSAK 219)")
st.caption("Otomasi Pengolahan & Rekonsiliasi Data Karyawan")
st.markdown("---")

# ==========================================
# SIDEBAR: PARAMETER UTAMA VALUASI
# ==========================================
st.sidebar.header("⚙️ Parameter & Asumsi Aktuaria")
upn = st.sidebar.number_input("Usia Pensiun Normal (UPN)", min_value=50, max_value=65, value=60)

# bunga_input = st.sidebar.number_input("Tingkat Diskonto (%)", min_value=1.0, max_value=15.0, value=6.63, step=0.01, format="%.2f")
# bunga_diskonto = bunga_input / 100.0
tingkat_diskonto_default = 0

gaji_input = st.sidebar.number_input("Estimasi Kenaikan Gaji Tahunan (%)", min_value=1.0, max_value=15.0, value=4.00, step=0.01, format="%.2f")
kenaikan_gaji = gaji_input / 100.0

cacat_input = st.sidebar.number_input("Tingkat Cacat (%)", min_value=0.0, max_value=100.0, value=5.00, step=0.01, format="%.2f")
tingkat_cacat = cacat_input / 100.0

tanggal_valuasi = datetime(2025, 12, 31)
tahun_berjalan = tanggal_valuasi.year
tahun_lalu_label = str(tahun_berjalan - 1)  # Otomatis mendeteksi "2024" jika valuasi 2025

# =========================================================================
# [BARU] DEFINISI WINDOW POP-UP DETAIL PUC PER KARYAWAN
# =========================================================================
@st.dialog("🔍 Detail Rumus PUC Aktuaria", width="large")
def tampilkan_modal_puc(row_karyawan):
    st.write(f"### Karyawan: **{row_karyawan['Nama Karyawan']}**")
    st.caption("Berikut adalah breakdown komponen formula PUC murni untuk dibandingkan dengan Excel:")

    # Fungsi pembantu internal untuk memformat mata uang Rupiah dengan pemisah titik
    def format_rupiah(angka):
        try:
            # Format desimal murni dua angka di belakang koma, ribuan dengan titik
            return f"{angka:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return f"{angka}"
    
    data_komparasi = {
        "Komponen PUC": [
            "Usia Saat Penilaian", "Usia Pensiun (UPN)", "Masa Kerja Aktual (Tahun)", "Masa Kerja ke Depan", "Gaji Terakhir",
            "Gaji Proyeksi Pensiun", "Total Manfaat Proyeksi",
            "Nilai Kini Kewajiban (PBO / Kewajiban Bersih)", "Biaya Jasa Kini (Current Service Cost)",
            "NK Pensiun", "NK Meninggal", "NK Cacat", "NK Resign"
        ],
        "Nilai di Aplikasi": [
            f"{row_karyawan['Usia']:.2f} Tahun",
            f"{int(row_karyawan['Usia Pensiun'])} Tahun",
            f"{row_karyawan['Masa Kerja']:.2f} Tahun",
            f"{row_karyawan['Masa Kerja ke Depan']:.2f} Tahun",
            format_rupiah(row_karyawan['Gaji Terakhir']),
            format_rupiah(row_karyawan['Gaji Proyeksi']),
            format_rupiah(row_karyawan['Total Manfaat Proyeksi']),
            format_rupiah(row_karyawan['Kewajiban Bersih']),
            format_rupiah(row_karyawan['Biaya Jasa Kini']),
            format_rupiah(row_karyawan['NK Pensiun']),
            format_rupiah(row_karyawan['NK Meninggal']),
            format_rupiah(row_karyawan['NK Cacat']),
            format_rupiah(row_karyawan['NK Resign'])
        ]
    }
    st.table(pd.DataFrame(data_komparasi))

# ==========================================
# PANEL UTAMA: MANAJEMEN BERKAS (4 UPLOADER)
# ==========================================
st.header("📂 Manajemen Berkas Valuasi Aktuaria")
file_karyawan, tabel_uuck, tabel_mortalita, tabel_spot_rate = st.columns(4)

df_uuck_loaded = None
df_tm_loaded = None
df_spot_rate_loaded = None

with file_karyawan:
    uploaded_file = st.file_uploader("1. Unggah Berkas Data Karyawan (.xlsx)", type=["xlsx"])

with tabel_uuck:
    uploaded_uuck = st.file_uploader("2. Unggah Berkas Template UUCK (.xlsx)", type=["xlsx"])
    if uploaded_uuck is not None:
        try:
            df_uuck_loaded = muat_template_uuck(uploaded_uuck)
            st.success("✅ Tabel Faktor Manfaat UUCK Aktif!")
        except Exception as e:
            st.error(f"Gagal memproses berkas UUCK: {e}")

with tabel_mortalita:
    uploaded_mortality = st.file_uploader("3. Unggah Berkas Tabel Mortalita (.xlsx)", type=["xlsx"])
    if uploaded_mortality is not None:
        try:
            df_tm_loaded = muat_tabel_mortalita_dinamis(uploaded_mortality, 
            tingkat_cacat_input=tingkat_cacat, 
            tingkat_diskonto_input=tingkat_diskonto_default,
            upn_input=upn,
            kenaikan_gaji_input=kenaikan_gaji)
            st.success("✅ Tabel Mortalita Dinamis Aktif!")
        except Exception as e:
            st.error(f"Gagal memproses berkas Tabel Mortalita: {e}")

with tabel_spot_rate:
    uploaded_spot_rate = st.file_uploader("4. Unggah Berkas Tabel IGSYC (.xlsx)", type=["xlsx"])
    if uploaded_spot_rate is not None:
        try:
            df_spot_rate_loaded = muat_tabel_spot_rate(uploaded_spot_rate)
            st.success("✅ Tabel IGSYC 31 Desember 2025 Aktif!")
        except Exception as e:
            st.error(f"Gagal memproses berkas Tabel IGSYC: {e}")

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, skiprows=10)
        df_mentah = pd.read_excel(uploaded_file, header=None)
        nama_perusahaan = ambil_nama_pt_dari_template(df_mentah)
        st.success("✅ File Data Karyawan Berhasil Dimuat!")
    except Exception as e:
        st.error(f"Gagal membaca file Excel Karyawan: {e}")
        st.stop()
else:
    st.warning("⚠️ Menunggu unggahan Berkas Data Karyawan untuk memulai perhitungan.")
    st.stop()

# ==========================================
# PROCESSING CORE ENGINE (CENTRALIZED DATAFRAME)
# ==========================================
df_aktif = df_raw.dropna(subset=['NIK', 'Aktif Tahun Ini']).copy()

with st.spinner("Menghitung matriks PUC komparatif seluruh karyawan..."):
    df_puc_final = proses_puc_seluruh_karyawan(
        df_aktif=df_aktif,
        df_tm=df_tm_loaded,
        df_uuck=df_uuck_loaded,
        df_spot_rate=df_spot_rate_loaded,
        kenaikan_gaji=kenaikan_gaji,
        tingkat_diskonto_default=tingkat_diskonto_default,
        tingkat_cacat=tingkat_cacat,
        upn=upn,
        uang_duka=0.0
    )

# Eksekusi kalkulasi Biaya Bunga
hasil_bunga_obj = hitung_biaya_bunga_dari_template(df_aktif, tahun_lalu_label)
total_biaya_bunga = hasil_bunga_obj["total_bunga"]

# Hitung Agregat Nilai Akhir Perusahaan
total_pbo = df_puc_final['Kewajiban Bersih'].sum()
total_csc = df_puc_final['Biaya Jasa Kini'].sum()
rata_rata_bunga_perusahaan = df_puc_final['Rate Diskonto Murni'].mean()

# =========================================================================
# TAMBAH WIDGET TINGKAT DISKONTO KE SIDEBAR SECARA DINAMIS SETELAH DIHITUNG
# =========================================================================
st.sidebar.markdown("---")
st.sidebar.subheader("📈 Hasil Output Tingkat Diskonto")
st.sidebar.metric(
    label="Tingkat Diskonto", 
    value=f"{rata_rata_bunga_perusahaan * 100:.2f}%",
    help="Dihitung secara otomatis dari rata-rata tingkat diskonto riil seluruh karyawan aktif"
)

# ==========================================
# DISPLAY DASBOR & OUTPUT VISUAL
# ==========================================
st.markdown("---")
st.subheader(f"📊 Hasil Penilaian Aktuaria PSAK 219 - **{nama_perusahaan}**")

col1, col2, col3, col4 = st.columns(4)
col1.metric(label="JUMLAH KARYAWAN AKTIF", value=f"{len(df_puc_final)} Jiwa")
col2.metric(label="TOTAL KEWAJIBAN BERSIH (PBO)", value=f"Rp {int(round(total_pbo)):,}".replace(",", "."))
col3.metric(label="BIAYA JASA KINI (CSC)", value=f"Rp {int(round(total_csc)):,}".replace(",", "."))
# col4.metric(label="RERATA TINGKAT DISKONTO", value=f"{rata_rata_bunga_perusahaan * 100:.2f}%")
col4.metric(label="BIAYA BUNGA (INTEREST COST)", value=f"Rp {total_biaya_bunga:,}".replace(",", "."))

st.markdown("---")

# LAYOUT KUSTOM: DAFTAR KARYAWAN DAN TOMBOL POP-UP DETAIL
st.subheader("📋 Laporan Perhitungan Per Karyawan")
st.caption("Klik tombol **🔍 Detail** untuk memverifikasi kalkulasi desimal murni dengan rumus lembar kerja Excel Anda.")

# Render Header Row
col_h1, col_h2, col_h3, col_h4, col_h5 = st.columns([3, 1.5, 2, 2, 1])
col_h1.markdown("**Nama Karyawan**")
col_h2.markdown("**Rate Diskonto**")
col_h3.markdown("**Kewajiban Bersih (PBO)**")
col_h4.markdown("**Biaya Jasa Kini (CSC)**")
col_h5.markdown("**Aksi**")
st.markdown("---")

# Render Body Row (Indeks otomatis dari 1 sesuai return core engine)
for idx, row in df_puc_final.iterrows():
    col_nama, col_rate, col_pbo, col_csc, col_aksi = st.columns([3, 1.5, 2, 2, 1])
    
    with col_nama:
        st.write(f"{idx}. **{row['Nama Karyawan']}**")
    with col_rate:
        st.write(f"{row['Rate Diskonto Murni'] * 100:.2f}%")
    with col_pbo:
        st.write(f"Rp {int(row['Kewajiban Bersih']):,}".replace(",", "."))
    with col_csc:
        st.write(f"Rp {int(row['Biaya Jasa Kini']):,}".replace(",", "."))
    with col_aksi:
        if st.button("🔍 Detail", key=f"btn_puc_{idx}"):
            tampilkan_modal_puc(row)

st.markdown("---")
st.subheader("📋 Detail Perhitungan Biaya Bunga per Karyawan (Historis)")
st.dataframe(hasil_bunga_obj["tabel_bunga"], use_container_width=True)