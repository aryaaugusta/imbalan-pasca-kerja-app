import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
import io

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

# ---------------------------------------------------------------------
# PARAMETER REKONSILIASI KEUANGAN & ARUS DANA NKKIP
# ---------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("💰 Parameter Perhitungan Aktuaria")

pbo_awal = st.sidebar.number_input(
    "Nilai Kini Kewajiban Pada Awal Tahun", 
    min_value=0.0, value=0.0, step=1000000.0
)
pbo_akhir = st.sidebar.number_input(
    "Nilai Kini Kewajiban Pada Akhir Tahun", 
    min_value=0.0, value=0.0, step=1000000.0
)
pembayaran_pesangon = st.sidebar.number_input(
    "Pembayaran Pesangon yang Diakui", 
    min_value=0.0, value=0.0, step=1000000.0
)

kelebihan_pembayaran = st.sidebar.number_input(
    "Kelebihan Pembayaran", 
    min_value=0.0, value=0.0, step=1000000.0
)

transfer_masuk_nkkip = st.sidebar.number_input(
    "Transfer Masuk NKKIP", 
    min_value=0.0, value=0.0, step=1000000.0
)

transfer_keluar_nkkip = st.sidebar.number_input(
    "Transfer Keluar NKKIP", 
    min_value=0.0, value=0.0, step=1000000.0, format="%.2f"
)

# =========================================================================
# DEFINISI WINDOW POP-UP DETAIL PUC PER KARYAWAN
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

def normalisasi_kolom_karyawan(df):
    """
    Mendeteksi dan me-rename nama kolom Excel secara dinamis.
    Mengunci kolom 'Aktif' dan 'NIK' pertama yang ditemukan dari kiri tabel.
    """
    df_clean = df.copy()
    
    kolom_nama_target = None
    kolom_nik_target = None
    
    for col in df_clean.columns:
        col_str = str(col).strip().lower()
        
        # 1. Kunci HANYA kolom Nama/Aktif yang PERTAMA KALI ditemukan
        if kolom_nama_target is None:
            if any(kw in col_str for kw in ['aktif', 'nama karyawan', 'nama']):
                kolom_nama_target = col
                # print(f"DEBUG: FIX KOLOM NAMA UTAMA DIKUNCI -> {col}")
                
        # 2. Kunci HANYA kolom NIK/NOPEG yang PERTAMA KALI ditemukan
        if kolom_nik_target is None:
            if any(kw in col_str for kw in ['nik', 'nopeg', 'no.peg', 'nip', 'id karyawan', 'no pegawai']):
                kolom_nik_target = col
                # print(f"DEBUG: FIX KOLOM NIK UTAMA DIKUNCI -> {col}")

    # Lakukan Rename ke format standar internal
    mapping_rename = {}
    if kolom_nama_target:
        mapping_rename[kolom_nama_target] = 'Aktif Tahun Ini'
    if kolom_nik_target:
        mapping_rename[kolom_nik_target] = 'NIK'
        
    df_clean = df_clean.rename(columns=mapping_rename)
    
    # Jika kolom NIK tidak ada di file Excel, siapkan kolom dummy
    if 'NIK' not in df_clean.columns:
        df_clean['NIK'] = "-"
        
    return df_clean

with file_karyawan:
    uploaded_file = st.file_uploader("1. Unggah Berkas Data Karyawan (.xlsx)", type=["xlsx","xls"])
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

with tabel_uuck:
    uploaded_uuck = st.file_uploader("2. Unggah Berkas Template UUCK (.xlsx)", type=["xlsx","xls"])
    if uploaded_uuck is not None:
        try:
            df_uuck_loaded = muat_template_uuck(uploaded_uuck)
            st.success("✅ Tabel Faktor Manfaat UUCK Aktif!")
        except Exception as e:
            st.error(f"Gagal memproses berkas UUCK: {e}")
            df_uuck_loaded = None

with tabel_mortalita:
    uploaded_mortality = st.file_uploader("3. Unggah Berkas Tabel Mortalita (.xlsx)", type=["xlsx","xls"])
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
            df_tm_loaded = None

with tabel_spot_rate:
    uploaded_spot_rate = st.file_uploader("4. Unggah Berkas Tabel IGSYC (.xlsx)", type=["xlsx","xls"])
    if uploaded_spot_rate is not None:
        try:
            df_spot_rate_loaded = muat_tabel_spot_rate(uploaded_spot_rate)
            st.success("✅ Tabel IGSYC 31 Desember 2025 Aktif!")
        except Exception as e:
            st.error(f"Gagal memproses berkas Tabel IGSYC: {e}")

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, skiprows=10)
        # Normalisasi nama kolom secara dinamis (mengubah NOPEG -> NIK, Aktif 2025 -> Aktif Tahun Ini)
        df_raw = normalisasi_kolom_karyawan(df_raw)
        df_ambil_nama_pt = pd.read_excel(uploaded_file, header=None)
        nama_perusahaan = ambil_nama_pt_dari_template(df_ambil_nama_pt)
        # st.success("✅ File Data Karyawan Berhasil Dimuat!")
        # Filter aman: pastikan kolom 'Aktif Tahun Ini' ada sebelum di-dropna
        if 'Aktif Tahun Ini' in df_raw.columns:
            df_aktif = df_raw.dropna(subset=['Aktif Tahun Ini']).copy()
            df_aktif['NIK'] = df_aktif['NIK'].fillna("-").astype(str).str.strip()
            # st.success("✅ File Data Karyawan Berhasil Dimuat & Kolom Disesuaikan Otomatis!")
        else:
            st.error("❌ Kolom Nama Karyawan/Aktif tidak ditemukan. Harap pastikan header tabel mengandung kata 'Aktif' atau 'Nama'.")
            df_aktif = None
    except Exception as e:
        st.error(f"Gagal membaca file Excel Karyawan: {e}")
        st.stop()
else:
    st.warning("⚠️ Menunggu unggahan Berkas Data Karyawan untuk memulai perhitungan.")
    st.stop()

# ==========================================
# PROCESSING CORE ENGINE
# ==========================================
# df_aktif = df_raw.dropna(subset=['NIK', 'Aktif Tahun Ini']).copy()

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

# ---------------------------------------------------------------------
# HITUNG BIAYA BERSIH (TOTAL BEBAN TAHUN BERJALAN)
# ---------------------------------------------------------------------
total_biaya_bersih = total_csc + total_biaya_bunga + kelebihan_pembayaran + transfer_masuk_nkkip - transfer_keluar_nkkip

# 1. Cek Gerbang Validasi: Pastikan Semua Berkas Utama Sudah Diunggah
berkas_lengkap = (
    df_raw is not None and 
    df_uuck_loaded is not None and 
    df_spot_rate_loaded is not None and 
    df_tm_loaded is not None
)

if berkas_lengkap:
    # df_aktif = df_raw.dropna(subset=['NIK', 'Aktif Tahun Ini']).copy()

    # print(f"DATA FRAME AKTIF: {df_aktif}")

    # Filter baris yang nama karyawannya ada (tidak NaN), NIK boleh kosong
    df_aktif = df_raw.dropna(subset=['Aktif Tahun Ini']).copy()

    # Opsional: Rapikan kolom NIK agar jika NaN diubah menjadi string kosong "" atau "-"
    df_aktif['NIK'] = df_aktif['NIK'].fillna("").astype(str).str.strip()

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

    # ---------------------------------------------------------------------
    # HITUNG BIAYA BERSIH (TOTAL BEBAN TAHUN BERJALAN)
    # ---------------------------------------------------------------------
    total_biaya_bersih = total_csc + total_biaya_bunga + kelebihan_pembayaran + transfer_masuk_nkkip - transfer_keluar_nkkip

    # Hitung Keuntungan / Kerugian Aktuaria
    total_pengurang_keuntungan_kerugian = (
        pbo_awal + 
        total_csc + 
        total_biaya_bunga + 
        pembayaran_pesangon + 
        transfer_masuk_nkkip - 
        transfer_keluar_nkkip + 
        kelebihan_pembayaran
    )

    keuntungan_kerugian_aktuaria = total_pbo - total_pengurang_keuntungan_kerugian

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

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric(label="JUMLAH KARYAWAN AKTIF", value=f"{len(df_puc_final)} Jiwa")
    col2.metric(label="TOTAL KEWAJIBAN BERSIH (PBO)", value=f"{int(round(total_pbo)):,}".replace(",", "."))
    col3.metric(label="BIAYA JASA KINI (CSC)", value=f"{int(round(total_csc)):,}".replace(",", "."))
    # col4.metric(label="RERATA TINGKAT DISKONTO", value=f"{rata_rata_bunga_perusahaan * 100:.2f}%")
    col4.metric(label="BIAYA BUNGA (INTEREST COST)", value=f"{total_biaya_bunga:,}".replace(",", "."))
    col5.metric(label="BIAYA BERSIH", value=f"{int(round(total_biaya_bersih)):,}".replace(",", "."))
    col6.metric(label="KEUNTUNGAN / KERUGIAN AKTUARIA", value=f"{int(round(keuntungan_kerugian_aktuaria)):,}".replace(",", "."))

    st.markdown("---")

    # TABEL RINGKASAN ARUS KAS MUTASI KANTOR KONSULTAN AKTUARIA VAB
    st.subheader("📋 Ringkasan Perhitungan Aktuaria")
    df_arus_kas = pd.DataFrame({
        "Komponen": [
            "Nilai Kini Kewajiban Awal Tahun",
            "Pembayaran Pesangon Yang Diakui", 
            "Kelebihan Pembayaran", 
            "Transfer Masuk NKKIP", 
            "Transfer Keluar NKKIP"
        ],
        "Nominal Riil": [
            f"{pbo_awal:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            f"{pembayaran_pesangon:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            f"{kelebihan_pembayaran:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            f"{transfer_masuk_nkkip:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            f"{transfer_keluar_nkkip:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        ]
    })
    st.table(df_arus_kas)

    st.markdown("---")

    # LAYOUT KUSTOM: DAFTAR KARYAWAN DAN TOMBOL POP-UP DETAIL
    st.subheader("📋 Laporan Perhitungan Per Karyawan - 31 Desember 2025")
    st.caption("Klik tombol **🔍 Detail** untuk melihat kalkulasi desimal murni dengan rumus lembar kerja Excel Anda.")

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
        col_nama, col_rate, col_pbo, col_csc, col_aksi = st.columns([3, 1.6, 1.6, 1.6, 1])
        
        with col_nama:
            st.write(f"{idx}. **{row['Nama Karyawan']}**")
        with col_rate:
            st.write(f"{row['Rate Diskonto Murni'] * 100:.2f}%")
        with col_pbo:
            st.write(f"{int(row['Kewajiban Bersih']):,}".replace(",", "."))
        with col_csc:
            st.write(f"{int(row['Biaya Jasa Kini']):,}".replace(",", "."))
        with col_aksi:
            if st.button("🔍 Detail", key=f"btn_puc_{idx}"):
                tampilkan_modal_puc(row)

    st.markdown("---")
    st.subheader("📋 Detail Perhitungan Biaya Bunga per Karyawan (Historis)")
    st.dataframe(hasil_bunga_obj["tabel_bunga"], use_container_width=True)
else:
    st.markdown("---")
    # st.info("👋 **Selamat Datang di Aplikasi Automasi PSAK 219 Kantor Konsultan Aktuaria VAB!**")
    
    st.warning("⚠️ **Perhitungan Belum Dapat Dimulai.** Mohon lengkapi pengunggahan berkas parameter pada tab di atas:")
    
    # Buat checklist status indikator berkas secara visual
    col_status1, col_status2 = st.columns(2)
    
    with col_status1:
        st.write("📁 **1. Data Karyawan:** " + ("🟢 Terunggah" if df_raw is not None else "🔴 Belum Ada"))
        st.write("📘 **2. Data UUCK:** " + ("🟢 Terunggah" if df_uuck_loaded is not None else "🔴 Belum Ada"))
        
    with col_status2:
        st.write("📊 **3. Tabel Mortalita Dinamis:** " + ("🟢 Terunggah" if df_tm_loaded is not None else "🔴 Belum Ada"))
        st.write("📈 **4. Data IGSYC:** " + ("🟢 Terunggah" if df_spot_rate_loaded is not None else "🔴 Belum Ada"))
        
    st.markdown("---")
    st.caption("Setelah seluruh berkas indikator di atas berwarna hijau (🟢), sistem *Core Engine* akan otomatis langsung mengeksekusi perhitungan PBO, CSC, dan perhitungan lainnya *real-time*.")

def generate_excel_output_from_template(template_bytes, context_data):
    """
    Mengisi template Excel (template_output.xlsx) secara presisi berdasarkan 
    hasil kalkulasi aktuaria dinamis dari aplikasi.
    """
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = wb.active
    
    # -------------------------------------------------------------------------
    # 1. BINDER HEADER UTAMA PERUSAHAAN & PERIODE
    # -------------------------------------------------------------------------
    raw_pt = str(context_data.get('nama_perusahaan', 'ARTHA SOLUTIONS INDONESIA')).strip()
    if raw_pt.upper().startswith("PT"):
        nama_pt = raw_pt.upper()
    else:
        nama_pt = f"PT {raw_pt.upper()}"
        
    ws['A5'] = nama_pt  # Nama PT (Baris 5, Kolom A)
    ws['A6'] = f"PER 31 DESEMBER {context_data.get('tahun_val', '2025')}"
    
    # Header Tanggal Kolom C (2024-12-31) dan D (2025-12-31)
    tahun_val = int(context_data.get('tahun_val', 2025))
    ws['C8'] = f"31 Desember 2024"
    ws['D8'] = f"31 Desember 2025"
    ws['C57'] = f"31 Desember 2024"
    ws['D57'] = f"31 Desember 2025"

    # -------------------------------------------------------------------------
    # 2. ASUMSI DAN METODE AKTUARIA
    # -------------------------------------------------------------------------
    ws['D12'] = context_data.get('bunga_diskonto', 0.0)         # Tingkat Diskonto
    ws['D13'] = context_data.get('kenaikan_gaji', 0.0)          # Kenaikan Gaji
    ws['D15'] = context_data.get('tingkat_cacat', 0.0)          # Tingkat Cacat
    ws['D18'] = context_data.get('upn', 60)                     # Usia Pensiun Normal

    # -------------------------------------------------------------------------
    # 3. STATISTIK DATA KARYAWAN
    # -------------------------------------------------------------------------
    rata_usia = context_data.get('rata_usia', 0.0)
    upn_val = context_data.get('upn', 60)
    
    # Formula Dinamis: Rata-Rata Sisa Masa Kerja = UPN - Rata-Rata Usia
    rata_sisa_mk_hitung = max(0.0, upn_val - rata_usia)

    ws['D22'] = context_data.get('jumlah_karyawan', 0)          # Jumlah Karyawan
    ws['D23'] = context_data.get('total_gaji_sebulan', 0.0)      # Total Gaji
    ws['D24'] = rata_usia                                       # Rata-rata Usia
    ws['D25'] = context_data.get('rata_masa_kerja', 0.0)        # Rata-rata Masa Kerja
    ws['D26'] = rata_sisa_mk_hitung                             # Rata-rata Sisa Masa Kerja

    # -------------------------------------------------------------------------
    # 4. REKONSILIASI ARUS DANA & MUTASI
    # -------------------------------------------------------------------------
    pembayaran_pesangon = context_data.get('pembayaran_pesangon', 0.0)
    kelebihan_pembayaran = context_data.get('kelebihan_pembayaran', 0.0)
    transfer_masuk = context_data.get('transfer_masuk_nkkip', 0.0)
    transfer_keluar = context_data.get('transfer_keluar_nkkip', 0.0)

    ws['D28'] = pembayaran_pesangon
    ws['D29'] = kelebihan_pembayaran

    # -------------------------------------------------------------------------
    # 5. NILAI PBO & BIAYA IMBALAN KERJA
    # -------------------------------------------------------------------------
    total_pbo_awal = context_data.get('total_pbo_awal', 0.0)
    total_pbo_akhir = context_data.get('total_pbo_akhir', 0.0)  # Total Kewajiban Bersih (PBO)
    total_csc = context_data.get('total_csc', 0.0)
    total_biaya_bunga = context_data.get('total_biaya_bunga', 0.0)
    
    # Validation PBO Awal
    if total_pbo_awal == 0 or total_pbo_awal is None:
        gain_loss_final = 0.0
    else:
        gain_loss_final = context_data.get('keuntungan_kerugian_aktuaria', 0.0)

    # PERHITUNGAN BIAYA JASA LALU (Cell D72) = D33 - D34 (PBO Akhir - CSC)
    biaya_jasa_lalu_calc = total_pbo_akhir - total_csc

    # Perhitungan Aktuaria Ringkasan Atas
    ws['D32'] = total_pbo_awal
    ws['D33'] = total_pbo_akhir                                 # Nilai Kini Kewajiban Akhir (2025)
    ws['D34'] = total_csc                                       # Biaya Jasa Kini
    ws['D35'] = total_biaya_bunga

    # Ringkasan Kewajiban Bersih
    ws['D61'] = total_pbo_akhir                                 # Nilai Kini Kewajiban Pada Akhir Tahun
    ws['D65'] = total_pbo_akhir                                 # Kewajiban Bersih

    # -------------------------------------------------------------------------
    # 🔥 FIX: RINGKASAN BIAYA BERSIH (P&L) LENGKAP KESELURUHAN (Cell D68 - D79)
    # -------------------------------------------------------------------------
    ws['D68'] = total_csc
    ws['D69'] = total_biaya_bunga
    ws['D72'] = biaya_jasa_lalu_calc                            # Biaya Jasa Lalu = D33 - D34
    ws['D74'] = transfer_masuk
    ws['D75'] = transfer_keluar
    ws['D76'] = kelebihan_pembayaran

    # Total Keseluruhan Nilai di Bagian Biaya Bersih (D79)
    total_biaya_bersih_keseluruhan = (
        total_csc + 
        total_biaya_bunga + 
        biaya_jasa_lalu_calc + 
        transfer_masuk - 
        transfer_keluar + 
        kelebihan_pembayaran
    )
    ws['D79'] = total_biaya_bersih_keseluruhan                  # Hasil total akurat (misal: 199.967.448)

    # -------------------------------------------------------------------------
    # 6. REKONSILIASI KEWAJIBAN
    # -------------------------------------------------------------------------
    ws['D82'] = total_pbo_awal                                  # Nilai Kini Kewajiban awal tahun
    ws['D83'] = total_csc                                       # Biaya Jasa Kini
    ws['D84'] = total_biaya_bunga                               # Biaya Bunga
    ws['D85'] = pembayaran_pesangon                             # Pembayaran Manfaat
    ws['D88'] = biaya_jasa_lalu_calc                            # Biaya Jasa Lalu
    ws['D89'] = transfer_masuk                                  # Transfer Masuk NKKIP
    ws['D90'] = transfer_keluar                                 # Transfer Keluar NKKIP
    ws['D91'] = kelebihan_pembayaran                            # Kelebihan Pembayaran
    ws['D93'] = gain_loss_final                                 # Gain/Loss Aktuaria
    ws['D94'] = total_pbo_akhir                                 # Nilai Kini Kewajiban Akhir (2025)

    # Simpan workbook ke memory buffer
    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)
    
    return output_buffer.getvalue()

# =========================================================================
# BLOK EKSPOR FILE EXCEL LAPORAN LENGKAP
# =========================================================================
st.markdown("---")
st.subheader("📥 Unduh Laporan Valuasi Aktuaria PSAK 219")

# Baca file template_output.xlsx dari folder lokal
try:
    with open("data/file/template_output.xlsx", "rb") as f:
        template_bytes = f.read()

    # Rakit kamus data kontekstual dari hasil kalkulasi aktif
    context_data = {
        'nama_perusahaan': nama_perusahaan if 'nama_perusahaan' in locals() and nama_perusahaan else "ARTHA SOLUTIONS INDONESIA",
        'tahun_val': 2025,
        'bunga_diskonto': rata_rata_bunga_perusahaan,
        'kenaikan_gaji': kenaikan_gaji,
        'tingkat_cacat': tingkat_cacat,
        'upn': upn,
        'jumlah_karyawan': len(df_puc_final),
        'total_gaji_sebulan': pd.to_numeric(df_aktif['Gaji'], errors='coerce').sum() if 'Gaji' in df_aktif.columns else 0.0,
        'rata_usia': df_puc_final['Usia'].mean() if 'Usia' in df_puc_final.columns else 0.0,
        'rata_masa_kerja': df_puc_final['Masa Kerja'].mean() if 'Masa Kerja' in df_puc_final.columns else 0.0,
        
        # Parameter Arus Dana & Mutasi
        'pembayaran_pesangon': pembayaran_pesangon,
        'kelebihan_pembayaran': kelebihan_pembayaran,
        'transfer_masuk_nkkip': transfer_masuk_nkkip,
        'transfer_keluar_nkkip': transfer_keluar_nkkip,
        
        # 🔥 FIX: Ambil PBO Awal & PBO Akhir Murni Hasil Kalkulasi Sistem
        'total_pbo_awal': pbo_awal,
        'total_pbo_akhir': total_pbo,   # <-- Menggunakan total_pbo hasil hitungan PUC murni
        'total_csc': total_csc,
        'total_biaya_bunga': total_biaya_bunga,
        'keuntungan_kerugian_aktuaria': keuntungan_kerugian_aktuaria
    }

    # Generate buffer file excel baru
    excel_data = generate_excel_output_from_template(template_bytes, context_data)

    # Render Tombol Unduh
    st.download_button(
        label="📄 Unduh Laporan Laporan Pengakuan & Pengukuran Excel (.xlsx)",
        data=excel_data,
        file_name=f"Laporan-Aktuaria-PSAK-Atribusi-UUCK-{nama_perusahaan.replace(' ', '_')}-31-Des-2025.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

except FileNotFoundError:
    st.error("⚠️ File 'template_output.xlsx' tidak ditemukan di folder proyek")