# app.py
import streamlit as st
import pandas as pd
from datetime import datetime

# Import logika baru dari core engine
from core.kalkulator_puc import (
    hitung_puc_karyawan_v19, 
    muat_tabel_mortalita_dinamis, 
    muat_template_uuck,
    muat_tabel_spot_rate
)

st.set_page_config(page_title="Sistem Aktuaria Terintegrasi PSAK 219", layout="wide")

st.title("🛡️ Sistem Perhitungan Imbalan Kerja Terintegrasi (PSAK 219)")
st.caption("Otomasi Pengolahan & Rekonsiliasi Data Karyawan")
st.markdown("---")

# ==========================================
# SIDEBAR: PARAMETER UTAMA VALUASI
# ==========================================
st.sidebar.header("⚙️ Parameter & Asumsi Aktuaria")
upn = st.sidebar.number_input("Usia Pensiun Normal (UPN)", min_value=50, max_value=65, value=58)

bunga_input = st.sidebar.number_input("Tingkat Bunga Diskonto (%)", min_value=1.0, max_value=15.0, value=6.63, step=0.01, format="%.2f")
bunga_diskonto = bunga_input / 100.0

gaji_input = st.sidebar.number_input("Estimasi Kenaikan Gaji Tahunan (%)", min_value=1.0, max_value=15.0, value=4.00, step=0.01, format="%.2f")
kenaikan_gaji = gaji_input / 100.0

cacat_input = st.sidebar.number_input("Tingkat Cacat (%)", min_value=0.0, max_value=100.0, value=5.00, step=0.01, format="%.2f")
tingkat_cacat = cacat_input / 100.0

tanggal_valuasi = datetime(2025, 12, 31)

# ==========================================
# PANEL UTAMA: MANAJEMEN BERKAS (3 UPLOADER)
# ==========================================
st.header("📂 Manajemen Berkas Valuasi Aktuaria")
file_karyawan, tabel_uuck, tabel_mortalita, tabel_spot_rate = st.columns(4)

df_uuck_loaded = None
df_tm_loaded = None
df_spot_rate_loaded = None

with file_karyawan:
    uploaded_file = st.file_uploader("1. Unggah Berkas Konsolidasi Karyawan (.xlsx)", type=["xlsx"])

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
            df_tm_loaded = muat_tabel_mortalita_dinamis(uploaded_mortality, tingkat_cacat, bunga_diskonto)
            st.success("✅ Tabel Mortalita Dinamis Aktif!")
        except Exception as e:
            st.error(f"Gagal memproses berkas Tabel Mortalita: {e}")

with tabel_spot_rate:
    uploaded_spot_rate = st.file_uploader("4. Unggah Berkas Tabel Spot Rate (.xlsx)", type=["xlsx"])
    if uploaded_spot_rate is not None:
        try:
            df_spot_rate_loaded = muat_tabel_spot_rate(uploaded_spot_rate)
            st.success("✅ Tabel Spot Rate Aktif!")
        except Exception as e:
            st.error(f"Gagal memproses berkas Tabel Spot Rate: {e}")

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, skiprows=10)
        st.success("✅ File Data Karyawan Berhasil Dimuat!")
    except Exception as e:
        st.error(f"Gagal membaca file Excel Karyawan: {e}")
        st.stop()
else:
    st.warning("⚠️ Menunggu unggahan Berkas Konsolidasi Karyawan untuk memulai valuasi.")
    st.stop()

# ==========================================
# PROCESSING CORE ENGINE WITH NEW MULTI-DATA
# ==========================================
df_aktif = df_raw.dropna(subset=['NIK', 'Aktif 2025']).copy()

total_dbo = 0
total_csc = 0
rows_hitung = []
chart_data_list = []

for index, kary in df_aktif.iterrows():
    try:
        nama = str(kary["Aktif 2025"])
        nik = str(kary["NIK"])
        gaji = float(kary["Gaji"])
        
        tgl_lahir = pd.to_datetime(kary["Tgl Lahir"])
        tgl_masuk = pd.to_datetime(kary["Tgl Masuk"])
        
        # Perhitungan presisi pecahan desimal (2 angka di belakang koma)
        usia = (tanggal_valuasi - tgl_lahir).days / 365.25
        masa_kerja = (tanggal_valuasi - tgl_masuk).days / 365.25
        
        usia = max(0.0, round(usia, 2))
        masa_kerja = max(0.0, round(masa_kerja, 2))
        
        # Eksekusi Engine Baru V2 yang menyerap data UUCK dan Mortalita
        res = hitung_puc_karyawan_v19(
            nama_karyawan=nama,
            usia_sekarang=usia,
            masa_kerja_sekarang=masa_kerja,
            gaji_sekarang=gaji,
            upn=upn,
            kenaikan_gaji=kenaikan_gaji,
            bunga_diskonto_default=bunga_diskonto,
            tingkat_cacat=tingkat_cacat,
            df_tm=df_tm_loaded,
            df_uuck=df_uuck_loaded,
            df_spot_rate=df_spot_rate_loaded
        )
        
        total_dbo += res["dbo"]
        total_csc += res["csc"]
        
        rows_hitung.append({
            "NIK": nik,
            "Nama Karyawan": nama,
            "Usia (Thn)": usia,         
            "Masa Kerja (Thn)": masa_kerja, 
            "Gaji": f"Rp {int(gaji):,}",
            "Kewajiban DBO": f"Rp {round(res['dbo']):,}",
            "Beban Berjalan CSC": f"Rp {round(res['csc']):,}"
        })
        
        if len(chart_data_list) < 15:
            chart_data_list.append({
                "Nama": nama,
                "Kewajiban (DBO)": res["dbo"],
                "Beban Berjalan (CSC)": res["csc"]
            })
    except Exception as e:
        print(f"DEBUG: Error occurred while processing {nama} (NIK: {nik}): {e}")
        continue

# ==========================================
# DISPLAY DASBOR & OUTPUT VISUAL
# ==========================================
st.markdown("---")
st.subheader("📊 Hasil Penilaian Aktuaria PSAK 219")

col1, col2, col3 = st.columns(3)
col1.metric(label="JUMLAH KARYAWAN AKTIF DIHITUNG", value=f"{len(rows_hitung)} Jiwa")
col2.metric(label="TOTAL KEWAJIBAN NERACA (DBO)", value=f"Rp {round(total_dbo):,}")
col3.metric(label="TOTAL BEBAN LABA/RUGI (CSC)", value=f"Rp {round(total_csc):,}")

st.subheader("📈 Grafik Perbandingan Komponen Aktuaria (Sampel 15 Karyawan Pertama)")
if chart_data_list:
    df_chart = pd.DataFrame(chart_data_list).set_index("Nama")
    st.bar_chart(df_chart)

st.subheader("📋 Laporan Perhitungan Riil per Karyawan (Siap Ekspor)")
if rows_hitung:
    df_hasil = pd.DataFrame(rows_hitung)
    df_hasil.index = df_hasil.index + 1 
    df_hasil.index.name = "No"  
    st.dataframe(df_hasil, use_container_width=True)

st.subheader("📑 Otomasi Draf Jurnal Penyesuaian (Akrual 2025)")
data_jurnal = [
    {"Kode Akun": "5.1.02.01", "Nama Akun": "Beban Imbalan Kerja Pasca Kerja (Laba/Rugi)", "Posisi": "DEBIT", "Nominal": f"Rp {round(total_csc):,}"},
    {"Kode Akun": "2.1.05.03", "Nama Akun": "Kewajiban Imbalan Pasti / Utang DBO (Neraca)", "Posisi": "KREDIT", "Nominal": f"Rp {round(total_csc):,}"}
]
df_jurnal = pd.DataFrame(data_jurnal)
df_jurnal.index = df_jurnal.index + 1
df_jurnal.index.name = "No"
st.table(df_jurnal)