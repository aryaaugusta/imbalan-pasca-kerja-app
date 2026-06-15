# app.py
import streamlit as st
import pandas as pd

# Import modul internal hasil pemisahan folder
from core.kalkulator_puc import hitung_puc_karyawan
from data.mock_data import dataset_karyawan

st.set_page_config(page_title="Kalkulator PSAK 219", layout="wide")

st.title("🧮 Prototipe Aplikasi Aktuaria - Imbalan Kerja PSAK 219")
st.subheader("Metode Projected Unit Credit (PUC) - Berbasis Modular")
st.markdown("---")

col_kontrol, col_hasil = st.columns([1, 2])

with col_kontrol:
    st.header("⚙️ Asumsi Aktuaria")
    upn = st.number_input("Usia Pensiun Normal (UPN)", min_value=50, max_value=65, value=55)
    bunga_diskonto = st.slider("Tingkat Bunga Diskonto (Dinamis)", min_value=0.01, max_value=0.15, value=0.06, step=0.005, format="%.3f")
    kenaikan_gaji = st.slider("Estimasi Kenaikan Gaji Tahunan", min_value=0.01, max_value=0.15, value=0.05, step=0.005, format="%.3f")

with col_hasil:
    st.header("📊 Ringkasan Eksekutif Perusahaan")
    
    total_dbo = 0
    total_csc = 0
    rows_hitung = []
    
    for kary in dataset_karyawan:
        res = hitung_puc_karyawan(
            usia_sekarang=kary["Usia"],
            masa_kerja_sekarang=kary["Masa Kerja"],
            gaji_sekarang=kary["Gaji"],
            upn=upn,
            kenaikan_gaji=kenaikan_gaji,
            bunga_diskonto=bunga_diskonto
        )
        total_dbo += res["dbo"]
        total_csc += res["csc"]
        
        rows_hitung.append({
            "NIK": kary["ID"],
            "Nama": kary["Nama"],
            "Usia": kary["Usia"],
            "Masa Kerja": kary["Masa Kerja"],
            "Gaji": f"Rp {kary['Gaji']:,}",
            "Kewajiban (DBO)": f"Rp {round(res['dbo']):,}",
            "Beban Berjalan (CSC)": f"Rp {round(res['csc']):,}"
        })
    
    col_metric1, col_metric2 = st.columns(2)
    col_metric1.metric(label="TOTAL KEWAJIBAN NERACA (DBO)", value=f"Rp {round(total_dbo):,}")
    col_metric2.metric(label="TOTAL BEBAN LABA/RUGI (CSC)", value=f"Rp {round(total_csc):,}")
    
    st.subheader("📋 Rincian Data Per Karyawan")
    st.dataframe(pd.DataFrame(rows_hitung), use_container_width=True)

st.markdown("---")
st.header("📑 Otomasi Draf Jurnal Akuntansi")
data_jurnal = [
    {"Kode Akun": "5.1.02.xx", "Nama Akun": "Beban Imbalan Kerja (Laba/Rugi)", "Posisi": "DEBIT", "Nominal": f"Rp {round(total_csc):,}"},
    {"Kode Akun": "2.1.05.xx", "Nama Akun": "Kewajiban Imbalan Pasti (Neraca)", "Posisi": "KREDIT", "Nominal": f"Rp {round(total_csc):,}"}
]
st.table(pd.DataFrame(data_jurnal))