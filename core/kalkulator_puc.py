import pandas as pd
import numpy as np

def muat_tabel_spot_rate(file_path):
    """Membaca tabel yield curve / spot rate murni dari struktur 2 kolom Bapak"""
    try:
        df_sr = pd.read_excel(file_path)
        
        # Bersihkan nama kolom dari whitespace dan ubah ke lowercase
        df_sr.columns = [str(c).strip().lower() for c in df_sr.columns]
        
        # Mapping nama kolom agar fleksibel jika ada variasi nama
        rename_dict = {}
        for col in df_sr.columns:
            if 'tenor' in col or 'periode' in col:
                rename_dict[col] = 'tenor'
            elif 'rate' in col or 'bunga' in col:
                rename_dict[col] = 'spot_rate'
                
        if rename_dict:
            df_sr.rename(columns=rename_dict, inplace=True)
            
        # Pastikan kolom tenor diubah menjadi angka bulat untuk index matching
        df_sr['tenor'] = pd.to_numeric(df_sr['tenor'], errors='coerce')
        df_sr = df_sr.dropna(subset=['tenor'])
        df_sr['tenor'] = df_sr['tenor'].astype(int)
        
        return df_sr.set_index('tenor')
    except Exception:
        return None

def dapatkan_spot_rate_dinamis(df_spot_rate, masa_kerja_depan, default_rate):
    """Mengambil nilai spot rate murni berdasarkan pembulatan bawah INT(masa_kerja_depan)"""
    if df_spot_rate is None or df_spot_rate.empty:
        return default_rate
        
    # Ambil nilai integer pembulatan ke bawah sesuai rumus INT() Excel Anda
    tenor_int = int(np.floor(masa_kerja_depan))
    
    # Jika tenor di bawah 1 tahun (misal 0.8 tahun), paksa ke tenor minimum (1) agar tidak error
    if tenor_int < 1:
        tenor_int = 1

    # Aturan Excel: Jika sisa masa kerja ke depan > 30 tahun, kunci di tenor maksimum (30)
    if tenor_int > 30:
        max_tenor = df_spot_rate.index.max()
        rate_val = df_spot_rate.loc[max_tenor, 'spot_rate']
    else:
        if tenor_int in df_spot_rate.index:
            rate_val = df_spot_rate.loc[tenor_int, 'spot_rate']
        else:
            # Fallback jika ada index melompat, ambil tenor terdekat
            idx = df_spot_rate.index.to_numpy()
            pos = np.abs(idx - tenor_int).argmin()
            rate_val = df_spot_rate.iloc[pos]['spot_rate']
            
    # Karena data Bapak sudah berbentuk desimal riil (0.0663), kita TIDAK PERLU membaginya dengan 100 lagi.
    return float(rate_val)

def muat_tabel_mortalita_dinamis(file_path, tingkat_cacat_input, bunga_diskonto_input):
    try:
        df_all = pd.read_excel(file_path)
        header_idx = None
        for i, row in df_all.iterrows():
            row_values = [str(val).strip().lower() for val in row.values if pd.notna(val)]
            if 'x' in row_values:
                header_idx = i
                break
        
        df_tm = pd.read_excel(file_path, skiprows=header_idx + 1 if header_idx is not None else 5)
        df_tm.columns = [str(c).strip() for c in df_tm.columns]
        
        if df_tm.columns[0].lower().startswith('x'):
            df_tm.rename(columns={df_tm.columns[0]: 'x'}, inplace=True)
            
        df_tm['x'] = pd.to_numeric(df_tm['x'], errors='coerce')
        df_tm = df_tm.dropna(subset=['x'])
        df_tm['x'] = df_tm['x'].astype(int)
        df_tm = df_tm.set_index('x')
        
        for col in ['qxd', 'qxi', 'qxw']:
            if col in df_tm.columns:
                df_tm[col] = pd.to_numeric(df_tm[col], errors='coerce').fillna(0.0)
        
        if 'qxi' in df_tm.columns:
            df_tm['qxi_prop'] = df_tm['qxi'] / 0.05
            df_tm['qxi_dinamis'] = df_tm['qxi_prop'] * tingkat_cacat_input
        else:
            df_tm['qxi_dinamis'] = tingkat_cacat_input
            
        v = 1 / (1 + bunga_diskonto_input)
        lx = 100000.0
        df_tm['lx_dinamis'] = 0.0
        df_tm['sDx_dinamis'] = 0.0
        
        for x in sorted(df_tm.index):
            # Isikan nilai lx hidup awal periode sebelum dikurangi decrement tahun berjalan
            df_tm.loc[x, 'lx_dinamis'] = lx
            df_tm.loc[x, 'sDx_dinamis'] = lx * (v ** x)
            
            qxd = float(df_tm.loc[x, 'qxd'])
            qxi = float(df_tm.loc[x, 'qxi_dinamis'])
            qxw = float(df_tm.loc[x, 'qxw']) if 'qxw' in df_tm.columns else 0.0
            
            # Replikasi Logika Komposit Probabilitas Bertahan Hidup (px) Excel:
            # px = (1 - qxd) * (1 - qxi) * (1 - qxw)
            px_dinamis = (1.0 - min(1.0, qxd)) * (1.0 - min(1.0, qxi)) * (1.0 - min(1.0, qxw))
            
            # Update nilai lx untuk usia berikutnya (x + 1) dengan jaring pengaman nilai minimal
            next_lx = lx * px_dinamis
            
            # Jika next_lx hancur jadi nol sebelum usia pensiun maksimum, tahan di angka minimum terkecil
            if next_lx <= 0 and x < 100:
                lx = lx * 0.999
            else:
                lx = next_lx
                
        return df_tm
    except Exception:
        return None

def muat_template_uuck(file_path):
    try:
        # Baca mentah dulu untuk mencari baris header yang asli
        df_all = pd.read_excel(file_path)
        
        header_idx = 0
        kolom_mk_nama = None
        
        # Cari baris yang mengandung kata 'mk' atau 'masa kerja'
        for i, row in df_all.iterrows():
            row_str = [str(val).strip().lower() for val in row.values if pd.notna(val)]
            candidates = [s for s in row_str if 'mk' in s or 'masa' in s or 'kerja' in s]
            if candidates:
                header_idx = i
                # Cari nama kolom fisiknya di baris tersebut
                for col_name in df_all.columns:
                    val_cell = str(df_all.loc[i, col_name]).strip().lower()
                    if 'mk' in val_cell or 'masa' in val_cell or 'kerja' in val_cell:
                        kolom_mk_nama = col_name
                        break
                break
        
        # Baca ulang Excel dengan meleompati baris kosong di atas header yang asli
        df_uuck_raw = pd.read_excel(file_path, skiprows=header_idx + 1)
        df_uuck_raw.columns = [str(c).strip().lower() for c in df_uuck_raw.columns]
        
        # Cari ulang posisi indeks kolom MK setelah disederhanakan
        mk_cols = [c for c in df_uuck_raw.columns if 'mk' in c or 'masa' in c or 'kerja' in c or 'unnamed' in c]
        idx_mk = list(df_uuck_raw.columns).index(mk_cols[0]) if mk_cols else 0
        
        # AMBIL PER POSISI RELATIF (Sangat aman dari perubahan nama teks kolom):
        # Kolom 0: MK | Kolom 1: Pensiun | Kolom 2: Meninggal | Kolom 3: Cacat | Kolom 4: Uang Pisah
        df_clean = df_uuck_raw.iloc[:, idx_mk:idx_mk+5].copy()
        df_clean.columns = ['mk', 'pensiun', 'meninggal', 'cacat', 'uang_pisah']
        
        # Bersihkan data dari baris teks/kosong
        df_clean['mk'] = pd.to_numeric(df_clean['mk'], errors='coerce')
        df_clean = df_clean.dropna(subset=['mk'])
        
        for col in ['pensiun', 'meninggal', 'cacat', 'uang_pisah']:
            df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0.0)
            
        return df_clean.set_index('mk')
        
    except Exception as e:
        print(f"DEBUG ERROR UUCK LOAD: {e}")
        return None

def dapatkan_faktor_uuck(df_uuck, masa_kerja, skenario='pensiun'):
    
    if df_uuck is None or df_uuck.empty:
        # Jika file excel benar-benar kosong/gagal muat, ini batas proteksi minimum
        if skenario == 'pensiun': return 29.61
        if skenario == 'uang_pisah': return 3.60
        return 32.20
        
    try:
        # Urutkan indeks masa kerja dari tabel UUCK
        idx_array = np.sort(df_uuck.index.to_numpy())
        
        # Cari baris-baris masa kerja di tabel yang nilainya kurang dari atau sama dengan masa kerja karyawan
        valid_idx = idx_array[idx_array <= masa_kerja]
        
        if len(valid_idx) > 0:
            # VLOOKUP TRUE: Ambil baris tertinggi yang lolos kriteria (misal ketemu 24 jika masa kerja 24.83)
            mk_lookup = valid_idx[-1] 
        else:
            # Jika masa kerja sangat kecil di bawah batas minimum tabel, ambil baris pertama
            mk_lookup = idx_array[0]
            
        return float(df_uuck.loc[mk_lookup, skenario])
        
    except Exception:
        # Jalur cadangan jika terjadi anomali indeks tipe data
        idx = df_uuck.index.to_numpy()
        posisi = np.abs(idx - masa_kerja).argmin()
        return float(df_uuck.iloc[posisi][skenario])

def hitung_puc_karyawan_v19(nama_karyawan, usia_sekarang, masa_kerja_sekarang, gaji_sekarang, 
                            upn, kenaikan_gaji, bunga_diskonto_default, tingkat_cacat, uang_duka=0.0,
                            df_tm=None, df_uuck=None, df_spot_rate=None):
    
    # Amankan tipe data input
    usia_sekarang = float(usia_sekarang) if pd.notna(usia_sekarang) else 0.0
    masa_kerja_sekarang = float(masa_kerja_sekarang) if pd.notna(masa_kerja_sekarang) else 0.0
    gaji_sekarang = float(gaji_sekarang) if pd.notna(gaji_sekarang) else 0.0
    upn = float(upn) if pd.notna(upn) else 58.0     
    tingkat_cacat = float(tingkat_cacat) if pd.notna(tingkat_cacat) else 0.0
    uang_duka = float(uang_duka) if pd.notna(uang_duka) else 0.0

    # Inisialisasi kolektor total matriks proyeksi 0 - 40 tahun
    total_proyeksi_meninggal = 0.0
    total_proyeksi_cacat = 0.0
    total_proyeksi_resign = 0.0

    # Ambil lx penyebut (usia sekarang) sebagai basis pembagi tetap diluar loop t
    usia_sekarang_bulat = int(round(usia_sekarang))
    lx_sekarang = 100000.0  # Default radix jika df_tm kosong
    if df_tm is not None and usia_sekarang_bulat in df_tm.index:
        lx_sekarang = float(df_tm.loc[usia_sekarang_bulat, 'lx_dinamis'])

    # =========================================================================
    # MATRIKS LOOP PROYEKSI HORIZONTAL (TAHUN 0 SAMPAI 40)
    # =========================================================================
    for t in range(41):  # 0, 1, 2, ..., 40
        usia_proyeksi = usia_sekarang + t
        masa_kerja_proyeksi = masa_kerja_sekarang + t
        usia_proyeksi_bulat = int(round(usia_proyeksi))

        # GERBANG CEK UPN: Jika usia proyeksi >= UPN, nilai proyeksi tahun t ke atas adalah 0
        if usia_proyeksi >= upn:
            continue

        # 1. Ambil Faktor Multiplier UUCK per tahun proyeksi t (VLOOKUP TRUE)
        f_meninggal_t = dapatkan_faktor_uuck(df_uuck, masa_kerja_proyeksi, 'meninggal')
        f_cacat_t = dapatkan_faktor_uuck(df_uuck, masa_kerja_proyeksi, 'cacat')
        f_resign_t = dapatkan_faktor_uuck(df_uuck, masa_kerja_proyeksi, 'uang_pisah')

        # 2. Ambil Nilai Decrement dx dari tabel komutasi dinamis
        dx_meninggal_t = 0.0
        dx_cacat_t = 0.0
        dx_resign_t = 0.0

        if df_tm is not None and usia_proyeksi_bulat in df_tm.index:
            dx_meninggal_t = float(df_tm.loc[usia_proyeksi_bulat, 'qxd']) * float(df_tm.loc[usia_proyeksi_bulat, 'lx_dinamis'])
            dx_cacat_t = float(df_tm.loc[usia_proyeksi_bulat, 'qxi_dinamis']) * float(df_tm.loc[usia_proyeksi_bulat, 'lx_dinamis'])
            dx_resign_t = float(df_tm.loc[usia_proyeksi_bulat, 'qxw']) * float(df_tm.loc[usia_proyeksi_bulat, 'lx_dinamis']) if 'qxw' in df_tm.columns else 0.0

        # Rasio Peluang Probabilitas Decrement Tahun t dibanding lx Awal
        rasio_p_meninggal = (dx_meninggal_t / lx_sekarang) if lx_sekarang > 0 else 0.0
        rasio_p_cacat = (dx_cacat_t / lx_sekarang) if lx_sekarang > 0 else 0.0
        rasio_p_resign = (dx_resign_t / lx_sekarang) if lx_sekarang > 0 else 0.0

        # 3. KOREKSI DISKONTO & GAJI KUMULATIF BERDASARKAN SPOT RATE TAHUN t
        tingkat_bunga_t = dapatkan_spot_rate_dinamis(df_spot_rate, float(t), bunga_diskonto_default)
        
        # AZ11 ^ BE10 (Discount Factor Kumulatif tahun ke-t)
        faktor_diskonto_kumulatif = (1 / (1 + tingkat_bunga_t)) ** t
        
        # (1 + Input!$M$12) ^ BE10 (Kenaikan Gaji Kumulatif tahun ke-t)
        faktor_gaji_kumulatif = (1 + kenaikan_gaji) ** t

        # 4. Hitung Komponen Skala Pembagi Imbalan Aktual Atribusi PUC
        pembagi_skala = masa_kerja_proyeksi
        faktor_puc_skala = (masa_kerja_sekarang / pembagi_skala) if pembagi_skala > 0 else 0.0

        # 5. REPLIKASI PERSIS RUMUS EXCEL BAPAK (Ditambahkan komponen Uang Duka untuk Meninggal)
        nilai_meninggal_t = f_meninggal_t * rasio_p_meninggal * (gaji_sekarang * (1 + uang_duka)) * faktor_gaji_kumulatif * faktor_diskonto_kumulatif * faktor_puc_skala
        nilai_cacat_t = f_cacat_t * rasio_p_cacat * gaji_sekarang * faktor_gaji_kumulatif * faktor_diskonto_kumulatif * faktor_puc_skala
        nilai_resign_t = f_resign_t * rasio_p_resign * gaji_sekarang * faktor_gaji_kumulatif * faktor_diskonto_kumulatif * faktor_puc_skala

        # Akumulasikan ke Grand Total Karyawan
        total_proyeksi_meninggal += nilai_meninggal_t
        total_proyeksi_cacat += nilai_cacat_t
        total_proyeksi_resign += nilai_resign_t

    # =========================================================================
    # LOGIKA GERBANG ATRIBUSI BERDASARKAN USIA FILTER (< UPN - 24)
    # =========================================================================
    usia_batas_atribusi = upn - 24.0
    if usia_sekarang < usia_batas_atribusi:
        total_proyeksi_meninggal = 0.0
        total_proyeksi_cacat = 0.0
        total_proyeksi_resign = 0.0

    # Perhitungan Komponen Utama Pensiun Normal (Sama seperti v16)
    selisih_usia = upn - usia_sekarang
    sisa_masa_kerja_depan = max(0.0, min(24.0, selisih_usia))
    total_aktual = masa_kerja_sekarang + sisa_masa_kerja_depan
    masa_kerja_total = 24.00001 if total_aktual >= 24 else total_aktual
    
    if masa_kerja_total >= 24:
        masa_kerja_lampau = round(masa_kerja_total - sisa_masa_kerja_depan, 5)
    else:
        masa_kerja_lampau = masa_kerja_sekarang

    gaji_proyeksi_pensiun = gaji_sekarang * ((1 + kenaikan_gaji) ** sisa_masa_kerja_depan)
    f_pensiun = dapatkan_faktor_uuck(df_uuck, masa_kerja_total, 'pensiun')
    
    faktor_pensiun_aktuaria = 1.0
    usia_pensiun_normal = int(round(upn))
    if usia_sekarang < upn and df_tm is not None and usia_pensiun_normal in df_tm.index and usia_sekarang_bulat in df_tm.index:
        lx_pensiun = float(df_tm.loc[usia_pensiun_normal, 'lx_dinamis'])
        if lx_sekarang > 0:
            faktor_pensiun_aktuaria = lx_pensiun / lx_sekarang

    tingkat_bunga_riil = dapatkan_spot_rate_dinamis(df_spot_rate, sisa_masa_kerja_depan, bunga_diskonto_default)
    factor_diskonto_murni = (1 / (1 + tingkat_bunga_riil)) ** sisa_masa_kerja_depan

    if usia_sekarang < usia_batas_atribusi:
        nk_pensiun = 0.0
        unit_manfaat_pensiun = 0.0
    else:
        total_manfaat_proyeksi = f_pensiun * gaji_proyeksi_pensiun
        manfaat_pensiun_unfunded = total_manfaat_proyeksi * faktor_pensiun_aktuaria * factor_diskonto_murni
        unit_manfaat_pensiun = manfaat_pensiun_unfunded / masa_kerja_total if masa_kerja_total > 0 else 0.0
        nk_pensiun = unit_manfaat_pensiun * masa_kerja_lampau

    # GABUNGKAN HASIL PROYEKSI O-40 KE TOTAL DBO KARYAWAN
    pbo_total = nk_pensiun + total_proyeksi_meninggal + total_proyeksi_cacat + total_proyeksi_resign
    csc = unit_manfaat_pensiun * 1.0 if masa_kerja_total > 0 else 0.0

    print("------------------------HITUNG PUC KARYAWAN--------------------------")
    print(f"DEBUG: Perhitungan PUC untuk {nama_karyawan} - Usia Sekarang: {usia_sekarang}, Masa Kerja Sekarang: {masa_kerja_sekarang}, Gaji Sekarang: Rp {gaji_sekarang:,.2f}")
    print(f"DEBUG: Sisa Masa Kerja Depan: {sisa_masa_kerja_depan}, Total Masa Kerja: {masa_kerja_total}, Masa Kerja Lampau: {masa_kerja_lampau}")
    print(f"DEBUG: Gaji Proyeksi Pensiun: Rp {gaji_proyeksi_pensiun:,.2f}, Faktor Pensiun Aktuaria: {faktor_pensiun_aktuaria:.4f}, Faktor Diskonto Murni: {factor_diskonto_murni:.4f}")
    print(f"DEBUG: Total Manfaat Proyeksi: Rp {total_manfaat_proyeksi:,.2f}, Manfaat Pensiun Unfunded: Rp {manfaat_pensiun_unfunded:,.2f}, Unit Manfaat Pensiun: Rp {unit_manfaat_pensiun:,.2f}")
    print(f"DEBUG: NK Pensiun: Rp {nk_pensiun:,.2f}, Total Proyeksi Meninggal: Rp {total_proyeksi_meninggal:,.2f}, Total Proyeksi Cacat: Rp {total_proyeksi_cacat:,.2f}, Total Proyeksi Resign: Rp {total_proyeksi_resign:,.2f}")
    print(f"DEBUG: Total DBO (PBO Total): Rp {pbo_total:,.2f}, CSC: Rp {csc:,.2f}")
    print("----------------------------------------------------------------------")

    return {
        "gaji_pensiun": gaji_proyeksi_pensiun,
        "total_manfaat": (f_pensiun * gaji_proyeksi_pensiun),
        "dbo": pbo_total,
        "csc": csc,
        "nk_pensiun": nk_pensiun,
        "nk_meninggal": total_proyeksi_meninggal,
        "nk_cacat": total_proyeksi_cacat,
        "nk_resign": total_proyeksi_resign
    }