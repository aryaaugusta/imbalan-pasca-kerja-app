import pandas as pd
import numpy as np

def muat_tabel_spot_rate(file_path):
    """Membaca tabel yield curve / spot rate murni dari struktur 2 kolom"""
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
        
    tenor_int = int(np.floor(masa_kerja_depan))
    
    if tenor_int < 1:
        tenor_int = 1

    if tenor_int > 30:
        max_tenor = df_spot_rate.index.max()
        rate_val = df_spot_rate.loc[max_tenor, 'spot_rate']
    else:
        if tenor_int in df_spot_rate.index:
            rate_val = df_spot_rate.loc[tenor_int, 'spot_rate']
        else:
            idx = df_spot_rate.index.to_numpy()
            pos = np.abs(idx - tenor_int).argmin()
            rate_val = df_spot_rate.iloc[pos]['spot_rate']
            
    # =========================================================================
    # JARING PENGAMAN: Jika rate_val berupa Series (karena tenor duplikat di Excel),
    # ambil baris pertama saja (.iloc[0]) agar tipenya kembali menjadi skalar/angka tunggal.
    # =========================================================================
    if isinstance(rate_val, pd.Series):
        rate_val = rate_val.iloc[0]
        
    return float(rate_val) / 100.0

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
    """
    Membaca file Excel UUCK secara spesifik mendeteksi header 'Pensiun', 'Meninggal', 
    'Cacat', dan 'Uang Pisah' pada sub-row agar tidak salah potong kolom.
    """
    try:
        df_raw = pd.read_excel(file_path, header=None)
        
        baris_header_idx = None
        # posisi_kolom_mk = None
        
        # 1. Cari baris yang mengandung header internal riil (Pensiun, Meninggal, Cacat, Uang Pisah)
        for idx, row in df_raw.iterrows():
            row_values = [str(val).strip().lower() for val in row.values if pd.notna(val)]
            if 'pensiun' in row_values and 'uang pisah' in row_values:
                baris_header_idx = idx
                break
                
        # Jika baris sub-header tidak ketemu, cari baris kata kunci 'mk (masa kerja)'
        if baris_header_idx is None:
            for idx, row in df_raw.iterrows():
                row_values = [str(val).strip().lower() for val in row.values if pd.notna(val)]
                for c_idx, cell_val in enumerate(row.values):
                    cell_str = str(cell_val).strip().lower()
                    if 'mk' in cell_str or 'masa kerja' in cell_str:
                        baris_header_idx = idx
                        # posisi_kolom_mk = c_idx
                        break
                if baris_header_idx is not None:
                    break

        # 2. Temukan indeks koordinat kolom masing-masing label secara dinamis pada baris header tersebut
        header_row = df_raw.iloc[baris_header_idx].values
        header_strings = [str(h).strip().lower() for h in header_row]
        
        # Cari posisi index kolom masing-masing (Replikasi VLOOKUP Kolom ke-6, 7, 8)
        col_idx_mk = [i for i, s in enumerate(header_strings) if 'mk' in s or 'masa kerja' in s or 'unnamed' in s][0]
        
        # Cari indeks kolom spesifik hasil akhir (mengabaikan kolom akumulasi pesangon/pmk di depan)
        col_idx_pensiun = [i for i, s in enumerate(header_strings) if 'pensiun' in s][-1]
        col_idx_meninggal = [i for i, s in enumerate(header_strings) if 'meninggal' in s][-1]
        col_idx_cacat = [i for i, s in enumerate(header_strings) if 'cacat' in s][-1]
        col_idx_resign = [i for i, s in enumerate(header_strings) if 'uang pisah' in s or 'pisah' in s][-1]
        
        # 3. Ekstrak data murni dari baris header ke bawah
        df_data = df_raw.iloc[baris_header_idx + 1:].copy()
        
        df_clean = pd.DataFrame()
        df_clean['mk'] = pd.to_numeric(df_data.iloc[:, col_idx_mk], errors='coerce')
        df_clean['pensiun'] = pd.to_numeric(df_data.iloc[:, col_idx_pensiun], errors='coerce')
        df_clean['meninggal'] = pd.to_numeric(df_data.iloc[:, col_idx_meninggal], errors='coerce')
        df_clean['cacat'] = pd.to_numeric(df_data.iloc[:, col_idx_cacat], errors='coerce')
        df_clean['uang_pisah'] = pd.to_numeric(df_data.iloc[:, col_idx_resign], errors='coerce')
        
        # Bersihkan baris kosong atau baris total di bawah tabel
        df_clean = df_clean.dropna(subset=['mk'])
        
        # Mengisi nilai kosong dengan 0.0
        for col in ['pensiun', 'meninggal', 'cacat', 'uang_pisah']:
            df_clean[col] = df_clean[col].fillna(0.0)
            
        # Set indeks berbasis Masa Kerja murni
        df_uuck_final = df_clean.set_index('mk')
        return df_uuck_final
        
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
    usia_sekarang = float(usia_sekarang) if pd.notna(usia_sekarang) else 0.00
    masa_kerja_sekarang = float(masa_kerja_sekarang) if pd.notna(masa_kerja_sekarang) else 0.00
    gaji_sekarang = float(gaji_sekarang) if pd.notna(gaji_sekarang) else 0.00
    upn = float(upn) if pd.notna(upn) else 58.0
    tingkat_cacat = float(tingkat_cacat) if pd.notna(tingkat_cacat) else 0.00
    uang_duka = float(uang_duka) if pd.notna(uang_duka) else 0.00

    # Inisialisasi kolektor total matriks proyeksi 0 - 40 tahun
    total_proyeksi_meninggal = 0.00
    total_proyeksi_cacat = 0.00
    total_proyeksi_resign = 0.00
    masa_kerja_proyeksi = 0.00
    masa_kerja_proyeksi_sesudah = 0.00

    if masa_kerja_sekarang < 1:
        masa_kerja_proyeksi_sesudah = 1
    else:
        masa_kerja_proyeksi_sesudah = masa_kerja_sekarang

    # Ambil lx penyebut (usia sekarang) sebagai basis pembagi tetap diluar loop t
    usia_sekarang_bulat = round(usia_sekarang, 1)
    usia_sekarang_mortalita = int(round(usia_sekarang))
    lx_sekarang = 100000.0  # Default radix jika df_tm kosong
    
    if df_tm is not None:
        # Sinkronisasi nama kolom lx murni milik (bisa 'lx', 'lx_dinamis', dll)
        # Cari lx untuk usia sekarang sebagai basis pembagi tetap
        if usia_sekarang_mortalita in df_tm.index:
            # Jika ada kolom murni lx hidup gunakan itu, jika tidak ada default ke 100000 atau total radix awal
            if 'lx' in df_tm.columns:
                lx_sekarang = float(df_tm.loc[usia_sekarang_mortalita, 'lx'])
            elif 'lx_dinamis' in df_tm.columns:
                lx_sekarang = float(df_tm.loc[usia_sekarang_mortalita, 'lx_dinamis'])

    selisih_usia = upn - usia_sekarang_bulat
    sisa_masa_kerja_depan = max(0.0, min(24.0, selisih_usia))            

    tingkat_bunga_t = dapatkan_spot_rate_dinamis(df_spot_rate, sisa_masa_kerja_depan, bunga_diskonto_default)
    faktor_diskonto = (1 / (1 + tingkat_bunga_t))

    # =========================================================================
    # MATRIKS LOOP PROYEKSI HORIZONTAL (TAHUN 0 SAMPAI 40)
    # =========================================================================
    detail_proyeksi = []

    for t in range(41):  # 0, 1, 2, ..., 40
        usia_proyeksi = usia_sekarang_bulat + t
        masa_kerja_proyeksi = masa_kerja_proyeksi_sesudah + t
        usia_proyeksi_bulat = int(round(usia_proyeksi))

        # GERBANG CEK UPN: Jika usia proyeksi >= UPN, nilai proyeksi tahun t ke atas adalah 0
        if usia_proyeksi >= upn:
            continue

        # 1. Ambil Faktor Multiplier UUCK per tahun proyeksi t (VLOOKUP TRUE)
        f_meninggal_t = dapatkan_faktor_uuck(df_uuck, masa_kerja_proyeksi, 'meninggal')
        f_cacat_t = dapatkan_faktor_uuck(df_uuck, masa_kerja_proyeksi, 'cacat')
        f_resign_t = dapatkan_faktor_uuck(df_uuck, masa_kerja_proyeksi, 'uang_pisah')

        # 2. Ambil Nilai Decrement dx dari tabel komutasi dinamis
        dx_meninggal_val = 0.00
        dx_cacat_val = 0.00
        dx_resign_val = 0.00

        if df_tm is not None and usia_proyeksi_bulat in df_tm.index:
            # Ambil nilai murni dari kolom 'dx' atau 'dx_meninggal' sesuai struktur file
            if 'dx' in df_tm.columns:
                dx_meninggal_val = float(df_tm.loc[usia_proyeksi_bulat, 'dx'])
            elif 'dx_meninggal' in df_tm.columns:
                dx_meninggal_val = float(df_tm.loc[usia_proyeksi_bulat, 'dx_meninggal'])
            # Ambil nilai murni dari kolom 'ix' atau 'ix_cacat' sesuai struktur file
            if 'ix' in df_tm.columns:
                dx_cacat_val = float(df_tm.loc[usia_proyeksi_bulat, 'ix'])
            elif 'ix_cacat' in df_tm.columns:
                dx_cacat_val = float(df_tm.loc[usia_proyeksi_bulat, 'ix_cacat'])
            # Ambil nilai murni dari kolom 'wx' atau 'wx_resign' sesuai struktur file
            if 'wx' in df_tm.columns:
                dx_resign_val = float(df_tm.loc[usia_proyeksi_bulat, 'wx'])
            elif 'wx_resign' in df_tm.columns:
                dx_resign_val = float(df_tm.loc[usia_proyeksi_bulat, 'wx_resign'])

        # Rasio Peluang Probabilitas Decrement Tahun t dibanding lx Awal
        rasio_p_meninggal = (dx_meninggal_val / lx_sekarang) if lx_sekarang > 0 else 0.00
        rasio_p_cacat = (dx_cacat_val / lx_sekarang) if lx_sekarang > 0 else 0.00
        rasio_p_resign = (dx_resign_val / lx_sekarang) if lx_sekarang > 0 else 0.00

        # 3. KOREKSI DISKONTO & GAJI KUMULATIF BERDASARKAN SPOT RATE TAHUN t
        # tingkat_bunga_t = dapatkan_spot_rate_dinamis(df_spot_rate, float(t), bunga_diskonto_default)
        tingkat_bunga_t = dapatkan_spot_rate_dinamis(df_spot_rate, sisa_masa_kerja_depan, bunga_diskonto_default)
        
        # AZ11 ^ BE10 (Discount Factor Kumulatif tahun ke-t)
        # faktor_diskonto = (1 / (1 + tingkat_bunga_t))
        faktor_diskonto_kumulatif = (faktor_diskonto) ** t
        
        # (1 + Input!$M$12) ^ BE10 (Kenaikan Gaji Kumulatif tahun ke-t)
        faktor_gaji_kumulatif = (1 + kenaikan_gaji) ** t

        # KHUSUS RESIGN: Pangkatnya dimulai/digeser dari 1 (menggunakan t + 1)
        # Sehingga saat t=0 (tahun berjalan), diskonto resign sudah berpangkat 1
        faktor_diskonto_kumulatif_resign = faktor_diskonto ** (t + 1)

        # 4. Hitung Komponen Skala Pembagi Imbalan Aktual Atribusi PUC
        pembagi_skala = masa_kerja_proyeksi

        hitung_nilai_meninggal = round(f_meninggal_t, 1) * rasio_p_meninggal * (gaji_sekarang * (1 + uang_duka)) * faktor_gaji_kumulatif * faktor_diskonto_kumulatif
        hitung_nilai_cacat = round(f_cacat_t, 1) * rasio_p_cacat * (gaji_sekarang * (1 + uang_duka)) * faktor_gaji_kumulatif * faktor_diskonto_kumulatif
        hitung_nilai_resign = round(f_resign_t, 1) * rasio_p_resign * gaji_sekarang * (1 + uang_duka) * faktor_gaji_kumulatif * faktor_diskonto_kumulatif_resign

        # 5. REPLIKASI RUMUS EXCEL (Ditambahkan komponen Uang Duka untuk Meninggal)
        nilai_meninggal_t_pembagi = (hitung_nilai_meninggal / pembagi_skala)
        nilai_cacat_t_pembagi = (hitung_nilai_cacat / pembagi_skala)
        nilai_resign_t_pembagi = (hitung_nilai_resign / pembagi_skala)

        nilai_meninggal_t = nilai_meninggal_t_pembagi * masa_kerja_proyeksi_sesudah
        nilai_cacat_t = nilai_cacat_t_pembagi * masa_kerja_proyeksi_sesudah
        nilai_resign_t = nilai_resign_t_pembagi * masa_kerja_proyeksi_sesudah

        # Akumulasikan ke Grand Total Karyawan
        total_proyeksi_meninggal += nilai_meninggal_t
        total_proyeksi_cacat += nilai_cacat_t
        total_proyeksi_resign += nilai_resign_t

        detail_proyeksi.append({
            "Tahun ke": t,
            "Usia Proyeksi": usia_proyeksi_bulat,
            "Masa Kerja Proyeksi": round(masa_kerja_proyeksi, 2),
            "Faktor UUCK Meninggal": round(f_meninggal_t, 2),
            "Faktor UUCK Cacat": round(f_cacat_t, 2),
            "Faktor UUCK Resign": round(f_resign_t, 2),
            "Rasio Meninggal": rasio_p_meninggal,
            "Rasio Cacat": rasio_p_cacat,
            "Rasio Resign": rasio_p_resign,
            "Diskonto": tingkat_bunga_t,
            "Faktor Diskonto": faktor_diskonto_kumulatif,
            "Faktor Gaji": faktor_gaji_kumulatif,
            "Nilai Meninggal": nilai_meninggal_t,
            "Nilai Cacat": nilai_cacat_t,
            "Nilai Resign": nilai_resign_t,
            "Total Proyeksi Meninggal": total_proyeksi_meninggal,
            "Total Proyeksi Cacat": total_proyeksi_cacat,
            "Total Proyeksi Resign": total_proyeksi_resign
        })

        print("------------------------START PROYEKSI MANFAAT--------------------------")
        print(f"DEBUG {nama_karyawan} | PROYEKSI TAHUN {t} - Usia: {usia_proyeksi_bulat:.2f} | Masa Kerja Skg: {masa_kerja_sekarang:.2f} | Masa Kerja Proyeksi Sesudah: {masa_kerja_proyeksi_sesudah:.2f}")
        print(f"       Faktor UUCK Meninggal: {round(f_meninggal_t, 1)} | Rasio Meninggal: {rasio_p_meninggal}")
        print(f"       Faktor UUCK Cacat: {round(f_cacat_t, 1)} | Rasio Cacat: {rasio_p_cacat}")
        print(f"       Faktor UUCK Resign: {round(f_resign_t, 1)} | Rasio Resign: {rasio_p_resign}")
        print(f"       Diskonto: {tingkat_bunga_t:.4f} | Faktor Diskonto Proyeksi: {faktor_diskonto:.4f} | Faktor Diskonto Kumulatif: {faktor_diskonto_kumulatif:.4f} | Faktor Diskonto Kumulatif Resign: {faktor_diskonto_kumulatif_resign:.4f}")
        print(f"       Pembagi Skala: {pembagi_skala} | Faktor Gaji Kumulatif: {faktor_gaji_kumulatif}")
        print(f"       Gaji Skg: {gaji_sekarang}")
        print(f"       Hitung Proyeksi Meninggal: {int(round(hitung_nilai_meninggal))}")
        print(f"       Nilai Proyeksi Meninggal Bagi Skala: {nilai_meninggal_t_pembagi}")
        print(f"       Nilai Proyeksi Meninggal: {nilai_meninggal_t}")
        print(f"       Total Nilai Proyeksi Meninggal: {total_proyeksi_meninggal}")
        print(f"       Hitung Proyeksi Cacat: {int(round(hitung_nilai_cacat))}")
        print(f"       Nilai Proyeksi Cacat Bagi Skala: {nilai_cacat_t_pembagi}")
        print(f"       Nilai Proyeksi Cacat: {nilai_cacat_t}")
        print(f"       Total Nilai Proyeksi Cacat: {total_proyeksi_cacat}")
        print(f"       Hitung Proyeksi Resign: {int(round(hitung_nilai_resign))}")
        print(f"       Nilai Proyeksi Resign Bagi Skala: {nilai_resign_t_pembagi}")
        print(f"       Nilai Proyeksi Resign: {nilai_resign_t}")
        print(f"       Total Nilai Proyeksi Resign: {total_proyeksi_resign}")
        print("------------------------------------------------------------------------")

    # =========================================================================
    # LOGIKA ATRIBUSI BERDASARKAN USIA FILTER (< UPN - 24)
    # =========================================================================
    usia_batas_atribusi = upn - 24.0    
    total_aktual = masa_kerja_sekarang + sisa_masa_kerja_depan
    masa_kerja_total = 24.00001 if total_aktual >= 24 else total_aktual
    
    if masa_kerja_total >= 24:
        masa_kerja_lampau = round(masa_kerja_total - sisa_masa_kerja_depan, 5)
    else:
        masa_kerja_lampau = masa_kerja_sekarang

    gaji_proyeksi_pensiun = gaji_sekarang * ((1 + kenaikan_gaji) ** sisa_masa_kerja_depan)
    f_pensiun = dapatkan_faktor_uuck(df_uuck, masa_kerja_total, 'pensiun')
    
    faktor_pensiun_aktuaria = 0.0
    usia_pensiun_normal = int(round(upn))
    if usia_sekarang_bulat >= upn:
        faktor_pensiun_aktuaria = 1.0
    else:
        lx_pensiun = float(df_tm.loc[usia_pensiun_normal, 'lx'])
        if lx_sekarang > 0:
            faktor_pensiun_aktuaria = lx_pensiun / lx_sekarang

    tingkat_bunga_riil = dapatkan_spot_rate_dinamis(df_spot_rate, sisa_masa_kerja_depan, bunga_diskonto_default)
    factor_diskonto_murni = (1 / (1 + tingkat_bunga_riil)) ** sisa_masa_kerja_depan

    nk_pensiun = 0.0
    nk_meninggal = 0.0
    nk_cacat = 0.0
    nk_resign = 0.0
    unit_manfaat_pensiun = 0.0
    manfaat_pensiun_unfunded = 0.0
    csc_final = 0.0

    total_manfaat_proyeksi = f_pensiun * gaji_proyeksi_pensiun

    if usia_sekarang_bulat < usia_batas_atribusi:
        # A. KOMPONEN UTAMA: PENSIUN NORMAL (Pembagi tetap menggunakan masa_kerja_total)
        manfaat_pensiun_unfunded = total_manfaat_proyeksi * faktor_pensiun_aktuaria * factor_diskonto_murni
        if masa_kerja_total == 24.00001:
            unit_manfaat_pensiun = 0.0
        else:
            unit_manfaat_pensiun = manfaat_pensiun_unfunded / masa_kerja_total if masa_kerja_total > 0 else 0.0
        nk_pensiun = unit_manfaat_pensiun * masa_kerja_lampau
    else:
        manfaat_pensiun_unfunded = total_manfaat_proyeksi * faktor_pensiun_aktuaria * factor_diskonto_murni
        unit_manfaat_pensiun = manfaat_pensiun_unfunded / masa_kerja_total if masa_kerja_total > 0 else 0.0
        nk_pensiun = unit_manfaat_pensiun * masa_kerja_lampau

    # Jika masa kerja sekarang < 1, pembagi = 1, jika tidak pembagi = masa_kerja_sekarang
    pembagi_risiko = 1.0 if masa_kerja_sekarang < 1.0 else masa_kerja_sekarang

    # B. KOMPONEN DECREMENT: MENINGGAL
    unit_manfaat_meninggal = total_proyeksi_meninggal / pembagi_risiko
    nk_meninggal = total_proyeksi_meninggal

    # C. KOMPONEN DECREMENT: CACAT
    unit_manfaat_cacat = total_proyeksi_cacat / pembagi_risiko
    nk_cacat = total_proyeksi_cacat

    # D. KOMPONEN DECREMENT: MENGUNDURKAN DIRI (RESIGN)
    unit_manfaat_resign = total_proyeksi_resign / pembagi_risiko
    nk_resign = total_proyeksi_resign

    # GABUNGKAN HASIL PROYEKSI O-40 KE TOTAL DBO KARYAWAN
    pbo_total = nk_pensiun + nk_meninggal + nk_cacat + nk_resign
    csc = unit_manfaat_pensiun + unit_manfaat_meninggal + unit_manfaat_cacat + unit_manfaat_resign
    
    if masa_kerja_sekarang < 1:
        csc_final = pbo_total / 2
    else:
        csc_final = csc

    # print("------------------------HITUNG PUC KARYAWAN--------------------------")
    # print(f"Usia Pensiun Normal: {usia_pensiun_normal}")
    # print(f"DEBUG: {nama_karyawan} | Usia: {usia_sekarang_bulat} | Gaji: {gaji_sekarang}")
    # # print(f"       LX PENSIUN: {lx_pensiun} | LX SKG: {lx_sekarang}")
    # print(f"       Tingkat Diskonto Spot Rate: {tingkat_bunga_riil} | Faktor Pensiun: {faktor_pensiun_aktuaria} | Faktor Diskonto Murni: {factor_diskonto_murni}")
    # print(f"       Masa Kerja: {masa_kerja_sekarang} | Masa Kerja Total: {masa_kerja_total} | Masa Kerja Ke Depan: {sisa_masa_kerja_depan} | Total Aktual: {total_aktual} | Masa Kerja Lampau: {masa_kerja_lampau}")
    # print(f"       Manfaat Pensiun belum didanai: {manfaat_pensiun_unfunded}")
    # print(f"       Faktor UUCK Pensiun: {f_pensiun}")
    # print(f"       Faktor UUCK Meninggal: {f_meninggal_t} | Rasio Meninggal: {rasio_p_meninggal}")
    # print(f"       Faktor UUCK Cacat: {f_cacat_t} | Rasio Cacat: {rasio_p_cacat}")
    # print(f"       Faktor UUCK Resign: {f_resign_t} | Rasio Resign: {rasio_p_resign}")
    # print(f"       Detail Proyeksi - Meninggal: {total_proyeksi_meninggal} | Cacat: {total_proyeksi_cacat} | Resign: {total_proyeksi_resign}")
    # print(f"       Proyeksi Upah: {gaji_proyeksi_pensiun} | Proyeksi UUK-13/2003: {total_manfaat_proyeksi}")
    # print(f"       Unit Manfaat Pensiun: {unit_manfaat_pensiun} | Unit Manfaat Meninggal: {unit_manfaat_meninggal} | Unit Manfaat Cacat: {unit_manfaat_cacat} | Unit Manfaat Resign: {unit_manfaat_resign}")
    # print(f"       NK Pensiun: {nk_pensiun} | NK Meninggal: {nk_meninggal} | NK Cacat: {nk_cacat} | NK Resign: {nk_resign}")
    # print(f"       PBO: {pbo_total} | CSC: {csc_final}")
    # print("---------------------------------------------------------------------")

    return {
        "gaji_pensiun": int(round(gaji_proyeksi_pensiun)),
        "total_manfaat": int(round(total_manfaat_proyeksi)),
        "pbo": int(round(pbo_total)),
        "csc_final": int(round(csc_final)),
        "nk_pensiun": int(round(total_proyeksi_resign)),
        "nk_meninggal": int(round(total_proyeksi_meninggal)),
        "nk_cacat": int(round(total_proyeksi_cacat)),
        "nk_resign": int(round(total_proyeksi_resign)),
        "detail_proyeksi": detail_proyeksi
    }

def muat_data_karyawan_dari_template(file_path):
    """
    Membaca data murni karyawan dari file upload awal, 
    termasuk jangkar data historis untuk Biaya Bunga.
    """
    try:
        df = pd.read_excel(file_path)
        
        # Bersihkan nama kolom agar seragam (lowercase & tanpa spasi)
        df.columns = [str(c).strip().lower() for c in df.columns]
        
        # Mapping nama kolom agar fleksibel jika ada variasi penamaan di Excel Bapak
        rename_dict = {}
        for col in df.columns:
            if 'nama' in col:
                rename_dict[col] = 'nama_karyawan'
            elif 'pbo_lalu' in col or 'pbo_2024' in col or 'pbo tahun lalu' in col:
                rename_dict[col] = 'pbo_tahun_lalu'
            elif 'rate_lalu' in col or 'diskonto_2024' in col or 'bunga tahun lalu' in col:
                rename_dict[col] = 'rate_diskonto_tahun_lalu'
                
        if rename_dict:
            df.rename(columns=rename_dict, inplace=True)
            
        # Pastikan tipe data finansial aman
        df['pbo_tahun_lalu'] = pd.to_numeric(df['pbo_tahun_lalu'], errors='coerce').fillna(0.0)
        df['rate_diskonto_tahun_lalu'] = pd.to_numeric(df['rate_diskonto_tahun_lalu'], errors='coerce').fillna(0.0)
        
        # JARING PENGAMAN: Jika di Excel rate ditulis bulat (misal 6.5), otomatis bagi 100
        df['rate_diskonto_tahun_lalu'] = df['rate_diskonto_tahun_lalu'].apply(
            lambda x: x / 100.0 if x > 1.0 else x
        )
        
        return df
    except Exception as e:
        print(f"Error membaca template karyawan: {e}")
        return None
    
def hitung_biaya_bunga_dari_template(df_karyawan, label_tahun_lalu):
    """
    Menghitung Biaya Bunga secara dinamis berdasarkan posisi kolom template terbaru:
    - Kolom C (Indeks 2): Nama Karyawan Aktif Tahun Ini
    - Kolom J (Indeks 9): Nama Karyawan Aktif Tahun Lalu
    - Kolom N (Indeks 13): PBO Tahun Lalu
    - Kolom O (Indeks 14): Diskonto Tahun Lalu
    
    Format output pada DataFrame langsung disesuaikan untuk kebutuhan UI (String Formatted).
    """
    df_src = df_karyawan.copy()
    
    # 1. Cari baris header utama (baris yang mengandung NIK atau AKTIF TAHUN INI)
    header_idx = 0
    for idx, row in df_src.iterrows():
        row_str = [str(x).strip().upper() for x in row.values]
        if 'NIK' in row_str or 'AKTIF TAHUN INI' in row_str or 'AKTIF' in str(row_str):
            header_idx = idx
            break
            
    # Potong data mengambil baris di bawah header agar murni data karyawan
    df_data = df_src.iloc[header_idx:].reset_index(drop=True)

    # 2. PROSES TABEL KARYAWAN AKTIF TAHUN INI (KOLOM C / INDEKS 2)
    df_aktif = pd.DataFrame()
    df_aktif['Nama_Kini'] = df_data.iloc[:, 2].astype(str).str.strip().str.upper()
    # Bersihkan dari baris kosong atau baris total/judul di bawah
    df_aktif = df_aktif[df_aktif['Nama_Kini'].notna() & ~df_aktif['Nama_Kini'].isin(['', 'NAN', 'NONE', 'TOTAL'])]

    # 3. PROSES TABEL HISTORIS TAHUN LALU (KOLOM J, N, O / INDEKS 9, 13, 14)
    df_lalu = pd.DataFrame()
    df_lalu['Nama_Lalu'] = df_data.iloc[:, 9].astype(str).str.strip().str.upper()
    df_lalu['pbo_raw'] = df_data.iloc[:, 13]
    df_lalu['rate_raw'] = df_data.iloc[:, 14]
    
    # Bersihkan data tahun lalu
    df_lalu = df_lalu[df_lalu['Nama_Lalu'].notna() & ~df_lalu['Nama_Lalu'].isin(['', 'NAN', 'NONE'])]
    df_lalu['pbo_clean'] = pd.to_numeric(df_lalu['pbo_raw'], errors='coerce').fillna(0.0)
    df_lalu['rate_clean'] = pd.to_numeric(df_lalu['rate_raw'], errors='coerce').fillna(0.0)
    
    # Hanya pertahankan data historis yang memiliki nilai PBO > 0
    df_lalu = df_lalu[df_lalu['pbo_clean'] > 0]

    # 4. REPLIKASI VLOOKUP KETAT (Strict Inner Join Berdasarkan Nama)
    df_matched = pd.merge(
        df_aktif,
        df_lalu[['Nama_Lalu', 'pbo_clean', 'rate_clean']],
        left_on='Nama_Kini',
        right_on='Nama_Lalu',
        how='inner'
    )
    
    # Jaring pengaman persentase tingkat diskonto (misal 7.05 menjadi 0.0705)
    df_matched['rate_clean'] = df_matched['rate_clean'].apply(
        lambda x: x / 100.0 if x > 1.0 else x
    )
    
    # 5. Hitung Biaya Bunga Akhir per Individu
    df_matched['biaya_bunga_mentah'] = df_matched['pbo_clean'] * df_matched['rate_clean']
    df_matched['Biaya Bunga'] = df_matched['biaya_bunga_mentah'].round().astype(int)

    # Simpan nilai total numerik sebelum kolom diubah menjadi teks berformat
    grand_total_bunga = df_matched['Biaya Bunga'].sum()

    # 6. Susun Output Akhir & Terapkan Formatter Finansial untuk Layar UI
    df_output = pd.DataFrame()
    df_output['Nama Karyawan'] = df_matched['Nama_Kini']
    
    # Format kolom PBO dengan pemisah ribuan titik
    df_output[f'PBO ({label_tahun_lalu})'] = df_matched['pbo_clean'].round().astype(int).apply(lambda x: f"{x:,.0f}".replace(",", "."))
    
    # Format kolom Rate Diskonto menjadi bentuk persentase (Contoh: 7.05%)
    df_output[f'Rate Diskonto ({label_tahun_lalu})'] = df_matched['rate_clean'].apply(lambda x: f"{x * 100:.2f}%")
    
    # Format kolom Biaya Bunga dengan pemisah ribuan titik
    df_output['Biaya Bunga'] = df_matched['Biaya Bunga'].apply(lambda x: f"{x:,.0f}".replace(",", "."))

    # Set indeks nomor urut rapi dari angka 1
    df_output = df_output.reset_index(drop=True)
    df_output.index = df_output.index + 1
    df_output.index.name = "No"

    return {
        "tabel_bunga": df_output,
        "total_bunga": int(grand_total_bunga)
    }

def ambil_nama_pt_dari_template(df_karyawan):
    """
    Mengambil nama PT secara dinamis dari baris pertama Kolom F (F1) di Excel.
    Jika kosong atau tidak terbaca, akan mengembalikan nilai Default 'Klien PT'
    """
    try:
        # Kolom F di Excel = Indeks ke-5 di Pandas
        # Baris 1 di Excel = Indeks ke-0 di Pandas
        nama_pt = str(df_karyawan.iloc[0, 5]).strip()
        
        # Jaring pengaman jika terbaca 'nan' atau kosong
        if nama_pt.lower() in ['nan', 'none', '']:
            return "PT XXXXX"
            
        return nama_pt
    except Exception:
        return "PT XXXXX"