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
            
    # Karena data sudah berbentuk desimal riil (0.0663), kita TIDAK PERLU membaginya dengan 100 lagi.
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
    usia_sekarang_bulat = int(round(usia_sekarang))
    lx_sekarang = 100000.0  # Default radix jika df_tm kosong
    
    if df_tm is not None:
        # Sinkronisasi nama kolom lx murni milik (bisa 'lx', 'lx_dinamis', dll)
        # Cari lx untuk usia sekarang sebagai basis pembagi tetap
        if usia_sekarang_bulat in df_tm.index:
            # Jika ada kolom murni lx hidup gunakan itu, jika tidak ada default ke 100000 atau total radix awal
            if 'lx' in df_tm.columns:
                lx_sekarang = float(df_tm.loc[usia_sekarang_bulat, 'lx'])
            elif 'lx_dinamis' in df_tm.columns:
                lx_sekarang = float(df_tm.loc[usia_sekarang_bulat, 'lx_dinamis'])

    # =========================================================================
    # MATRIKS LOOP PROYEKSI HORIZONTAL (TAHUN 0 SAMPAI 40)
    # =========================================================================
    for t in range(41):  # 0, 1, 2, ..., 40
        usia_proyeksi = usia_sekarang + t
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
            
            # Sesuaikan juga untuk cacat dan resign jika menggunakan kolom komutasi serupa
            # dx_cacat_val = float(df_tm.loc[usia_proyeksi_bulat, 'dx_cacat']) if 'dx_cacat' in df_tm.columns else 0.00
            # dx_resign_val = float(df_tm.loc[usia_proyeksi_bulat, 'dx_resign']) if 'dx_resign' in df_tm.columns else 0.00

        # Rasio Peluang Probabilitas Decrement Tahun t dibanding lx Awal
        rasio_p_meninggal = (dx_meninggal_val / lx_sekarang) if lx_sekarang > 0 else 0.00
        rasio_p_cacat = (dx_cacat_val / lx_sekarang) if lx_sekarang > 0 else 0.00
        rasio_p_resign = (dx_resign_val / lx_sekarang) if lx_sekarang > 0 else 0.00

        # 3. KOREKSI DISKONTO & GAJI KUMULATIF BERDASARKAN SPOT RATE TAHUN t
        tingkat_bunga_t = dapatkan_spot_rate_dinamis(df_spot_rate, float(t), bunga_diskonto_default)
        
        # AZ11 ^ BE10 (Discount Factor Kumulatif tahun ke-t)
        faktor_diskonto = (1 / (1 + tingkat_bunga_t))
        faktor_diskonto_kumulatif = (faktor_diskonto) ** t
        
        # (1 + Input!$M$12) ^ BE10 (Kenaikan Gaji Kumulatif tahun ke-t)
        faktor_gaji_kumulatif = (1 + kenaikan_gaji) ** t

        # KHUSUS RESIGN: Pangkatnya dimulai/digeser dari 1 (menggunakan t + 1)
        # Sehingga saat t=0 (tahun berjalan), diskonto resign sudah berpangkat 1
        faktor_diskonto_kumulatif_resign = faktor_diskonto ** (t + 1)

        # 4. Hitung Komponen Skala Pembagi Imbalan Aktual Atribusi PUC
        pembagi_skala = masa_kerja_proyeksi
        faktor_puc_skala = round((masa_kerja_sekarang / pembagi_skala), 2) if pembagi_skala > 0 else 0.00

        hitung_uang_duka = gaji_sekarang * (1 + uang_duka)

        hitung_nilai_meninggal = round(f_meninggal_t * rasio_p_meninggal * (gaji_sekarang * (1 + uang_duka)) * faktor_gaji_kumulatif * faktor_diskonto_kumulatif, 2)
        hitung_nilai_cacat = round(f_cacat_t * rasio_p_cacat * (gaji_sekarang * (1 + uang_duka)) * faktor_gaji_kumulatif * faktor_diskonto_kumulatif, 2)
        hitung_nilai_resign_1 = f_resign_t * rasio_p_resign * hitung_uang_duka
        hitung_nilai_resign = hitung_nilai_resign_1 * faktor_gaji_kumulatif * faktor_diskonto_kumulatif_resign

        # 5. REPLIKASI PERSIS RUMUS EXCEL (Ditambahkan komponen Uang Duka untuk Meninggal)
        nilai_meninggal_t = round((hitung_nilai_meninggal / pembagi_skala) * masa_kerja_proyeksi_sesudah, 2)
        nilai_cacat_t = round((hitung_nilai_cacat / pembagi_skala) * masa_kerja_proyeksi_sesudah, 2)
        nilai_resign_t = round((hitung_nilai_resign / pembagi_skala) * masa_kerja_proyeksi_sesudah, 2)

        # Akumulasikan ke Grand Total Karyawan
        total_proyeksi_meninggal += nilai_meninggal_t
        total_proyeksi_cacat += nilai_cacat_t
        total_proyeksi_resign += nilai_resign_t

        print("------------------------START PROYEKSI MANFAAT--------------------------")
        print(f"DEBUG {nama_karyawan} | PROYEKSI TAHUN {t} - Usia: {usia_proyeksi:.2f} | Masa Kerja Skg: {masa_kerja_sekarang:.2f} | Masa Kerja Proyeksi Sesudah: {masa_kerja_proyeksi_sesudah:.2f}")
        # print(f"       DX MENINGGAL VAL: {dx_meninggal_val} | LX SKG: {lx_sekarang}")
        print(f"       Faktor UUCK Meninggal: {f_meninggal_t} | Rasio Meninggal: {rasio_p_meninggal}")
        print(f"       Faktor UUCK Cacat: {f_cacat_t} | Rasio Cacat: {rasio_p_cacat}")
        print(f"       Faktor UUCK Resign: {f_resign_t} | Rasio Resign: {rasio_p_resign}")
        print(f"       Diskonto: {tingkat_bunga_t:.4f} | Faktor Diskonto Proyeksi: {faktor_diskonto:.4f} | Faktor Diskonto Kumulatif: {faktor_diskonto_kumulatif:.4f} | Faktor Diskonto Kumulatif Resign: {faktor_diskonto_kumulatif_resign:.4f}")
        print(f"       Pembagi Skala: {pembagi_skala} | Faktor Gaji Kumulatif: {faktor_gaji_kumulatif} | Faktor PUC Skala: {faktor_puc_skala}")
        print(f"       Gaji Skg: {gaji_sekarang} | Hitung Uang Duka: {hitung_uang_duka}")
        print(f"       Hitung Proyeksi Meninggal: {hitung_nilai_meninggal}")
        print(f"       Nilai Proyeksi Meninggal: {nilai_meninggal_t}")
        print(f"       Total Nilai Proyeksi Meninggal: {total_proyeksi_meninggal}")
        print(f"       Hitung Proyeksi Cacat: {hitung_nilai_cacat}")
        print(f"       Nilai Proyeksi Cacat: {nilai_cacat_t}")
        print(f"       Total Nilai Proyeksi Cacat: {total_proyeksi_cacat}")
        print(f"       Hitung Proyeksi Resign 1: {hitung_nilai_resign_1}")
        print(f"       Hitung Proyeksi Resign 2: {hitung_nilai_resign}")
        print(f"       Nilai Proyeksi Resign: {nilai_resign_t}")
        print(f"       Total Nilai Proyeksi Resign: {total_proyeksi_resign}")
        print("------------------------------------------------------------------------")

    # =========================================================================
    # LOGIKA ATRIBUSI BERDASARKAN USIA FILTER (< UPN - 24)
    # =========================================================================
    usia_batas_atribusi = upn - 24.0
    # print(f"       Usia Sekarang: {usia_sekarang} | Usia Batas Atribusi: {usia_batas_atribusi}")
    # if usia_sekarang < usia_batas_atribusi:
    #     total_proyeksi_meninggal = 0.0
    #     total_proyeksi_cacat = 0.0
    #     total_proyeksi_resign = 0.0
    
    # print(f"       Detail Proyeksi Masuk Logika Atribusi - Meninggal: {total_proyeksi_meninggal} | Cacat: {total_proyeksi_cacat} | Resign: {total_proyeksi_resign}")

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

    nk_pensiun = 0.0
    nk_meninggal = 0.0
    nk_cacat = 0.0
    nk_resign = 0.0
    unit_manfaat_pensiun = 0.0

    if usia_sekarang < usia_batas_atribusi:
        # A. KOMPONEN UTAMA: PENSIUN NORMAL (Pembagi tetap menggunakan masa_kerja_total)
        total_manfaat_proyeksi = f_pensiun * gaji_proyeksi_pensiun
        manfaat_pensiun_unfunded = total_manfaat_proyeksi * faktor_pensiun_aktuaria * factor_diskonto_murni
        if masa_kerja_total == 24.00001:
            unit_manfaat_pensiun = 0.0
        else:
            unit_manfaat_pensiun = manfaat_pensiun_unfunded / masa_kerja_total if masa_kerja_total > 0 else 0.0
        nk_pensiun = unit_manfaat_pensiun * masa_kerja_lampau

        # KOREKSI UTAMA: PENENTUAN SKALA PEMBAGI RISIKO TAMBAHAN
        # Jika masa kerja sekarang < 1, pembagi = 1, jika tidak pembagi = masa_kerja_sekarang
        pembagi_risiko = 1.0 if masa_kerja_sekarang < 1.0 else masa_kerja_sekarang

        # B. KOMPONEN DECREMENT: MENINGGAL
        # manfaat_meninggal_unfunded = (f_meninggal_t * gaji_proyeksi_pensiun) * qxd_sekarang * factor_diskonto_murni
        unit_manfaat_meninggal = total_proyeksi_meninggal / pembagi_risiko
        nk_meninggal = total_proyeksi_meninggal

        # C. KOMPONEN DECREMENT: CACAT
        # manfaat_cacat_unfunded = (f_cacat_t * gaji_proyeksi_pensiun) * qxi_sekarang * factor_diskonto_murni
        unit_manfaat_cacat = total_proyeksi_cacat / pembagi_risiko
        nk_cacat = total_proyeksi_cacat

        # D. KOMPONEN DECREMENT: MENGUNDURKAN DIRI (RESIGN)
        # manfaat_resign_unfunded = (f_resign_t * gaji_proyeksi_pensiun) * qxw_sekarang * factor_diskonto_murni
        unit_manfaat_resign = total_proyeksi_resign / pembagi_risiko
        nk_resign = total_proyeksi_resign
    else:
        # KOREKSI UTAMA: PENENTUAN SKALA PEMBAGI RISIKO TAMBAHAN
        # Jika masa kerja sekarang < 1, pembagi = 1, jika tidak pembagi = masa_kerja_sekarang
        pembagi_risiko = 1.0 if masa_kerja_sekarang < 1.0 else masa_kerja_sekarang

        # B. KOMPONEN DECREMENT: MENINGGAL
        # manfaat_meninggal_unfunded = (f_meninggal_t * gaji_proyeksi_pensiun) * qxd_sekarang * factor_diskonto_murni
        unit_manfaat_meninggal = total_proyeksi_meninggal / pembagi_risiko
        nk_meninggal = total_proyeksi_meninggal

        # C. KOMPONEN DECREMENT: CACAT
        # manfaat_cacat_unfunded = (f_cacat_t * gaji_proyeksi_pensiun) * qxi_sekarang * factor_diskonto_murni
        unit_manfaat_cacat = total_proyeksi_cacat / pembagi_risiko
        nk_cacat = total_proyeksi_cacat

        # D. KOMPONEN DECREMENT: MENGUNDURKAN DIRI (RESIGN)
        # manfaat_resign_unfunded = (f_resign_t * gaji_proyeksi_pensiun) * qxw_sekarang * factor_diskonto_murni
        unit_manfaat_resign = total_proyeksi_resign / pembagi_risiko
        nk_resign = total_proyeksi_resign

    # GABUNGKAN HASIL PROYEKSI O-40 KE TOTAL DBO KARYAWAN
    pbo_total = round(nk_pensiun + nk_meninggal + nk_cacat + nk_resign, 2)
    csc = round(unit_manfaat_pensiun + unit_manfaat_meninggal + unit_manfaat_cacat + unit_manfaat_resign, 2)

    print("------------------------HITUNG PUC KARYAWAN--------------------------")
    print(f"DEBUG: {nama_karyawan} | Usia: {usia_sekarang} | Masa Kerja: {masa_kerja_sekarang} | Gaji: {gaji_sekarang} | DBO: {pbo_total} | CSC: {csc}")
    # print(f"       Detail Proyeksi - Meninggal: {total_proyeksi_meninggal} | Cacat: {total_proyeksi_cacat} | Resign: {total_proyeksi_resign} | Pensiun: {total_proyeksi_pensiun}")
    print(f"       Detail Proyeksi - Meninggal: {total_proyeksi_meninggal} | Cacat: {total_proyeksi_cacat} | Resign: {total_proyeksi_resign}")
    print(f"       Proyeksi Upah: {gaji_proyeksi_pensiun} | Proyeksi UUK-13/2003: {total_manfaat_proyeksi}")
    print(f"       Unit Manfaat Pensiun: {unit_manfaat_pensiun} | Unit Manfaat Meninggal: {unit_manfaat_meninggal} | Unit Manfaat Cacat: {unit_manfaat_cacat} | Unit Manfaat Resign: {unit_manfaat_resign}")
    print(f"       NK Pensiun: {nk_pensiun} | NK Meninggal: {nk_meninggal} | NK Cacat: {nk_cacat} | NK Resign: {nk_resign}")
    print(f"       PBO: {pbo_total}")
    print("---------------------------------------------------------------------")

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