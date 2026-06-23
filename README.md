# Aplikasi Perhitungan Imbalan Pasca Kerja Karyawan (PSAK 24 / UUCK)

Repositori ini berisi core engine berbasis Python untuk melakukan otomasi proyeksi dan perhitungan kewajiban imbalan pasca-kerja (kewajiban imbalan pasti) sesuai dengan regulasi **PSAK 24** dan ketentuan **Undang-Undang Cipta Kerja (UUCK)** menggunakan metode **Projected Unit Credit (PUC)**.

## 💻 Cara Menjalankan Aplikasi

Pastikan Anda sudah berada di dalam folder proyek melalui Terminal / Command Prompt lalu ketikan : **py -m streamlit run app.py**

## 🚀 Fitur Utama
- **Metode Atribusi PUC**: Mengimplementasikan perhitungan skala atribusi masa kerja secara proporsional (`masa_kerja_sekarang / masa_kerja_proyeksi`).
- **Gerbang Batas Atribusi**: Pembatasan otomatis untuk karyawan dengan usia di bawah kriteria minimum (`UPN - 24 tahun`).
- **Multi-Decrement Dinamis**: Mengakomodasi 3 risiko pemberhentian kerja sekaligus secara horizontal ($t=0$ hingga $t=40$):
  - Proyeksi Meninggal Dunia (menggunakan basis $t=0$).
  - Proyeksi Cacat/Invaliditas (menggunakan basis $t=0$).
  - Proyeksi Mengundurkan Diri / Resign (dikondisikan murni efektif dari $t \ge 1$).
- **Spot Rate / Yield Curve Dinamis**: Pembacaan tingkat diskonto secara spesifik per tenor berdasarkan pembulatan bawah `INT()` masa kerja ke depan.
- **VLOOKUP TRUE Replication**: Pencarian indeks tabel UUCK yang fleksibel dan akurat berbasis algoritma pencarian baris terdekat (VLOOKUP pendekatan eksak/interpolasi aktuaria).
- **Hasil Akhir Presisi**: Semua output keuangan (DBO, CSC, Nilai Kini) disajikan dalam angka bulat utuh (integer) sesuai standar kertas kerja akuntansi.

## 📂 Struktur Modul & Berkas
- `app.py`: Script inti yang memuat fungsi pembacaan data (`muat_template_uuck`, `muat_tabel_mortalita_dinamis`, `muat_tabel_spot_rate`) dan fungsi kalkulasi PUC (`hitung_puc_karyawan_v19`).

## 🛠️ Logika Kalkulasi Penting

### 1. Rumus Proyeksi & Diskonto
Untuk komponen Meninggal dan Cacat, faktor diskonto kumulatif dihitung berbasis tahun berjalan menggunakan rumus:

$$\text{Faktor Diskonto Kumulatif} = \left(\frac{1}{1 + \text{spot\_rate}_t}\right)^t$$

Khusus untuk komponen **Mengundurkan Diri (Resign)**, perhitungan tahun berjalan ($t=0$) diabaikan/di-force ke angka `0`, dan akumulasi baru dihitung secara normal dari tahun pertama ($t \ge 1$).

### 2. Pembulatan Hasil Akhir
Fungsi `hitung_puc_karyawan_v19` mengembalikan nilai yang sudah dibersihkan dari angka desimal (koma) menggunakan pembulatan terdekat:
```python
"dbo": int(round(pbo_total)),
"csc": int(round(csc))

### 2. Pembulatan Hasil Akhir
Fungsi `hitung_puc_karyawan_v19` mengembalikan nilai yang sudah dibersihkan dari angka desimal (koma) menggunakan pembulatan terdekat:
```python
"dbo": int(round(pbo_total)),
"csc": int(round(csc)).
