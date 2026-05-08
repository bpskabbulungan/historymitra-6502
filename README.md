# History Mitra 

## Ringkasan

**Aplikasi CLI berbasis Python** ini dirancang untuk mengambil daftar mitra BPS Bulungan beserta histori survei yang pernah diikuti tiap mitra, menampilkan ringkasan rapi di terminal, dan mengekspor hasilnya dalam bentuk **Excel** yang detail dan mudah dibaca.

Alur penggunaan: lakukan [Persiapan (Python dan Git)](#persiapan-python-dan-git), lanjut ke [Unduh Project](#unduh-project), lakukan [Instal Dependensi](#instal-dependensi), isi [Konfigurasi `.env`](#konfigurasi-env), lalu jalankan dari [Cara Menjalankan](#cara-menjalankan).

## Daftar Isi

- [Ringkasan](#ringkasan)
- [Daftar Isi](#daftar-isi)
- [Struktur Folder](#struktur-folder)
- [Persiapan (Python dan Git)](#persiapan-python-dan-git)
- [Unduh Project](#unduh-project)
- [Instal Dependensi](#instal-dependensi)
- [Konfigurasi `.env`](#konfigurasi-env)
- [Cara Menjalankan](#cara-menjalankan)
- [Opsi CLI](#opsi-cli)
- [Alur Kerja Singkat](#alur-kerja-singkat)
- [Output yang Dihasilkan](#output-yang-dihasilkan)
- [Troubleshooting](#troubleshooting)
- [Catatan Etika Penggunaan](#catatan-etika-penggunaan)
- [Kredit](#kredit)

## Struktur Folder

```text
.
|- src/
|  `- historymitra/
|     |- __init__.py          # Inisialisasi modul
|     |- __main__.py          # Entry point aplikasi (python -m historymitra)
|     |- cli.py               # Parsing argumen CLI & orkestrasi utama
|     |- api.py               # Client API (fetch data mitra & histori)
|     |- parsers.py           # Parsing data mentah (khusus mode sample)
|     `- reporting.py         # Ekspor ke format Excel & JSON
|- output/                    # Hasil ekspor (dibuat otomatis saat dijalankan)
|  |- raw/                    # Raw response data dalam format JSON
|  `- mitra_history_report.xlsx # File hasil akhir dengan multi-sheet
|- .env.example               # Contoh file konfigurasi environment
|- .gitignore                 # Konfigurasi file/folder yang diabaikan Git
|- pyproject.toml             # Konfigurasi build & metadata project Python
|- requirements.txt           # Daftar dependensi package Python
`- README.md                  # Dokumentasi utama project
```

## Persiapan (Python dan Git)

Jika Python dan Git sudah terpasang, langsung ke [Unduh Project](#unduh-project).

### 1) Install Python

Unduh Python 3.8+ dari:
https://www.python.org/downloads/

### 2) Install Git

Unduh Git dari:
https://git-scm.com/downloads

### 3) Cek instalasi

Jalankan di PowerShell/CMD:

```bash
python --version
```

```bash
git --version
```

## Unduh Project

Jika project belum ada di mesin lokal:

```bash
git clone <url-repository-anda>
cd historymitra-6502
```

Jika folder project sudah ada, cukup masuk ke folder project:

```bash
cd historymitra-6502
```

## Instal Dependensi

Pastikan terminal berada di folder project. Sangat disarankan untuk menggunakan **virtual environment**:

```bash
python -m venv .venv
```

Aktifkan virtual environment (di PowerShell/Windows):

```powershell
.venv\Scripts\activate
```

Atau (di Git Bash/Linux/macOS):

```bash
source .venv/bin/activate
```

Lalu jalankan instalasi:

```bash
pip install -e .
```

Atau melalui requirements:

```bash
pip install -r requirements.txt
```

## Konfigurasi `.env`

Copy file contoh env:

```powershell
Copy-Item .env.example .env
```

Atau (Git Bash/Linux/macOS):

```bash
cp .env.example .env
```

Lalu isi kredensial/header yang dibutuhkan (terutama untuk mode `live`):

```env
MITRA_API_USER_AGENT="Mozilla/5.0"
MITRA_API_COOKIE="cookie_dari_browser_anda"
```

Variabel lain yang didukung:
- `MITRA_API_AUTHORIZATION`
- `MITRA_API_HEADERS_JSON`
- `MITRA_API_REFERER`

## Cara Menjalankan

Ada dua pilihan mode:
- `sample`: memakai file dump teks (berguna untuk testing offline tanpa koneksi API).
- `live`: memanggil API langsung memakai header atau cookie dari sesi yang sah di browser.

### Menjalankan mode sample

```bash
python -m historymitra run ^
  --source sample ^
  --table-file "C:\Users\ASUS\Downloads\tabel-mitra.txt" ^
  --detail-file "C:\Users\ASUS\Downloads\id-mitra.txt" ^
  --history-file "C:\Users\ASUS\Downloads\history-mitra.txt"
```

*(Catatan: sesuaikan path file dump teks dengan lokasi di komputer Anda)*

### Menjalankan mode live

Pastikan variabel `.env` sudah diisi dengan benar, lalu jalankan:

```bash
historymitra run --source live --year 2026 --prov 65 --kab 02
```

## Opsi CLI

- `--source <sample|live>`: sumber data (wajib).
- `--year <number>`: tahun survei (digunakan pada mode live, default: tahun berjalan).
- `--prov <string>`: kode provinsi (digunakan pada mode live).
- `--kab <string>`: kode kabupaten/kota (digunakan pada mode live).
- `--table-file <path>`: path file dump list mitra (mode sample).
- `--detail-file <path>`: path file dump id mitra (mode sample).
- `--history-file <path>`: path file dump histori mitra (mode sample).
- `--output-dir <path>`: mengubah folder target untuk output.

## Alur Kerja Singkat

1. Load konfigurasi dari `.env` dan argumen CLI.
2. Siapkan folder kerja (`output` atau sesuai argumen).
3. **Mode Live**: Ambil data mitra dan histori langsung dari endpoint API menggunakan sesi `.env`.
   **Mode Sample**: Baca dan parsing data dari file dump teks lokal.
4. Normalisasi dan gabungkan data mitra dengan historinya.
5. Tampilkan ringkasan progres di terminal.
6. Ekspor hasil menjadi file Excel `.xlsx` dengan beberapa sheet terstruktur dan simpan data mentah (raw) dalam format JSON.

## Output yang Dihasilkan

Secara default hasil disimpan ke folder `output/`:

### File Excel
- `output/mitra_history_report.xlsx`

Sheet Excel yang dibuat:
- `Ringkasan`
- `Mitra`
- `History`
- `Rekap Mitra`
- `Rekap Survey`

### Folder Raw (`output/raw/`)
Menyimpan file JSON untuk keperluan debug atau inspeksi:
- `mitra_list.json`
- `mitra_histories.json`

## Troubleshooting

1. **Endpoint live mengembalikan `401 Unauthorized`:**
   - Endpoint membutuhkan autentikasi yang sah. Pastikan variabel lingkungan (`MITRA_API_COOKIE` dsb) diset menggunakan cookie/sesi yang valid dari browser yang sudah login.
2. **Mode `live` selalu gagal:**
   - Mode `sample` disediakan agar alur normalisasi dan ekspor tetap bisa diuji dari file dump teks jika koneksi ke API bermasalah.
3. **Data sensitif tidak muncul di output:**
   - Aplikasi ini sengaja tidak memfokuskan ekstraksi data sensitif seperti rekening, NPWP, foto, dan dokumen pendukung demi keamanan data.

## Catatan Etika Penggunaan

1. Gunakan tool ini hanya untuk akun dan data yang Anda miliki izin aksesnya.
2. Mohon beri jeda antar-request dan hindari peningkatan agresivitas fetch secara berlebihan.
3. Gunakan pengaturan delay yang wajar agar tidak membebani server saat menggunakan mode `live`.
4. Tool ini ditujukan untuk mempermudah pekerjaan Anda, bukan untuk penggunaan yang tidak semestinya.
5. Segala penggunaan menjadi tanggung jawab masing-masing pengguna.

## Kredit

Semoga panduan ini membantu. Jika ada pertanyaan, hubungi tim IPDS BPS Kabupaten Bulungan.
