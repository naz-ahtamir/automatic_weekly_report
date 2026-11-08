# Agen Laporan Mingguan HSE

Script [generate_report.py](generate_report.py) menyusun laporan mingguan HSE (file `.xlsx`) secara otomatis dari dua Google Sheet.

## Sumber Data

Hanya dua sumber ini yang dipakai:

| Sumber | Dipakai untuk |
|---|---|
| Google Sheet inspeksi | Sheet `Inspeksi-Hazard` (Bahaya Tersembunyi) dan `Inspeksi-6S` (Pelanggaran 6S) |
| Google Sheet pelatihan | Sheet `Induksi`, `SafetyTalk`, `Pelatihan`, `Drill` |

Kedua sheet harus bisa diakses publik (siapa saja yang punya link), karena diunduh sebagai CSV tanpa login.

## Kebutuhan

- Python 3.10 atau lebih baru
- Library: `pandas`, `requests`, `openpyxl`, `pillow`

```bash
python3 -m pip install pandas requests openpyxl pillow
```

## Cara Pakai

Jalankan dari folder ini.

```bash
# Periode default: Jumat minggu lalu s/d Kamis minggu ini
python3 generate_report.py

# Periode sesuai tanggal yang dimasukkan
python3 generate_report.py --start 2026-10-30 --end 2026-11-05

# Pakai data yang sudah diunduh sebelumnya (lebih cepat, tidak unduh ulang)
python3 generate_report.py --use-cache

# Simpan hasil di folder lain
python3 generate_report.py --outdir output
```

| Opsi | Keterangan |
|---|---|
| `--start YYYY-MM-DD` | Tanggal mulai. Jika `--end` tidak diisi, akhir = mulai + 6 hari. |
| `--end YYYY-MM-DD` | Tanggal akhir. Jika `--start` tidak diisi, mulai = akhir - 6 hari. |
| `--use-cache` | Pakai CSV di `cache/` tanpa mengunduh ulang. |
| `--outdir PATH` | Folder hasil. Default: folder ini. |

### Periode Default

Periode default berakhir di Kamis terakhir (hari ini jika hari ini Kamis) dan dimulai 6 hari sebelumnya, yaitu hari Jumat.
Contoh: dijalankan Minggu 8 Nov 2026 menghasilkan periode 30 Okt - 5 Nov 2026.

## Hasil

File disimpan di folder ini, kecuali `--outdir` diisi. Nama file menyesuaikan tanggal laporan, mengikuti contoh:

```
HSE周报 2026年10月30日11月5日 (Weekly Report HSE, 30 October-5 November 2026).xlsx
```

Jika nama file sama sudah ada (misalnya periode yang sama dengan file contoh), file lama akan tertimpa. Gunakan `--outdir` agar aman.

### Isi Laporan

| Sheet | Isi |
|---|---|
| `Summary` | Rekap temuan (jumlah, perbaikan, tingkat perbaikan), rekap pelatihan (sesi, peserta), dan 2 grafik |
| `Induksi` | Kegiatan bertema Induksi |
| `SafetyTalk` | Kegiatan bertema Safety Talk |
| `Pelatihan` | **Hanya kategori In-house Training** (tema mengandung "In-house Training/Pelatihan Internal") |
| `Drill` | Simulasi tanggap darurat |
| `Inspeksi-Hazard` | Temuan bertipe Bahaya Tersembunyi |
| `Inspeksi-6S` | Temuan bertipe Pelanggaran 6S |

Kegiatan pelatihan di luar empat kategori itu (misalnya Test Alat Calon Operator, Weekly Meeting) tidak dimasukkan.

### Foto

- Foto ditempelkan langsung di dalam file Excel, jadi tetap muncul tanpa koneksi internet.
- Jika satu data punya beberapa foto, hanya **1 foto** yang diambil (foto pertama yang berhasil diunduh).
- Foto diperkecil agar ukuran file tidak membengkak.
- Jika foto gagal diunduh, selnya diisi `(foto tidak tersedia)`.

Kolom foto:
- Inspeksi: foto sebelum (`HAZARD IMAGE`) dan foto perbaikan (`RECTIFICATION IMAGE`).
- Pelatihan: foto dokumentasi kegiatan dan foto absensi.

## Struktur Folder

```
weekly_report/
├── generate_report.py   # script utama
├── README.md
├── cache/
│   ├── inspeksi.csv     # unduhan sheet inspeksi
│   ├── pelatihan.csv    # unduhan sheet pelatihan
│   └── photos/          # foto yang sudah diunduh (dipakai ulang)
└── *.xlsx               # laporan hasil
```

Folder `cache/` aman dihapus. Isinya akan diunduh ulang saat script dijalankan lagi.

## Catatan

- Pengunduhan CSV inspeksi (sekitar 20 MB) dan foto bisa memakan waktu beberapa menit, tergantung koneksi.
- Foto yang sudah pernah diunduh disimpan di `cache/photos/`, jadi proses berikutnya lebih cepat.
- Tingkat Perbaikan di `Summary` dihitung sebagai jumlah perbaikan (status `CLSD`) dibagi total temuan, sama seperti rumus di contoh laporan.
