# Generator & Dashboard Tabel PODES 2025 Desa

Aplikasi lokal untuk **mengisi template tabel publikasi (gaya "Kecamatan Dalam Angka") secara otomatis dari file `.sav` PODES 2025 Desa**, lalu menampilkannya sebagai dashboard web dan mengekspornya kembali ke Excel per kecamatan × per bab.

Cukup **upload satu file `.sav`** — sistem akan:
1. Membuat template per kecamatan sesuai daftar desa di data (nama kabupaten/kecamatan & baris desa mengikuti `.sav`).
2. Mengisi tabel dari variabel PODES (`r304`, `r701`, `r601`, `r902`, dst).
3. Menampilkan dashboard (setia format template) dan menyediakan tombol **Compile ZIP**.

> Instrumen yang didukung: **PODES 2025 Desa** (nama variabel `r101`…`r1203`). Data yang tidak ada di `.sav` (mis. jumlah penduduk Bab 3, batas wilayah 1.1.3, nama pejabat) akan tampil `–`.

## Persyaratan

- Python 3.9+
- Paket: lihat `requirements.txt`

```bash
pip install -r requirements.txt
```

## Menjalankan (mode web — disarankan)

```bash
PORT=8090 python3 webapp/app.py
```

Buka **http://localhost:8090** → halaman upload muncul → pilih file `.sav` PODES 2025 Desa → **Proses & Muat**.

Setelah selesai (±1 menit untuk 1 kabupaten):
- Ringkasan kabupaten + pilih kecamatan (sidebar kiri).
- Tiap kecamatan punya tab **Bab 1–7**; tabel dirender persis mengikuti template (judul dwibahasa, header bertingkat, merge).
- **📦 Compile ZIP** — unduh seluruh hasil (per kecamatan × per bab) dalam satu `.zip`.
- **🗑 Clear Project** — bersihkan hasil untuk berpindah ke file `.sav` lain.

## Menjalankan (mode CLI)

Taruh `.sav` di folder `data/` dan folder template di `template tabel/`, lalu:

```bash
python3 run_all.py
```

Menjalankan pipeline lengkap dan menghasilkan folder `template tabel (terisi)/` + arsip zip.

Konfigurasi path bisa lewat environment variable:

```bash
PODES_SAV=/path/ke/data.sav \
PODES_TEMPLATE="/path/ke/template tabel" \
PODES_OUTPUT=/path/ke/output \
python3 run_all.py
```

## Struktur

| Berkas | Fungsi |
|--------|--------|
| `webapp/app.py` | Server web: upload `.sav`, dashboard, compile, clear |
| `generate.py` | Membuat template per kecamatan dari `.sav` (kloning master + sesuaikan baris desa + ganti nama kab/kec + hapus angka hardcode) |
| `fill_bab7.py`, `fill_bab6.py`, `fill_bab4.py` | Mengisi tabel spesifik Bab 7/6/4 |
| `fill_extra.py` | Mengisi Bab 1/2/3/5, menghitung rumus, mengisi `–` untuk sel tanpa sumber data |
| `run_all.py` | Menjalankan seluruh pipeline (CLI) |
| `config.py` | Konfigurasi path (`.sav`, template, output) |
| `template tabel/` | Template master (struktur tabel BPS) |
| `MAPPING.md` | Pemetaan tabel → variabel `.sav` |

## Cara kerja pengisian

- **Data-driven**: semua acuan pakai nama variabel PODES, jadi berlaku untuk satker mana pun (selama instrumennya PODES 2025 Desa).
- **Setia template**: dashboard merender langsung file Excel hasil; tabel bertumpuk (mis. 3.3, 4.2.2) digabung menyamping agar bisa di-scroll.
- **Konvensi BPS**: sel nihil/tanpa data ditulis `–`.

## Catatan

- File `.sav` **tidak** disertakan di repo ini (data survei mentah) — bawa file Anda sendiri.
- Aplikasi berjalan **lokal** (`localhost`). Tidak memerlukan koneksi/tunnel eksternal.
- Untuk jumlah desa di luar set master bawaan, generator otomatis memakai master terdekat lalu menyesuaikan jumlah baris (sisip/hapus).
