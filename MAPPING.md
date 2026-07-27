# Mapping Template Tabel PODES 2025 → Variabel `.sav`

Dokumen ini memetakan setiap tabel di template (`template tabel/<kec>/Bab N.xlsx`) ke
variabel sumber di file `.sav`. **Fokus: hanya tabel yang datanya tersedia di `.sav`.**
Silakan review; setelah disetujui, mapping ini jadi dasar script pengisian otomatis.

## Sumber data

| File | Bentuk | Peran |
|------|--------|-------|
| `data/33_podes2025-desa_rev_20.sav` | 195 desa × 347 var | **Data utama** — semua angka per desa (`r101`…`r1203`) |
| `data/33_podes2025-desa[infrastruktur]_20.sav` | 3.418 baris (long) | **Daftar nama** fasilitas; kolom `jenis` = kode variabel (mis. `r701ck2`). Dipakai bila perlu jumlah fasilitas per desa hasil pencacahan langsung |

> **FOKUS: Bab 4, 6, 7** (per keputusan user).

## Aturan umum (sudah disepakati)

1. **Filter kecamatan**: pilih baris `.sav` dengan `nama_kec` sesuai folder (mis. `KEDUNG`), lalu urutkan `r104` (kode desa) → cocok dengan urutan baris template (001, 002, …).
2. **Kode keberadaan**: `1 = Ada`, `2 = Tidak ada` (beberapa var pakai `1=ada … 4=tidak ada`, mis. `r805a`, `r805c`).
3. **Variabel jumlah** bersifat numerik → dipakai apa adanya (per desa) atau dijumlah (total kecamatan).
4. **Tabel ringkasan "Banyaknya Desa/Kelurahan dengan…"** = **jumlah desa** yang memenuhi syarat (mis. `count(var > 0)` atau `count(var == 1)`), bukan jumlah unit.
5. **Kolom tahun**: hanya kolom **2025** yang diisi. Kolom 2023/2024 (kecuali 4.4.2) **diisi `-`**. ✔
6. **Sel tanpa data diisi `-`**. ✔
7. **4.4.2** (korban jiwa) → **DITUNDA / TBD**, akan dielaborasi lagi bersama user. Belum diisi.
8. `r503` (bahan bakar masak) = `4` (Elpiji 3 kg) untuk semua 195 desa → tabel 4.3.2 valid, terkonsentrasi di baris Elpiji 3 kg.

---

## ✅ Tabel yang BISA diisi dari `.sav`

Legenda: **[desa]** = satu baris per desa · **[ringkas]** = baris = kategori, nilai = jumlah desa.

### Bab 2 — Pemerintahan
| Tabel | Bentuk | Kolom template | Variabel / cara hitung |
|-------|--------|----------------|------------------------|
| 2.1.1 | [desa] | (2) RW, (3) RT | `r304a`, `r304b` |
| 2.2.5 | [desa] | (2) Petinggi/Lurah, (3) Carik, (4) Perangkat Lainnya | (2) `1 jika r1201ak2=1` · (3) `1 jika r1201bk2=1` · (4) `r1202a+r1202b+r1202c+r1202d` — **perlu konfirmasi definisi** |

### Bab 4 — Sosial
| Tabel | Bentuk | Isi | Variabel / cara hitung |
|-------|--------|-----|------------------------|
| 4.1.1 | [ringkas] | Jumlah **desa** yang punya tiap jenjang pendidikan (kol 2025) | TK `r701ak2+ak3>0`; RA/BA `r701b*`; SD `r701c*`; MI `r701d*`; SMP `r701e*`; MTs `r701f*`; SMA `r701g*`; MA `r701h*`; SMK `r701i*`; PT `r701j*` → `count(negeri+swasta>0)` |
| 4.1.2 | [ringkas] | Jumlah **sekolah** per jenjang: Negeri / Swasta / Total (kol 2025/2026 saja) | Negeri = Σ`r701Xk2`; Swasta = Σ`r701Xk3` per kecamatan. **Kolom 2024/2025 (tahun lalu) tidak tersedia** |
| 4.2.1 | [ringkas] | Jumlah **desa** punya tiap sarana kesehatan (kol 2025) | RS `r702ak2>0`; Pusk. rawat inap `r702dk2>0`; Pusk. tanpa rawat inap `r702ek2>0`; Pustu `r702fk2>0`; Klinik `r702b/g`; Apotek `r702lk2>0`; Poskesdes `r702jk2>0`; Polindes `r702kk2>0`; dst → `count(>0)`. **Kolom 2023/2024 tidak tersedia** |
| 4.2.2 | [desa] | Jumlah sarana kesehatan per desa (RS, Puskesmas, Pustu, …) | `r702ak2` (RS), `r702dk2+r702ek2` (Puskesmas), `r702fk2` (Pustu), dst — sesuai kolom |
| 4.3.1 | [ringkas] | Jumlah **desa** menurut sumber penerangan jalan utama (kol 2025) | kategori `r502b` → `count` per kode. **Perlu codebook arti kode r502b** |
| 4.3.2 | [ringkas] | Jumlah **desa** menurut bahan bakar memasak | kategori `r503` → `count` per kode (saat ini semua = 4) |
| 4.4.1 | [ringkas] | Jumlah **desa** mengalami tiap jenis bencana | per jenis `r601Xk2 == 1` → `count`. (a=longsor, b=banjir, c=banjir bandang, d=gempa, e=tsunami, f=gelombang pasang, g=puting beliung, h=letusan, i=karhutla, j=kekeringan, k=abrasi) |
| 4.4.2 | [ringkas] | Jumlah **desa** ada korban jiwa per jenis bencana (kol 2024 & 2025) | ⏸️ **TBD — dielaborasi nanti**. Kandidat: 2024=`r601Xk4>0`, 2025=`r601Xk7>0` |
| 4.4.3 | [ringkas] | Jumlah **desa** dengan fasilitas/upaya antisipasi bencana (kol 2025) | `r602a`…`r602e == 1` → `count` |

### Bab 6 — Akomodasi, Transportasi, Komunikasi
| Tabel | Bentuk | Isi | Variabel / cara hitung |
|-------|--------|-----|------------------------|
| 6.1.1 | [ringkas] | Jumlah **desa** ada Hotel / Penginapan (kol 2025) | Hotel `r905hk2>0`; Penginapan `r905ik2>0` → `count` |
| 6.2.1 | [ringkas] | Jumlah **desa** menurut prasarana/sarana transportasi (kol 2025) | Darat/air `r801a`; dapat dilalui roda-4 `r801b2`; angkutan umum `r801c1`; dll → `count` sesuai baris. **Perlu cocokkan tiap baris ke var** |
| 6.3.1 | [ringkas] | Jumlah **desa** ada fasilitas pos & ekspedisi (kol 2025) | Kantor pos `r805a==1`; Pos keliling `r805b==1`; Ekspedisi swasta `r805c==1` → `count` |

### Bab 7 — Ekonomi
| Tabel | Bentuk | Isi | Variabel / cara hitung |
|-------|--------|-----|------------------------|
| 7.1 | [desa] | Jumlah koperasi per desa: KUD, Industri/Kerajinan, Simpan Pinjam, Lainnya | `r903a`, `r903b`, `r903c`, `r903d` |
| 7.3 | [ringkas] | Jumlah **desa** ada tiap jenis bank (kol 2025) | Bank Umum Pemerintah `r902a1>0`; Bank Umum Swasta `r902a2>0`; BPR `r902a3>0` → `count` |
| 7.4 | [ringkas] | Jumlah **desa** ada tiap jenis koperasi aktif (kol 2025) | `r903a>0`, `r903b>0`, `r903c>0`, `r903d>0` → `count` |
| 7.5 | [ringkas] | Jumlah **desa** ada tiap sarana perdagangan (kol 2025) | Pertokoan `r905ak2>0`; Pasar permanen `r905bk2>0`; Pasar semi `r905ck2>0`; Pasar tanpa bangunan `r905dk2>0`; Minimarket `r905ek2>0`; Restoran `r905fk2>0`; Warung `r905gk2>0`; Toko kelontong `r905jk2>0` → `count` |

---

## ❌ Tabel yang TIDAK BISA diisi (data tidak ada di `.sav`) — dilewati

| Tabel | Kebutuhan data | Sumber sebenarnya |
|-------|----------------|-------------------|
| **Bab 1** (1.1.1–1.1.5) | Luas daerah, jarak, letak, ketinggian, topografi | Geospasial / BPS wilayah |
| **Bab 3** (3.1–3.8) | Penduduk per umur & jenis kelamin, kepadatan, KTP-el, KIA, agama | Dukcapil |
| 2.2.1–2.2.3 | Jumlah PNS, nama petinggi & carik | Kepegawaian / pemerintah kec |
| 2.2.4 | Jumlah anggota BPD, nama ketua | Pemerintah desa |
| 2.2.6 | Jumlah Hansip | Pemerintah desa |
| 4.1.3, 4.1.4 | Jumlah **guru** & **murid** per jenjang | Dinas Pendidikan / Kemenag |
| 4.1.5–4.1.35 | **Murid, guru, ruang kelas** per jenjang | Dinas Pendidikan / Kemenag (kolom "Sekolah" saja bisa dari infrastruktur, tapi tabel butuh murid/guru/ruang kelas) |
| 4.2.3 | Pasien puskesmas | Dinas Kesehatan |
| 4.2.4–4.2.6 | Layanan KB, PUS, peserta KB, alat kontrasepsi | DP2KB |
| 4.5.1 | Jumlah tempat peribadatan | Kemenag |
| 4.5.2 | Jemaah haji | Kemenag |
| 4.5.3 | Nikah, talak, cerai | KUA / Pengadilan Agama |
| 4.5.4 | Penduduk menurut agama | Dukcapil / Kemenag |
| **Bab 5** (5.1–5.13) | Luas panen & produksi tanaman/ternak (2022–2025) | Dinas Pertanian |
| 7.2 | Sertifikat tanah (PTSL, wakaf) | BPN/ATR |

---

## Lampiran: Codebook (dari Kuesioner PODES 2025 Desa)

**r502b — Sumber penerangan jalan utama** (baris tabel 4.3.1):
`1` Listrik diusahakan pemerintah · `2` Listrik diusahakan non-pemerintah · `3` Non listrik

**r503 — Bahan bakar memasak** (baris tabel 4.3.2):
`1` Listrik · `2` Elpiji 5,5 kg · `3` Elpiji 12 kg · `4` Elpiji 3 kg · `5` Gas kota · `6` Biogas · `7` Minyak tanah · `8` Briket · `9` Arang · `10` Kayu bakar · `11` Lainnya

**r701 — Jenjang pendidikan** (kolom 2=Negeri, 3=Swasta):
`a` TK · `b` RA/BA · `c` SD · `d` MI · `e` SMP · `f` MTs · `g` SMA · `h` MA · `i` SMK · `j` Akademi/PT

**r702 — Jenis faskes** (kolom 2=Jumlah):
`a` RS · `b` Klinik utama · `c` Balai kesehatan · `d` Puskesmas rawat inap · `e` Puskesmas tanpa rawat inap · `f` Pustu · `g` Klinik pratama · `h` Praktik dokter · `i` Praktik bidan · `j` Poskesdes · `k` Polindes · `l` Apotek · `m` Toko obat/jamu

**r601 — Bencana** (a=longsor, b=banjir, c=banjir bandang, d=gempa, e=tsunami, f=gelombang pasang, g=puting beliung, h=letusan gunung, i=karhutla, j=kekeringan, k=abrasi):
`k2` Kejadian (1=Ada) · `k3` Jml kejadian 2024 · `k4` Korban meninggal 2024 · `k5` Keberadaan korban 2024 · `k6` Jml kejadian 2025 · `k7` Korban meninggal 2025 · `k8` Keberadaan korban 2025

**r602 — Mitigasi bencana** (semua 1=Ada): a=peringatan dini, b=peringatan tsunami, c=perlengkapan keselamatan, d=rambu/jalur evakuasi, e=normalisasi sungai dll

**r801 — Transportasi**: `a` lalu lintas (1=Darat, 2=Air, 3=Darat&air, 4=Udara) · `b2` roda-4 (1=sepanjang tahun…4=tidak) · `c1` angkutan umum (A=trayek tetap, B=tanpa trayek, X=tidak ada)

**r805 — Pos**: `a` kantor pos (1=Beroperasi…4=Tidak ada) · `b` pos keliling (1=Ada, 2=Tidak) · `c` ekspedisi swasta (1=Beroperasi…4=Tidak ada)

**r902 — Bank** (jumlah): `a1` Bank Umum Pemerintah · `a2` Bank Umum Swasta · `a3` BPR
**r903 — Koperasi** (jumlah): `a` KUD · `b` Kopinkra/usaha mikro · `c` KSP/Kospin · `d` Lainnya
**r905 — Fasilitas ekonomi** (jumlah): `a` Kelompok pertokoan · `b` Pasar permanen · `c` Pasar semi permanen · `d` Pasar tanpa bangunan · `e` Minimarket · `f` Restoran · `g` Warung makan · `h` Hotel · `i` Penginapan · `j` Toko/warung kelontong
