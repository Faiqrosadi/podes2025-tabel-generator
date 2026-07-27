"""Konfigurasi pipeline pengisian tabel PODES 2025 Desa.
Ubah di sini (atau lewat environment variable) untuk memakai satker/kabupaten lain.

  PODES_SAV       path file .sav utama (data desa, BUKAN file infrastruktur)
  PODES_TEMPLATE  folder berisi subfolder per-kecamatan (Bab 1..7.xlsx)
  PODES_OUTPUT    folder hasil pengisian

Syarat: .sav adalah instrumen PODES 2025 Desa (nama variabel r101..r1203 sama),
dan nama subfolder template = nama kecamatan (dicocokkan via kolom nama_kec).
"""
import os, glob

BASE = os.path.dirname(os.path.abspath(__file__))

def _find_sav():
    env = os.environ.get("PODES_SAV")
    if env:
        return env
    cands = glob.glob(os.path.join(BASE, "data", "*.sav"))
    main = [f for f in cands if "infra" not in os.path.basename(f).lower()]
    # utamakan yang mengandung 'rev', lalu berdasarkan ukuran terbesar
    rev = [f for f in main if "rev" in os.path.basename(f).lower()]
    pick = (rev or main or cands)
    return max(pick, key=os.path.getsize) if pick else ""

SAV      = _find_sav()
TEMPLATE = os.environ.get("PODES_TEMPLATE", os.path.join(BASE, "template tabel"))
OUTPUT   = os.environ.get("PODES_OUTPUT",   os.path.join(BASE, "hasil"))
