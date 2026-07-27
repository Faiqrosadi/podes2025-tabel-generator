#!/usr/bin/env python3
"""Jalankan seluruh pipeline pengisian tabel PODES 2025 Desa, lalu compile.

  python3 run_all.py

Membaca konfigurasi dari config.py (path .sav, folder template, folder output).
Untuk satker lain: taruh .sav di data/ dan folder template di 'template tabel/'
(atau set env PODES_SAV / PODES_TEMPLATE / PODES_OUTPUT), lalu jalankan ini.
Urutan penting: bab7/6/4 (isian spesifik) dulu, baru fill_extra (bab 1,2,3,5 +
hitung rumus + isi '-' generik untuk sel tanpa sumber data).
"""
import os, sys, shutil, subprocess
import config

BASE = os.path.dirname(os.path.abspath(__file__))
STEPS = ["fill_bab7.py", "fill_bab6.py", "fill_bab4.py", "fill_extra.py"]
GEN_TPL = os.path.join(BASE, "workspace", "template_gen")

def main():
    print(f"SAV    : {config.SAV}")
    print(f"MASTER : {config.TEMPLATE}")
    print(f"OUTPUT : {config.OUTPUT}\n")
    if not config.SAV or not os.path.exists(config.SAV):
        sys.exit("! File .sav tidak ditemukan. Set PODES_SAV atau taruh di data/")
    if not os.path.isdir(config.TEMPLATE):
        sys.exit("! Folder master tidak ditemukan. Set PODES_TEMPLATE.")

    # 1) generate template per kecamatan dari .sav (master generik -> template_gen)
    print("=== generate.py ===")
    r = subprocess.run([sys.executable, os.path.join(BASE, "generate.py"),
                        config.SAV, config.TEMPLATE, GEN_TPL])
    if r.returncode != 0: sys.exit("! Gagal generate")
    print()
    # 2) isi tabel dari .sav (pakai template_gen)
    env = os.environ.copy()
    env.update(PODES_SAV=config.SAV, PODES_TEMPLATE=GEN_TPL, PODES_OUTPUT=config.OUTPUT)
    for s in STEPS:
        print(f"=== {s} ===")
        r = subprocess.run([sys.executable, os.path.join(BASE, s)], env=env)
        if r.returncode != 0:
            sys.exit(f"! Gagal di {s}")
        print()

    # compile folder deliverable + zip
    dest = os.path.join(BASE, "template tabel (terisi)")
    if os.path.isdir(dest): shutil.rmtree(dest)
    shutil.copytree(config.OUTPUT, dest)
    zip_base = os.path.join(BASE, "PODES_tabel_terisi")
    if os.path.exists(zip_base + ".zip"): os.remove(zip_base + ".zip")
    shutil.make_archive(zip_base, "zip", root_dir=BASE, base_dir="template tabel (terisi)")
    n_kec = len([d for d in os.listdir(dest) if os.path.isdir(os.path.join(dest, d))])
    print(f"SELESAI. {n_kec} kecamatan -> {dest}")
    print(f"Arsip   : {zip_base}.zip")

if __name__ == "__main__":
    main()
