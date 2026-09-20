#!/usr/bin/env python3
"""
Generate Pas Foto 3x4 in Canva 6x7 Grid Template (42 Photos per Page)
Matching exact coordinates from 'Salinan dari Cara Pakai QR Code!.pdf'
"""

import os
import json
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from PIL import Image

BASE_DIR = "/home/abuyyy/Foto_Teman_3x4_A4"
DATA_DIR = os.path.join(BASE_DIR, "data")
PROC_IMG_DIR = os.path.join(DATA_DIR, "processed_3x4")
STATE_FILE = os.path.join(DATA_DIR, "students_state.json")
DOWNLOADS_DIR = "/home/abuyyy/Downloads"

# Exact frame coordinates in points (72 points = 1 inch = 2.54 cm)
# Page size: A4 = 595.5 x 842.25 pt
PAGE_W = 595.5
PAGE_H = 842.25

XS = [14.41, 101.17, 187.76, 274.52, 361.11, 447.87]
YS = [717.97, 600.97, 483.97, 366.97, 249.97, 132.97, 16.33] # row 0 (top) to row 6 (bottom)
FRAME_W = 79.09
FRAME_H = 108.00
COLS = 6
ROWS = 7
ITEMS_PER_PAGE = COLS * ROWS # 42

def load_students():
    with open(STATE_FILE, "r", encoding="utf-8") as f:
        students = json.load(f)
    return students

def generate_pdf(students, output_path, with_label=False, title_text=None):
    c = canvas.Canvas(output_path, pagesize=(PAGE_W, PAGE_H))
    total_items = len(students)
    total_pages = (total_items + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE

    for page_idx in range(total_pages):
        page_num = page_idx + 1
        page_students = students[page_idx * ITEMS_PER_PAGE : (page_idx + 1) * ITEMS_PER_PAGE]

        for idx, s in enumerate(page_students):
            col = idx % COLS
            row = idx // COLS
            x = XS[col]
            y = YS[row]

            nrp = s.get("short_nrp") or str(s.get("nrp", ""))[-3:].zfill(3)
            nama = s.get("nama", "")
            img_path = os.path.join(PROC_IMG_DIR, f"{nrp}_3x4.jpg")
            has_real_photo = s.get("has_photo", False) and os.path.exists(img_path)

            if has_real_photo:
                # Gambar foto pas
                c.drawImage(img_path, x, y, width=FRAME_W, height=FRAME_H, preserveAspectRatio=False)
                
                # Garis panduan potong tipis (0.3 pt)
                c.saveState()
                c.setStrokeColor(colors.HexColor("#CBD5E1"))
                c.setLineWidth(0.3)
                c.rect(x, y, FRAME_W, FRAME_H, stroke=1, fill=0)
                c.restoreState()

                # Label identitas opsional di bawah foto
                if with_label:
                    c.saveState()
                    # Semi-transparent overlay bar at bottom
                    c.setFillColor(colors.HexColor("#0F172A"))
                    c.setFillAlpha(0.72)
                    c.rect(x, y, FRAME_W, 11.5, stroke=0, fill=1)
                    c.setFillAlpha(1.0)
                    c.setFillColor(colors.white)
                    c.setFont("Helvetica-Bold", 5.5)
                    short_name = nama.split()[0] if len(nama.split()) == 1 else " ".join(nama.split()[:2])
                    if len(short_name) > 16:
                        short_name = short_name[:14] + ".."
                    c.drawCentredString(x + FRAME_W / 2.0, y + 3.0, f"{nrp} - {short_name}")
                    c.restoreState()

            else:
                # Placeholder kotak kosong jika belum ada foto
                c.saveState()
                c.setFillColor(colors.HexColor("#F8FAFC"))
                c.rect(x, y, FRAME_W, FRAME_H, stroke=0, fill=1)
                c.setStrokeColor(colors.HexColor("#94A3B8"))
                c.setLineWidth(0.6)
                c.setDash(2, 2)
                c.rect(x, y, FRAME_W, FRAME_H, stroke=1, fill=0)

                # Teks informasi mahasiswa
                c.setFont("Helvetica-Bold", 6.5)
                c.setFillColor(colors.HexColor("#475569"))
                c.drawCentredString(x + FRAME_W / 2.0, y + FRAME_H / 2.0 + 4, f"{nrp}")
                c.setFont("Helvetica", 5.0)
                c.setFillColor(colors.HexColor("#64748B"))
                first_name = nama.split()[0] if nama else "-"
                c.drawCentredString(x + FRAME_W / 2.0, y + FRAME_H / 2.0 - 5, first_name)
                c.setFont("Helvetica-Oblique", 4.2)
                c.setFillColor(colors.HexColor("#94A3B8"))
                c.drawCentredString(x + FRAME_W / 2.0, y + FRAME_H / 2.0 - 13, "Belum Upload")
                c.restoreState()

        c.showPage()

    c.save()
    print(f"[OK] Selesai: {output_path} ({total_pages} halaman, {total_items} mahasiswa)")

def main():
    students = load_students()
    
    # 1. Versi Mahasiswa yang SUDAH memiliki foto (Siap Cetak & Potong, tanpa kotak kosong)
    available_students = [s for s in students if s.get("has_photo")]
    print(f"Total mahasiswa dengan foto tersedia: {len(available_students)}")

    out_clean_downloads = os.path.join(DOWNLOADS_DIR, "Foto_3x4_Grid_Canva_Siap_Cetak_Polos.pdf")
    generate_pdf(available_students, out_clean_downloads, with_label=False)

    out_label_downloads = os.path.join(DOWNLOADS_DIR, "Foto_3x4_Grid_Canva_Siap_Cetak_Dengan_Label.pdf")
    generate_pdf(available_students, out_label_downloads, with_label=True)

    # 2. Versi LENGKAP 106 Mahasiswa (berurutan NRP 001 - 120, ada placeholder untuk yang belum upload)
    out_all_downloads = os.path.join(DOWNLOADS_DIR, "Foto_3x4_Grid_Canva_Lengkap_106_Mahasiswa.pdf")
    generate_pdf(students, out_all_downloads, with_label=False)

    out_all_label = os.path.join(DOWNLOADS_DIR, "Foto_3x4_Grid_Canva_Lengkap_106_Dengan_Label.pdf")
    generate_pdf(students, out_all_label, with_label=True)

    # Simpan juga salinan di folder Foto_Teman_3x4_A4
    for fname in [
        "Foto_3x4_Grid_Canva_Siap_Cetak_Polos.pdf",
        "Foto_3x4_Grid_Canva_Siap_Cetak_Dengan_Label.pdf",
        "Foto_3x4_Grid_Canva_Lengkap_106_Mahasiswa.pdf",
        "Foto_3x4_Grid_Canva_Lengkap_106_Dengan_Label.pdf"
    ]:
        src = os.path.join(DOWNLOADS_DIR, fname)
        dst = os.path.join(BASE_DIR, fname)
        with open(src, "rb") as f_in, open(dst, "wb") as f_out:
            f_out.write(f_in.read())

    # 3. Buat file Peta Indeks Posisi Foto (Cheatsheet Cetak)
    cheat_path = os.path.join(DOWNLOADS_DIR, "PANDUAN_POSISI_FOTO_GRID_CANVA.txt")
    with open(cheat_path, "w", encoding="utf-8") as f:
        f.write("PETA INDEKS POSISI PAS FOTO 3X4 - GRID CANVA (6x7 = 42 KOTAK / HALAMAN)\n")
        f.write("="*75 + "\n\n")
        for idx, s in enumerate(available_students):
            page = (idx // ITEMS_PER_PAGE) + 1
            pos_in_page = idx % ITEMS_PER_PAGE
            row = (pos_in_page // COLS) + 1
            col = (pos_in_page % COLS) + 1
            f.write(f"No {idx+1:2d} | Hal {page}, Baris {row}, Kolom {col} : NRP {s['short_nrp']} - {s['nama']}\n")
    
    print(f"[OK] Cheatsheet indeks dibuat di: {cheat_path}")

if __name__ == "__main__":
    main()
