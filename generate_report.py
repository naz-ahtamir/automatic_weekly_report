#!/usr/bin/env python3
"""Agen penyusun Laporan Mingguan HSE.

Sumber data (hanya dua ini):
  1. Google Sheet inspeksi  (Bahaya Tersembunyi & Pelanggaran 6S)
  2. Google Sheet pelatihan (Induksi, Safety Talk, In-house Training, Drill)

Pemakaian:
  python3 generate_report.py                          # default: Jumat minggu lalu - Kamis minggu ini
  python3 generate_report.py --start 2026-10-30 --end 2026-11-05
  python3 generate_report.py --use-cache              # pakai data unduhan sebelumnya (lebih cepat)
  python3 generate_report.py --outdir /path/lain

Semua file (cache, hasil) berada di folder script ini.
"""
import argparse
import io
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image

BASE = Path(__file__).resolve().parent
CACHE = BASE / "cache"
PHOTO_CACHE = CACHE / "photos"

INSPECTION_ID = "1vJVOkzRPLIuAWntBHyTFFR4b_95qUggZqVzUYA85R5I"
INSPECTION_GID = "0"
TRAINING_ID = "1mglnfdiGn4-mShHS_RO4ul3q7C6bshkJCEPbAc23tL8"
TRAINING_GID = "1399125371"

MONTHS_EN = ["January", "February", "March", "April", "May", "June", "July",
             "August", "September", "October", "November", "December"]

# ----------------------------------------------------------------- tanggal ---

def default_range(today=None):
    """Jumat minggu lalu s/d Kamis minggu ini (Kamis terakhir <= hari ini)."""
    today = today or date.today()
    end = today - timedelta(days=(today.weekday() - 3) % 7)  # Kamis = weekday 3
    return end - timedelta(days=6), end


def report_filename(start, end):
    """Contoh: HSE周报 2026年10月30日11月5日 (Weekly Report HSE, 30 October-5 November 2026).xlsx"""
    if start.year == end.year:
        zh = f"{start.year}年{start.month}月{start.day}日{end.month}月{end.day}日"
        if start.month == end.month:
            en = f"{start.day}-{end.day} {MONTHS_EN[end.month - 1]} {end.year}"
        else:
            en = (f"{start.day} {MONTHS_EN[start.month - 1]}-"
                  f"{end.day} {MONTHS_EN[end.month - 1]} {end.year}")
    else:
        zh = f"{start.year}年{start.month}月{start.day}日{end.year}年{end.month}月{end.day}日"
        en = (f"{start.day} {MONTHS_EN[start.month - 1]} {start.year}-"
              f"{end.day} {MONTHS_EN[end.month - 1]} {end.year}")
    return f"HSE周报 {zh} (Weekly Report HSE, {en}).xlsx"


# ------------------------------------------------------------------- unduh ---

def fetch_csv(sheet_id, gid, name, use_cache):
    CACHE.mkdir(exist_ok=True)
    path = CACHE / f"{name}.csv"
    if use_cache and path.exists():
        print(f"[cache] {path.name}")
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    print(f"[unduh] {name} (bisa beberapa menit)...")
    last = None
    for attempt in range(3):
        try:
            r = requests.get(url, timeout=(30, 600))
            r.raise_for_status()
            path.write_bytes(r.content)
            return pd.read_csv(path, dtype=str, keep_default_na=False)
        except Exception as e:  # noqa: BLE001
            last = e
            print(f"  percobaan {attempt + 1} gagal: {e}")
            time.sleep(3)
    if path.exists():
        print(f"  [peringatan] memakai cache lama untuk {name}")
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    raise RuntimeError(f"Gagal mengunduh {name}: {last}")


def split_urls(cell):
    return re.findall(r"https?://[^\s,]+(?:,(?!\s*https?://)[^\s,]*)*", cell or "")


def first_valid(*cells):
    """Gabungkan semua URL foto dari beberapa sel, sesuai urutan."""
    urls = []
    for c in cells:
        urls += split_urls(c)
    return urls


def drive_id(url):
    m = re.search(r"[?&]id=([\w-]+)", url) or re.search(r"/d/([\w-]+)", url)
    return m.group(1) if m else None


def candidate_urls(url):
    if "drive.google.com" in url:
        fid = drive_id(url)
        if not fid:
            return [url]
        return [f"https://drive.google.com/uc?export=download&id={fid}",
                f"https://drive.usercontent.google.com/download?id={fid}&export=download"]
    return [url]


def cache_key(url):
    if "drive.google.com" in url:
        return "gd_" + (drive_id(url) or re.sub(r"\W", "", url)[-40:])
    m = re.search(r"tally\.so/private/([^?]+)\?id=(\w+)", url)
    if m:
        return "tally_" + re.sub(r"\W", "_", m.group(1)) + "_" + m.group(2)
    return "u_" + re.sub(r"\W", "", url)[-60:]


def get_photo(urls):
    """Ambil SATU foto: URL pertama yang berhasil. Return bytes JPEG terkompres / None."""
    PHOTO_CACHE.mkdir(parents=True, exist_ok=True)
    for url in urls:
        out = PHOTO_CACHE / (cache_key(url) + ".jpg")
        if out.exists() and out.stat().st_size > 0:
            return out.read_bytes()
        for cand in candidate_urls(url):
            for _ in range(3):
                try:
                    r = requests.get(cand, timeout=(20, 90))
                    if r.status_code != 200:
                        continue
                    img = Image.open(io.BytesIO(r.content))
                    img = img.convert("RGB")
                    img.thumbnail((700, 700))
                    buf = io.BytesIO()
                    img.save(buf, "JPEG", quality=80)
                    out.write_bytes(buf.getvalue())
                    return buf.getvalue()
                except Exception:  # noqa: BLE001
                    time.sleep(1)
    return None


def fetch_photos(url_lists):
    with ThreadPoolExecutor(max_workers=8) as ex:
        return list(ex.map(get_photo, url_lists))


# -------------------------------------------------------------- data logika ---

def col(df, pattern):
    """Cari kolom berdasarkan awalan/regex nama (case-insensitive)."""
    rx = re.compile(pattern, re.I)
    for c in df.columns:
        if rx.search(c):
            return c
    raise KeyError(f"Kolom tidak ditemukan: {pattern}")


def clean_type(header):
    h = re.sub(r"^\s*\d+\.\s*", "", header)
    h = re.sub(r"\s*\(.*$", "", h)
    return h.strip().upper()


def prepare_inspection(df, start, end):
    c_date = col(df, r"^TANGGAL INSPEKSI")
    d = df.copy()
    d["_date"] = pd.to_datetime(d[c_date], errors="coerce")
    d = d[(d["_date"] >= pd.Timestamp(start)) & (d["_date"] <= pd.Timestamp(end))]
    d = d.sort_values("_date", kind="stable")

    c_type = col(d, r"^TIPE TEMUAN")
    c_hier = col(d, r"^HIRARKI PENGENDALIAN.*\)$|^HIRARKI PENGENDALIAN \(")
    cols = list(d.columns)
    cat_cols = [c for c in cols[cols.index(c_type) + 1:] if re.match(r"^\s*\d+\.\s", c)]

    c_group = col(d, r"^QMB GROUP")
    c_dept = col(d, r"^NAMA DEPARTEMEN")
    c_img = col(d, r"^DOKUMENTASI HAZARD")
    c_level = col(d, r"^LEVEL HAZARD")
    c_pic = col(d, r"^PIC PERBAIKAN")
    c_so = col(d, r"^SAFETY OFFICER")
    c_dead = col(d, r"^DEADLINE PERBAIKAN")
    c_fix = col(d, r"^DOKUMENTASI PERBAIKAN")
    c_stat = col(d, r"^STATUS PERBAIKAN")
    c_rem = col(d, r"^KETERANGAN TAMBAHAN")

    rows = []
    for _, r in d.iterrows():
        htype, desc = "", ""
        for cc in cat_cols:
            if r[cc].strip():
                htype, desc = clean_type(cc), r[cc].strip().upper()
                break
        rows.append({
            "date": r["_date"].to_pydatetime(),
            "unit": r[c_dept].strip().upper(),
            "kind": r[c_type].strip().upper(),
            "area": "KONTRAKTOR" if "KONTRAKTOR" in r[c_group].upper() else "QMB",
            "before": first_valid(r[c_img]),
            "desc": desc,
            "htype": htype,
            "level": r[c_level].strip(),
            "pic": r[c_pic].strip().upper(),
            "so": r[c_so].strip(),
            "deadline": pd.to_datetime(r[c_dead], errors="coerce"),
            "after": first_valid(r[c_fix]),
            "status": r[c_stat].strip().upper(),
            "remark": r[c_rem].strip(),
        })
    return rows


def prepare_training(df, start, end):
    c_date = col(df, r"^TANGGAL KEGIATAN")
    d = df.copy()
    d["_date"] = pd.to_datetime(d[c_date], errors="coerce")
    d = d[(d["_date"] >= pd.Timestamp(start)) & (d["_date"] <= pd.Timestamp(end))]
    d = d.sort_values("_date", kind="stable")

    c_dept = col(d, r"^车间/部门")
    c_for = col(d, r"^PELATIHAN/ KEGIATAN DITUJUKAN")
    c_tema = col(d, r"^TEMA PELATIHAN")
    c_konten = col(d, r"^KONTEN PELATIHAN")
    c_jam = col(d, r"^JAM KEGIATAN")
    c_tka = col(d, r"^Jumlah Peserta TKA")
    c_tki = col(d, r"^Jumlah Peserta TKI")
    c_dur = col(d, r"^DURASI KEGIATAN")
    c_obj = col(d, r"^NAMA DEPARTEMEN/KONTRAKTOR")
    c_prob = col(d, r"培训存在的问题")
    c_hse = col(d, r"^NAMA PERSONIL HSE")
    c_foto = [c for c in d.columns if c.startswith("Foto Dokumentasi")]
    c_abs = [c for c in d.columns if c.startswith("Foto Absensi")]

    def num(v):
        try:
            return int(float(v))
        except (TypeError, ValueError):
            return 0

    out = {"Induksi": [], "SafetyTalk": [], "Pelatihan": [], "Drill": []}
    for _, r in d.iterrows():
        tema = r[c_tema].strip()
        low = tema.lower()
        for_ = r[c_for].strip()
        if "in-house" in low or "inhouse" in low or "pelatihan internal" in low:
            sheet = "Pelatihan"                      # permintaan: hanya kategori in-house training
        elif "induksi" in low:
            sheet = "Induksi"
        elif "safety talk" in low:
            sheet = "SafetyTalk"
        elif "simulasi tanggap darurat" in low or for_.lower().startswith("simulasi darurat"):
            sheet = "Drill"
        else:
            continue
        side = "Kontraktor" if ("kontraktor" in for_.lower() or "kontraktor" in low) else "QMB"
        out[sheet].append({
            "dept": r[c_dept].strip(),
            "tema": tema,
            "konten": r[c_konten].strip(),
            "date": r["_date"].to_pydatetime(),
            "jam": r[c_jam].strip(),
            "obj": r[c_obj].strip(),
            "tka": num(r[c_tka]),
            "tki": num(r[c_tki]),
            "dur": r[c_dur].strip(),
            "prob": "" if r[c_prob].strip() in ("-", "") else r[c_prob].strip(),
            "photo": first_valid(*[r[c] for c in c_foto]),
            "absen": first_valid(*[r[c] for c in c_abs]),
            "hse": r[c_hse].strip(),
            "for": for_,
            "side": side,
            "prod": "Konstruksi" if "konstruksi" in for_.lower() else "Produksi",
        })
    return out


# ------------------------------------------------------------------ excel ---

ARIAL = Font(name="Arial", size=12)
ARIAL_B = Font(name="Arial", size=12, bold=True)
THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")

INSP_HEAD = ["NO\n数字", "TANGGAL\n日期", "CONTRACTOR UNIT\n承包商单位", "HAZARD IMAGE\n隐患图片",
             "HAZARD DESCRIPTION\n隐患描述", "HAZARD TYPE\n隐患类型", "HAZARD LEVEL\n隐患级别",
             "CONTRACTOR PIC\n应急部安全员", "SAFETY OFFICER\n承包商负责人",
             "RECTIFICATION DEADLINE\n整改改期", "RECTIFICATION IMAGE\n整改图片",
             "RECTIFICATION STATUS\n整改情况", "REMARK\n备注"]
INSP_WIDTH = [7, 15, 28, 46, 46, 24, 34, 30, 34, 26, 46, 22, 22]

TRAIN_HEAD = ["车间/部门\nBengkel/\nDepartemen", "序号\nNo.", "培训主题\nTema Pelatihan",
              "培训内容\nKonten Pelatihan", "培训日期\nTanggal Pelatihan", "培训时间\nWaktu Pelatihan",
              "培训对象\nObjek Pelatihan", "培训人数\nPartisipan Pelatihan", None,
              "培训时长\n（小时/h）\nDurasi Pelatihan (jam)",
              "培训存在的问题及解决措施 \nMasalah pelatihan dan solusinya",
              "培训照片\nDokumentasi Pelatihan", "备注\nCatatan (Foto Absensi)", "NAMA PERSONIL HSE"]
TRAIN_WIDTH = [22, 8, 40, 60, 16, 20, 36, 11, 11, 18, 40, 46, 46, 30]

IMG_W, IMG_H = 300, 185          # px, foto inspeksi
TIMG_W, TIMG_H = 300, 210        # px, foto pelatihan


def style_cell(c, head=False, fmt=None):
    c.font = ARIAL_B if head else ARIAL
    c.alignment = CENTER
    c.border = BORDER
    if head:
        c.fill = HEAD_FILL
    if fmt:
        c.number_format = fmt


def place_image(ws, data, cell, max_w, max_h, keep):
    if not data:
        ws[cell] = "(foto tidak tersedia)"
        return
    img = XLImage(io.BytesIO(data))
    ratio = min(max_w / img.width, max_h / img.height, 1.0)
    img.width, img.height = int(img.width * ratio), int(img.height * ratio)
    img.anchor = cell
    ws.add_image(img)
    keep.append(img)


def build_inspection_sheet(wb, name, rows, photos, keep):
    ws = wb.create_sheet(name)
    ws.row_dimensions[1].height = 45
    for i, (h, w) in enumerate(zip(INSP_HEAD, INSP_WIDTH), 1):
        style_cell(ws.cell(1, i, h), head=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    for n, (r, (pb, pa)) in enumerate(zip(rows, photos), 1):
        ri = n + 1
        ws.row_dimensions[ri].height = 144
        dl = r["deadline"]
        vals = [n, r["date"], r["unit"], None, r["desc"], r["htype"], r["level"], r["pic"],
                r["so"], dl.to_pydatetime() if pd.notna(dl) else None, None, r["status"],
                r["remark"] or None]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(ri, ci, v)
            style_cell(c, fmt="yyyy-mm-dd" if ci in (2, 10) else None)
        place_image(ws, pb, f"D{ri}", IMG_W, IMG_H, keep)
        place_image(ws, pa, f"K{ri}", IMG_W, IMG_H, keep)
    ws.freeze_panes = "A2"
    return ws


def build_training_sheet(wb, name, rows, photos, keep):
    ws = wb.create_sheet(name)
    ws.row_dimensions[1].height = 48
    ws.row_dimensions[2].height = 30
    for i, (h, w) in enumerate(zip(TRAIN_HEAD, TRAIN_WIDTH), 1):
        style_cell(ws.cell(1, i, h), head=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    style_cell(ws.cell(2, 8, "中国籍\nTKA"), head=True)
    style_cell(ws.cell(2, 9, "印尼籍\nTKI"), head=True)
    for i in range(1, 15):
        if i not in (8, 9):
            style_cell(ws.cell(2, i), head=True)
    ws.merge_cells("H1:I1")
    for i in range(1, 15):
        if i not in (8, 9):
            ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
    for n, (r, (pf, pa)) in enumerate(zip(rows, photos), 1):
        ri = n + 3
        ws.row_dimensions[ri].height = min(409, max(170, 15 * (r["konten"].count("\n") + 2)))
        vals = [r["dept"], n, r["tema"], r["konten"], r["date"], r["jam"], r["obj"],
                r["tka"], r["tki"], r["dur"], r["prob"] or None, None, None, r["hse"]]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(ri, ci, v)
            style_cell(c, fmt="yyyy-mm-dd" if ci == 5 else None)
        ws.cell(ri, 4).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        place_image(ws, pf, f"L{ri}", TIMG_W, TIMG_H, keep)
        place_image(ws, pa, f"M{ri}", TIMG_W, TIMG_H, keep)
    return ws


def build_summary(wb, hazard, s6, train):
    ws = wb.active
    ws.title = "Summary"
    ws["B2"] = "SUMMARY WEEKLY REPORT"
    ws["B2"].font = Font(name="Arial", size=16, bold=True)

    # --- tabel temuan
    heads = ["Jenis 类型", "", "Area 区域", "Jumlah Temuan \n问题数量",
             "Jumlah Perbaikan\n整改数量", "Tingkat Perbaikan\n整改率"]
    for i, h in enumerate(heads):
        style_cell(ws.cell(7, 2 + i, h), head=True)
    ws.merge_cells("B7:C7")
    combos = [("Bahaya Tersembunyi 隐患", "青美邦 QMB", hazard, "QMB"),
              ("Bahaya Tersembunyi 隐患", "承包商 Kontraktor", hazard, "KONTRAKTOR"),
              ("Pelanggaran 6S  6S违规", "青美邦 QMB", s6, "QMB"),
              ("Pelanggaran 6S  6S违规", "承包商 Kontraktor", s6, "KONTRAKTOR")]
    for i, (jenis, area, rows, key) in enumerate(combos):
        r = 8 + i
        sel = [x for x in rows if x["area"] == key]
        ws.cell(r, 2, i // 2 + 1 if i % 2 == 0 else None)
        ws.cell(r, 3, jenis if i % 2 == 0 else None)
        ws.cell(r, 4, area)
        ws.cell(r, 5, len(sel))
        ws.cell(r, 6, sum(1 for x in sel if x["status"] == "CLSD"))
        ws.cell(r, 7, f"=IF($E$12=0,0,F{r}/$E$12)").number_format = "0.0%"
        for c in range(2, 8):
            style_cell(ws.cell(r, c), fmt=ws.cell(r, c).number_format)
    ws.merge_cells("B8:B9"); ws.merge_cells("C8:C9")
    ws.merge_cells("B10:B11"); ws.merge_cells("C10:C11")
    ws.cell(10, 2, 2)
    ws["D12"] = "Total"
    ws["E12"] = "=SUM(E8:E11)"
    ws["F12"] = "=SUM(F8:F11)"
    ws["G12"] = "=IF(E12=0,0,F12/E12)"
    ws["G12"].number_format = "0.0%"
    for c in range(2, 8):
        style_cell(ws.cell(12, c), head=(c in (4, 5, 6, 7)),
                   fmt="0.0%" if c == 7 else None)
    ws["E19"] = "Keterangan: Tingkat Perbaikan = Jumlah Perbaikan (status CLSD) / Total Temuan"
    ws["E19"].font = Font(name="Arial", size=10, italic=True)

    # --- tabel pelatihan
    for i, h in enumerate(["Program Pelatihan\n培训项目", "", "Sasaran Pelatihan\n培训对象",
                           "场次\n Jumlah Sesi", "人数\n Partisipan"]):
        style_cell(ws.cell(6, 11 + i - 1, h) if False else ws.cell(6, 10 + i, h), head=True)
    ws.merge_cells("J6:K6")

    def stat(rows, pred):
        sel = [x for x in rows if pred(x)]
        return len(sel), sum(x["tka"] + x["tki"] for x in sel)

    tr = [
        ("Induksi Karyawan baru\n新员工培训", "青美邦 QMB",
         stat(train["Induksi"], lambda x: x["side"] == "QMB")),
        ("Induksi Karyawan baru\n新员工培训", "承包商 Kontraktor",
         stat(train["Induksi"], lambda x: x["side"] == "Kontraktor")),
        ("Safety Talk\n现场培训", "施工培训 Konstruksi",
         stat(train["SafetyTalk"], lambda x: x["prod"] == "Konstruksi")),
        ("Safety Talk\n现场培训", "生产培训 Produksi",
         stat(train["SafetyTalk"], lambda x: x["prod"] == "Produksi")),
        ("Pelatihan Khusus (In-house)\n专项培训", "青美邦 QMB",
         stat(train["Pelatihan"], lambda x: x["side"] == "QMB")),
        ("Pelatihan Khusus (In-house)\n专项培训", "承包商 Kontraktor",
         stat(train["Pelatihan"], lambda x: x["side"] == "Kontraktor")),
        ("Emergency Drill\n应急演练", "青美邦 QMB", stat(train["Drill"], lambda x: True)),
    ]
    for i, (prog, sasaran, (sesi, org)) in enumerate(tr):
        r = 7 + i
        ws.cell(r, 10, prog if i % 2 == 0 or i == 6 else None)
        ws.cell(r, 12, sasaran)
        ws.cell(r, 13, sesi)
        ws.cell(r, 14, org)
        for c in range(10, 15):
            style_cell(ws.cell(r, c))
    for a, b in [(7, 8), (9, 10), (11, 12)]:
        ws.merge_cells(start_row=a, start_column=10, end_row=b, end_column=11)
    ws.merge_cells("J13:K13")
    ws["L14"] = "Total"
    ws["M14"] = "=SUM(M7:M13)"
    ws["N14"] = "=SUM(N7:N13)"
    for c in range(12, 15):
        style_cell(ws.cell(14, c), head=True)

    for letter, w in {"A": 3, "B": 8, "C": 26, "D": 24, "E": 16, "F": 16, "G": 16, "H": 4,
                      "I": 4, "J": 22, "K": 10, "L": 24, "M": 14, "N": 14}.items():
        ws.column_dimensions[letter].width = w
    for r in range(6, 15):
        ws.row_dimensions[r].height = 36

    # --- grafik
    ch1 = BarChart()
    ch1.title = "Temuan & Perbaikan"
    ch1.add_data(Reference(ws, min_col=5, max_col=6, min_row=7, max_row=11), titles_from_data=True)
    ch1.set_categories(Reference(ws, min_col=4, min_row=8, max_row=11))
    ch1.height, ch1.width = 8, 15
    ws.add_chart(ch1, "B21")
    ch2 = BarChart()
    ch2.title = "Pelatihan: Sesi & Peserta"
    ch2.add_data(Reference(ws, min_col=13, max_col=14, min_row=6, max_row=13), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=12, min_row=7, max_row=13))
    ch2.height, ch2.width = 8, 16
    ws.add_chart(ch2, "J21")
    return ws


def build(start, end, use_cache, outdir):
    insp_df = fetch_csv(INSPECTION_ID, INSPECTION_GID, "inspeksi", use_cache)
    train_df = fetch_csv(TRAINING_ID, TRAINING_GID, "pelatihan", use_cache)

    insp = prepare_inspection(insp_df, start, end)
    hazard = [r for r in insp if "BAHAYA" in r["kind"]]
    s6 = [r for r in insp if "6S" in r["kind"]]
    train = prepare_training(train_df, start, end)

    print(f"Periode {start} s/d {end}")
    print(f"  Bahaya Tersembunyi: {len(hazard)} | Pelanggaran 6S: {len(s6)}")
    print("  Pelatihan: " + ", ".join(f"{k}={len(v)}" for k, v in train.items()))

    print("Mengunduh foto (1 foto per entri)...")

    def photos_for(rows, a, b):
        la = fetch_photos([r[a] for r in rows])
        lb = fetch_photos([r[b] for r in rows])
        return list(zip(la, lb))

    p_h = photos_for(hazard, "before", "after")
    p_6 = photos_for(s6, "before", "after")
    p_t = {k: photos_for(v, "photo", "absen") for k, v in train.items()}
    missing = (sum(1 for p in p_h + p_6 for x in p if x is None)
               + sum(1 for v in p_t.values() for p in v for x in p if x is None))
    if missing:
        print(f"  [info] {missing} foto tidak tersedia/tidak terunduh (ditandai di laporan)")

    wb = Workbook()
    keep = []
    build_summary(wb, hazard, s6, train)
    for name in ["Induksi", "SafetyTalk", "Pelatihan", "Drill"]:
        build_training_sheet(wb, name, train[name], p_t[name], keep)
    build_inspection_sheet(wb, "Inspeksi-Hazard", hazard, p_h, keep)
    build_inspection_sheet(wb, "Inspeksi-6S", s6, p_6, keep)

    outdir.mkdir(parents=True, exist_ok=True)
    out = outdir / report_filename(start, end)
    wb.save(out)
    print(f"Selesai: {out}")
    return out


def parse_date(s):
    return datetime.strptime(s, "%Y-%m-%d").date()


def main():
    ap = argparse.ArgumentParser(description="Agen laporan mingguan HSE")
    ap.add_argument("--start", type=parse_date, help="Tanggal mulai YYYY-MM-DD (default: Jumat minggu lalu)")
    ap.add_argument("--end", type=parse_date, help="Tanggal akhir YYYY-MM-DD (default: Kamis minggu ini)")
    ap.add_argument("--use-cache", action="store_true", help="Pakai data unduhan sebelumnya")
    ap.add_argument("--outdir", type=Path, default=BASE, help="Folder hasil (default: folder script)")
    a = ap.parse_args()
    d_start, d_end = default_range()
    start, end = a.start or d_start, a.end or d_end
    if a.start and not a.end:
        end = start + timedelta(days=6)
    if a.end and not a.start:
        start = end - timedelta(days=6)
    if start > end:
        sys.exit("Tanggal mulai harus <= tanggal akhir")
    build(start, end, a.use_cache, a.outdir)


if __name__ == "__main__":
    main()
