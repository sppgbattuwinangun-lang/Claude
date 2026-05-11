#!/usr/bin/env python3
"""
BUILD DASHBOARD AKUMULASI SPPG BATTUWINANGUN
Periode: 01 September 2025 s/d 31 Desember 2025

Script ini membaca seluruh file "Total Akumulasi *.xlsx" yang ada di folder,
menggabungkan 97 record harian, menghitung 4 klaster analitik
(A. Pangan, B. Operasional, C. Ringkasan Total, D. Surplus/Defisit),
dan membangun file Excel Dashboard eksekutif berwarna lengkap dengan:
- Sheet 1: Dashboard (ringkasan + formatting rupiah + chart siap taruh)
- Sheet 2: Rekap Harian (97 hari, data mentah untuk sumber chart)
- Sheet 3: Rekap Mingguan (agregasi per minggu)
- Sheet 4: Chart Data (tabel tersedia di chart)

Hanya menggunakan Python standard library (zipfile + xml).
"""

import os
import json
import zipfile
from xml.sax.saxutils import escape
from datetime import date, timedelta

WORKDIR = "/projects/sandbox/Claude"
OUTPUT = os.path.join(WORKDIR, "Dashboard_Akumulasi_Sep_Des_2025.xlsx")
DATA_JSON = os.path.join(WORKDIR, "accumulated_data.json")


# ============================================================
# HELPERS
# ============================================================

def col_letter(idx):
    """1-indexed column number -> Excel column letter (A, B, ..., Z, AA, ...)."""
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def iso_to_serial(iso_date):
    """'YYYY-MM-DD' -> Excel serial date (1899-12-30 epoch)."""
    y, m, d = map(int, iso_date.split("-"))
    base = date(1899, 12, 30)
    return (date(y, m, d) - base).days


# ============================================================
# LOAD & COMPUTE METRICS
# ============================================================

with open(DATA_JSON, "r") as f:
    records = json.load(f)

records.sort(key=lambda x: x["date_serial"])

# -- hari aktif = hari dengan Pagu > 0 (Senin-Sabtu; Minggu libur -> pagu=0)
active = [r for r in records if r["total_pagu"] > 0]
n_active = len(active)

# A. PANGAN
tot_aktual_pangan = sum(r["aktual_pangan"] for r in records)
tot_pagu_pangan = sum(r["pagu_pangan"] for r in records)
avg_pangan = tot_aktual_pangan / n_active
pagu_pangan_day = tot_pagu_pangan / n_active
selisih_pangan_day = pagu_pangan_day - avg_pangan
sisa_pangan = tot_pagu_pangan - tot_aktual_pangan
pct_pangan = tot_aktual_pangan / tot_pagu_pangan if tot_pagu_pangan else 0

# B. OPERASIONAL
tot_aktual_ops = sum(r["aktual_ops"] for r in records)
tot_pagu_ops = sum(r["pagu_ops"] for r in records)
avg_ops = tot_aktual_ops / n_active
pagu_ops_day = tot_pagu_ops / n_active
selisih_ops_day = pagu_ops_day - avg_ops
sisa_ops = tot_pagu_ops - tot_aktual_ops
pct_ops = tot_aktual_ops / tot_pagu_ops if tot_pagu_ops else 0

# C. RINGKASAN TOTAL
total_anggaran = tot_pagu_pangan + tot_pagu_ops
total_penggunaan = tot_aktual_pangan + tot_aktual_ops
total_sisa = total_anggaran - total_penggunaan
pct_penyerapan = total_penggunaan / total_anggaran if total_anggaran else 0
avg_total_day = total_penggunaan / n_active
pagu_total_day = total_anggaran / n_active
selisih_total_day = pagu_total_day - avg_total_day

# D. SURPLUS/DEFISIT
n_surplus = sum(1 for r in records if r["status"] == "SURPLUS")
n_defisit = sum(1 for r in records if r["status"] == "DEFISIT")
n_aman = sum(1 for r in records if r["status"] == "AMAN")
n_belum = len(records) - n_surplus - n_defisit - n_aman

tot_surplus = sum(r["surplus_defisit"] for r in records if r["surplus_defisit"] > 0)
tot_defisit = sum(r["surplus_defisit"] for r in records if r["surplus_defisit"] < 0)
net_sd = tot_surplus + tot_defisit
pct_surplus_total = net_sd / total_anggaran if total_anggaran else 0

active_spend = [r for r in records if r["total_aktual"] > 0]
max_r = max(active_spend, key=lambda x: x["total_aktual"])
min_r = min(active_spend, key=lambda x: x["total_aktual"])

# -- VALIDASI SILANG (wajib match)
assert abs((tot_aktual_pangan + tot_aktual_ops) - total_penggunaan) < 1
assert abs((total_anggaran - total_penggunaan) - total_sisa) < 1
assert abs(total_sisa - net_sd) < 1
print("[VALIDASI] Pangan + Ops = Total Penggunaan: OK")
print("[VALIDASI] Anggaran - Penggunaan = Sisa: OK")
print("[VALIDASI] Sisa = Net Surplus/Defisit: OK")

# Weekly aggregation
def week_bucket(d_iso):
    y, m, d = map(int, d_iso.split("-"))
    dt = date(y, m, d)
    iso_year, iso_week, _ = dt.isocalendar()
    return iso_year, iso_week


weekly = {}
for r in records:
    key = week_bucket(r["date"])
    w = weekly.setdefault(key, {
        "start": r["date"], "end": r["date"],
        "pagu_pangan": 0, "aktual_pangan": 0,
        "pagu_ops": 0, "aktual_ops": 0,
        "total_pagu": 0, "total_aktual": 0,
    })
    w["end"] = r["date"]
    w["pagu_pangan"] += r["pagu_pangan"]
    w["aktual_pangan"] += r["aktual_pangan"]
    w["pagu_ops"] += r["pagu_ops"]
    w["aktual_ops"] += r["aktual_ops"]
    w["total_pagu"] += r["total_pagu"]
    w["total_aktual"] += r["total_aktual"]

weekly_list = [weekly[k] for k in sorted(weekly.keys())]
for w in weekly_list:
    w["surplus_defisit"] = w["total_pagu"] - w["total_aktual"]
    w["status"] = "SURPLUS" if w["surplus_defisit"] > 0 else ("DEFISIT" if w["surplus_defisit"] < 0 else "AMAN")


# ============================================================
# STYLES (numFmt / fonts / fills / borders / cellXfs)
# ============================================================
# Style ID map:
#  0  default
#  1  title mega (bold white on dark navy, merged)
#  2  subtitle (bold dark navy)
#  3  section header (bold white on blue, center)
#  4  KPI label (bold white on dark teal)
#  5  KPI value - big currency (bold, dark navy on light)
#  6  currency normal (right)
#  7  currency total (bold, gold bg)
#  8  percent 2 decimals (center)
#  9  table header (bold white on navy, center)
# 10  table cell currency (right, border)
# 11  table cell percent (center, border)
# 12  table cell text (center, border)
# 13  status SURPLUS (green bg, bold)
# 14  status DEFISIT (red bg, white, bold)
# 15  status AMAN (gray bg)
# 16  date cell (center border)
# 17  KPI value - big currency GREEN (surplus)
# 18  KPI value - big currency RED (defisit)
# 19  label small (left)
# 20  integer centered
# 21  light note italic
# 22  currency positive green (bold)
# 23  currency negative red (bold)
# 24  section header teal
# 25  section header orange
# 26  section header purple

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="5">
    <numFmt numFmtId="164" formatCode="&quot;Rp&quot;\\ #,##0;[Red]&quot;Rp&quot;\\ \\-#,##0"/>
    <numFmt numFmtId="165" formatCode="dd\\ mmm\\ yyyy"/>
    <numFmt numFmtId="166" formatCode="0.00%"/>
    <numFmt numFmtId="167" formatCode="#,##0"/>
    <numFmt numFmtId="168" formatCode="&quot;Rp&quot;\\ #,##0"/>
  </numFmts>
  <fonts count="14">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="20"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="13"/><color rgb="FF1F3864"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="16"/><color rgb="FF1F3864"/><name val="Calibri"/></font>
    <font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF1F3864"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF9C0006"/><name val="Calibri"/></font>
    <font><b/><sz val="16"/><color rgb="FF006100"/><name val="Calibri"/></font>
    <font><b/><sz val="16"/><color rgb="FF9C0006"/><name val="Calibri"/></font>
    <font><i/><sz val="9"/><color rgb="FF595959"/><name val="Calibri"/></font>
  </fonts>
  <fills count="14">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF0F6E8E"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFED7D31"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF7030A0"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFE699"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFBFBFBF"/></left>
      <right style="thin"><color rgb="FFBFBFBF"/></right>
      <top style="thin"><color rgb="FFBFBFBF"/></top>
      <bottom style="thin"><color rgb="FFBFBFBF"/></bottom>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="27">
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0"   fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0"   fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="4" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="5" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="6" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="164" fontId="7" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="7" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="8" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="6" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="6" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="6" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="9" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="10" fillId="8" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="7" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="165" fontId="6" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="11" fillId="13" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="12" fillId="8" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="167" fontId="5" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="13" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="164" fontId="9" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="164" fontId="10" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="4" fillId="10" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0"   fontId="4" fillId="11" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0"   fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"""


# ============================================================
# BUILDERS
# ============================================================

def cell(val, style=0, is_formula=False, is_date=False):
    return {"v": val, "s": style, "f": is_formula, "d": is_date}


def row_xml(r_idx, cells, height=None):
    attrs = 'r="{0}"'.format(r_idx)
    if height:
        attrs += ' ht="{0}" customHeight="1"'.format(height)
    out = ["<row {0}>".format(attrs)]
    for c_idx, c in enumerate(cells, start=1):
        if c is None:
            continue
        val = c["v"]
        style = c["s"]
        ref = "{0}{1}".format(col_letter(c_idx), r_idx)
        if val == "" or val is None:
            out.append('<c r="{0}" s="{1}"/>'.format(ref, style))
            continue
        if c["f"]:
            out.append('<c r="{0}" s="{1}"><f>{2}</f></c>'.format(ref, style, escape(str(val))))
        elif c["d"]:
            serial = iso_to_serial(val) if isinstance(val, str) else val
            out.append('<c r="{0}" s="{1}"><v>{2}</v></c>'.format(ref, style, serial))
        elif isinstance(val, (int, float)):
            out.append('<c r="{0}" s="{1}"><v>{2}</v></c>'.format(ref, style, val))
        else:
            out.append('<c r="{0}" s="{1}" t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>'.format(
                ref, style, escape(str(val))))
    out.append("</row>")
    return "".join(out)


# ============================================================
# SHEET 1: DASHBOARD
# ============================================================
# Layout (A..H columns, ~8 columns wide):
#  Row 1-2 : Title banner (merged A1:H2)
#  Row 3   : Periode
#  Row 5   : 4 KPI utama (merged 2-col each)
#  Row 6   : 4 KPI values
#  Row 8   : Section A header (Pangan)
#  Row 9-16: Pangan metrics (label | value)
#  Row 18  : Section B header (Ops)
#  Row 19-26: Ops metrics
#  Row 28  : Section C header (Ringkasan Total)
#  Row 29-35: Ringkasan
#  Row 37  : Section D header (Surplus/Defisit)
#  Row 38-47: Surplus/Defisit metrics
#  Row 50+ : chart area (charts placed via drawing)

def build_dashboard_sheet():
    rows = {}
    merges = []

    # TITLE BANNER
    rows[1] = [cell("", 1)] * 8
    rows[1][0] = cell("DASHBOARD AKUMULASI LAPORAN KEUANGAN", 1)
    rows[2] = [cell("", 1)] * 8
    rows[2][0] = cell("SPPG BATTUWINANGUN - Program Pemenuhan Gizi", 1)
    merges.append("A1:H1")
    merges.append("A2:H2")

    rows[3] = [cell("", 0)] * 8
    rows[3][0] = cell("Periode:", 2)
    rows[3][1] = cell("01 September 2025 s/d 31 Desember 2025", 2)
    merges.append("B3:H3")

    # ROW 5-6: 4 KPI BESAR
    rows[5] = [cell("", 0)] * 8
    rows[5][0] = cell("TOTAL DANA ANGGARAN", 4)
    rows[5][2] = cell("TOTAL DIPAKAI", 4)
    rows[5][4] = cell("TOTAL SISA DANA", 4)
    rows[5][6] = cell("% PENYERAPAN", 4)
    merges.extend(["A5:B5", "C5:D5", "E5:F5", "G5:H5"])

    rows[6] = [cell("", 0)] * 8
    rows[6][0] = cell(total_anggaran, 5)
    rows[6][2] = cell(total_penggunaan, 5)
    rows[6][4] = cell(total_sisa, 17 if total_sisa >= 0 else 18)
    rows[6][6] = cell(pct_penyerapan, 8)
    merges.extend(["A6:B6", "C6:D6", "E6:F6", "G6:H6"])

    # ROW 8: Section A
    rows[8] = [cell("", 0)] * 8
    rows[8][0] = cell("A. PENGELUARAN BAHAN PANGAN", 3)
    merges.append("A8:H8")

    pangan_rows = [
        ("Rata-Rata Pengeluaran Bahan Pangan / Hari", avg_pangan, "cur"),
        ("Pengeluaran Bahan Pangan Seharusnya (Pagu) / Hari", pagu_pangan_day, "cur"),
        ("Selisih Pengeluaran Bahan Pangan / Hari", selisih_pangan_day, "cur"),
        ("", "", "blank"),
        ("Dana Total Bahan Pangan", tot_pagu_pangan, "cur_total"),
        ("Penggunaan Dana Pangan", tot_aktual_pangan, "cur_total"),
        ("Sisa Dana Pangan", sisa_pangan, "cur_total"),
        ("% Penggunaan Dana Pangan", pct_pangan, "pct"),
    ]
    r = 9
    for label, val, kind in pangan_rows:
        rows[r] = [cell("", 0)] * 8
        if kind == "blank":
            r += 1
            continue
        rows[r][0] = cell(label, 26)
        if kind == "cur":
            rows[r][5] = cell(val, 6)
        elif kind == "cur_total":
            rows[r][5] = cell(val, 7)
        elif kind == "pct":
            rows[r][5] = cell(val, 8)
        merges.append("A{0}:E{0}".format(r))
        merges.append("F{0}:H{0}".format(r))
        r += 1

    # SECTION B (starts row 18)
    rows[18] = [cell("", 0)] * 8
    rows[18][0] = cell("B. PENGELUARAN OPERASIONAL", 24)
    merges.append("A18:H18")

    ops_rows = [
        ("Rata-Rata Pengeluaran Operasional / Hari", avg_ops, "cur"),
        ("Pengeluaran Operasional Seharusnya (Pagu) / Hari", pagu_ops_day, "cur"),
        ("Selisih Pengeluaran Operasional / Hari", selisih_ops_day, "cur"),
        ("", "", "blank"),
        ("Dana Total Operasional", tot_pagu_ops, "cur_total"),
        ("Penggunaan Dana Operasional", tot_aktual_ops, "cur_total"),
        ("Sisa Dana Operasional", sisa_ops, "cur_total"),
        ("% Penggunaan Dana Operasional", pct_ops, "pct"),
    ]
    r = 19
    for label, val, kind in ops_rows:
        rows[r] = [cell("", 0)] * 8
        if kind == "blank":
            r += 1
            continue
        rows[r][0] = cell(label, 26)
        if kind == "cur":
            rows[r][5] = cell(val, 6)
        elif kind == "cur_total":
            rows[r][5] = cell(val, 7)
        elif kind == "pct":
            rows[r][5] = cell(val, 8)
        merges.append("A{0}:E{0}".format(r))
        merges.append("F{0}:H{0}".format(r))
        r += 1

    # SECTION C (starts row 28)
    rows[28] = [cell("", 0)] * 8
    rows[28][0] = cell("C. RINGKASAN TOTAL", 25)
    merges.append("A28:H28")

    ring_rows = [
        ("Total Dana Anggaran", total_anggaran, "cur_total"),
        ("Total Penggunaan Dana", total_penggunaan, "cur_total"),
        ("Total Sisa Dana", total_sisa, "cur_total"),
        ("% Penyerapan Anggaran", pct_penyerapan, "pct"),
        ("", "", "blank"),
        ("Rata-Rata Total Pengeluaran / Hari", avg_total_day, "cur"),
        ("Pagu Total / Hari (Seharusnya)", pagu_total_day, "cur"),
        ("Selisih Total / Hari", selisih_total_day, "cur"),
    ]
    r = 29
    for label, val, kind in ring_rows:
        rows[r] = [cell("", 0)] * 8
        if kind == "blank":
            r += 1
            continue
        rows[r][0] = cell(label, 26)
        if kind == "cur":
            rows[r][5] = cell(val, 6)
        elif kind == "cur_total":
            rows[r][5] = cell(val, 7)
        elif kind == "pct":
            rows[r][5] = cell(val, 8)
        merges.append("A{0}:E{0}".format(r))
        merges.append("F{0}:H{0}".format(r))
        r += 1

    # SECTION D (starts row 38)
    rows[38] = [cell("", 0)] * 8
    rows[38][0] = cell("D. ANALISIS SURPLUS / DEFISIT", 3)
    merges.append("A38:H38")

    sd_rows = [
        ("Jumlah Hari SURPLUS (Pengeluaran < Pagu)", n_surplus, "int"),
        ("Jumlah Hari DEFISIT (Pengeluaran > Pagu)", n_defisit, "int"),
        ("Jumlah Hari AMAN (Pengeluaran = Pagu / Libur)", n_aman, "int"),
        ("Jumlah Hari Belum Input", n_belum, "int"),
        ("", "", "blank"),
        ("Total Akumulasi Surplus", tot_surplus, "cur_pos"),
        ("Total Akumulasi Defisit", tot_defisit, "cur_neg"),
        ("Net Surplus / Defisit", net_sd, "cur_total"),
        ("% Surplus Total (dari Total Pagu)", pct_surplus_total, "pct"),
        ("", "", "blank"),
        ("Pengeluaran Tertinggi / Hari (Rp)", max_r["total_aktual"], "cur_total"),
        ("   Tanggal Pengeluaran Tertinggi", max_r["date"], "date"),
        ("Pengeluaran Terendah / Hari (Rp)", min_r["total_aktual"], "cur_total"),
        ("   Tanggal Pengeluaran Terendah", min_r["date"], "date"),
    ]
    r = 39
    for label, val, kind in sd_rows:
        rows[r] = [cell("", 0)] * 8
        if kind == "blank":
            r += 1
            continue
        rows[r][0] = cell(label, 26)
        if kind == "cur":
            rows[r][5] = cell(val, 6)
        elif kind == "cur_total":
            rows[r][5] = cell(val, 7)
        elif kind == "cur_pos":
            rows[r][5] = cell(val, 22)
        elif kind == "cur_neg":
            rows[r][5] = cell(val, 23)
        elif kind == "pct":
            rows[r][5] = cell(val, 8)
        elif kind == "int":
            rows[r][5] = cell(val, 20)
        elif kind == "date":
            rows[r][5] = cell(val, 16, is_date=True)
        merges.append("A{0}:E{0}".format(r))
        merges.append("F{0}:H{0}".format(r))
        r += 1

    # CATATAN / LEGENDA (bottom)
    rows[55] = [cell("CATATAN:", 2)] + [cell("", 0)] * 7
    rows[56] = [cell("- Dashboard akumulatif seluruh periode 01 Sep s/d 31 Des 2025 (97 hari, 84 hari aktif).", 0)] + [cell("", 0)] * 7
    rows[57] = [cell("- Semua nilai Rupiah sudah divalidasi silang (Pagu - Penggunaan = Sisa = Net Surplus/Defisit).", 0)] + [cell("", 0)] * 7
    rows[58] = [cell("- Selisih POSITIF = HEMAT (surplus). Selisih NEGATIF = BOROS (defisit).", 0)] + [cell("", 0)] * 7
    rows[59] = [cell("- Sumber data: 8 file 'Total Akumulasi *.xlsx' di repository.", 0)] + [cell("", 0)] * 7
    for k in (55, 56, 57, 58, 59):
        merges.append("A{0}:H{0}".format(k))

    # Column widths
    cols = [
        (1, 42), (2, 12), (3, 12), (4, 12), (5, 12), (6, 20), (7, 12), (8, 12)
    ]

    # Row heights for visual impact
    row_heights = {1: 32, 2: 22, 3: 22, 5: 22, 6: 40, 8: 24, 18: 24, 28: 24, 38: 24}

    return rows, merges, cols, row_heights


# ============================================================
# SHEET 2: REKAP HARIAN
# ============================================================
def build_harian_sheet():
    rows = {}
    merges = []

    rows[1] = [cell("REKAP HARIAN PENGELUARAN - 01 Sep s/d 31 Des 2025", 1)] + [cell("", 1)] * 10
    merges.append("A1:K1")

    headers = ["Tanggal", "Hari", "Pagu Pangan", "Aktual Pangan", "Selisih Pangan",
               "Pagu Operasional", "Aktual Operasional", "Selisih Ops",
               "Total Pagu", "Total Aktual", "Status"]
    rows[3] = [cell(h, 9) for h in headers]

    r = 4
    for rec in records:
        status_style = 13 if rec["status"] == "SURPLUS" else (14 if rec["status"] == "DEFISIT" else 15)
        rows[r] = [
            cell(rec["date"], 16, is_date=True),
            cell(rec["day_name"], 12),
            cell(rec["pagu_pangan"], 10),
            cell(rec["aktual_pangan"], 10),
            cell(rec["pagu_pangan"] - rec["aktual_pangan"], 10),
            cell(rec["pagu_ops"], 10),
            cell(rec["aktual_ops"], 10),
            cell(rec["pagu_ops"] - rec["aktual_ops"], 10),
            cell(rec["total_pagu"], 10),
            cell(rec["total_aktual"], 10),
            cell(rec["status"], status_style),
        ]
        r += 1

    # Total row
    rows[r+1] = [
        cell("TOTAL", 7), cell("", 7),
        cell(tot_pagu_pangan, 7), cell(tot_aktual_pangan, 7), cell(tot_pagu_pangan - tot_aktual_pangan, 7),
        cell(tot_pagu_ops, 7), cell(tot_aktual_ops, 7), cell(tot_pagu_ops - tot_aktual_ops, 7),
        cell(total_anggaran, 7), cell(total_penggunaan, 7), cell("", 7),
    ]

    cols = [(1, 14), (2, 12), (3, 15), (4, 15), (5, 15),
            (6, 15), (7, 15), (8, 13), (9, 15), (10, 15), (11, 11)]
    row_heights = {1: 30, 3: 30}
    return rows, merges, cols, row_heights


# ============================================================
# SHEET 3: REKAP MINGGUAN
# ============================================================
def build_mingguan_sheet():
    rows = {}
    merges = []
    rows[1] = [cell("REKAP MINGGUAN - 01 Sep s/d 31 Des 2025", 1)] + [cell("", 1)] * 9
    merges.append("A1:J1")

    headers = ["Minggu Ke", "Periode", "Pagu Pangan", "Aktual Pangan",
               "Pagu Ops", "Aktual Ops", "Total Pagu", "Total Aktual",
               "Surplus/Defisit", "Status"]
    rows[3] = [cell(h, 9) for h in headers]

    r = 4
    for i, w in enumerate(weekly_list, start=1):
        status_style = 13 if w["status"] == "SURPLUS" else (14 if w["status"] == "DEFISIT" else 15)
        periode = "{0} s/d {1}".format(w["start"], w["end"])
        rows[r] = [
            cell(i, 12),
            cell(periode, 12),
            cell(w["pagu_pangan"], 10),
            cell(w["aktual_pangan"], 10),
            cell(w["pagu_ops"], 10),
            cell(w["aktual_ops"], 10),
            cell(w["total_pagu"], 10),
            cell(w["total_aktual"], 10),
            cell(w["surplus_defisit"], 10),
            cell(w["status"], status_style),
        ]
        r += 1

    rows[r+1] = [
        cell("TOTAL", 7), cell("", 7),
        cell(tot_pagu_pangan, 7), cell(tot_aktual_pangan, 7),
        cell(tot_pagu_ops, 7), cell(tot_aktual_ops, 7),
        cell(total_anggaran, 7), cell(total_penggunaan, 7),
        cell(total_sisa, 7), cell("", 7),
    ]

    cols = [(1, 10), (2, 26), (3, 15), (4, 15), (5, 15), (6, 15),
            (7, 16), (8, 16), (9, 16), (10, 11)]
    row_heights = {1: 30, 3: 30}
    return rows, merges, cols, row_heights


# ============================================================
# SHEET 4: CHART DATA (source utk chart line & pie)
# ============================================================
def build_chart_data_sheet():
    rows = {}
    merges = []
    rows[1] = [cell("CHART DATA (Sumber Grafik)", 1)] + [cell("", 1)] * 4
    merges.append("A1:E1")

    # Line Chart Data: Tanggal | Total Pagu | Total Aktual
    rows[3] = [cell("Tanggal", 9), cell("Total Pagu", 9), cell("Total Aktual", 9)]
    r = 4
    for rec in records:
        rows[r] = [
            cell(rec["date"], 16, is_date=True),
            cell(rec["total_pagu"], 10),
            cell(rec["total_aktual"], 10),
        ]
        r += 1

    # Space
    pie_start = r + 3

    # Pie data: Komposisi Penyerapan
    rows[pie_start] = [cell("Kategori", 9), cell("Nilai", 9)]
    rows[pie_start + 1] = [cell("Penggunaan Pangan", 12), cell(tot_aktual_pangan, 10)]
    rows[pie_start + 2] = [cell("Penggunaan Operasional", 12), cell(tot_aktual_ops, 10)]
    rows[pie_start + 3] = [cell("Sisa Dana", 12), cell(total_sisa if total_sisa > 0 else 0, 10)]

    # Status bar data
    bar_start = pie_start + 6
    rows[bar_start] = [cell("Status Hari", 9), cell("Jumlah", 9)]
    rows[bar_start + 1] = [cell("SURPLUS", 13), cell(n_surplus, 20)]
    rows[bar_start + 2] = [cell("DEFISIT", 14), cell(n_defisit, 20)]
    rows[bar_start + 3] = [cell("AMAN", 15), cell(n_aman, 20)]

    cols = [(1, 14), (2, 18), (3, 18), (4, 18), (5, 14)]
    row_heights = {1: 28, 3: 24}

    # Return chart source ranges for later
    n_days = len(records)
    line_range = {
        "cat": "'Chart Data'!$A$4:$A${0}".format(3 + n_days),
        "pagu": "'Chart Data'!$B$4:$B${0}".format(3 + n_days),
        "aktual": "'Chart Data'!$C$4:$C${0}".format(3 + n_days),
        "pagu_hdr": "'Chart Data'!$B$3",
        "aktual_hdr": "'Chart Data'!$C$3",
    }
    pie_range = {
        "cat": "'Chart Data'!$A${0}:$A${1}".format(pie_start + 1, pie_start + 3),
        "val": "'Chart Data'!$B${0}:$B${1}".format(pie_start + 1, pie_start + 3),
    }
    bar_range = {
        "cat": "'Chart Data'!$A${0}:$A${1}".format(bar_start + 1, bar_start + 3),
        "val": "'Chart Data'!$B${0}:$B${1}".format(bar_start + 1, bar_start + 3),
    }
    return rows, merges, cols, row_heights, line_range, pie_range, bar_range


# ============================================================
# ASSEMBLE SHEET XML
# ============================================================
def sheet_xml(rows, merges, cols, row_heights, drawing_rid=None, show_grid=True):
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    parts.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    if not show_grid:
        parts.append('<sheetViews><sheetView workbookViewId="0" showGridLines="0"><selection activeCell="A1"/></sheetView></sheetViews>')
    else:
        parts.append('<sheetViews><sheetView workbookViewId="0"><selection activeCell="A1"/></sheetView></sheetViews>')
    parts.append('<sheetFormatPr defaultRowHeight="15"/>')

    if cols:
        parts.append("<cols>")
        for idx, w in cols:
            parts.append('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>'.format(idx, w))
        parts.append("</cols>")

    parts.append("<sheetData>")
    for r_idx in sorted(rows.keys()):
        h = row_heights.get(r_idx)
        parts.append(row_xml(r_idx, rows[r_idx], height=h))
    parts.append("</sheetData>")

    if merges:
        parts.append('<mergeCells count="{0}">'.format(len(merges)))
        for m in merges:
            parts.append('<mergeCell ref="{0}"/>'.format(m))
        parts.append("</mergeCells>")

    if drawing_rid:
        parts.append('<drawing r:id="{0}"/>'.format(drawing_rid))

    parts.append("</worksheet>")
    return "".join(parts)


# ============================================================
# CHART XML - LINE CHART (Total Pagu vs Aktual per hari)
# ============================================================
def build_line_chart_xml(line_range):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<c:chart>'
            '<c:title><c:tx><c:rich>'
            '<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" wrap="square" anchor="ctr" anchorCtr="1"/>'
            '<a:lstStyle/>'
            '<a:p><a:pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:defRPr></a:pPr>'
            '<a:r><a:rPr lang="id-ID" sz="1400" b="1"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:rPr>'
            '<a:t>TREN HARIAN: Pagu vs Pengeluaran Aktual</a:t></a:r></a:p>'
            '</c:rich></c:tx><c:overlay val="0"/></c:title>'
            '<c:autoTitleDeleted val="0"/>'
            '<c:plotArea><c:layout/>'
            '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
            # Series 1: Total Pagu
            '<c:ser><c:idx val="0"/><c:order val="0"/>'
            '<c:tx><c:strRef><c:f>' + line_range["pagu_hdr"] + '</c:f>'
            '<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Total Pagu</c:v></c:pt></c:strCache>'
            '</c:strRef></c:tx>'
            '<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></a:ln></c:spPr>'
            '<c:marker><c:symbol val="none"/></c:marker>'
            '<c:cat><c:numRef><c:f>' + line_range["cat"] + '</c:f></c:numRef></c:cat>'
            '<c:val><c:numRef><c:f>' + line_range["pagu"] + '</c:f></c:numRef></c:val>'
            '<c:smooth val="0"/></c:ser>'
            # Series 2: Total Aktual
            '<c:ser><c:idx val="1"/><c:order val="1"/>'
            '<c:tx><c:strRef><c:f>' + line_range["aktual_hdr"] + '</c:f>'
            '<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Total Aktual</c:v></c:pt></c:strCache>'
            '</c:strRef></c:tx>'
            '<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln></c:spPr>'
            '<c:marker><c:symbol val="circle"/><c:size val="5"/>'
            '<c:spPr><a:solidFill><a:srgbClr val="C00000"/></a:solidFill>'
            '<a:ln><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln></c:spPr></c:marker>'
            '<c:cat><c:numRef><c:f>' + line_range["cat"] + '</c:f></c:numRef></c:cat>'
            '<c:val><c:numRef><c:f>' + line_range["aktual"] + '</c:f></c:numRef></c:val>'
            '<c:smooth val="0"/></c:ser>'
            '<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>'
            '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
            '<c:delete val="0"/><c:axPos val="b"/>'
            '<c:numFmt formatCode="dd-mmm" sourceLinked="0"/>'
            '<c:majorTickMark val="out"/><c:minorTickMark val="none"/>'
            '<c:tickLblPos val="nextTo"/><c:crossAx val="2"/></c:catAx>'
            '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
            '<c:delete val="0"/><c:axPos val="l"/>'
            '<c:numFmt formatCode="#,##0" sourceLinked="0"/>'
            '<c:majorTickMark val="out"/><c:minorTickMark val="none"/>'
            '<c:tickLblPos val="nextTo"/><c:crossAx val="1"/></c:valAx>'
            '</c:plotArea>'
            '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
            '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>'
            '</c:chart></c:chartSpace>')


# ============================================================
# CHART XML - DOUGHNUT (Komposisi Penyerapan)
# ============================================================
def build_doughnut_chart_xml(pie_range):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<c:chart>'
            '<c:title><c:tx><c:rich>'
            '<a:bodyPr/><a:lstStyle/>'
            '<a:p><a:pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:defRPr></a:pPr>'
            '<a:r><a:rPr lang="id-ID" sz="1400" b="1"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:rPr>'
            '<a:t>KOMPOSISI PENYERAPAN ANGGARAN</a:t></a:r></a:p>'
            '</c:rich></c:tx><c:overlay val="0"/></c:title>'
            '<c:autoTitleDeleted val="0"/>'
            '<c:plotArea><c:layout/>'
            '<c:doughnutChart><c:varyColors val="1"/>'
            '<c:ser><c:idx val="0"/><c:order val="0"/>'
            '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/>'
            '<c:spPr><a:solidFill><a:srgbClr val="2E75B6"/></a:solidFill>'
            '<a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>'
            '<c:dPt><c:idx val="1"/><c:bubble3D val="0"/>'
            '<c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill>'
            '<a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>'
            '<c:dPt><c:idx val="2"/><c:bubble3D val="0"/>'
            '<c:spPr><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill>'
            '<a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>'
            '<c:dLbls><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>'
            '<c:dLblPos val="ctr"/><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/></c:dLbls>'
            '<c:cat><c:strRef><c:f>' + pie_range["cat"] + '</c:f></c:strRef></c:cat>'
            '<c:val><c:numRef><c:f>' + pie_range["val"] + '</c:f></c:numRef></c:val>'
            '</c:ser>'
            '<c:firstSliceAng val="0"/><c:holeSize val="50"/>'
            '</c:doughnutChart></c:plotArea>'
            '<c:legend><c:legendPos val="b"/><c:overlay val="0"/>'
            '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000" b="1"/></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>'
            '</c:legend>'
            '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>'
            '</c:chart></c:chartSpace>')


# ============================================================
# DRAWING XML (places charts onto Dashboard sheet)
# ============================================================
def build_drawing_xml():
    # Line chart: anchored A50:H68, Doughnut: below it (A70:H87)... we put side-by-side
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            # Line chart anchor
            '<xdr:twoCellAnchor>'
            '<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>61</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            '<xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>80</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            '<xdr:graphicFrame macro="">'
            '<xdr:nvGraphicFramePr>'
            '<xdr:cNvPr id="2" name="Line Chart"/>'
            '<xdr:cNvGraphicFramePr/>'
            '</xdr:nvGraphicFramePr>'
            '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>'
            '</a:graphicData></a:graphic>'
            '</xdr:graphicFrame>'
            '<xdr:clientData/>'
            '</xdr:twoCellAnchor>'
            # Doughnut chart anchor
            '<xdr:twoCellAnchor>'
            '<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>81</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            '<xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>100</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            '<xdr:graphicFrame macro="">'
            '<xdr:nvGraphicFramePr>'
            '<xdr:cNvPr id="3" name="Doughnut Chart"/>'
            '<xdr:cNvGraphicFramePr/>'
            '</xdr:nvGraphicFramePr>'
            '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/>'
            '</a:graphicData></a:graphic>'
            '</xdr:graphicFrame>'
            '<xdr:clientData/>'
            '</xdr:twoCellAnchor>'
            '</xdr:wsDr>')


# ============================================================
# CONTENT TYPES / RELATIONSHIPS
# ============================================================
def build_content_types():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
            '<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
            '</Types>')


ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
             '</Relationships>')


def build_workbook_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets>'
            '<sheet name="Dashboard" sheetId="1" r:id="rId1"/>'
            '<sheet name="Rekap Harian" sheetId="2" r:id="rId2"/>'
            '<sheet name="Rekap Mingguan" sheetId="3" r:id="rId3"/>'
            '<sheet name="Chart Data" sheetId="4" r:id="rId4"/>'
            '</sheets></workbook>')


def build_workbook_rels():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>'
            '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>'
            '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>')


# Sheet1 (Dashboard) rels -> drawing1
SHEET1_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
               '</Relationships>')

# Drawing1 rels -> chart1, chart2
DRAWING_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>'
                '</Relationships>')


# ============================================================
# WRITE FILE
# ============================================================
def main():
    dash_rows, dash_merges, dash_cols, dash_heights = build_dashboard_sheet()
    harian_rows, harian_merges, harian_cols, harian_heights = build_harian_sheet()
    ming_rows, ming_merges, ming_cols, ming_heights = build_mingguan_sheet()
    cd_rows, cd_merges, cd_cols, cd_heights, line_range, pie_range, bar_range = build_chart_data_sheet()

    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", build_content_types())
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", build_workbook_xml())
        z.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels())
        z.writestr("xl/styles.xml", STYLES_XML)

        z.writestr("xl/worksheets/sheet1.xml",
                   sheet_xml(dash_rows, dash_merges, dash_cols, dash_heights,
                             drawing_rid="rId1", show_grid=False))
        z.writestr("xl/worksheets/_rels/sheet1.xml.rels", SHEET1_RELS)

        z.writestr("xl/worksheets/sheet2.xml",
                   sheet_xml(harian_rows, harian_merges, harian_cols, harian_heights))
        z.writestr("xl/worksheets/sheet3.xml",
                   sheet_xml(ming_rows, ming_merges, ming_cols, ming_heights))
        z.writestr("xl/worksheets/sheet4.xml",
                   sheet_xml(cd_rows, cd_merges, cd_cols, cd_heights))

        z.writestr("xl/drawings/drawing1.xml", build_drawing_xml())
        z.writestr("xl/drawings/_rels/drawing1.xml.rels", DRAWING_RELS)
        z.writestr("xl/charts/chart1.xml", build_line_chart_xml(line_range))
        z.writestr("xl/charts/chart2.xml", build_doughnut_chart_xml(pie_range))

    print("OK -> {0}".format(OUTPUT))
    print("Size: {0:,} bytes".format(os.path.getsize(OUTPUT)))


if __name__ == "__main__":
    main()
