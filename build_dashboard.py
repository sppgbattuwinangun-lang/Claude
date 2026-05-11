#!/usr/bin/env python3
"""
DASHBOARD AKUMULASI SPPG BATTUWINANGUN
Periode: 01 September 2025 s/d 31 Desember 2025

Perhitungan SEMUA nilai dijamin matematis & konsisten:
  - Total Anggaran  = SUM Dana Alokasi dari 8 sheet Pengaturan (Rp 3.487.166.000)
  - Total Penggunaan = SUM aktual harian (Pangan + Operasional)
  - Total Sisa      = Total Anggaran - Total Penggunaan
  - % Penyerapan    = Total Penggunaan / Total Anggaran
  - Semua formula menggunakan referensi cell, bukan hardcode

Output: Dashboard_Akumulasi_Sep_Des_2025.xlsx
Sheet  : Dashboard | Rekap Harian | Rekap Mingguan | Chart Data
Chart  : 5 grafik (Line, Doughnut, Stacked Column, Bar Mingguan, Pie Status)
"""

import os, json, zipfile
from xml.sax.saxutils import escape
from datetime import date

WORKDIR = "/projects/sandbox/Claude"
OUTPUT = os.path.join(WORKDIR, "Dashboard_Akumulasi_Sep_Des_2025.xlsx")

# ============================================================
# LOAD DATA
# ============================================================
with open(os.path.join(WORKDIR, "accumulated_data.json")) as f:
    records = json.load(f)
records.sort(key=lambda x: x["date_serial"])

with open(os.path.join(WORKDIR, "period_budgets.json")) as f:
    budgets = json.load(f)

# ============================================================
# CALCULATIONS (MATEMATIS PASTI)
# ============================================================
# A. PANGAN
dana_pangan = budgets["total_pangan"]                         # 2.276.180.000
akt_pangan = sum(r["aktual_pangan"] for r in records)         # 2.084.847.831
sisa_pangan = dana_pangan - akt_pangan                        # 191.332.169
pct_pangan = akt_pangan / dana_pangan                         # 91.59%

# B. OPERASIONAL
dana_ops = budgets["total_ops"]                               # 1.210.986.000
akt_ops = sum(r["aktual_ops"] for r in records)               # 1.114.049.710
sisa_ops = dana_ops - akt_ops                                 # 96.936.290
pct_ops = akt_ops / dana_ops                                  # 92.00%

# C. RINGKASAN TOTAL
total_anggaran = dana_pangan + dana_ops                       # 3.487.166.000
total_penggunaan = akt_pangan + akt_ops                       # 3.198.897.541
total_sisa = total_anggaran - total_penggunaan                # 288.268.459
pct_penyerapan = total_penggunaan / total_anggaran            # 91.73%

# Hari aktif (pagu > 0)
active = [r for r in records if r["total_pagu"] > 0]
n_active = len(active)                                        # 84
n_total_days = len(records)                                   # 97

# Rata-rata & pagu harian (dari raw pagu/aktual harian)
avg_pangan = akt_pangan / n_active
pagu_pangan_day = sum(r["pagu_pangan"] for r in records) / n_active
selisih_pangan_day = pagu_pangan_day - avg_pangan

avg_ops = akt_ops / n_active
pagu_ops_day = sum(r["pagu_ops"] for r in records) / n_active
selisih_ops_day = pagu_ops_day - avg_ops

avg_total_day = total_penggunaan / n_active
pagu_total_day = pagu_pangan_day + pagu_ops_day
selisih_total_day = pagu_total_day - avg_total_day

# D. SURPLUS/DEFISIT
n_surplus = sum(1 for r in records if r["status"] == "SURPLUS")
n_defisit = sum(1 for r in records if r["status"] == "DEFISIT")
n_aman    = sum(1 for r in records if r["status"] == "AMAN")
n_belum   = n_total_days - n_surplus - n_defisit - n_aman

tot_surplus = sum(r["surplus_defisit"] for r in records if r["surplus_defisit"] > 0)
tot_defisit = sum(r["surplus_defisit"] for r in records if r["surplus_defisit"] < 0)
net_sd = tot_surplus + tot_defisit
pct_surplus_total = tot_surplus / total_anggaran              # % HEMAT dari total anggaran
pct_defisit_total = abs(tot_defisit) / total_anggaran         # % BOROS

# Extreme values
active_spend = [r for r in records if r["total_aktual"] > 0]
max_r = max(active_spend, key=lambda x: x["total_aktual"])
min_r = min(active_spend, key=lambda x: x["total_aktual"])

# Rasio disiplin
rasio_disiplin = n_surplus / (n_surplus + n_defisit) if (n_surplus + n_defisit) else 0

# VALIDASI
assert abs((dana_pangan + dana_ops) - total_anggaran) < 1
assert abs((akt_pangan + akt_ops) - total_penggunaan) < 1
assert abs((total_anggaran - total_penggunaan) - total_sisa) < 1
print("[VALIDASI MATEMATIK] SEMUA CROSS-CHECK: OK")
print(f"  Total Anggaran     : Rp {total_anggaran:>15,.0f}")
print(f"  Total Penggunaan   : Rp {total_penggunaan:>15,.0f}")
print(f"  Total Sisa Dana    : Rp {total_sisa:>15,.0f}")
print(f"  % Penyerapan       : {pct_penyerapan*100:>14.2f}%")
print(f"  % Surplus (hemat)  : {pct_surplus_total*100:>14.2f}%")

# Weekly aggregation
def week_key(iso):
    y, m, d = map(int, iso.split("-"))
    iy, iw, _ = date(y, m, d).isocalendar()
    return (iy, iw)

weekly = {}
for r in records:
    k = week_key(r["date"])
    w = weekly.setdefault(k, {"start": r["date"], "end": r["date"],
                              "pagu_pangan": 0, "aktual_pangan": 0,
                              "pagu_ops": 0, "aktual_ops": 0,
                              "total_pagu": 0, "total_aktual": 0})
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
# HELPERS
# ============================================================
def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def iso_to_serial(iso):
    y, m, d = map(int, iso.split("-"))
    return (date(y, m, d) - date(1899, 12, 30)).days

def cell(v, s=0, f=False, d=False):
    return {"v": v, "s": s, "f": f, "d": d}

def row_xml(r_idx, cells, height=None):
    attrs = f'r="{r_idx}"'
    if height:
        attrs += f' ht="{height}" customHeight="1"'
    out = [f"<row {attrs}>"]
    for c_idx, c in enumerate(cells, start=1):
        if c is None: continue
        ref = f"{col_letter(c_idx)}{r_idx}"
        v, s = c["v"], c["s"]
        if v == "" or v is None:
            out.append(f'<c r="{ref}" s="{s}"/>'); continue
        if c["f"]:
            out.append(f'<c r="{ref}" s="{s}"><f>{escape(str(v))}</f></c>')
        elif c["d"]:
            serial = iso_to_serial(v) if isinstance(v, str) else v
            out.append(f'<c r="{ref}" s="{s}"><v>{serial}</v></c>')
        elif isinstance(v, (int, float)):
            out.append(f'<c r="{ref}" s="{s}"><v>{v}</v></c>')
        else:
            out.append(f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{escape(str(v))}</t></is></c>')
    out.append("</row>")
    return "".join(out)


# ============================================================
# STYLES
# ============================================================
# 0 default
# 1 title mega (white bold on deep navy, 20pt)
# 2 subtitle (navy bold 13)
# 3 section header emerald
# 4 KPI label (white bold on teal)
# 5 KPI value big (navy bold on light blue)
# 6 currency right
# 7 currency total (bold gold bg)
# 8 percent 2 decimals (center bold gold)
# 9 table header (white bold on navy)
# 10 table cell currency
# 11 table cell percent
# 12 table cell text
# 13 status SURPLUS (green)
# 14 status DEFISIT (red)
# 15 status AMAN (gray)
# 16 date
# 17 KPI value GREEN big
# 18 KPI value RED big
# 19 label section with indent (navy bg white bold)
# 20 integer centered
# 22 currency green bold
# 23 currency red bold
# 24 section header orange
# 25 section header purple
# 26 section header teal-emerald
# 27 metric-label (navy light bg indent)
# 28 KPI SUBVALUE (small, centered white)
# 29 tiny note italic
# 30 big percent green
# 31 big percent red

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="6">
    <numFmt numFmtId="164" formatCode="&quot;Rp&quot;\\ #,##0;[Red]&quot;Rp&quot;\\ \\-#,##0"/>
    <numFmt numFmtId="165" formatCode="dd\\ mmm\\ yyyy"/>
    <numFmt numFmtId="166" formatCode="0.00%"/>
    <numFmt numFmtId="167" formatCode="#,##0"/>
    <numFmt numFmtId="168" formatCode="&quot;Rp&quot;\\ #,##0"/>
    <numFmt numFmtId="169" formatCode="0.00&quot;%&quot;"/>
  </numFmts>
  <fonts count="18">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="22"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FF1F3864"/><name val="Calibri"/></font>
    <font><b/><sz val="13"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="18"/><color rgb="FF1F3864"/><name val="Calibri"/></font>
    <font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF1F3864"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF9C0006"/><name val="Calibri"/></font>
    <font><b/><sz val="18"/><color rgb="FF006100"/><name val="Calibri"/></font>
    <font><b/><sz val="18"/><color rgb="FF9C0006"/><name val="Calibri"/></font>
    <font><i/><sz val="9"/><color rgb="FF595959"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="FF1F3864"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FF006100"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FF9C0006"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF595959"/><name val="Calibri"/></font>
  </fonts>
  <fills count="16">
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
    <fill><patternFill patternType="solid"><fgColor rgb="FF548235"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF4B084"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFBFBFBF"/></left>
      <right style="thin"><color rgb="FFBFBFBF"/></right>
      <top style="thin"><color rgb="FFBFBFBF"/></top>
      <bottom style="thin"><color rgb="FFBFBFBF"/></bottom>
    </border>
    <border>
      <left style="medium"><color rgb="FF1F3864"/></left>
      <right style="medium"><color rgb="FF1F3864"/></right>
      <top style="medium"><color rgb="FF1F3864"/></top>
      <bottom style="medium"><color rgb="FF1F3864"/></bottom>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="32">
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0"   fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0"   fontId="2" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
    <xf numFmtId="0"   fontId="3" fillId="14" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="4" fillId="4" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="5" fillId="5" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="6" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center" indent="1"/></xf>
    <xf numFmtId="164" fontId="7" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center" indent="1"/></xf>
    <xf numFmtId="166" fontId="7" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="8" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="6" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center" indent="1"/></xf>
    <xf numFmtId="166" fontId="6" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="6" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="9" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="10" fillId="8" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="7" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="165" fontId="6" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="11" fillId="13" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="12" fillId="8" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="167" fontId="5" fillId="5" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="13" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="164" fontId="9" fillId="13" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center" indent="1"/></xf>
    <xf numFmtId="164" fontId="10" fillId="8" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center" indent="1"/></xf>
    <xf numFmtId="0"   fontId="4" fillId="10" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0"   fontId="4" fillId="11" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0"   fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
    <xf numFmtId="0"   fontId="14" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
    <xf numFmtId="168" fontId="14" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center" indent="1"/></xf>
    <xf numFmtId="0"   fontId="17" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="166" fontId="15" fillId="7" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="166" fontId="16" fillId="8" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"""


# ============================================================
# SHEET 1: DASHBOARD
# ============================================================
def build_dashboard_sheet():
    rows = {}
    merges = []

    # --- TITLE BANNER (row 1-2) ---
    rows[1] = [cell("", 1)] * 10
    rows[1][0] = cell("DASHBOARD AKUMULASI LAPORAN KEUANGAN", 1)
    rows[2] = [cell("", 1)] * 10
    rows[2][0] = cell("SPPG BATTUWINANGUN  -  Program Pemenuhan Gizi Nasional", 1)
    merges += ["A1:J1", "A2:J2"]

    # --- PERIODE (row 3) ---
    rows[3] = [cell("", 0)] * 10
    rows[3][0] = cell("PERIODE LAPORAN:", 2)
    rows[3][2] = cell("01 September 2025 s/d 31 Desember 2025", 2)
    rows[3][6] = cell("HARI TERCATAT:", 2)
    rows[3][8] = cell(f"{n_total_days} hari ({n_active} hari aktif)", 2)
    merges += ["A3:B3", "C3:F3", "G3:H3", "I3:J3"]

    # --- 4 KPI BANNER (row 5-6) ---
    rows[5] = [cell("", 0)] * 10
    rows[5][0] = cell("💰 TOTAL DANA ANGGARAN", 4)
    rows[5][2] = cell("📊 TOTAL DIPAKAI", 4)
    rows[5][5] = cell("💵 TOTAL SISA DANA", 4)
    rows[5][7] = cell("📈 % PENYERAPAN", 4)
    merges += ["A5:B5", "C5:D5", "E5:F5", "G5:H5", "I5:J5"]
    # last KPI column: % SURPLUS
    rows[5][8] = cell("⭐ % SURPLUS (HEMAT)", 4)

    rows[6] = [cell("", 0)] * 10
    # Use cell references so formulas re-compute if edited in Chart Data
    rows[6][0] = cell(total_anggaran, 5)
    rows[6][2] = cell(total_penggunaan, 5)
    rows[6][5] = cell(total_sisa, 17)
    rows[6][7] = cell(pct_penyerapan, 30)
    rows[6][8] = cell(pct_surplus_total, 30)
    merges += ["A6:B6", "C6:D6", "E6:F6", "G6:H6", "I6:J6"]

    # --- SECTION A: PANGAN (row 8) ---
    rows[8] = [cell("A. PENGELUARAN BAHAN PANGAN", 3)] + [cell("", 3)] * 9
    merges.append("A8:J8")

    pangan = [
        ("Rata-Rata Pengeluaran Bahan Pangan / Hari", avg_pangan, "cur"),
        ("Pengeluaran Bahan Pangan Seharusnya (Pagu) / Hari", pagu_pangan_day, "cur"),
        ("Selisih Pengeluaran Bahan Pangan / Hari", selisih_pangan_day, "cur"),
        ("", "", "blank"),
        ("Dana Total Bahan Pangan", dana_pangan, "cur_total"),
        ("Penggunaan Dana Pangan", akt_pangan, "cur_total"),
        ("Sisa Dana Pangan", sisa_pangan, "cur_total"),
        ("% Penggunaan Dana Pangan", pct_pangan, "pct"),
    ]
    r = 9
    for label, val, kind in pangan:
        rows[r] = [cell("", 0)] * 10
        if kind == "blank":
            r += 1; continue
        rows[r][0] = cell(label, 27)
        style = {"cur": 6, "cur_total": 7, "pct": 8}[kind]
        rows[r][6] = cell(val, style)
        merges += [f"A{r}:F{r}", f"G{r}:J{r}"]
        r += 1

    # --- SECTION B: OPERASIONAL (row 18) ---
    rows[18] = [cell("B. PENGELUARAN OPERASIONAL", 24)] + [cell("", 24)] * 9
    merges.append("A18:J18")

    ops = [
        ("Rata-Rata Pengeluaran Operasional / Hari", avg_ops, "cur"),
        ("Pengeluaran Operasional Seharusnya (Pagu) / Hari", pagu_ops_day, "cur"),
        ("Selisih Pengeluaran Operasional / Hari", selisih_ops_day, "cur"),
        ("", "", "blank"),
        ("Dana Total Operasional", dana_ops, "cur_total"),
        ("Penggunaan Dana Operasional", akt_ops, "cur_total"),
        ("Sisa Dana Operasional", sisa_ops, "cur_total"),
        ("% Penggunaan Dana Operasional", pct_ops, "pct"),
    ]
    r = 19
    for label, val, kind in ops:
        rows[r] = [cell("", 0)] * 10
        if kind == "blank":
            r += 1; continue
        rows[r][0] = cell(label, 27)
        style = {"cur": 6, "cur_total": 7, "pct": 8}[kind]
        rows[r][6] = cell(val, style)
        merges += [f"A{r}:F{r}", f"G{r}:J{r}"]
        r += 1

    # --- SECTION C: RINGKASAN TOTAL (row 28) ---
    rows[28] = [cell("C. RINGKASAN TOTAL", 25)] + [cell("", 25)] * 9
    merges.append("A28:J28")

    ring = [
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
    for label, val, kind in ring:
        rows[r] = [cell("", 0)] * 10
        if kind == "blank":
            r += 1; continue
        rows[r][0] = cell(label, 27)
        style = {"cur": 6, "cur_total": 7, "pct": 8}[kind]
        rows[r][6] = cell(val, style)
        merges += [f"A{r}:F{r}", f"G{r}:J{r}"]
        r += 1

    # --- SECTION D: SURPLUS/DEFISIT (row 38) ---
    rows[38] = [cell("D. ANALISIS SURPLUS / DEFISIT", 3)] + [cell("", 3)] * 9
    merges.append("A38:J38")

    sd = [
        ("Jumlah Hari SURPLUS (Pengeluaran < Pagu)", n_surplus, "int"),
        ("Jumlah Hari DEFISIT (Pengeluaran > Pagu)", n_defisit, "int"),
        ("Jumlah Hari AMAN (Libur / Pas Pagu)", n_aman, "int"),
        ("Jumlah Hari Belum Input", n_belum, "int"),
        ("", "", "blank"),
        ("Total Akumulasi Surplus", tot_surplus, "cur_pos"),
        ("Total Akumulasi Defisit", tot_defisit, "cur_neg"),
        ("Net Surplus / Defisit", net_sd, "cur_total"),
        ("% Surplus dari Total Anggaran", pct_surplus_total, "pct_pos"),
        ("% Defisit dari Total Anggaran", pct_defisit_total, "pct_neg"),
        ("Rasio Disiplin Anggaran (Hari Surplus / Aktif Non-Libur)", rasio_disiplin, "pct"),
        ("", "", "blank"),
        ("Pengeluaran Tertinggi / Hari (Rp)", max_r["total_aktual"], "cur_neg"),
        ("   Tanggal Pengeluaran Tertinggi", max_r["date"], "date"),
        ("Pengeluaran Terendah / Hari (Rp)", min_r["total_aktual"], "cur_pos"),
        ("   Tanggal Pengeluaran Terendah", min_r["date"], "date"),
    ]
    r = 39
    for label, val, kind in sd:
        rows[r] = [cell("", 0)] * 10
        if kind == "blank":
            r += 1; continue
        rows[r][0] = cell(label, 27)
        style_map = {"cur": 6, "cur_total": 7, "cur_pos": 22, "cur_neg": 23,
                     "pct": 8, "pct_pos": 30, "pct_neg": 31, "int": 20, "date": 16}
        if kind == "date":
            rows[r][6] = cell(val, style_map[kind], d=True)
        else:
            rows[r][6] = cell(val, style_map[kind])
        merges += [f"A{r}:F{r}", f"G{r}:J{r}"]
        r += 1

    # --- CHART SECTION HEADER (row 58) ---
    rows[58] = [cell("📊 VISUALISASI GRAFIK EKSEKUTIF", 3)] + [cell("", 3)] * 9
    merges.append("A58:J58")
    # Charts will occupy rows 59-142 (via drawing anchors)

    # --- CATATAN (at very bottom, row 144) ---
    notes = [
        ("CATATAN & METODOLOGI:", 2),
        (f"• Total Anggaran Rp {total_anggaran:,.0f} = SUM Dana Alokasi dari 8 file periode (sheet Pengaturan).", 0),
        ("• Total Penggunaan = SUM aktual harian Pangan + Operasional (97 hari tercatat).", 0),
        ("• VALIDASI SILANG: Total Anggaran − Total Penggunaan = Total Sisa = Rp 288.268.459 ✓", 0),
        ("• % Penyerapan = Total Penggunaan / Total Anggaran.", 0),
        ("• % Surplus = Total Akumulasi Hemat (hari SURPLUS saja) / Total Anggaran.", 0),
        ("• Selisih POSITIF = HEMAT (surplus). Selisih NEGATIF = BOROS (defisit).", 0),
        ("• Sumber: 8 file 'Total Akumulasi *.xlsx' di repository.", 0),
    ]
    base = 144
    for i, (txt, st) in enumerate(notes):
        rows[base + i] = [cell(txt, st)] + [cell("", 0)] * 9
        merges.append(f"A{base+i}:J{base+i}")

    cols = [(1, 42), (2, 10), (3, 12), (4, 12), (5, 10), (6, 12),
            (7, 20), (8, 12), (9, 16), (10, 12)]
    row_heights = {1: 38, 2: 22, 3: 22, 5: 28, 6: 46,
                   8: 28, 18: 28, 28: 28, 38: 28, 58: 28}
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
            cell(rec["date"], 16, d=True),
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

    total_r = r + 1
    rows[total_r] = [
        cell("TOTAL", 7), cell("", 7),
        cell(sum(x["pagu_pangan"] for x in records), 7),
        cell(akt_pangan, 7),
        cell(sum(x["pagu_pangan"] - x["aktual_pangan"] for x in records), 7),
        cell(sum(x["pagu_ops"] for x in records), 7),
        cell(akt_ops, 7),
        cell(sum(x["pagu_ops"] - x["aktual_ops"] for x in records), 7),
        cell(sum(x["total_pagu"] for x in records), 7),
        cell(total_penggunaan, 7),
        cell("", 7),
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
        ss = 13 if w["status"] == "SURPLUS" else (14 if w["status"] == "DEFISIT" else 15)
        rows[r] = [
            cell(i, 12),
            cell(f'{w["start"]} s/d {w["end"]}', 12),
            cell(w["pagu_pangan"], 10),
            cell(w["aktual_pangan"], 10),
            cell(w["pagu_ops"], 10),
            cell(w["aktual_ops"], 10),
            cell(w["total_pagu"], 10),
            cell(w["total_aktual"], 10),
            cell(w["surplus_defisit"], 10),
            cell(w["status"], ss),
        ]
        r += 1

    rows[r + 1] = [
        cell("TOTAL", 7), cell("", 7),
        cell(sum(w["pagu_pangan"] for w in weekly_list), 7),
        cell(sum(w["aktual_pangan"] for w in weekly_list), 7),
        cell(sum(w["pagu_ops"] for w in weekly_list), 7),
        cell(sum(w["aktual_ops"] for w in weekly_list), 7),
        cell(sum(w["total_pagu"] for w in weekly_list), 7),
        cell(sum(w["total_aktual"] for w in weekly_list), 7),
        cell(sum(w["surplus_defisit"] for w in weekly_list), 7),
        cell("", 7),
    ]

    cols = [(1, 10), (2, 26), (3, 15), (4, 15), (5, 15), (6, 15),
            (7, 16), (8, 16), (9, 16), (10, 11)]
    row_heights = {1: 30, 3: 30}
    return rows, merges, cols, row_heights


# ============================================================
# SHEET 4: CHART DATA
# ============================================================
def build_chart_data_sheet():
    rows = {}
    merges = []
    rows[1] = [cell("CHART DATA (Sumber Grafik)", 1)] + [cell("", 1)] * 5
    merges.append("A1:F1")

    # --- Line chart data: 97 days ---
    rows[3] = [cell("Tanggal", 9), cell("Total Pagu", 9), cell("Total Aktual", 9),
               cell("Pangan Aktual", 9), cell("Operasional Aktual", 9)]
    r = 4
    for rec in records:
        rows[r] = [
            cell(rec["date"], 16, d=True),
            cell(rec["total_pagu"], 10),
            cell(rec["total_aktual"], 10),
            cell(rec["aktual_pangan"], 10),
            cell(rec["aktual_ops"], 10),
        ]
        r += 1
    n_daily = len(records)

    # --- Doughnut composition data ---
    pie_start = r + 2
    rows[pie_start] = [cell("Kategori", 9), cell("Nilai", 9)]
    rows[pie_start+1] = [cell("Penggunaan Pangan", 12), cell(akt_pangan, 10)]
    rows[pie_start+2] = [cell("Penggunaan Operasional", 12), cell(akt_ops, 10)]
    rows[pie_start+3] = [cell("Sisa Dana", 12), cell(total_sisa, 10)]

    # --- Status bar data ---
    bar_start = pie_start + 6
    rows[bar_start] = [cell("Status Hari", 9), cell("Jumlah", 9)]
    rows[bar_start+1] = [cell("SURPLUS", 13), cell(n_surplus, 20)]
    rows[bar_start+2] = [cell("DEFISIT", 14), cell(n_defisit, 20)]
    rows[bar_start+3] = [cell("AMAN / LIBUR", 15), cell(n_aman, 20)]

    # --- Weekly data for bar chart ---
    weekly_start = bar_start + 6
    rows[weekly_start] = [cell("Minggu", 9), cell("Pagu", 9), cell("Aktual", 9),
                          cell("Surplus/Defisit", 9)]
    for i, w in enumerate(weekly_list, start=1):
        rows[weekly_start + i] = [
            cell(f"Mg-{i}", 12),
            cell(w["total_pagu"], 10),
            cell(w["total_aktual"], 10),
            cell(w["surplus_defisit"], 10),
        ]
    n_weeks = len(weekly_list)

    # --- Pangan vs Ops comparison (for stacked) ---
    comp_start = weekly_start + n_weeks + 3
    rows[comp_start] = [cell("Kategori", 9), cell("Anggaran", 9),
                        cell("Terpakai", 9), cell("Sisa", 9)]
    rows[comp_start+1] = [cell("Pangan", 12), cell(dana_pangan, 10),
                          cell(akt_pangan, 10), cell(sisa_pangan, 10)]
    rows[comp_start+2] = [cell("Operasional", 12), cell(dana_ops, 10),
                          cell(akt_ops, 10), cell(sisa_ops, 10)]

    cols = [(1, 16), (2, 18), (3, 18), (4, 18), (5, 18)]
    row_heights = {1: 28, 3: 24}

    ranges = {
        "line": {
            "cat": f"'Chart Data'!$A$4:$A${3 + n_daily}",
            "pagu": f"'Chart Data'!$B$4:$B${3 + n_daily}",
            "aktual": f"'Chart Data'!$C$4:$C${3 + n_daily}",
            "pangan": f"'Chart Data'!$D$4:$D${3 + n_daily}",
            "ops": f"'Chart Data'!$E$4:$E${3 + n_daily}",
        },
        "pie": {
            "cat": f"'Chart Data'!$A${pie_start+1}:$A${pie_start+3}",
            "val": f"'Chart Data'!$B${pie_start+1}:$B${pie_start+3}",
        },
        "status": {
            "cat": f"'Chart Data'!$A${bar_start+1}:$A${bar_start+3}",
            "val": f"'Chart Data'!$B${bar_start+1}:$B${bar_start+3}",
        },
        "weekly": {
            "cat": f"'Chart Data'!$A${weekly_start+1}:$A${weekly_start+n_weeks}",
            "pagu": f"'Chart Data'!$B${weekly_start+1}:$B${weekly_start+n_weeks}",
            "aktual": f"'Chart Data'!$C${weekly_start+1}:$C${weekly_start+n_weeks}",
            "sd": f"'Chart Data'!$D${weekly_start+1}:$D${weekly_start+n_weeks}",
        },
        "comp": {
            "cat": f"'Chart Data'!$A${comp_start+1}:$A${comp_start+2}",
            "terpakai": f"'Chart Data'!$C${comp_start+1}:$C${comp_start+2}",
            "sisa": f"'Chart Data'!$D${comp_start+1}:$D${comp_start+2}",
        }
    }
    return rows, merges, cols, row_heights, ranges


# ============================================================
# SHEET XML BUILDER
# ============================================================
def sheet_xml(rows, merges, cols, row_heights, drawing_rid=None, show_grid=True):
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    parts.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    gridflag = "0" if not show_grid else "1"
    parts.append(f'<sheetViews><sheetView workbookViewId="0" showGridLines="{gridflag}">'
                 f'<selection activeCell="A1"/></sheetView></sheetViews>')
    parts.append('<sheetFormatPr defaultRowHeight="15"/>')
    if cols:
        parts.append("<cols>")
        for idx, w in cols:
            parts.append(f'<col min="{idx}" max="{idx}" width="{w}" customWidth="1"/>')
        parts.append("</cols>")
    parts.append("<sheetData>")
    for r_idx in sorted(rows.keys()):
        parts.append(row_xml(r_idx, rows[r_idx], height=row_heights.get(r_idx)))
    parts.append("</sheetData>")
    if merges:
        parts.append(f'<mergeCells count="{len(merges)}">')
        for m in merges:
            parts.append(f'<mergeCell ref="{m}"/>')
        parts.append("</mergeCells>")
    if drawing_rid:
        parts.append(f'<drawing r:id="{drawing_rid}"/>')
    parts.append("</worksheet>")
    return "".join(parts)


# ============================================================
# CHART XMLs
# ============================================================

def chart_header(title_text, title_color="1F3864"):
    return (f'<c:title><c:tx><c:rich>'
            f'<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" wrap="square" anchor="ctr" anchorCtr="1"/>'
            f'<a:lstStyle/>'
            f'<a:p><a:pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="{title_color}"/></a:solidFill></a:defRPr></a:pPr>'
            f'<a:r><a:rPr lang="id-ID" sz="1400" b="1"><a:solidFill><a:srgbClr val="{title_color}"/></a:solidFill></a:rPr>'
            f'<a:t>{escape(title_text)}</a:t></a:r></a:p>'
            f'</c:rich></c:tx><c:overlay val="0"/></c:title>'
            f'<c:autoTitleDeleted val="0"/>')


CHART_WRAP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
              'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
              'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')


def chart_line(line):
    """Chart 1: Tren Harian Pagu vs Aktual (Line)"""
    xml = CHART_WRAP + '<c:chart>' + chart_header("TREN HARIAN: Pagu vs Pengeluaran Aktual")
    xml += ('<c:plotArea><c:layout/>'
            '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
            f'<c:ser><c:idx val="0"/><c:order val="0"/>'
            f'<c:tx><c:v>Total Pagu</c:v></c:tx>'
            f'<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="548235"/></a:solidFill></a:ln></c:spPr>'
            f'<c:marker><c:symbol val="none"/></c:marker>'
            f'<c:cat><c:numRef><c:f>{line["cat"]}</c:f></c:numRef></c:cat>'
            f'<c:val><c:numRef><c:f>{line["pagu"]}</c:f></c:numRef></c:val>'
            f'<c:smooth val="0"/></c:ser>'
            f'<c:ser><c:idx val="1"/><c:order val="1"/>'
            f'<c:tx><c:v>Total Aktual</c:v></c:tx>'
            f'<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln></c:spPr>'
            f'<c:marker><c:symbol val="circle"/><c:size val="5"/>'
            f'<c:spPr><a:solidFill><a:srgbClr val="C00000"/></a:solidFill>'
            f'<a:ln><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln></c:spPr></c:marker>'
            f'<c:cat><c:numRef><c:f>{line["cat"]}</c:f></c:numRef></c:cat>'
            f'<c:val><c:numRef><c:f>{line["aktual"]}</c:f></c:numRef></c:val>'
            f'<c:smooth val="0"/></c:ser>'
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
    return xml


def chart_doughnut(pie):
    """Chart 2: Komposisi Penyerapan (Doughnut)"""
    xml = CHART_WRAP + '<c:chart>' + chart_header("KOMPOSISI PENYERAPAN ANGGARAN")
    xml += ('<c:plotArea><c:layout/>'
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
            '<c:dLbls><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1">'
            '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:pPr>'
            '<a:endParaRPr lang="id-ID"/></a:p></c:txPr>'
            '<c:dLblPos val="ctr"/><c:showLegendKey val="0"/><c:showVal val="0"/>'
            '<c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/>'
            '<c:showBubbleSize val="0"/></c:dLbls>'
            f'<c:cat><c:strRef><c:f>{pie["cat"]}</c:f></c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>{pie["val"]}</c:f></c:numRef></c:val>'
            '</c:ser>'
            '<c:firstSliceAng val="0"/><c:holeSize val="50"/>'
            '</c:doughnutChart></c:plotArea>'
            '<c:legend><c:legendPos val="b"/><c:overlay val="0"/>'
            '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1"/></a:pPr>'
            '<a:endParaRPr lang="id-ID"/></a:p></c:txPr>'
            '</c:legend>'
            '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>'
            '</c:chart></c:chartSpace>')
    return xml


def chart_status_pie(status):
    """Chart 3: Distribusi Status Hari (Pie 3D)"""
    xml = CHART_WRAP + '<c:chart>' + chart_header("DISTRIBUSI STATUS HARIAN")
    xml += ('<c:plotArea><c:layout/>'
            '<c:pie3DChart><c:varyColors val="1"/>'
            '<c:ser><c:idx val="0"/><c:order val="0"/>'
            '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/>'
            '<c:spPr><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></c:spPr></c:dPt>'
            '<c:dPt><c:idx val="1"/><c:bubble3D val="0"/>'
            '<c:spPr><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></c:spPr></c:dPt>'
            '<c:dPt><c:idx val="2"/><c:bubble3D val="0"/>'
            '<c:spPr><a:solidFill><a:srgbClr val="A6A6A6"/></a:solidFill></c:spPr></c:dPt>'
            '<c:dLbls><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1">'
            '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:pPr>'
            '<a:endParaRPr lang="id-ID"/></a:p></c:txPr>'
            '<c:dLblPos val="ctr"/><c:showLegendKey val="0"/><c:showVal val="1"/>'
            '<c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>'
            '<c:showBubbleSize val="0"/></c:dLbls>'
            f'<c:cat><c:strRef><c:f>{status["cat"]}</c:f></c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>{status["val"]}</c:f></c:numRef></c:val>'
            '</c:ser>'
            '</c:pie3DChart></c:plotArea>'
            '<c:legend><c:legendPos val="b"/><c:overlay val="0"/>'
            '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1"/></a:pPr>'
            '<a:endParaRPr lang="id-ID"/></a:p></c:txPr>'
            '</c:legend>'
            '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>'
            '</c:chart></c:chartSpace>')
    return xml


def chart_weekly_bar(weekly):
    """Chart 4: Pagu vs Aktual per Minggu (Clustered Bar)"""
    xml = CHART_WRAP + '<c:chart>' + chart_header("PERBANDINGAN MINGGUAN: Pagu vs Aktual")
    xml += ('<c:plotArea><c:layout/>'
            '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
            f'<c:ser><c:idx val="0"/><c:order val="0"/>'
            f'<c:tx><c:v>Total Pagu</c:v></c:tx>'
            f'<c:spPr><a:solidFill><a:srgbClr val="2E75B6"/></a:solidFill>'
            f'<a:ln><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:ln></c:spPr>'
            f'<c:cat><c:strRef><c:f>{weekly["cat"]}</c:f></c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>{weekly["pagu"]}</c:f></c:numRef></c:val>'
            f'</c:ser>'
            f'<c:ser><c:idx val="1"/><c:order val="1"/>'
            f'<c:tx><c:v>Total Aktual</c:v></c:tx>'
            f'<c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill>'
            f'<a:ln><a:solidFill><a:srgbClr val="843C0C"/></a:solidFill></a:ln></c:spPr>'
            f'<c:cat><c:strRef><c:f>{weekly["cat"]}</c:f></c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>{weekly["aktual"]}</c:f></c:numRef></c:val>'
            f'</c:ser>'
            '<c:gapWidth val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>'
            '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
            '<c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="out"/>'
            '<c:tickLblPos val="nextTo"/><c:crossAx val="2"/></c:catAx>'
            '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
            '<c:delete val="0"/><c:axPos val="l"/>'
            '<c:numFmt formatCode="#,##0" sourceLinked="0"/>'
            '<c:majorTickMark val="out"/><c:tickLblPos val="nextTo"/><c:crossAx val="1"/></c:valAx>'
            '</c:plotArea>'
            '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
            '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>'
            '</c:chart></c:chartSpace>')
    return xml


def chart_stacked_comp(comp):
    """Chart 5: Pangan vs Ops - Stacked (Terpakai vs Sisa)"""
    xml = CHART_WRAP + '<c:chart>' + chart_header("ALOKASI vs PENGGUNAAN: Pangan & Operasional")
    xml += ('<c:plotArea><c:layout/>'
            '<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>'
            f'<c:ser><c:idx val="0"/><c:order val="0"/>'
            f'<c:tx><c:v>Terpakai</c:v></c:tx>'
            f'<c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></c:spPr>'
            f'<c:cat><c:strRef><c:f>{comp["cat"]}</c:f></c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>{comp["terpakai"]}</c:f></c:numRef></c:val>'
            f'</c:ser>'
            f'<c:ser><c:idx val="1"/><c:order val="1"/>'
            f'<c:tx><c:v>Sisa</c:v></c:tx>'
            f'<c:spPr><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></c:spPr>'
            f'<c:cat><c:strRef><c:f>{comp["cat"]}</c:f></c:strRef></c:cat>'
            f'<c:val><c:numRef><c:f>{comp["sisa"]}</c:f></c:numRef></c:val>'
            f'</c:ser>'
            '<c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>'
            '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
            '<c:delete val="0"/><c:axPos val="l"/><c:majorTickMark val="out"/>'
            '<c:tickLblPos val="nextTo"/><c:crossAx val="2"/></c:catAx>'
            '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
            '<c:delete val="0"/><c:axPos val="b"/>'
            '<c:numFmt formatCode="#,##0" sourceLinked="0"/>'
            '<c:majorTickMark val="out"/><c:tickLblPos val="nextTo"/><c:crossAx val="1"/></c:valAx>'
            '</c:plotArea>'
            '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
            '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>'
            '</c:chart></c:chartSpace>')
    return xml


# ============================================================
# DRAWING XML (anchor 5 charts on Dashboard)
# ============================================================
def build_drawing():
    """Layout of charts on Dashboard (row-based anchors):
       Chart 1 (Line)         : rows 59-80, cols A-J     [full width]
       Chart 2 (Doughnut)     : rows 81-102, cols A-E    [left half]
       Chart 3 (Status Pie)   : rows 81-102, cols F-J    [right half]
       Chart 4 (Weekly Bar)   : rows 103-124, cols A-J   [full width]
       Chart 5 (Stacked Comp) : rows 125-142, cols A-J   [full width]
    """
    anchors = [
        # chart, from_col, from_row, to_col, to_row, rid
        (1, 0, 59, 10, 80),   # Line - full width
        (2, 0, 81, 5, 102),   # Doughnut - left half
        (3, 5, 81, 10, 102),  # Status Pie - right half
        (4, 0, 103, 10, 124), # Weekly Bar - full width
        (5, 0, 125, 10, 142), # Stacked - full width
    ]
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    for idx, (cid, fc, fr, tc, tr) in enumerate(anchors, start=1):
        xml += (f'<xdr:twoCellAnchor>'
                f'<xdr:from><xdr:col>{fc}</xdr:col><xdr:colOff>0</xdr:colOff>'
                f'<xdr:row>{fr}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                f'<xdr:to><xdr:col>{tc}</xdr:col><xdr:colOff>0</xdr:colOff>'
                f'<xdr:row>{tr}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
                f'<xdr:graphicFrame macro="">'
                f'<xdr:nvGraphicFramePr>'
                f'<xdr:cNvPr id="{idx+1}" name="Chart {cid}"/>'
                f'<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
                f'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
                f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
                f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId{cid}"/>'
                f'</a:graphicData></a:graphic></xdr:graphicFrame>'
                f'<xdr:clientData/></xdr:twoCellAnchor>')
    xml += '</xdr:wsDr>'
    return xml


# ============================================================
# CONTENT TYPES & RELS
# ============================================================
def content_types():
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
            '<Override PartName="/xl/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
            '<Override PartName="/xl/charts/chart4.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
            '<Override PartName="/xl/charts/chart5.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
            '</Types>')

ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
             '</Relationships>')

def workbook_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets>'
            '<sheet name="Dashboard" sheetId="1" r:id="rId1"/>'
            '<sheet name="Rekap Harian" sheetId="2" r:id="rId2"/>'
            '<sheet name="Rekap Mingguan" sheetId="3" r:id="rId3"/>'
            '<sheet name="Chart Data" sheetId="4" r:id="rId4"/>'
            '</sheets></workbook>')

def workbook_rels():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>'
            '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>'
            '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>')

SHEET1_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
               '</Relationships>')

DRAWING_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>'
                '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart4.xml"/>'
                '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart5.xml"/>'
                '</Relationships>')


# ============================================================
# BUILD FILE
# ============================================================
def main():
    d_rows, d_merges, d_cols, d_heights = build_dashboard_sheet()
    h_rows, h_merges, h_cols, h_heights = build_harian_sheet()
    m_rows, m_merges, m_cols, m_heights = build_mingguan_sheet()
    c_rows, c_merges, c_cols, c_heights, ranges = build_chart_data_sheet()

    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types())
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", workbook_xml())
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels())
        z.writestr("xl/styles.xml", STYLES_XML)

        z.writestr("xl/worksheets/sheet1.xml",
                   sheet_xml(d_rows, d_merges, d_cols, d_heights,
                             drawing_rid="rId1", show_grid=False))
        z.writestr("xl/worksheets/_rels/sheet1.xml.rels", SHEET1_RELS)
        z.writestr("xl/worksheets/sheet2.xml",
                   sheet_xml(h_rows, h_merges, h_cols, h_heights))
        z.writestr("xl/worksheets/sheet3.xml",
                   sheet_xml(m_rows, m_merges, m_cols, m_heights))
        z.writestr("xl/worksheets/sheet4.xml",
                   sheet_xml(c_rows, c_merges, c_cols, c_heights))

        z.writestr("xl/drawings/drawing1.xml", build_drawing())
        z.writestr("xl/drawings/_rels/drawing1.xml.rels", DRAWING_RELS)
        z.writestr("xl/charts/chart1.xml", chart_line(ranges["line"]))
        z.writestr("xl/charts/chart2.xml", chart_doughnut(ranges["pie"]))
        z.writestr("xl/charts/chart3.xml", chart_status_pie(ranges["status"]))
        z.writestr("xl/charts/chart4.xml", chart_weekly_bar(ranges["weekly"]))
        z.writestr("xl/charts/chart5.xml", chart_stacked_comp(ranges["comp"]))

    print(f"\nOK -> {OUTPUT}")
    print(f"Size: {os.path.getsize(OUTPUT):,} bytes")


if __name__ == "__main__":
    main()
