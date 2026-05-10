"""
Generator Template Laporan Keuangan SPPG (Excel .xlsx)
Professional edition - dengan Pagu Harian variabel, rekap mingguan, dan
5 grafik profesional. Dibangun hanya dengan Python standard library.

Sheet:
  1. Dashboard         - KPI cards professional & ringkasan
  2. Rekap Harian      - Per-hari dengan pagu variabel & status
  3. Rekap Mingguan    - Per-minggu (Minggu 1 s/d 5) dengan surplus/defisit
  4. Grafik            - 5 chart profesional
  5. Pengaturan        - Master setting (dana, periode)
  6. Pagu Harian       - Pagu per tanggal (editable, beda-beda)
  7. Pengeluaran Pangan     - Log transaksi pangan
  8. Pengeluaran Operasional - Log transaksi operasional
  9. Kategori          - Referensi

Output: Template_Keuangan.xlsx
"""

import zipfile
from datetime import date
from xml.sax.saxutils import escape

OUTPUT = "Template_Keuangan.xlsx"

# =========================================================================
# CONFIG
# =========================================================================
MAX_DAYS = 31        # baris di Rekap Harian & Pagu Harian
MAX_WEEKS = 5        # minggu di Rekap Mingguan


# =========================================================================
# CELL HELPERS (val, style_id, is_formula)
# =========================================================================

def c(v, s=0, f=False):
    return (v, s, f)

def txt(v, s=0): return c(v, s, False)
def num(v, s=11): return c(v, s, False)
def cur(v, s=3): return c(v, s, False)
def hdr(v, s=1): return c(v, s, False)
def fx(f, s=3): return c(f, s, True)
def dt(v, s=4): return c(v, s, False)
def sec(v, s=13): return c(v, s, False)  # section header


# =========================================================================
# STYLE INDEX (lihat STYLES_XML)
# =========================================================================
# 0  default
# 1  header (biru tua + putih)
# 2  title besar
# 3  currency IDR
# 4  date
# 5  subtitle
# 6  percent
# 7  total currency (bold, bg kuning)
# 8  center wrap
# 9  currency merah
# 10 currency hijau
# 11 integer center
# 12 label bold italic
# 13 section hijau muda
# 14 currency bold
# 15 center bold
# 16 KPI big currency (bg kuning)
# 17 selisih hijau (bg abu)
# 18 selisih merah (bg abu)
# 19 small italic gray
# 20 header oranye
# 21 header hijau tua
# 22 KPI hero besar biru
# 23 KPI hero besar hijau
# 24 KPI hero besar merah
# 25 KPI hero besar kuning
# 26 status SURPLUS (hijau dengan bg)
# 27 status DEFISIT (merah dengan bg)
# 28 status AMAN (biru)
# 29 center bold white
# 30 label hero (abu kecil)
# 31 subtitle italic abu
# 32 section biru tua
# 33 angka besar KPI
# 34 section oranye
# 35 currency bold hijau besar


# =========================================================================
# SHEET: PENGATURAN
# =========================================================================
pengaturan_rows = [
    [txt("PENGATURAN LAPORAN KEUANGAN", 2)],
    [txt("Master konfigurasi - edit bagian ini untuk menyesuaikan laporan", 31)],
    [txt("Periode Mulai", 12),       dt("2026-05-01"),      txt("<- edit tanggal", 19)],
    [txt("Periode Selesai", 12),     dt("2026-05-31"),      txt("<- edit tanggal", 19)],
    [txt("Jumlah Hari Operasi", 12), fx("B4-B3+1", 11),     txt("(otomatis)", 19)],
    [txt("")],
    [sec("DANA ANGGARAN", 32)],
    [txt("Dana Total Bahan Pangan", 12),   cur(60000000, 3), txt("<- edit total dana", 19)],
    [txt("Dana Total Operasional", 12),    cur(25000000, 3), txt("<- edit total dana", 19)],
    [txt("Total Dana Keseluruhan", 12),    fx("B8+B9", 7)],
    [txt("")],
    [sec("PAGU DEFAULT / HARI (SEHARUSNYA)", 32)],
    [txt("Pagu Default Pangan/Hari", 12),      fx("IF(B5=0,0,B8/B5)", 14), txt("= Dana Pangan / Jumlah Hari", 19)],
    [txt("Pagu Default Operasional/Hari", 12), fx("IF(B5=0,0,B9/B5)", 14), txt("= Dana Operasional / Jumlah Hari", 19)],
    [txt("Pagu Default Total/Hari", 12),       fx("B13+B14", 14),          txt("= Pangan + Operasional", 19)],
    [txt("")],
    [sec("INFORMASI ORGANISASI", 32)],
    [txt("Nama Instansi", 12),      txt("SPPG Battuwinangun")],
    [txt("Alamat", 12),             txt("-")],
    [txt("Penanggung Jawab", 12),   txt("-")],
    [txt("Periode Laporan", 12),    fx('TEXT(B3,"dd mmm yyyy")&" s/d "&TEXT(B4,"dd mmm yyyy")', 15)],
    [txt("")],
    [txt("CATATAN:", 12)],
    [txt("- Edit hanya di kolom B pada baris yang bertanda <- edit", 19)],
    [txt("- Pagu per hari bisa berbeda-beda, diatur di sheet 'Pagu Harian'", 19)],
    [txt("- Jika Pagu Harian kosong untuk tanggal tertentu, dipakai pagu default", 19)],
    [txt("- Dashboard, Rekap Harian, & Grafik update otomatis", 19)],
]


# =========================================================================
# SHEET: PAGU HARIAN (pagu per tanggal, bisa beda-beda)
# =========================================================================
pagu_rows = [
    [txt("PAGU HARIAN PENGELUARAN", 2)],
    [txt("Set pagu per tanggal. Biarkan kosong -> pakai pagu default dari sheet Pengaturan", 31)],
    [hdr("Tanggal"), hdr("Pagu Pangan (Rp)"), hdr("Pagu Operasional (Rp)"),
     hdr("Pagu Total (Rp)"), hdr("Keterangan")],
]
PAGU_DATA_START = 4
for i in range(MAX_DAYS):
    r = PAGU_DATA_START + i
    pagu_rows.append([
        fx('IF(Pengaturan!$B$3+{i}>Pengaturan!$B$4,"",Pengaturan!$B$3+{i})'.format(i=i), 4),
        fx('IF(A{r}="","",Pengaturan!$B$13)'.format(r=r), 3),
        fx('IF(A{r}="","",Pengaturan!$B$14)'.format(r=r), 3),
        fx('IF(A{r}="","",B{r}+C{r})'.format(r=r), 14),
        txt(""),
    ])
PAGU_DATA_END = PAGU_DATA_START + MAX_DAYS - 1
# Total
pagu_rows.append([txt("")])
pagu_rows.append([
    txt("TOTAL", 5),
    fx('SUM(B{s}:B{e})'.format(s=PAGU_DATA_START, e=PAGU_DATA_END), 7),
    fx('SUM(C{s}:C{e})'.format(s=PAGU_DATA_START, e=PAGU_DATA_END), 7),
    fx('SUM(D{s}:D{e})'.format(s=PAGU_DATA_START, e=PAGU_DATA_END), 7),
    txt(""),
])


# =========================================================================
# SHEET: PENGELUARAN PANGAN
# =========================================================================
pangan_rows = [
    [txt("LOG PENGELUARAN BAHAN PANGAN", 2)],
    [txt("Catat semua transaksi bahan pangan di sini", 31)],
    [hdr("Tanggal"), hdr("Item / Bahan"), hdr("Kategori"),
     hdr("Qty"), hdr("Satuan"), hdr("Harga Satuan (Rp)"),
     hdr("Jumlah (Rp)"), hdr("Keterangan")],
]
contoh_pangan = [
    ("2026-05-01", "Beras premium",   "Karbohidrat",    50, "kg",    15000, "Stok awal"),
    ("2026-05-01", "Ayam potong",     "Protein Hewani", 20, "kg",    38000, "Menu hari 1"),
    ("2026-05-02", "Sayur bayam",     "Sayuran",        15, "ikat",   3500, ""),
    ("2026-05-02", "Telur ayam",      "Protein Hewani",  8, "kg",    28000, ""),
    ("2026-05-03", "Ikan nila",       "Protein Hewani", 18, "kg",    32000, ""),
    ("2026-05-03", "Wortel",          "Sayuran",        10, "kg",     8000, ""),
    ("2026-05-04", "Tempe",           "Protein Nabati", 12, "kg",    12000, ""),
    ("2026-05-04", "Minyak goreng",   "Bumbu",           5, "liter", 18000, ""),
    ("2026-05-05", "Buah pisang",     "Buah",           20, "sisir", 15000, ""),
    ("2026-05-05", "Bumbu dapur",     "Bumbu",           1, "paket",150000, "Bawang, cabe"),
]
for t, itm, kat, qty, sat, harga, ket in contoh_pangan:
    pangan_rows.append([
        dt(t), txt(itm), txt(kat), num(qty, 11), txt(sat), cur(harga, 3),
        fx("D{r}*F{r}".format(r=len(pangan_rows) + 1), 9),
        txt(ket),
    ])
PANGAN_DATA_START = 4
for _ in range(110):
    pangan_rows.append([txt(""), txt(""), txt(""), txt(""), txt(""), txt(""), txt(""), txt("")])
PANGAN_LAST_DATA_ROW = len(pangan_rows)
pangan_rows.append([txt("")])
pangan_rows.append([
    txt(""), txt(""), txt(""), txt(""), txt(""),
    txt("TOTAL PENGELUARAN PANGAN", 5),
    fx("SUM(G{s}:G{e})".format(s=PANGAN_DATA_START, e=PANGAN_LAST_DATA_ROW), 7),
    txt(""),
])
PANGAN_TOTAL_ROW = len(pangan_rows)


# =========================================================================
# SHEET: PENGELUARAN OPERASIONAL
# =========================================================================
operasional_rows = [
    [txt("LOG PENGELUARAN OPERASIONAL", 2)],
    [txt("Catat semua biaya operasional non-bahan pangan di sini", 31)],
    [hdr("Tanggal"), hdr("Deskripsi"), hdr("Kategori"),
     hdr("Qty"), hdr("Satuan"), hdr("Harga Satuan (Rp)"),
     hdr("Jumlah (Rp)"), hdr("Keterangan")],
]
contoh_op = [
    ("2026-05-01", "Gas LPG 12kg",           "Bahan Bakar",   2,  "tabung", 185000, "Memasak"),
    ("2026-05-02", "Gaji juru masak",         "Tenaga Kerja",  1,  "orang",  150000, "Harian"),
    ("2026-05-02", "Listrik & air",           "Utilitas",      1,  "bulan",  450000, ""),
    ("2026-05-03", "Sabun cuci & pembersih",  "Sanitasi",      5,  "pcs",     25000, ""),
    ("2026-05-04", "Kemasan makan",           "Packaging",    200, "pcs",      1500, ""),
    ("2026-05-05", "Transportasi distribusi", "Transportasi",  1,  "trip",   120000, ""),
    ("2026-05-06", "Alat masak (spatula)",    "Peralatan",     3,  "pcs",     35000, ""),
    ("2026-05-07", "Perawatan kompor",        "Pemeliharaan",  1,  "kali",   200000, ""),
]
for t, desc, kat, qty, sat, harga, ket in contoh_op:
    operasional_rows.append([
        dt(t), txt(desc), txt(kat), num(qty, 11), txt(sat), cur(harga, 3),
        fx("D{r}*F{r}".format(r=len(operasional_rows) + 1), 9),
        txt(ket),
    ])
OP_DATA_START = 4
for _ in range(110):
    operasional_rows.append([txt(""), txt(""), txt(""), txt(""), txt(""), txt(""), txt(""), txt("")])
OP_LAST_DATA_ROW = len(operasional_rows)
operasional_rows.append([txt("")])
operasional_rows.append([
    txt(""), txt(""), txt(""), txt(""), txt(""),
    txt("TOTAL PENGELUARAN OPERASIONAL", 5),
    fx("SUM(G{s}:G{e})".format(s=OP_DATA_START, e=OP_LAST_DATA_ROW), 7),
    txt(""),
])
OP_TOTAL_ROW = len(operasional_rows)


# =========================================================================
# SHEET: REKAP HARIAN
# =========================================================================
# Kolom:
#   A  No
#   B  Tanggal
#   C  Hari
#   D  Minggu Ke
#   E  Pagu Pangan (dari Pagu Harian, fallback ke default)
#   F  Aktual Pangan
#   G  Selisih Pangan (E-F)
#   H  Pagu Operasional
#   I  Aktual Operasional
#   J  Selisih Operasional (H-I)
#   K  Total Pagu
#   L  Total Aktual
#   M  Surplus/Defisit (K-L)
#   N  Status (SURPLUS/DEFISIT/AMAN)

rekap_rows = [
    [txt("REKAP HARIAN PENGELUARAN", 2)],
    [txt("Setiap hari akan dihitung otomatis: pagu, aktual, selisih, dan status surplus/defisit", 31)],
    [txt("Periode:", 12), fx('TEXT(Pengaturan!B3,"dd mmm yyyy")&" s/d "&TEXT(Pengaturan!B4,"dd mmm yyyy")', 15)],
    [txt("")],
    [hdr("No"), hdr("Tanggal"), hdr("Hari"), hdr("Mg Ke"),
     hdr("Pagu Pangan", 21), hdr("Aktual Pangan", 21), hdr("Selisih Pangan", 21),
     hdr("Pagu Operasional", 20), hdr("Aktual Operasional", 20), hdr("Selisih Operasional", 20),
     hdr("Total Pagu"), hdr("Total Aktual"), hdr("Surplus/Defisit"), hdr("Status")],
]
REKAP_DATA_START = 6
for i in range(MAX_DAYS):
    r = REKAP_DATA_START + i
    rekap_rows.append([
        num(i + 1, 15),
        fx('IF(Pengaturan!$B$3+{i}>Pengaturan!$B$4,"",Pengaturan!$B$3+{i})'.format(i=i), 4),
        fx('IF(B{r}="","",TEXT(B{r},"dddd"))'.format(r=r), 15),
        fx('IF(B{r}="","",INT((B{r}-Pengaturan!$B$3)/7)+1)'.format(r=r), 15),
        # Pagu Pangan (lookup dari Pagu Harian, fallback ke default)
        fx("IFERROR(VLOOKUP(B{r},'Pagu Harian'!$A$4:$D$40,2,FALSE),Pengaturan!$B$13)".format(r=r), 3),
        # Aktual Pangan
        fx('IF(B{r}="",0,SUMIF(\'Pengeluaran Pangan\'!A:A,B{r},\'Pengeluaran Pangan\'!G:G))'.format(r=r), 3),
        # Selisih Pangan
        fx('IF(B{r}="",0,E{r}-F{r})'.format(r=r), 17),
        # Pagu Operasional
        fx("IFERROR(VLOOKUP(B{r},'Pagu Harian'!$A$4:$D$40,3,FALSE),Pengaturan!$B$14)".format(r=r), 3),
        # Aktual Operasional
        fx('IF(B{r}="",0,SUMIF(\'Pengeluaran Operasional\'!A:A,B{r},\'Pengeluaran Operasional\'!G:G))'.format(r=r), 3),
        # Selisih Operasional
        fx('IF(B{r}="",0,H{r}-I{r})'.format(r=r), 17),
        # Total Pagu
        fx('E{r}+H{r}'.format(r=r), 14),
        # Total Aktual
        fx('F{r}+I{r}'.format(r=r), 14),
        # Surplus/Defisit
        fx('K{r}-L{r}'.format(r=r), 17),
        # Status
        fx('IF(B{r}="","",IF(M{r}>0,"SURPLUS",IF(M{r}<0,"DEFISIT","AMAN")))'.format(r=r), 26),
    ])
REKAP_DATA_END = REKAP_DATA_START + MAX_DAYS - 1
# Total & rata-rata
rekap_rows.append([txt("")])
rekap_rows.append([
    txt(""), txt("TOTAL", 5), txt(""), txt(""),
    fx('SUM(E{s}:E{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('SUM(F{s}:F{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('SUM(G{s}:G{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('SUM(H{s}:H{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('SUM(I{s}:I{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('SUM(J{s}:J{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('SUM(K{s}:K{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('SUM(L{s}:L{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('SUM(M{s}:M{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    txt(""),
])
rekap_rows.append([
    txt(""), txt("RATA-RATA/HARI", 5), txt(""), txt(""),
    fx('IFERROR(AVERAGEIF(B{s}:B{e},"<>",E{s}:E{e}),0)'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('IFERROR(AVERAGEIF(B{s}:B{e},"<>",F{s}:F{e}),0)'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('IFERROR(AVERAGEIF(B{s}:B{e},"<>",G{s}:G{e}),0)'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('IFERROR(AVERAGEIF(B{s}:B{e},"<>",H{s}:H{e}),0)'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('IFERROR(AVERAGEIF(B{s}:B{e},"<>",I{s}:I{e}),0)'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('IFERROR(AVERAGEIF(B{s}:B{e},"<>",J{s}:J{e}),0)'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('IFERROR(AVERAGEIF(B{s}:B{e},"<>",K{s}:K{e}),0)'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('IFERROR(AVERAGEIF(B{s}:B{e},"<>",L{s}:L{e}),0)'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('IFERROR(AVERAGEIF(B{s}:B{e},"<>",M{s}:M{e}),0)'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    txt(""),
])


# =========================================================================
# SHEET: REKAP MINGGUAN
# =========================================================================
# Minggu 1: hari 1-7, Minggu 2: 8-14, dst.
# Untuk setiap minggu: SUMIF ke Rekap Harian!D (kolom minggu ke)

mingguan_rows = [
    [txt("REKAP MINGGUAN PENGELUARAN", 2)],
    [txt("Agregasi mingguan otomatis. Surplus = hemat, Defisit = boros", 31)],
    [hdr("Minggu Ke"), hdr("Periode"),
     hdr("Pagu Pangan", 21), hdr("Aktual Pangan", 21), hdr("Selisih Pangan", 21),
     hdr("Pagu Operasional", 20), hdr("Aktual Operasional", 20), hdr("Selisih Operasional", 20),
     hdr("Total Pagu"), hdr("Total Aktual"), hdr("Surplus/Defisit"), hdr("Status")],
]
MINGGUAN_DATA_START = 4
for w in range(1, MAX_WEEKS + 1):
    r = MINGGUAN_DATA_START + w - 1
    start_offset = (w - 1) * 7
    end_offset = w * 7 - 1
    mingguan_rows.append([
        num(w, 15),
        fx(
            'IF(Pengaturan!$B$3+{s}>Pengaturan!$B$4,"",'
            'TEXT(Pengaturan!$B$3+{s},"dd mmm")&" - "&'
            'TEXT(MIN(Pengaturan!$B$3+{e},Pengaturan!$B$4),"dd mmm"))'.format(s=start_offset, e=end_offset),
            15
        ),
        fx("SUMIF('Rekap Harian'!$D${s}:$D${e},A{r},'Rekap Harian'!$E${s}:$E${e})".format(s=REKAP_DATA_START, e=REKAP_DATA_END, r=r), 3),
        fx("SUMIF('Rekap Harian'!$D${s}:$D${e},A{r},'Rekap Harian'!$F${s}:$F${e})".format(s=REKAP_DATA_START, e=REKAP_DATA_END, r=r), 3),
        fx("C{r}-D{r}".format(r=r), 17),
        fx("SUMIF('Rekap Harian'!$D${s}:$D${e},A{r},'Rekap Harian'!$H${s}:$H${e})".format(s=REKAP_DATA_START, e=REKAP_DATA_END, r=r), 3),
        fx("SUMIF('Rekap Harian'!$D${s}:$D${e},A{r},'Rekap Harian'!$I${s}:$I${e})".format(s=REKAP_DATA_START, e=REKAP_DATA_END, r=r), 3),
        fx("F{r}-G{r}".format(r=r), 17),
        fx("C{r}+F{r}".format(r=r), 14),
        fx("D{r}+G{r}".format(r=r), 14),
        fx("I{r}-J{r}".format(r=r), 17),
        fx('IF(I{r}=0,"",IF(K{r}>0,"SURPLUS",IF(K{r}<0,"DEFISIT","AMAN")))'.format(r=r), 26),
    ])
MINGGUAN_DATA_END = MINGGUAN_DATA_START + MAX_WEEKS - 1
mingguan_rows.append([txt("")])
mingguan_rows.append([
    txt("TOTAL", 5), txt("", 5),
    fx('SUM(C{s}:C{e})'.format(s=MINGGUAN_DATA_START, e=MINGGUAN_DATA_END), 7),
    fx('SUM(D{s}:D{e})'.format(s=MINGGUAN_DATA_START, e=MINGGUAN_DATA_END), 7),
    fx('SUM(E{s}:E{e})'.format(s=MINGGUAN_DATA_START, e=MINGGUAN_DATA_END), 7),
    fx('SUM(F{s}:F{e})'.format(s=MINGGUAN_DATA_START, e=MINGGUAN_DATA_END), 7),
    fx('SUM(G{s}:G{e})'.format(s=MINGGUAN_DATA_START, e=MINGGUAN_DATA_END), 7),
    fx('SUM(H{s}:H{e})'.format(s=MINGGUAN_DATA_START, e=MINGGUAN_DATA_END), 7),
    fx('SUM(I{s}:I{e})'.format(s=MINGGUAN_DATA_START, e=MINGGUAN_DATA_END), 7),
    fx('SUM(J{s}:J{e})'.format(s=MINGGUAN_DATA_START, e=MINGGUAN_DATA_END), 7),
    fx('SUM(K{s}:K{e})'.format(s=MINGGUAN_DATA_START, e=MINGGUAN_DATA_END), 7),
    txt(""),
])


# =========================================================================
# SHEET: DASHBOARD
# =========================================================================
# Layout (row number):
#  1  TITLE
#  2  subtitle
#  3  periode
#  4  -
#  5  HERO KPI BAR (4 cards) -- label
#  6  HERO KPI BAR (4 cards) -- values
#  7  -
#  8  SECTION A. PENGELUARAN BAHAN PANGAN
# (rows 9-17)
# 18  SECTION B. PENGELUARAN OPERASIONAL
# (rows 19-27)
# 28  SECTION C. RINGKASAN TOTAL
# (rows 29-37)
# 38  -
# 39  SECTION D. ANALISIS HARIAN & MINGGUAN
# (rows 40-48)

dashboard_rows = [
    [txt("DASHBOARD LAPORAN KEUANGAN", 2)],
    [txt("SPPG Battuwinangun - Program Pemenuhan Gizi", 31)],
    [txt("Periode:", 12), fx('TEXT(Pengaturan!B3,"dd mmm yyyy")&" s/d "&TEXT(Pengaturan!B4,"dd mmm yyyy")', 15)],
    [txt("")],

    # HERO KPI CARDS (row 5 label, row 6 value) - 4 kolom
    [txt("TOTAL DANA ANGGARAN", 30), txt("TOTAL DIPAKAI", 30), txt("TOTAL SISA DANA", 30), txt("% PENYERAPAN", 30)],
    [fx("Pengaturan!B10", 22),
     fx("'Pengeluaran Pangan'!G{p}+'Pengeluaran Operasional'!G{o}".format(p=PANGAN_TOTAL_ROW, o=OP_TOTAL_ROW), 24),
     fx("A6-B6", 23),
     fx("IF(A6=0,0,B6/A6)", 25)],
    [txt("")],                                                                                                       # row 7

    # A. PANGAN
    [hdr("A. PENGELUARAN BAHAN PANGAN", 21)],                                                                        # row 8
    [txt("Rata-Rata Pengeluaran Bahan Pangan / Hari", 12),
     fx("IFERROR('Pengeluaran Pangan'!G{p}/Pengaturan!B5,0)".format(p=PANGAN_TOTAL_ROW), 16)],                      # row 9
    [txt("Pengeluaran Bahan Pangan Seharusnya / Hari", 12), fx("Pengaturan!B13", 16)],                               # row 10
    [txt("Selisih Pengeluaran Bahan Pangan / Hari", 12), fx("B10-B9", 17), txt("(positif = hemat)", 19)],           # row 11
    [txt("")],                                                                                                       # row 12
    [txt("Dana Total Bahan Pangan", 12),     fx("Pengaturan!B8", 16)],                                              # row 13
    [txt("Penggunaan Dana Pangan", 12),      fx("'Pengeluaran Pangan'!G{p}".format(p=PANGAN_TOTAL_ROW), 9)],        # row 14
    [txt("Sisa Dana Pangan", 12),            fx("B13-B14", 17)],                                                    # row 15
    [txt("% Penggunaan Dana Pangan", 12),    fx("IF(B13=0,0,B14/B13)", 6)],                                         # row 16
    [txt("")],                                                                                                       # row 17

    # B. OPERASIONAL
    [hdr("B. PENGELUARAN OPERASIONAL", 20)],                                                                         # row 18
    [txt("Rata-Rata Pengeluaran Operasional / Hari", 12),
     fx("IFERROR('Pengeluaran Operasional'!G{o}/Pengaturan!B5,0)".format(o=OP_TOTAL_ROW), 16)],                     # row 19
    [txt("Pengeluaran Operasional Seharusnya / Hari", 12), fx("Pengaturan!B14", 16)],                                # row 20
    [txt("Selisih Pengeluaran Operasional / Hari", 12), fx("B20-B19", 17), txt("(positif = hemat)", 19)],           # row 21
    [txt("")],                                                                                                       # row 22
    [txt("Dana Total Operasional", 12),      fx("Pengaturan!B9", 16)],                                              # row 23
    [txt("Penggunaan Dana Operasional", 12), fx("'Pengeluaran Operasional'!G{o}".format(o=OP_TOTAL_ROW), 9)],       # row 24
    [txt("Sisa Dana Operasional", 12),       fx("B23-B24", 17)],                                                    # row 25
    [txt("% Penggunaan Dana Operasional", 12), fx("IF(B23=0,0,B24/B23)", 6)],                                       # row 26
    [txt("")],                                                                                                       # row 27

    # C. RINGKASAN TOTAL
    [hdr("C. RINGKASAN TOTAL", 1)],                                                                                  # row 28
    [txt("Total Dana Anggaran", 12),         fx("B13+B23", 16)],                                                    # row 29
    [txt("Total Penggunaan Dana", 12),       fx("B14+B24", 16)],                                                    # row 30
    [txt("Total Sisa Dana", 12),             fx("B29-B30", 7)],                                                     # row 31
    [txt("% Penyerapan Anggaran", 12),       fx("IF(B29=0,0,B30/B29)", 6)],                                         # row 32
    [txt("")],                                                                                                       # row 33
    [txt("Rata-Rata Total Pengeluaran / Hari", 12), fx("B9+B19", 16)],                                              # row 34
    [txt("Pagu Total / Hari (Seharusnya)", 12), fx("Pengaturan!B15", 16)],                                          # row 35
    [txt("Selisih Total / Hari", 12), fx("B35-B34", 17), txt("(positif = hemat)", 19)],                             # row 36
    [txt("")],                                                                                                       # row 37

    # D. ANALISIS SURPLUS/DEFISIT
    [hdr("D. ANALISIS SURPLUS/DEFISIT", 34)],                                                                        # row 38
    [txt("Jumlah Hari SURPLUS (hemat)", 12),
     fx('COUNTIF(\'Rekap Harian\'!N{s}:N{e},"SURPLUS")'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 15)],          # row 39
    [txt("Jumlah Hari DEFISIT (boros)", 12),
     fx('COUNTIF(\'Rekap Harian\'!N{s}:N{e},"DEFISIT")'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 15)],          # row 40
    [txt("Jumlah Hari AMAN (pas pagu)", 12),
     fx('COUNTIF(\'Rekap Harian\'!N{s}:N{e},"AMAN")'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 15)],             # row 41
    [txt("")],                                                                                                       # row 42
    [txt("Total Akumulasi Surplus", 12),
     fx('SUMIF(\'Rekap Harian\'!M{s}:M{e},">0")'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 10)],                 # row 43
    [txt("Total Akumulasi Defisit", 12),
     fx('SUMIF(\'Rekap Harian\'!M{s}:M{e},"<0")'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 9)],                  # row 44
    [txt("Net Surplus/Defisit", 12), fx("B43+B44", 7)],                                                              # row 45
    [txt("")],
    [txt("Pengeluaran Tertinggi/Hari", 12),
     fx("MAX('Rekap Harian'!L{s}:L{e})".format(s=REKAP_DATA_START, e=REKAP_DATA_END), 9)],                           # row 47
    [txt("Pengeluaran Terendah/Hari", 12),
     fx('MINIFS(\'Rekap Harian\'!L{s}:L{e},\'Rekap Harian\'!L{s}:L{e},">0")'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 10)],  # row 48
    [txt("")],

    [txt("CATATAN:", 12)],
    [txt("- Dashboard update otomatis dari sheet lain. Edit hanya di Pengaturan, Pagu Harian, & Log Pengeluaran.", 19)],
    [txt("- Selisih/Surplus POSITIF = hemat, NEGATIF = boros.", 19)],
    [txt("- Lihat detail harian di 'Rekap Harian', detail mingguan di 'Rekap Mingguan'.", 19)],
    [txt("- Visualisasi tersedia di sheet 'Grafik'.", 19)],
]


# =========================================================================
# SHEET: GRAFIK (data chart + drawing)
# =========================================================================
# Tabel data pendukung chart
grafik_rows = [
    [txt("VISUALISASI LAPORAN KEUANGAN", 2)],
    [txt("Grafik profesional di sebelah kanan dan bawah tabel data", 31)],
    [txt("")],

    # ----------------- DATA 1: Tren Harian -----------------
    [sec("1. Tren Harian - Pangan & Operasional vs Pagu")],
    [hdr("Tgl"), hdr("Aktual Pangan"), hdr("Pagu Pangan"),
     hdr("Aktual Operasional"), hdr("Pagu Operasional")],
]
# rows 6..5+MAX_DAYS contain daily data
GRAFIK_TREN_START = 6
for i in range(MAX_DAYS):
    r_rekap = REKAP_DATA_START + i
    grafik_rows.append([
        fx("'Rekap Harian'!B{r}".format(r=r_rekap), 4),
        fx("'Rekap Harian'!F{r}".format(r=r_rekap), 3),
        fx("'Rekap Harian'!E{r}".format(r=r_rekap), 3),
        fx("'Rekap Harian'!I{r}".format(r=r_rekap), 3),
        fx("'Rekap Harian'!H{r}".format(r=r_rekap), 3),
    ])
GRAFIK_TREN_END = GRAFIK_TREN_START + MAX_DAYS - 1
grafik_rows.append([txt("")])

# ----------------- DATA 2: Surplus/Defisit Harian -----------------
grafik_rows.append([sec("2. Surplus vs Defisit Per Hari")])
grafik_rows.append([hdr("Tgl"), hdr("Surplus (Rp)"), hdr("Defisit (Rp)")])
GRAFIK_SD_START = len(grafik_rows) + 1
for i in range(MAX_DAYS):
    r_rekap = REKAP_DATA_START + i
    grafik_rows.append([
        fx("'Rekap Harian'!B{r}".format(r=r_rekap), 4),
        fx("IF('Rekap Harian'!M{r}>0,'Rekap Harian'!M{r},0)".format(r=r_rekap), 10),
        fx("IF('Rekap Harian'!M{r}<0,-'Rekap Harian'!M{r},0)".format(r=r_rekap), 9),
    ])
GRAFIK_SD_END = GRAFIK_SD_START + MAX_DAYS - 1
grafik_rows.append([txt("")])

# ----------------- DATA 3: Rekap Mingguan -----------------
grafik_rows.append([sec("3. Rekap Mingguan - Aktual vs Pagu")])
grafik_rows.append([hdr("Minggu"), hdr("Aktual Pangan"), hdr("Pagu Pangan"),
                    hdr("Aktual Operasional"), hdr("Pagu Operasional")])
GRAFIK_MG_START = len(grafik_rows) + 1
for w in range(1, MAX_WEEKS + 1):
    r = MINGGUAN_DATA_START + w - 1
    grafik_rows.append([
        fx('"Mg "&A{r}'.format(r=r), 15),
        fx("'Rekap Mingguan'!D{r}".format(r=r), 3),
        fx("'Rekap Mingguan'!C{r}".format(r=r), 3),
        fx("'Rekap Mingguan'!G{r}".format(r=r), 3),
        fx("'Rekap Mingguan'!F{r}".format(r=r), 3),
    ])
GRAFIK_MG_END = GRAFIK_MG_START + MAX_WEEKS - 1
grafik_rows.append([txt("")])

# ----------------- DATA 4: Komposisi Pengeluaran -----------------
grafik_rows.append([sec("4. Komposisi Pengeluaran")])
grafik_rows.append([hdr("Kategori"), hdr("Jumlah (Rp)"), hdr("Persentase")])
GRAFIK_KOMP_START = len(grafik_rows) + 1
grafik_rows.append([
    txt("Bahan Pangan"),
    fx("'Pengeluaran Pangan'!G{p}".format(p=PANGAN_TOTAL_ROW), 3),
    fx("IF(('Pengeluaran Pangan'!G{p}+'Pengeluaran Operasional'!G{o})=0,0,"
       "B{k}/('Pengeluaran Pangan'!G{p}+'Pengeluaran Operasional'!G{o}))".format(
           p=PANGAN_TOTAL_ROW, o=OP_TOTAL_ROW, k=GRAFIK_KOMP_START), 6),
])
grafik_rows.append([
    txt("Operasional"),
    fx("'Pengeluaran Operasional'!G{o}".format(o=OP_TOTAL_ROW), 3),
    fx("IF(('Pengeluaran Pangan'!G{p}+'Pengeluaran Operasional'!G{o})=0,0,"
       "B{k}/('Pengeluaran Pangan'!G{p}+'Pengeluaran Operasional'!G{o}))".format(
           p=PANGAN_TOTAL_ROW, o=OP_TOTAL_ROW, k=GRAFIK_KOMP_START + 1), 6),
])
GRAFIK_KOMP_END = GRAFIK_KOMP_START + 1
grafik_rows.append([txt("")])

# ----------------- DATA 5: Anggaran vs Realisasi -----------------
grafik_rows.append([sec("5. Anggaran vs Realisasi vs Sisa")])
grafik_rows.append([hdr("Kategori"), hdr("Dana Total"), hdr("Terpakai"), hdr("Sisa")])
GRAFIK_AR_START = len(grafik_rows) + 1
grafik_rows.append([
    txt("Bahan Pangan"),
    fx("Pengaturan!B8", 3),
    fx("'Pengeluaran Pangan'!G{p}".format(p=PANGAN_TOTAL_ROW), 3),
    fx("B{r}-C{r}".format(r=GRAFIK_AR_START), 3),
])
grafik_rows.append([
    txt("Operasional"),
    fx("Pengaturan!B9", 3),
    fx("'Pengeluaran Operasional'!G{o}".format(o=OP_TOTAL_ROW), 3),
    fx("B{r}-C{r}".format(r=GRAFIK_AR_START + 1), 3),
])
GRAFIK_AR_END = GRAFIK_AR_START + 1


# =========================================================================
# SHEET: KATEGORI
# =========================================================================
kategori_rows = [
    [txt("DAFTAR KATEGORI (REFERENSI)", 2)],
    [txt("Gunakan nama kategori yang konsisten agar SUMIF akurat", 31)],
    [hdr("Kategori Bahan Pangan"), hdr("Kategori Operasional")],
]
kat_pangan_list = ["Karbohidrat", "Protein Hewani", "Protein Nabati",
                   "Sayuran", "Buah", "Susu & Olahan", "Bumbu", "Minuman", "Lain-lain"]
kat_op_list = ["Bahan Bakar", "Tenaga Kerja", "Utilitas", "Sanitasi",
               "Packaging", "Transportasi", "Peralatan", "Pemeliharaan",
               "Administrasi", "Lain-lain"]
for i in range(max(len(kat_pangan_list), len(kat_op_list))):
    a = kat_pangan_list[i] if i < len(kat_pangan_list) else ""
    b = kat_op_list[i] if i < len(kat_op_list) else ""
    kategori_rows.append([txt(a), txt(b)])


# =========================================================================
# SHEETS CONFIG
# (name, rows, col_widths, merge_A1, tab_color)
# =========================================================================
SHEETS = [
    ("Dashboard",              dashboard_rows,   [(1, 44), (2, 24), (3, 24), (4, 24)], "A1:D1", "1F4E79"),
    ("Rekap Harian",           rekap_rows,       [(1,5),(2,14),(3,14),(4,7),(5,16),(6,16),(7,16),(8,16),(9,16),(10,16),(11,16),(12,16),(13,18),(14,14)], "A1:N1", "548235"),
    ("Rekap Mingguan",         mingguan_rows,    [(1,10),(2,22),(3,16),(4,16),(5,16),(6,16),(7,16),(8,16),(9,16),(10,16),(11,18),(12,14)], "A1:L1", "548235"),
    ("Grafik",                 grafik_rows,      [(1, 14), (2, 20), (3, 20), (4, 20), (5, 20)], "A1:E1", "C00000"),
    ("Pengaturan",             pengaturan_rows,  [(1, 38), (2, 22), (3, 30)], "A1:C1", "7030A0"),
    ("Pagu Harian",            pagu_rows,        [(1, 14), (2, 22), (3, 22), (4, 22), (5, 28)], "A1:E1", "BF8F00"),
    ("Pengeluaran Pangan",     pangan_rows,      [(1, 14), (2, 28), (3, 18), (4, 8), (5, 10), (6, 18), (7, 18), (8, 22)], "A1:H1", "548235"),
    ("Pengeluaran Operasional", operasional_rows,[(1, 14), (2, 28), (3, 18), (4, 8), (5, 10), (6, 18), (7, 18), (8, 22)], "A1:H1", "ED7D31"),
    ("Kategori",               kategori_rows,    [(1, 28), (2, 28)], "A1:B1", "808080"),
]


# =========================================================================
# XML BUILDERS
# =========================================================================

def col_letter(idx):
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def iso_to_serial(iso):
    y, m, d = map(int, iso.split("-"))
    base = date(1899, 12, 30)
    return (date(y, m, d) - base).days


def build_sheet_xml(rows, merge_range, tab_color=None, has_drawing=False, freeze=None):
    out = []
    out.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    out.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    if tab_color:
        out.append('<sheetPr><tabColor rgb="FF{0}"/></sheetPr>'.format(tab_color))
    # sheet views with freeze panes
    if freeze:
        fr_row, fr_col = freeze
        top_left = "{0}{1}".format(col_letter(fr_col + 1), fr_row + 1)
        out.append('<sheetViews><sheetView workbookViewId="0">'
                   '<pane xSplit="{xs}" ySplit="{ys}" topLeftCell="{tl}" activePane="bottomRight" state="frozen"/>'
                   '<selection pane="bottomRight" activeCell="{tl}" sqref="{tl}"/>'
                   '</sheetView></sheetViews>'.format(xs=fr_col, ys=fr_row, tl=top_left))
    else:
        out.append('<sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>')
    out.append('<sheetFormatPr defaultRowHeight="17"/>')
    out.append("__COLS__")
    out.append("<sheetData>")
    for r_idx, row in enumerate(rows, start=1):
        if not row:
            continue
        # special row height for title (row 1)
        if r_idx == 1:
            out.append('<row r="1" ht="32" customHeight="1">')
        elif r_idx == 2:
            out.append('<row r="2" ht="20" customHeight="1">')
        else:
            out.append('<row r="{0}">'.format(r_idx))
        for col_idx, cell in enumerate(row, start=1):
            val, style, is_formula = cell
            ref = "{0}{1}".format(col_letter(col_idx), r_idx)
            if val == "" and not is_formula:
                continue
            if is_formula:
                out.append('<c r="{0}" s="{1}"><f>{2}</f></c>'.format(ref, style, escape(str(val))))
            else:
                if style == 4 and isinstance(val, str):
                    serial = iso_to_serial(val)
                    out.append('<c r="{0}" s="{1}"><v>{2}</v></c>'.format(ref, style, serial))
                elif isinstance(val, (int, float)):
                    out.append('<c r="{0}" s="{1}"><v>{2}</v></c>'.format(ref, style, val))
                else:
                    out.append('<c r="{0}" s="{1}" t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>'.format(
                        ref, style, escape(str(val))))
        out.append("</row>")
    out.append("</sheetData>")
    if merge_range:
        out.append('<mergeCells count="1"><mergeCell ref="{0}"/></mergeCells>'.format(merge_range))
    if has_drawing:
        out.append('<drawing r:id="rId1"/>')
    out.append("</worksheet>")
    return "\n".join(out)


def build_cols_block(widths):
    if not widths:
        return ""
    parts = ["<cols>"]
    for idx, w in widths:
        parts.append('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>'.format(idx, w))
    parts.append("</cols>")
    return "\n".join(parts)


# =========================================================================
# STYLES XML - Professional Edition
# =========================================================================
STYLES_XML = r'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="4">
    <numFmt numFmtId="164" formatCode="&quot;Rp&quot;\ #,##0;[Red]&quot;Rp&quot;\ \-#,##0"/>
    <numFmt numFmtId="165" formatCode="dd\ mmm\ yyyy"/>
    <numFmt numFmtId="166" formatCode="0.0%"/>
    <numFmt numFmtId="167" formatCode="#,##0"/>
  </numFmts>
  <fonts count="18">
    <font><sz val="11"/><color rgb="FF262626"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="20"/><color rgb="FF1F4E79"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF262626"/><name val="Calibri"/></font>
    <font><sz val="11"/><color rgb="FFC00000"/><name val="Calibri"/></font>
    <font><sz val="11"/><color rgb="FF375623"/><name val="Calibri"/></font>
    <font><i/><sz val="9"/><color rgb="FF808080"/><name val="Calibri"/></font>
    <font><b/><i/><sz val="11"/><color rgb="FF262626"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FF1F4E79"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="18"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="9"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><i/><sz val="10"/><color rgb="FF595959"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF375623"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF833C0C"/><name val="Calibri"/></font>
    <font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="FF1F4E79"/><name val="Calibri"/></font>
    <font><b/><sz val="9"/><color rgb="FF404040"/><name val="Calibri"/></font>
  </fonts>
  <fills count="16">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF548235"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFED7D31"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF70AD47"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFC00000"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFBF8F00"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFEBEB"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill>
  </fills>
  <borders count="5">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFBFBFBF"/></left>
      <right style="thin"><color rgb="FFBFBFBF"/></right>
      <top style="thin"><color rgb="FFBFBFBF"/></top>
      <bottom style="thin"><color rgb="FFBFBFBF"/></bottom>
    </border>
    <border>
      <left style="medium"><color rgb="FF1F4E79"/></left>
      <right style="medium"><color rgb="FF1F4E79"/></right>
      <top style="medium"><color rgb="FF1F4E79"/></top>
      <bottom style="medium"><color rgb="FF1F4E79"/></bottom>
    </border>
    <border>
      <bottom style="medium"><color rgb="FF1F4E79"/></bottom>
    </border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FFD9D9D9"/></top>
      <bottom style="thin"><color rgb="FFD9D9D9"/></bottom>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="36">
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0"   fontId="1" fillId="2" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0"   fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="4" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="4" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0"   fontId="3" fillId="3" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="166" fontId="3" fillId="0" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="164" fontId="3" fillId="4" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="4" fillId="0" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="5" fillId="0" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
    <xf numFmtId="167" fontId="0" fillId="0" borderId="4" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0"   fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0"   fontId="9" fillId="5" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
    <xf numFmtId="164" fontId="3" fillId="0" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
    <xf numFmtId="0"   fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="164" fontId="8" fillId="4" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="13" fillId="8" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="14" fillId="8" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0"   fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0"   fontId="9" fillId="7" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
    <xf numFmtId="0"   fontId="9" fillId="6" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
    <xf numFmtId="164" fontId="10" fillId="9" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="10" fillId="10" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="10" fillId="11" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="166" fontId="10" fillId="12" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="11" fillId="10" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="11" fillId="11" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="11" fillId="9" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="1" fillId="2" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="17" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0"   fontId="12" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0"   fontId="9" fillId="2" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
    <xf numFmtId="167" fontId="16" fillId="15" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="9" fillId="7" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
    <xf numFmtId="164" fontId="13" fillId="5" borderId="4" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'''


# =========================================================================
# CHART XML BUILDERS (Professional Edition)
# =========================================================================

def chart_tren_pangan():
    """Line chart: Aktual Pangan vs Pagu Pangan harian"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/>
        <a:p><a:r><a:rPr lang="id-ID" b="1" sz="1600"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:rPr><a:t>Tren Harian: Bahan Pangan (Aktual vs Pagu)</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Grafik!$B$5</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Aktual Pangan</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="548235"/></a:solidFill><a:round/></a:ln></c:spPr>
          <c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="548235"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="548235"/></a:solidFill></a:ln></c:spPr></c:marker>
          <c:cat><c:numRef><c:f>Grafik!$A$''' + str(GRAFIK_TREN_START) + ''':$A$''' + str(GRAFIK_TREN_END) + '''</c:f></c:numRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$B$''' + str(GRAFIK_TREN_START) + ''':$B$''' + str(GRAFIK_TREN_END) + '''</c:f></c:numRef></c:val>
          <c:smooth val="0"/>
        </c:ser>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/>
          <c:tx><c:strRef><c:f>Grafik!$C$5</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Pagu Pangan</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:ln w="22225" cap="rnd"><a:solidFill><a:srgbClr val="A9D18E"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>
          <c:marker><c:symbol val="none"/></c:marker>
          <c:cat><c:numRef><c:f>Grafik!$A$''' + str(GRAFIK_TREN_START) + ''':$A$''' + str(GRAFIK_TREN_END) + '''</c:f></c:numRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$C$''' + str(GRAFIK_TREN_START) + ''':$C$''' + str(GRAFIK_TREN_END) + '''</c:f></c:numRef></c:val>
          <c:smooth val="0"/>
        </c:ser>
        <c:marker val="1"/>
        <c:axId val="101"/><c:axId val="102"/>
      </c:lineChart>
      <c:catAx><c:axId val="101"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="dd/mm" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:txPr><a:bodyPr rot="-2700000"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr><c:crossAx val="102"/></c:catAx>
      <c:valAx><c:axId val="102"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E7E6E6"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="&quot;Rp&quot;#,##0" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="101"/></c:valAx>
      <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>
  <c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>
</c:chartSpace>
'''


def chart_tren_operasional():
    """Line chart: Aktual Operasional vs Pagu Operasional"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:rPr lang="id-ID" b="1" sz="1600"><a:solidFill><a:srgbClr val="BF6A1E"/></a:solidFill></a:rPr><a:t>Tren Harian: Operasional (Aktual vs Pagu)</a:t></a:r></a:p>
    </c:rich></c:tx><c:overlay val="0"/></c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Grafik!$D$5</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Aktual Operasional</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></a:ln></c:spPr>
          <c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></a:ln></c:spPr></c:marker>
          <c:cat><c:numRef><c:f>Grafik!$A$''' + str(GRAFIK_TREN_START) + ''':$A$''' + str(GRAFIK_TREN_END) + '''</c:f></c:numRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$D$''' + str(GRAFIK_TREN_START) + ''':$D$''' + str(GRAFIK_TREN_END) + '''</c:f></c:numRef></c:val>
          <c:smooth val="0"/>
        </c:ser>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/>
          <c:tx><c:strRef><c:f>Grafik!$E$5</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Pagu Operasional</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:ln w="22225" cap="rnd"><a:solidFill><a:srgbClr val="F4B183"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>
          <c:marker><c:symbol val="none"/></c:marker>
          <c:cat><c:numRef><c:f>Grafik!$A$''' + str(GRAFIK_TREN_START) + ''':$A$''' + str(GRAFIK_TREN_END) + '''</c:f></c:numRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$E$''' + str(GRAFIK_TREN_START) + ''':$E$''' + str(GRAFIK_TREN_END) + '''</c:f></c:numRef></c:val>
          <c:smooth val="0"/>
        </c:ser>
        <c:marker val="1"/>
        <c:axId val="201"/><c:axId val="202"/>
      </c:lineChart>
      <c:catAx><c:axId val="201"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="dd/mm" sourceLinked="0"/><c:majorTickMark val="out"/><c:tickLblPos val="nextTo"/><c:txPr><a:bodyPr rot="-2700000"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr><c:crossAx val="202"/></c:catAx>
      <c:valAx><c:axId val="202"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E7E6E6"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="&quot;Rp&quot;#,##0" sourceLinked="0"/><c:majorTickMark val="out"/><c:tickLblPos val="nextTo"/><c:crossAx val="201"/></c:valAx>
      <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>
</c:chartSpace>
'''


def chart_surplus_defisit():
    """Column chart: Surplus vs Defisit harian"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:rPr lang="id-ID" b="1" sz="1600"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:rPr><a:t>Surplus vs Defisit Harian</a:t></a:r></a:p>
    </c:rich></c:tx><c:overlay val="0"/></c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Grafik!$B$''' + str(GRAFIK_SD_START - 1) + '''</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Surplus</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></c:spPr>
          <c:cat><c:numRef><c:f>Grafik!$A$''' + str(GRAFIK_SD_START) + ''':$A$''' + str(GRAFIK_SD_END) + '''</c:f></c:numRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$B$''' + str(GRAFIK_SD_START) + ''':$B$''' + str(GRAFIK_SD_END) + '''</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/>
          <c:tx><c:strRef><c:f>Grafik!$C$''' + str(GRAFIK_SD_START - 1) + '''</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Defisit</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></c:spPr>
          <c:cat><c:numRef><c:f>Grafik!$A$''' + str(GRAFIK_SD_START) + ''':$A$''' + str(GRAFIK_SD_END) + '''</c:f></c:numRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$C$''' + str(GRAFIK_SD_START) + ''':$C$''' + str(GRAFIK_SD_END) + '''</c:f></c:numRef></c:val>
        </c:ser>
        <c:gapWidth val="80"/>
        <c:overlap val="0"/>
        <c:axId val="301"/><c:axId val="302"/>
      </c:barChart>
      <c:catAx><c:axId val="301"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="dd/mm" sourceLinked="0"/><c:majorTickMark val="out"/><c:tickLblPos val="nextTo"/><c:txPr><a:bodyPr rot="-2700000"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr><c:crossAx val="302"/></c:catAx>
      <c:valAx><c:axId val="302"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E7E6E6"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="&quot;Rp&quot;#,##0" sourceLinked="0"/><c:majorTickMark val="out"/><c:tickLblPos val="nextTo"/><c:crossAx val="301"/></c:valAx>
      <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
  <c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>
</c:chartSpace>
'''


def chart_mingguan():
    """Clustered bar chart: rekap mingguan"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:rPr lang="id-ID" b="1" sz="1600"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:rPr><a:t>Rekap Mingguan - Aktual vs Pagu</a:t></a:r></a:p>
    </c:rich></c:tx><c:overlay val="0"/></c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Grafik!$B$''' + str(GRAFIK_MG_START - 1) + '''</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Aktual Pangan</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="548235"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Grafik!$A$''' + str(GRAFIK_MG_START) + ''':$A$''' + str(GRAFIK_MG_END) + '''</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$B$''' + str(GRAFIK_MG_START) + ''':$B$''' + str(GRAFIK_MG_END) + '''</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/>
          <c:tx><c:strRef><c:f>Grafik!$C$''' + str(GRAFIK_MG_START - 1) + '''</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Pagu Pangan</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="A9D18E"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Grafik!$A$''' + str(GRAFIK_MG_START) + ''':$A$''' + str(GRAFIK_MG_END) + '''</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$C$''' + str(GRAFIK_MG_START) + ''':$C$''' + str(GRAFIK_MG_END) + '''</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2"/><c:order val="2"/>
          <c:tx><c:strRef><c:f>Grafik!$D$''' + str(GRAFIK_MG_START - 1) + '''</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Aktual Operasional</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Grafik!$A$''' + str(GRAFIK_MG_START) + ''':$A$''' + str(GRAFIK_MG_END) + '''</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$D$''' + str(GRAFIK_MG_START) + ''':$D$''' + str(GRAFIK_MG_END) + '''</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="3"/><c:order val="3"/>
          <c:tx><c:strRef><c:f>Grafik!$E$''' + str(GRAFIK_MG_START - 1) + '''</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Pagu Operasional</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="F4B183"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Grafik!$A$''' + str(GRAFIK_MG_START) + ''':$A$''' + str(GRAFIK_MG_END) + '''</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$E$''' + str(GRAFIK_MG_START) + ''':$E$''' + str(GRAFIK_MG_END) + '''</c:f></c:numRef></c:val>
        </c:ser>
        <c:gapWidth val="100"/>
        <c:overlap val="-10"/>
        <c:axId val="401"/><c:axId val="402"/>
      </c:barChart>
      <c:catAx><c:axId val="401"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="out"/><c:tickLblPos val="nextTo"/><c:crossAx val="402"/></c:catAx>
      <c:valAx><c:axId val="402"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E7E6E6"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="&quot;Rp&quot;#,##0" sourceLinked="0"/><c:majorTickMark val="out"/><c:tickLblPos val="nextTo"/><c:crossAx val="401"/></c:valAx>
      <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
  <c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>
</c:chartSpace>
'''


def chart_donut():
    """Donut chart: komposisi pengeluaran"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:rPr lang="id-ID" b="1" sz="1600"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:rPr><a:t>Komposisi Pengeluaran</a:t></a:r></a:p>
    </c:rich></c:tx><c:overlay val="0"/></c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:doughnutChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Grafik!$B$''' + str(GRAFIK_KOMP_START - 1) + '''</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Jumlah</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="548235"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>
          <c:dPt><c:idx val="1"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>
          <c:dLbls>
            <c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>
            <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1"/></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>
            <c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/>
          </c:dLbls>
          <c:cat><c:strRef><c:f>Grafik!$A$''' + str(GRAFIK_KOMP_START) + ''':$A$''' + str(GRAFIK_KOMP_END) + '''</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$B$''' + str(GRAFIK_KOMP_START) + ''':$B$''' + str(GRAFIK_KOMP_END) + '''</c:f></c:numRef></c:val>
        </c:ser>
        <c:firstSliceAng val="0"/>
        <c:holeSize val="50"/>
      </c:doughnutChart>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100"/></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr></c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
  <c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>
</c:chartSpace>
'''


# =========================================================================
# DRAWING (posisi chart di sheet Grafik)
# =========================================================================
# Layout chart di grid kolom-G sampai kolom-O:
#   Chart 1 (G2:O22)   - Tren Pangan
#   Chart 2 (G23:O43)  - Tren Operasional
#   Chart 3 (G44:O64)  - Surplus/Defisit
#   Chart 4 (G65:O85)  - Rekap Mingguan
#   Chart 5 (G86:M105) - Donut Komposisi

def _anchor(col_f, row_f, col_t, row_t, chart_num):
    return '''<xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>{cf}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{rf}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>{ct}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{rt}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="{n}" name="Chart {n}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId{n}"/>
      </a:graphicData></a:graphic>
    </xdr:graphicFrame><xdr:clientData/>
  </xdr:twoCellAnchor>'''.format(cf=col_f, rf=row_f, ct=col_t, rt=row_t, n=chart_num)


DRAWING_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
''' + _anchor(6, 1, 18, 22, 1) + '''
''' + _anchor(6, 23, 18, 44, 2) + '''
''' + _anchor(6, 45, 18, 66, 3) + '''
''' + _anchor(6, 67, 18, 88, 4) + '''
''' + _anchor(6, 89, 14, 110, 5) + '''
</xdr:wsDr>
'''


DRAWING_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart4.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart5.xml"/>
</Relationships>
'''


# =========================================================================
# PACKAGE
# =========================================================================

def build_content_types(n_sheets, n_charts):
    overrides = []
    for i in range(1, n_sheets + 1):
        overrides.append(
            '<Override PartName="/xl/worksheets/sheet{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(i))
    overrides.append('<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>')
    for i in range(1, n_charts + 1):
        overrides.append(
            '<Override PartName="/xl/charts/chart{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'.format(i))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + "".join(overrides)
        + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    )


ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '</Relationships>'
)


def build_workbook_xml(sheets):
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets>'
    ]
    for i, (name, _, _, _, _) in enumerate(sheets, start=1):
        parts.append('<sheet name="{0}" sheetId="{1}" r:id="rId{1}"/>'.format(escape(name), i))
    parts.append('</sheets></workbook>')
    return "".join(parts)


def build_workbook_rels(n_sheets):
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    ]
    for i in range(1, n_sheets + 1):
        parts.append(
            '<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{0}.xml"/>'.format(i))
    parts.append(
        '<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'.format(n_sheets + 1))
    parts.append('</Relationships>')
    return "".join(parts)


def build_sheet_rels_drawing():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
            '</Relationships>')


# =========================================================================
# MAIN
# =========================================================================
def main():
    grafik_idx = None
    for i, (name, _, _, _, _) in enumerate(SHEETS, start=1):
        if name == "Grafik":
            grafik_idx = i
            break

    # freeze panes mapping: row, col to freeze
    FREEZE = {
        "Dashboard":              (4, 0),   # freeze first 4 rows (periode info)
        "Rekap Harian":           (5, 2),   # freeze header (5 rows) & 2 cols (No, Tgl)
        "Rekap Mingguan":         (3, 1),
        "Pagu Harian":            (3, 1),
        "Pengeluaran Pangan":     (3, 0),
        "Pengeluaran Operasional": (3, 0),
    }

    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", build_content_types(len(SHEETS), 5))
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", build_workbook_xml(SHEETS))
        z.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels(len(SHEETS)))
        z.writestr("xl/styles.xml", STYLES_XML)

        for i, (name, rows, widths, merge, tab_color) in enumerate(SHEETS, start=1):
            has_drawing = (i == grafik_idx)
            freeze = FREEZE.get(name)
            xml = build_sheet_xml(rows, merge, tab_color=tab_color,
                                  has_drawing=has_drawing, freeze=freeze)
            xml = xml.replace("__COLS__", build_cols_block(widths))
            z.writestr("xl/worksheets/sheet{0}.xml".format(i), xml)

        z.writestr("xl/worksheets/_rels/sheet{0}.xml.rels".format(grafik_idx), build_sheet_rels_drawing())
        z.writestr("xl/drawings/drawing1.xml", DRAWING_XML)
        z.writestr("xl/drawings/_rels/drawing1.xml.rels", DRAWING_RELS)
        z.writestr("xl/charts/chart1.xml", chart_tren_pangan())
        z.writestr("xl/charts/chart2.xml", chart_tren_operasional())
        z.writestr("xl/charts/chart3.xml", chart_surplus_defisit())
        z.writestr("xl/charts/chart4.xml", chart_mingguan())
        z.writestr("xl/charts/chart5.xml", chart_donut())

    print("OK -> " + OUTPUT)


if __name__ == "__main__":
    main()
