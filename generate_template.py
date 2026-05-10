"""
Generator Template Laporan Keuangan SPPG (Excel .xlsx)
Dibangun hanya dengan Python standard library (zipfile + XML).

Sheet yang dihasilkan:
    1. Dashboard         - Ringkasan & KPI utama (semua tersambung)
    2. Pengaturan        - Input: dana total, target harian, periode
    3. Pengeluaran Pangan     - Log transaksi bahan pangan
    4. Pengeluaran Operasional - Log transaksi operasional
    5. Rekap Harian      - Tabel harian otomatis per kategori
    6. Grafik            - Chart native Excel (batang & garis)
    7. Kategori          - Referensi

Output: Template_Keuangan.xlsx
"""

import zipfile
from datetime import date
from xml.sax.saxutils import escape

OUTPUT = "Template_Keuangan.xlsx"

# =========================================================================
#   STYLE MAP (index sesuai cellXfs di styles.xml)
# =========================================================================
# 0  default
# 1  header (bold putih, bg biru tua)
# 2  title besar bold
# 3  currency IDR
# 4  date dd mmm yyyy
# 5  sub-header (bold, bg biru muda)
# 6  percent
# 7  total currency (bold, bg kuning)
# 8  text wrap center
# 9  currency merah (negatif / pengeluaran)
# 10 currency hijau (positif / pemasukan)
# 11 number integer
# 12 label kiri (bold italic)
# 13 title section (bold, bg hijau muda)
# 14 currency bold (bold saja)
# 15 center bold
# 16 big bold currency KPI
# 17 currency selisih (kondisional hijau)
# 18 currency selisih (kondisional merah)
# 19 small italic gray
# 20 header warna oranye (untuk operasional)
# 21 sub-header warna hijau (untuk pangan)


def cell(val, style=0, formula=False):
    return (val, style, formula)


def txt(v, s=0):
    return cell(v, s, False)


def num(v, s=11):
    return cell(v, s, False)


def cur(v, s=3):
    return cell(v, s, False)


def hdr(v, s=1):
    return cell(v, s, False)


def fx(f, s=3):
    return cell(f, s, True)


def dt(v, s=4):
    return cell(v, s, False)


def section(v):
    return cell(v, 13, False)


# =========================================================================
#   SHEET DEFINITIONS
# =========================================================================

# ---------- 2. Pengaturan (dibuat duluan karena jadi sumber acuan) ----------
# Layout:
#   Baris 3: Periode Mulai | =date
#   Baris 4: Periode Selesai | =date
#   Baris 5: Jumlah Hari | =formula
#
#   Baris 7: Dana Total Bahan Pangan
#   Baris 8: Dana Total Operasional
#   Baris 9: Total Dana
#
#   Baris 11: Target Pengeluaran Pangan/Hari (Seharusnya)
#   Baris 12: Target Pengeluaran Operasional/Hari (Seharusnya)

pengaturan_rows = [
    [txt("PENGATURAN LAPORAN KEUANGAN", 2)],
    [txt("")],
    [txt("Periode Mulai", 12),    dt("2026-05-01"), txt("<- edit sesuai periode", 19)],
    [txt("Periode Selesai", 12),  dt("2026-05-31"), txt("<- edit sesuai periode", 19)],
    [txt("Jumlah Hari Operasi", 12), fx("B4-B3+1", 11), txt("(otomatis)", 19)],
    [txt("")],
    [section("DANA ANGGARAN"), txt(""), txt("")],
    [txt("Dana Total Bahan Pangan", 12),    cur(60000000, 3), txt("<- edit dana pangan", 19)],
    [txt("Dana Total Operasional", 12),     cur(25000000, 3), txt("<- edit dana operasional", 19)],
    [txt("Total Dana Keseluruhan", 12),     fx("B8+B9", 7)],
    [txt("")],
    [section("TARGET PENGELUARAN / HARI (SEHARUSNYA)"), txt(""), txt("")],
    [txt("Target Pengeluaran Pangan/Hari", 12),      fx("IF(B5=0,0,B8/B5)", 3), txt("= Dana Pangan / Jumlah Hari", 19)],
    [txt("Target Pengeluaran Operasional/Hari", 12), fx("IF(B5=0,0,B9/B5)", 3), txt("= Dana Operasional / Jumlah Hari", 19)],
    [txt("Target Pengeluaran Total/Hari", 12),       fx("B13+B14", 14), txt("= Pangan + Operasional", 19)],
    [txt("")],
    [section("INFORMASI ORGANISASI"), txt(""), txt("")],
    [txt("Nama Instansi", 12), txt("SPPG Battuwinangun")],
    [txt("Alamat", 12),        txt("-")],
    [txt("Penanggung Jawab", 12), txt("-")],
    [txt("Periode Laporan", 12), fx('TEXT(B3,"dd mmm yyyy")&" s/d "&TEXT(B4,"dd mmm yyyy")', 15)],
    [txt("")],
    [txt("CATATAN:", 12)],
    [txt("1. Edit nilai di kolom B pada baris yang bertanda <- edit.", 19)],
    [txt("2. Semua sheet lain otomatis mengikuti nilai di sheet ini.", 19)],
    [txt("3. Jangan menghapus baris 3-14 (berisi referensi formula).", 19)],
]

# ---------- 3. Pengeluaran Pangan ----------
pangan_rows = [
    [txt("LOG PENGELUARAN BAHAN PANGAN", 2)],
    [txt("")],
    [hdr("Tanggal"), hdr("Item / Bahan"), hdr("Kategori"),
     hdr("Qty"), hdr("Satuan"), hdr("Harga Satuan (Rp)"), hdr("Jumlah (Rp)"), hdr("Keterangan")],
]
contoh_pangan = [
    ("2026-05-01", "Beras premium",      "Karbohidrat", 50, "kg",    15000, "Stok awal bulan"),
    ("2026-05-01", "Ayam potong",        "Protein Hewani", 20, "kg", 38000, "Menu hari 1"),
    ("2026-05-02", "Sayur bayam",        "Sayuran",     15, "ikat",  3500, ""),
    ("2026-05-02", "Telur ayam",         "Protein Hewani", 8, "kg",  28000, ""),
    ("2026-05-03", "Ikan nila",          "Protein Hewani", 18, "kg", 32000, ""),
    ("2026-05-03", "Wortel",             "Sayuran",     10, "kg",    8000, ""),
    ("2026-05-04", "Tempe",              "Protein Nabati", 12, "kg", 12000, ""),
    ("2026-05-04", "Minyak goreng",      "Bumbu",        5, "liter", 18000, ""),
    ("2026-05-05", "Buah pisang",        "Buah",        20, "sisir", 15000, ""),
    ("2026-05-05", "Bumbu dapur",        "Bumbu",        1, "paket", 150000, "Bawang, cabe, dll"),
]
for t, itm, kat, qty, sat, harga, ket in contoh_pangan:
    pangan_rows.append([
        dt(t), txt(itm), txt(kat), num(qty, 11), txt(sat), cur(harga, 3),
        fx("D{r}*F{r}".format(r=len(pangan_rows) + 1), 9),
        txt(ket)
    ])

# Baris kosong untuk diisi user
PANGAN_DATA_START = 4  # baris data pertama
for _ in range(100):
    pangan_rows.append([txt(""), txt(""), txt(""), txt(""), txt(""), txt(""), txt(""), txt("")])

# Total di bagian bawah
PANGAN_LAST_DATA_ROW = len(pangan_rows)  # termasuk baris kosong
pangan_rows.append([txt("")])
pangan_rows.append([
    txt(""), txt(""), txt(""), txt(""), txt(""),
    txt("TOTAL PENGELUARAN PANGAN", 5),
    fx("SUM(G{s}:G{e})".format(s=PANGAN_DATA_START, e=PANGAN_LAST_DATA_ROW), 7),
    txt("")
])


# ---------- 4. Pengeluaran Operasional ----------
operasional_rows = [
    [txt("LOG PENGELUARAN OPERASIONAL", 2)],
    [txt("")],
    [hdr("Tanggal"), hdr("Deskripsi"), hdr("Kategori"),
     hdr("Qty"), hdr("Satuan"), hdr("Harga Satuan (Rp)"), hdr("Jumlah (Rp)"), hdr("Keterangan")],
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
        txt(ket)
    ])

OP_DATA_START = 4
for _ in range(100):
    operasional_rows.append([txt(""), txt(""), txt(""), txt(""), txt(""), txt(""), txt(""), txt("")])
OP_LAST_DATA_ROW = len(operasional_rows)
operasional_rows.append([txt("")])
operasional_rows.append([
    txt(""), txt(""), txt(""), txt(""), txt(""),
    txt("TOTAL PENGELUARAN OPERASIONAL", 5),
    fx("SUM(G{s}:G{e})".format(s=OP_DATA_START, e=OP_LAST_DATA_ROW), 7),
    txt("")
])


# ---------- 5. Rekap Harian ----------
# Buat tabel dari Pengaturan!B3 (tanggal mulai) s/d Pengaturan!B4 (tanggal selesai)
# Maksimum 31 hari ditampilkan (diperluas jika perlu)
rekap_rows = [
    [txt("REKAP HARIAN PENGELUARAN", 2)],
    [txt("")],
    [txt("Periode:", 12), fx('TEXT(Pengaturan!B3,"dd mmm yyyy")&" s/d "&TEXT(Pengaturan!B4,"dd mmm yyyy")', 15)],
    [txt("")],
    [hdr("No"), hdr("Tanggal"),
     hdr("Pengeluaran Pangan (Rp)", 21), hdr("Target Pangan/Hari", 21), hdr("Selisih Pangan (Rp)", 21),
     hdr("Pengeluaran Operasional (Rp)", 20), hdr("Target Operasional/Hari", 20), hdr("Selisih Operasional (Rp)", 20),
     hdr("Total Hari (Rp)")],
]
REKAP_DATA_START = 6
MAX_DAYS = 31
for i in range(MAX_DAYS):
    r = REKAP_DATA_START + i
    rekap_rows.append([
        num(i + 1, 15),
        # Tanggal = tanggal mulai + i, tapi hanya jika <= tanggal selesai
        fx('IF(Pengaturan!$B$3+{i}>Pengaturan!$B$4,"",Pengaturan!$B$3+{i})'.format(i=i), 4),
        # Pengeluaran pangan hari ini
        fx('IF(B{r}="",0,SUMIF(\'Pengeluaran Pangan\'!A:A,B{r},\'Pengeluaran Pangan\'!G:G))'.format(r=r), 3),
        # Target pangan/hari
        fx('IF(B{r}="",0,Pengaturan!$B$13)'.format(r=r), 3),
        # Selisih pangan = target - aktual (positif = hemat)
        fx('IF(B{r}="",0,D{r}-C{r})'.format(r=r), 17),
        # Pengeluaran operasional hari ini
        fx('IF(B{r}="",0,SUMIF(\'Pengeluaran Operasional\'!A:A,B{r},\'Pengeluaran Operasional\'!G:G))'.format(r=r), 3),
        # Target operasional/hari
        fx('IF(B{r}="",0,Pengaturan!$B$14)'.format(r=r), 3),
        # Selisih operasional
        fx('IF(B{r}="",0,G{r}-F{r})'.format(r=r), 17),
        # Total hari
        fx('C{r}+F{r}'.format(r=r), 14),
    ])

REKAP_DATA_END = REKAP_DATA_START + MAX_DAYS - 1
# Baris total
rekap_rows.append([txt("")])
total_row_rekap = REKAP_DATA_END + 2
rekap_rows.append([
    txt(""),
    txt("TOTAL", 5),
    fx('SUM(C{s}:C{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('SUM(D{s}:D{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('SUM(E{s}:E{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('SUM(F{s}:F{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('SUM(G{s}:G{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('SUM(H{s}:H{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('SUM(I{s}:I{e})'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
])
# Baris rata-rata
rekap_rows.append([
    txt(""),
    txt("RATA-RATA", 5),
    fx('IFERROR(AVERAGEIF(B{s}:B{e},"<>",C{s}:C{e}),0)'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('Pengaturan!$B$13', 7),
    fx('IFERROR(AVERAGEIF(B{s}:B{e},"<>",E{s}:E{e}),0)'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('IFERROR(AVERAGEIF(B{s}:B{e},"<>",F{s}:F{e}),0)'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('Pengaturan!$B$14', 7),
    fx('IFERROR(AVERAGEIF(B{s}:B{e},"<>",H{s}:H{e}),0)'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
    fx('IFERROR(AVERAGEIF(B{s}:B{e},"<>",I{s}:I{e}),0)'.format(s=REKAP_DATA_START, e=REKAP_DATA_END), 7),
])


# ---------- 1. Dashboard ----------
# Layout baris (1-indexed):
#  1  TITLE
#  2  subtitle
#  3  -
#  4  Periode
#  5  Jumlah Hari
#  6  -
#  7  SECTION A. PENGELUARAN BAHAN PANGAN
#  8  -
#  9  Rata-Rata Pengeluaran Pangan/Hari    (B9)
# 10  Pengeluaran Pangan Seharusnya/Hari   (B10)
# 11  Selisih Pengeluaran Pangan/Hari      (B11 = B10-B9)
# 12  -
# 13  Dana Total Bahan Pangan              (B13)
# 14  Penggunaan Dana Pangan               (B14)
# 15  Sisa Dana Pangan                     (B15 = B13-B14)
# 16  % Penggunaan Dana Pangan             (B16 = B14/B13)
# 17  -
# 18  SECTION B. PENGELUARAN OPERASIONAL
# 19  -
# 20  Rata-Rata Pengeluaran Operasional/Hari   (B20)
# 21  Pengeluaran Operasional Seharusnya/Hari  (B21)
# 22  Selisih Pengeluaran Operasional/Hari     (B22 = B21-B20)
# 23  -
# 24  Dana Total Operasional                   (B24)
# 25  Penggunaan Dana Operasional              (B25)
# 26  Sisa Dana Operasional                    (B26 = B24-B25)
# 27  % Penggunaan Dana Operasional            (B27 = B25/B24)
# 28  -
# 29  SECTION C. RINGKASAN TOTAL
# 30  -
# 31  Total Dana Anggaran                      (B31 = B13+B24)
# 32  Total Penggunaan Dana                    (B32 = B14+B25)
# 33  Total Sisa Dana                          (B33 = B31-B32)
# 34  % Penyerapan Anggaran                    (B34 = B32/B31)
# 35  -
# 36  Rata-Rata Total Pengeluaran/Hari         (B36 = B9+B20)
# 37  Target Total Pengeluaran/Hari            (B37 = Pengaturan!B15)
# 38  Selisih Total/Hari                       (B38 = B37-B36)

PANGAN_TOTAL_ROW = len(pangan_rows)          # row di sheet Pangan yang berisi SUM(G)
OP_TOTAL_ROW = len(operasional_rows)         # row di sheet Operasional yang berisi SUM(G)

dashboard_rows = [
    [txt("DASHBOARD LAPORAN KEUANGAN", 2)],
    [txt("SPPG Battuwinangun - Program Pemenuhan Gizi", 19)],
    [txt("")],
    [txt("Periode:", 12), fx('TEXT(Pengaturan!B3,"dd mmm yyyy")&" s/d "&TEXT(Pengaturan!B4,"dd mmm yyyy")', 15)],
    [txt("Jumlah Hari:", 12), fx("Pengaturan!B5", 11)],
    [txt("")],

    # ============================ PANGAN ============================
    [hdr("A. PENGELUARAN BAHAN PANGAN", 21), txt(""), txt(""), txt("")],                                       # row 7
    [txt("")],                                                                                                  # row 8
    [txt("Rata-Rata Pengeluaran Bahan Pangan / Hari", 12),                                                      # row 9
     fx("IFERROR('Pengeluaran Pangan'!G{r}/Pengaturan!B5,0)".format(r=PANGAN_TOTAL_ROW), 16)],
    [txt("Pengeluaran Bahan Pangan Seharusnya / Hari", 12),                                                     # row 10
     fx("Pengaturan!B13", 16)],
    [txt("Selisih Pengeluaran Bahan Pangan / Hari", 12),                                                        # row 11
     fx("B10-B9", 17),
     txt("(positif = hemat)", 19)],
    [txt("")],                                                                                                  # row 12
    [txt("Dana Total Bahan Pangan", 12),           fx("Pengaturan!B8", 16)],                                    # row 13
    [txt("Penggunaan Dana Pangan", 12),                                                                         # row 14
     fx("'Pengeluaran Pangan'!G{r}".format(r=PANGAN_TOTAL_ROW), 9)],
    [txt("Sisa Dana Pangan", 12),                  fx("B13-B14", 17)],                                          # row 15
    [txt("% Penggunaan Dana Pangan", 12),          fx("IF(B13=0,0,B14/B13)", 6)],                               # row 16
    [txt("")],                                                                                                  # row 17

    # ============================ OPERASIONAL ============================
    [hdr("B. PENGELUARAN OPERASIONAL", 20), txt(""), txt(""), txt("")],                                        # row 18
    [txt("")],                                                                                                  # row 19
    [txt("Rata-Rata Pengeluaran Operasional / Hari", 12),                                                       # row 20
     fx("IFERROR('Pengeluaran Operasional'!G{r}/Pengaturan!B5,0)".format(r=OP_TOTAL_ROW), 16)],
    [txt("Pengeluaran Operasional Seharusnya / Hari", 12),                                                      # row 21
     fx("Pengaturan!B14", 16)],
    [txt("Selisih Pengeluaran Operasional / Hari", 12),                                                         # row 22
     fx("B21-B20", 17),
     txt("(positif = hemat)", 19)],
    [txt("")],                                                                                                  # row 23
    [txt("Dana Total Operasional", 12),            fx("Pengaturan!B9", 16)],                                    # row 24
    [txt("Penggunaan Dana Operasional", 12),                                                                    # row 25
     fx("'Pengeluaran Operasional'!G{r}".format(r=OP_TOTAL_ROW), 9)],
    [txt("Sisa Dana Operasional", 12),             fx("B24-B25", 17)],                                          # row 26
    [txt("% Penggunaan Dana Operasional", 12),     fx("IF(B24=0,0,B25/B24)", 6)],                               # row 27
    [txt("")],                                                                                                  # row 28

    # ============================ RINGKASAN TOTAL ============================
    [hdr("C. RINGKASAN TOTAL", 1), txt(""), txt(""), txt("")],                                                 # row 29
    [txt("")],                                                                                                  # row 30
    [txt("Total Dana Anggaran", 12),               fx("B13+B24", 16)],                                          # row 31
    [txt("Total Penggunaan Dana", 12),             fx("B14+B25", 16)],                                          # row 32
    [txt("Total Sisa Dana", 12),                   fx("B31-B32", 7)],                                           # row 33
    [txt("% Penyerapan Anggaran", 12),             fx("IF(B31=0,0,B32/B31)", 6)],                               # row 34
    [txt("")],                                                                                                  # row 35
    [txt("Rata-Rata Total Pengeluaran / Hari", 12), fx("B9+B20", 16)],                                          # row 36
    [txt("Target Total Pengeluaran / Hari", 12),    fx("Pengaturan!B15", 16)],                                  # row 37
    [txt("Selisih Total / Hari", 12),               fx("B37-B36", 17)],                                         # row 38
    [txt("")],
    [txt("CATATAN:", 12)],
    [txt("- Semua angka otomatis dari sheet Pengeluaran Pangan & Operasional.", 19)],
    [txt("- Edit dana & target harian di sheet 'Pengaturan'.", 19)],
    [txt("- Selisih positif = hemat (di bawah target).", 19)],
    [txt("- Selisih negatif = boros (di atas target).", 19)],
]


# ---------- 6. Grafik ----------
# Sheet ini akan berisi:
#   - Tabel data pendukung chart
#   - Chart batang (Budget vs Aktual)
#   - Chart garis (tren harian)
grafik_rows = [
    [txt("VISUALISASI KEUANGAN", 2)],
    [txt("")],
    [section("1. Data Perbandingan Anggaran vs Realisasi"), txt(""), txt(""), txt("")],
    [hdr("Kategori"), hdr("Dana Total (Rp)"), hdr("Terpakai (Rp)"), hdr("Sisa (Rp)")],
    [txt("Bahan Pangan"),   fx("Pengaturan!B8", 3), fx("Dashboard!B14", 3), fx("B5-C5", 3)],
    [txt("Operasional"),    fx("Pengaturan!B9", 3), fx("Dashboard!B25", 3), fx("B6-C6", 3)],
    [txt("")],
    [section("2. Data Rata-Rata Pengeluaran vs Target Harian"), txt(""), txt(""), txt("")],
    [hdr("Kategori"), hdr("Rata-Rata Aktual/Hari"), hdr("Target/Hari"), hdr("Selisih/Hari")],
    [txt("Bahan Pangan"),  fx("Dashboard!B9", 3),  fx("Dashboard!B10", 3),  fx("C10-B10", 17)],
    [txt("Operasional"),   fx("Dashboard!B20", 3), fx("Dashboard!B21", 3),  fx("C11-B11", 17)],
    [txt("")],
    [section("3. Komposisi Pengeluaran"), txt(""), txt(""), txt("")],
    [hdr("Kategori"), hdr("Jumlah (Rp)"), hdr("Persentase"), txt("")],
    [txt("Bahan Pangan"),  fx("Dashboard!B14", 3), fx("IF(Dashboard!B32=0,0,B15/Dashboard!B32)", 6), txt("")],
    [txt("Operasional"),   fx("Dashboard!B25", 3), fx("IF(Dashboard!B32=0,0,B16/Dashboard!B32)", 6), txt("")],
    [txt("TOTAL", 5),      fx("SUM(B15:B16)", 7),  fx("SUM(C15:C16)", 7), txt("")],
    [txt("")],
    [txt("")],
    [txt("Grafik akan ditampilkan di sebelah kanan tabel ini:", 12)],
    [txt("- Bar Chart: Anggaran vs Realisasi (A5:D6)", 19)],
    [txt("- Bar Chart: Rata-Rata vs Target Harian (A10:C11)", 19)],
    [txt("- Pie Chart: Komposisi Pengeluaran (A15:B16)", 19)],
]


# ---------- 7. Kategori ----------
kategori_rows = [
    [txt("DAFTAR KATEGORI (REFERENSI)", 2)],
    [txt("")],
    [hdr("Kategori Bahan Pangan"), hdr("Kategori Operasional")],
]
kat_pangan = ["Karbohidrat", "Protein Hewani", "Protein Nabati",
              "Sayuran", "Buah", "Susu & Olahan", "Bumbu", "Minuman", "Lain-lain"]
kat_op = ["Bahan Bakar", "Tenaga Kerja", "Utilitas", "Sanitasi",
          "Packaging", "Transportasi", "Peralatan", "Pemeliharaan",
          "Administrasi", "Lain-lain"]
for i in range(max(len(kat_pangan), len(kat_op))):
    a = kat_pangan[i] if i < len(kat_pangan) else ""
    b = kat_op[i] if i < len(kat_op) else ""
    kategori_rows.append([txt(a), txt(b)])


# =========================================================================
#   KONFIGURASI SHEET
# =========================================================================
SHEETS = [
    ("Dashboard",              dashboard_rows,   [(1, 42), (2, 22), (3, 26)], "A1:D1"),
    ("Pengaturan",             pengaturan_rows,  [(1, 38), (2, 22), (3, 30)], "A1:C1"),
    ("Pengeluaran Pangan",     pangan_rows,      [(1, 14), (2, 28), (3, 18), (4, 8), (5, 10), (6, 18), (7, 18), (8, 24)], "A1:H1"),
    ("Pengeluaran Operasional", operasional_rows,[(1, 14), (2, 28), (3, 18), (4, 8), (5, 10), (6, 18), (7, 18), (8, 24)], "A1:H1"),
    ("Rekap Harian",           rekap_rows,       [(1, 5), (2, 15), (3, 22), (4, 20), (5, 20), (6, 22), (7, 22), (8, 20), (9, 18)], "A1:I1"),
    ("Grafik",                 grafik_rows,      [(1, 28), (2, 22), (3, 22), (4, 22)], "A1:D1"),
    ("Kategori",               kategori_rows,    [(1, 28), (2, 28)], "A1:B1"),
]


# =========================================================================
#   BUILDERS
# =========================================================================

def col_letter(idx):  # 1-based
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def iso_to_serial(iso):
    """Konversi 'YYYY-MM-DD' ke serial date Excel (1900 date system)."""
    y, m, d = map(int, iso.split("-"))
    base = date(1899, 12, 30)
    return (date(y, m, d) - base).days


def build_sheet_xml(rows, merge_range, has_drawing=False, sheet_idx=None):
    out = []
    out.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    out.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    out.append('<sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>')
    out.append('<sheetFormatPr defaultRowHeight="16"/>')
    out.append("__COLS__")
    out.append("<sheetData>")
    for r_idx, row in enumerate(rows, start=1):
        if not row:
            continue
        out.append('<row r="{0}">'.format(r_idx))
        for c_idx, c in enumerate(row, start=1):
            val, style, is_formula = c
            ref = "{0}{1}".format(col_letter(c_idx), r_idx)
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
#   STYLES
# =========================================================================
STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="4">
    <numFmt numFmtId="164" formatCode="&quot;Rp&quot;\\ #,##0;[Red]&quot;Rp&quot;\\ \\-#,##0"/>
    <numFmt numFmtId="165" formatCode="dd\\ mmm\\ yyyy"/>
    <numFmt numFmtId="166" formatCode="0.0%"/>
    <numFmt numFmtId="167" formatCode="#,##0"/>
  </numFmts>
  <fonts count="10">
    <font><sz val="11"/><name val="Calibri"/></font>                                                  <!-- 0 default -->
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>                       <!-- 1 header white -->
    <font><b/><sz val="18"/><color rgb="FF1F4E79"/><name val="Calibri"/></font>                       <!-- 2 title -->
    <font><b/><sz val="11"/><name val="Calibri"/></font>                                              <!-- 3 bold -->
    <font><sz val="11"/><color rgb="FFC00000"/><name val="Calibri"/></font>                           <!-- 4 red -->
    <font><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/></font>                           <!-- 5 green -->
    <font><i/><sz val="9"/><color rgb="FF808080"/><name val="Calibri"/></font>                        <!-- 6 small italic gray -->
    <font><b/><i/><sz val="11"/><name val="Calibri"/></font>                                          <!-- 7 bold italic -->
    <font><b/><sz val="14"/><color rgb="FF1F4E79"/><name val="Calibri"/></font>                       <!-- 8 big bold KPI -->
    <font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>                       <!-- 9 section header white -->
  </fonts>
  <fills count="9">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill>  <!-- 2 biru tua -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>  <!-- 3 biru muda -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>  <!-- 4 kuning -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/><bgColor indexed="64"/></patternFill></fill>  <!-- 5 hijau muda -->
    <fill><patternFill patternType="solid"><fgColor rgb="FF548235"/><bgColor indexed="64"/></patternFill></fill>  <!-- 6 hijau tua -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFED7D31"/><bgColor indexed="64"/></patternFill></fill>  <!-- 7 oranye -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>  <!-- 8 abu sangat muda -->
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
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="22">
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"/>                                                                                           <!-- 0 default -->
    <xf numFmtId="0"   fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>  <!-- 1 header biru -->
    <xf numFmtId="0"   fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>                       <!-- 2 title -->
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>                                                     <!-- 3 currency -->
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>  <!-- 4 date -->
    <xf numFmtId="0"   fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>                                               <!-- 5 sub-header -->
    <xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>  <!-- 6 percent -->
    <xf numFmtId="164" fontId="3" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>                         <!-- 7 total -->
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf>       <!-- 8 wrap center -->
    <xf numFmtId="164" fontId="4" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>                                       <!-- 9 currency merah -->
    <xf numFmtId="164" fontId="5" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>                                       <!-- 10 currency hijau -->
    <xf numFmtId="167" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>  <!-- 11 integer -->
    <xf numFmtId="0"   fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1"/>                                                                             <!-- 12 label bold italic -->
    <xf numFmtId="0"   fontId="9" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>  <!-- 13 section hijau muda -->
    <xf numFmtId="164" fontId="3" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>                                       <!-- 14 currency bold -->
    <xf numFmtId="0"   fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>                       <!-- 15 center bold -->
    <xf numFmtId="164" fontId="8" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>                         <!-- 16 KPI besar -->
    <xf numFmtId="164" fontId="5" fillId="8" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>                         <!-- 17 selisih default (akan di-cf) -->
    <xf numFmtId="164" fontId="4" fillId="8" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>                         <!-- 18 selisih merah -->
    <xf numFmtId="0"   fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"/>                                                                             <!-- 19 small italic -->
    <xf numFmtId="0"   fontId="9" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>  <!-- 20 header oranye -->
    <xf numFmtId="0"   fontId="9" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>  <!-- 21 header hijau tua -->
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
'''


# =========================================================================
#   CHART XML (untuk sheet Grafik)
# =========================================================================

def build_chart_bar_anggaran():
    """Bar chart: Anggaran vs Realisasi vs Sisa (data di Grafik!A4:D6)"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich>
        <a:bodyPr/><a:lstStyle/>
        <a:p><a:r><a:rPr lang="id-ID" b="1" sz="1400"/><a:t>Anggaran vs Realisasi vs Sisa (Rp)</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Grafik!$B$4</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Dana Total</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Grafik!$A$5:$A$6</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Bahan Pangan</c:v></c:pt><c:pt idx="1"><c:v>Operasional</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$B$5:$B$6</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/>
          <c:tx><c:strRef><c:f>Grafik!$C$4</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Terpakai</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Grafik!$A$5:$A$6</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$C$5:$C$6</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2"/><c:order val="2"/>
          <c:tx><c:strRef><c:f>Grafik!$D$4</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Sisa</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Grafik!$A$5:$A$6</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$D$5:$D$6</c:f></c:numRef></c:val>
        </c:ser>
        <c:gapWidth val="150"/>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>
'''


def build_chart_bar_harian():
    """Bar chart: Rata-rata aktual vs target harian (Grafik!A9:C11)"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/>
        <a:p><a:r><a:rPr lang="id-ID" b="1" sz="1400"/><a:t>Rata-Rata Aktual vs Target Harian (Rp)</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="bar"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Grafik!$B$9</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Rata-Rata Aktual/Hari</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="5B9BD5"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Grafik!$A$10:$A$11</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Bahan Pangan</c:v></c:pt><c:pt idx="1"><c:v>Operasional</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$B$10:$B$11</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/>
          <c:tx><c:strRef><c:f>Grafik!$C$9</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Target/Hari</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="A5A5A5"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Grafik!$A$10:$A$11</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$C$10:$C$11</c:f></c:numRef></c:val>
        </c:ser>
        <c:gapWidth val="150"/>
        <c:axId val="3"/><c:axId val="4"/>
      </c:barChart>
      <c:catAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="4"/></c:catAx>
      <c:valAx><c:axId val="4"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="3"/></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>
'''


def build_chart_pie():
    """Pie chart: Komposisi pengeluaran (Grafik!A14:B16)"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/>
        <a:p><a:r><a:rPr lang="id-ID" b="1" sz="1400"/><a:t>Komposisi Pengeluaran</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:pie3DChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Grafik!$B$14</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Jumlah</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="1"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></c:spPr></c:dPt>
          <c:dLbls>
            <c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></c:spPr>
            <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000" b="1"/></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>
            <c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/>
          </c:dLbls>
          <c:cat><c:strRef><c:f>Grafik!$A$15:$A$16</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Bahan Pangan</c:v></c:pt><c:pt idx="1"><c:v>Operasional</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Grafik!$B$15:$B$16</c:f></c:numRef></c:val>
        </c:ser>
      </c:pie3DChart>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>
'''


# ---- drawing1.xml (posisi 3 chart di sheet Grafik) ----
DRAWING_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>13</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>

  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>21</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>13</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>39</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="3" name="Chart 2"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>

  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>40</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>13</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>58</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="4" name="Chart 3"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId3"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
'''

DRAWING_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>
</Relationships>
'''


# =========================================================================
#   PACKAGE PARTS
# =========================================================================

def build_content_types(n_sheets, grafik_idx):
    overrides = []
    for i in range(1, n_sheets + 1):
        overrides.append(
            '<Override PartName="/xl/worksheets/sheet{0}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(i))
    # drawing + charts
    overrides.append(
        '<Override PartName="/xl/drawings/drawing1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>')
    for i in range(1, 4):
        overrides.append(
            '<Override PartName="/xl/charts/chart{0}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'.format(i))
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
    for i, (name, _, _, _) in enumerate(sheets, start=1):
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


def build_sheet_rels_for_grafik():
    """Sheet grafik menunjuk ke drawing1.xml"""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
            '</Relationships>')


# =========================================================================
#   MAIN
# =========================================================================

def main():
    # cari index sheet Grafik (1-based)
    grafik_idx = None
    for i, (name, _, _, _) in enumerate(SHEETS, start=1):
        if name == "Grafik":
            grafik_idx = i
            break

    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as z:
        # package-level
        z.writestr("[Content_Types].xml", build_content_types(len(SHEETS), grafik_idx))
        z.writestr("_rels/.rels", ROOT_RELS)

        # workbook
        z.writestr("xl/workbook.xml", build_workbook_xml(SHEETS))
        z.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels(len(SHEETS)))
        z.writestr("xl/styles.xml", STYLES_XML)

        # worksheets
        for i, (name, rows, widths, merge) in enumerate(SHEETS, start=1):
            has_drawing = (i == grafik_idx)
            xml = build_sheet_xml(rows, merge, has_drawing=has_drawing, sheet_idx=i)
            xml = xml.replace("__COLS__", build_cols_block(widths))
            z.writestr("xl/worksheets/sheet{0}.xml".format(i), xml)

        # sheet rels for Grafik (links to drawing)
        z.writestr("xl/worksheets/_rels/sheet{0}.xml.rels".format(grafik_idx),
                   build_sheet_rels_for_grafik())

        # drawings + charts
        z.writestr("xl/drawings/drawing1.xml", DRAWING_XML)
        z.writestr("xl/drawings/_rels/drawing1.xml.rels", DRAWING_RELS)
        z.writestr("xl/charts/chart1.xml", build_chart_bar_anggaran())
        z.writestr("xl/charts/chart2.xml", build_chart_bar_harian())
        z.writestr("xl/charts/chart3.xml", build_chart_pie())

    print("OK -> " + OUTPUT)


if __name__ == "__main__":
    main()
