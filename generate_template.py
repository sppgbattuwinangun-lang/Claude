"""
Generator Template Keuangan (.xlsx) menggunakan Python standard library saja.
Tidak memerlukan openpyxl / pandas / xlsxwriter.

Menghasilkan: Template_Keuangan.xlsx

Sheet yang dibuat:
1. Dashboard       - Ringkasan total pemasukan, pengeluaran, saldo
2. Pemasukan       - Catatan pemasukan
3. Pengeluaran     - Catatan pengeluaran
4. Anggaran        - Anggaran bulanan vs realisasi
5. Kategori        - Referensi daftar kategori
"""

import zipfile
from xml.sax.saxutils import escape

OUTPUT = "Template_Keuangan.xlsx"

# ------------------------- DEFINISI SHEET ------------------------- #

# Setiap sheet: list of rows, setiap row list of (value, style_id, is_formula)
# style_id merujuk pada daftar style di styles.xml
# Style map:
#   0 = default
#   1 = header (bold white bg biru)
#   2 = title (bold besar)
#   3 = currency IDR
#   4 = date
#   5 = sub-header (bold bg abu)
#   6 = percent
#   7 = currency bold (total)
#   8 = text wrap center
#   9 = currency merah (minus/pengeluaran)
#  10 = currency hijau (plus/pemasukan)

def cur(v, style=3, formula=False):
    return (v, style, formula)

def txt(v, style=0):
    return (v, style, False)

def hdr(v):
    return (v, 1, False)

def sub(v):
    return (v, 5, False)

def title(v):
    return (v, 2, False)

def dt(v):
    return (v, 4, False)

def fx(f, style=3):
    return (f, style, True)


# --- Sheet Pemasukan ---
pemasukan_rows = [
    [title("CATATAN PEMASUKAN"), txt(""), txt(""), txt(""), txt("")],
    [txt("")],
    [hdr("Tanggal"), hdr("Sumber"), hdr("Kategori"), hdr("Deskripsi"), hdr("Jumlah (Rp)")],
]
# contoh data
contoh_pemasukan = [
    ("2026-05-01", "Gaji Bulanan", "Gaji", "Gaji Mei 2026", 8500000),
    ("2026-05-05", "Freelance",    "Usaha Sampingan", "Proyek desain logo", 1500000),
    ("2026-05-10", "Dividen",      "Investasi", "Dividen saham", 250000),
]
for tgl, sumber, kat, desc, jml in contoh_pemasukan:
    pemasukan_rows.append([dt(tgl), txt(sumber), txt(kat), txt(desc), cur(jml, 10)])

# Tambah baris kosong untuk diisi user (sampai baris ke ~54)
for _ in range(50):
    pemasukan_rows.append([txt(""), txt(""), txt(""), txt(""), txt("")])

# Baris total di akhir (baris 57 di Excel - 1-indexed)
total_row_pemasukan = len(pemasukan_rows) + 2
pemasukan_rows.append([txt("")])
pemasukan_rows.append([
    txt(""), txt(""), txt(""), sub("TOTAL PEMASUKAN"),
    fx("SUM(E4:E{0})".format(total_row_pemasukan - 2), 7)
])


# --- Sheet Pengeluaran ---
pengeluaran_rows = [
    [title("CATATAN PENGELUARAN"), txt(""), txt(""), txt(""), txt("")],
    [txt("")],
    [hdr("Tanggal"), hdr("Kategori"), hdr("Deskripsi"), hdr("Metode Bayar"), hdr("Jumlah (Rp)")],
]
contoh_pengeluaran = [
    ("2026-05-02", "Makanan",      "Belanja bulanan di supermarket", "Debit", 750000),
    ("2026-05-03", "Transportasi", "Isi bensin mobil",               "Tunai", 300000),
    ("2026-05-04", "Tagihan",      "Listrik & air",                  "Transfer", 450000),
    ("2026-05-06", "Hiburan",      "Nonton bioskop",                 "E-Wallet", 120000),
    ("2026-05-08", "Kesehatan",    "Beli obat & vitamin",            "Tunai", 180000),
]
for tgl, kat, desc, metode, jml in contoh_pengeluaran:
    pengeluaran_rows.append([dt(tgl), txt(kat), txt(desc), txt(metode), cur(jml, 9)])

for _ in range(60):
    pengeluaran_rows.append([txt(""), txt(""), txt(""), txt(""), txt("")])

total_row_pengeluaran = len(pengeluaran_rows) + 2
pengeluaran_rows.append([txt("")])
pengeluaran_rows.append([
    txt(""), txt(""), txt(""), sub("TOTAL PENGELUARAN"),
    fx("SUM(E4:E{0})".format(total_row_pengeluaran - 2), 7)
])


# --- Sheet Anggaran ---
anggaran_rows = [
    [title("ANGGARAN BULANAN"), txt(""), txt(""), txt(""), txt("")],
    [txt("")],
    [hdr("Kategori"), hdr("Anggaran (Rp)"), hdr("Realisasi (Rp)"), hdr("Sisa (Rp)"), hdr("% Terpakai")],
]
kategori_anggaran = [
    ("Makanan",       2500000),
    ("Transportasi",  1000000),
    ("Tagihan",       1200000),
    ("Hiburan",        500000),
    ("Kesehatan",      400000),
    ("Belanja",        800000),
    ("Pendidikan",     500000),
    ("Tabungan",      2000000),
    ("Lain-lain",      300000),
]
anggaran_start_row = 4
for i, (kat, anggaran) in enumerate(kategori_anggaran):
    r = anggaran_start_row + i  # row number di excel (1-indexed)
    anggaran_rows.append([
        txt(kat),
        cur(anggaran, 3),
        fx('SUMIF(Pengeluaran!B:B,A{0},Pengeluaran!E:E)'.format(r), 3),
        fx('B{0}-C{0}'.format(r), 3),
        fx('IF(B{0}=0,0,C{0}/B{0})'.format(r), 6),
    ])

total_anggaran_row = anggaran_start_row + len(kategori_anggaran)
anggaran_rows.append([txt("")])
anggaran_rows.append([
    sub("TOTAL"),
    fx('SUM(B{0}:B{1})'.format(anggaran_start_row, total_anggaran_row - 1), 7),
    fx('SUM(C{0}:C{1})'.format(anggaran_start_row, total_anggaran_row - 1), 7),
    fx('SUM(D{0}:D{1})'.format(anggaran_start_row, total_anggaran_row - 1), 7),
    fx('IF(B{0}=0,0,C{0}/B{0})'.format(total_anggaran_row + 1), 6),
])


# --- Sheet Kategori ---
kategori_rows = [
    [title("DAFTAR KATEGORI"), txt("")],
    [txt("")],
    [hdr("Kategori Pemasukan"), hdr("Kategori Pengeluaran")],
]
kat_in = ["Gaji", "Bonus", "Usaha Sampingan", "Investasi", "Hadiah", "Lain-lain"]
kat_out = ["Makanan", "Transportasi", "Tagihan", "Hiburan", "Kesehatan",
           "Belanja", "Pendidikan", "Tabungan", "Cicilan", "Lain-lain"]
max_len = max(len(kat_in), len(kat_out))
for i in range(max_len):
    a = kat_in[i] if i < len(kat_in) else ""
    b = kat_out[i] if i < len(kat_out) else ""
    kategori_rows.append([txt(a), txt(b)])


# --- Sheet Dashboard ---
# Dashboard berisi formula yang mereferensi sheet lain
dashboard_rows = [
    [title("DASHBOARD KEUANGAN"), txt("")],
    [txt("")],
    [sub("Periode"), txt("Mei 2026")],
    [txt("")],
    [hdr("Ringkasan"), hdr("Jumlah (Rp)")],
    [txt("Total Pemasukan"),   fx('SUM(Pemasukan!E:E)', 10)],
    [txt("Total Pengeluaran"), fx('SUM(Pengeluaran!E:E)', 9)],
    [sub("SALDO AKHIR"),       fx('B6-B7', 7)],
    [txt("")],
    [hdr("Anggaran"), hdr("Jumlah (Rp)")],
    [txt("Total Anggaran"),        fx("SUM(Anggaran!B:B)", 3)],
    [txt("Total Realisasi"),       fx("SUM(Anggaran!C:C)", 3)],
    [sub("Sisa Anggaran"),         fx("B11-B12", 7)],
    [txt("% Anggaran Terpakai"),   fx("IF(B11=0,0,B12/B11)", 6)],
    [txt("")],
    [txt("")],
    [sub("Tips:")],
    [txt("- Isi transaksi di sheet 'Pemasukan' dan 'Pengeluaran'.")],
    [txt("- Total & saldo otomatis terhitung di sheet ini.")],
    [txt("- Atur target bulanan di sheet 'Anggaran'.")],
    [txt("- Gunakan nama kategori yang konsisten agar SUMIF akurat.")],
]

SHEETS = [
    ("Dashboard",   dashboard_rows,   [(1, 28), (2, 22)]),                      # col widths
    ("Pemasukan",   pemasukan_rows,   [(1, 14), (2, 22), (3, 20), (4, 32), (5, 18)]),
    ("Pengeluaran", pengeluaran_rows, [(1, 14), (2, 20), (3, 32), (4, 16), (5, 18)]),
    ("Anggaran",    anggaran_rows,    [(1, 22), (2, 18), (3, 18), (4, 18), (5, 14)]),
    ("Kategori",    kategori_rows,    [(1, 26), (2, 26)]),
]

# ------------------------- BUILD XML ------------------------- #

def col_letter(idx):  # 1-based
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def iso_to_serial(iso):
    """Konversi 'YYYY-MM-DD' ke serial date Excel."""
    from datetime import date
    y, m, d = map(int, iso.split("-"))
    # Excel epoch 1900-01-01 = 1 (with the 1900 leap year bug)
    base = date(1899, 12, 30)
    return (date(y, m, d) - base).days


def build_sheet_xml(rows):
    out = []
    out.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    out.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')

    # Column widths — akan di-set per sheet, ditambahkan di luar
    out.append("__COLS__")
    out.append("<sheetData>")
    for r_idx, row in enumerate(rows, start=1):
        if not row:
            continue
        out.append('<row r="{0}">'.format(r_idx))
        for c_idx, cell in enumerate(row, start=1):
            val, style, is_formula = cell
            ref = "{0}{1}".format(col_letter(c_idx), r_idx)
            if val == "" and not is_formula:
                # baris kosong tetap perlu style? skip
                continue
            if is_formula:
                out.append('<c r="{0}" s="{1}"><f>{2}</f></c>'.format(
                    ref, style, escape(str(val))))
            else:
                if style == 4 and isinstance(val, str):
                    # date
                    serial = iso_to_serial(val)
                    out.append('<c r="{0}" s="{1}"><v>{2}</v></c>'.format(ref, style, serial))
                elif isinstance(val, (int, float)):
                    out.append('<c r="{0}" s="{1}"><v>{2}</v></c>'.format(ref, style, val))
                else:
                    out.append('<c r="{0}" s="{1}" t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>'.format(
                        ref, style, escape(str(val))))
        out.append("</row>")
    out.append("</sheetData>")
    out.append("__MERGE__")
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


# ---------- styles.xml ----------
STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3">
    <numFmt numFmtId="164" formatCode="&quot;Rp&quot;\\ #,##0;[Red]&quot;Rp&quot;\\ \\-#,##0"/>
    <numFmt numFmtId="165" formatCode="dd\\ mmm\\ yyyy"/>
    <numFmt numFmtId="166" formatCode="0.0%"/>
  </numFmts>
  <fonts count="6">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="16"/><color rgb="FF1F4E79"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
    <font><sz val="11"/><color rgb="FFC00000"/><name val="Calibri"/></font>
    <font><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/></font>
  </fonts>
  <fills count="5">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
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
  <cellXfs count="11">
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"/>                                                      <!-- 0 default -->
    <xf numFmtId="0"   fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf> <!-- 1 header -->
    <xf numFmtId="0"   fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>                                           <!-- 2 title -->
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>                <!-- 3 currency -->
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf> <!-- 4 date -->
    <xf numFmtId="0"   fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>          <!-- 5 sub-header -->
    <xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf> <!-- 6 percent -->
    <xf numFmtId="164" fontId="3" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>  <!-- 7 total currency -->
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf>          <!-- 8 text wrap center -->
    <xf numFmtId="164" fontId="4" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>  <!-- 9 currency merah -->
    <xf numFmtId="164" fontId="5" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>  <!-- 10 currency hijau -->
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
'''


# ---------- content types & rels ----------
def build_content_types(n_sheets):
    overrides = []
    for i in range(1, n_sheets + 1):
        overrides.append(
            '<Override PartName="/xl/worksheets/sheet{0}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(i))
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
    for i, (name, _, _) in enumerate(sheets, start=1):
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


# ------------------------- WRITE XLSX ------------------------- #

def main():
    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", build_content_types(len(SHEETS)))
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", build_workbook_xml(SHEETS))
        z.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels(len(SHEETS)))
        z.writestr("xl/styles.xml", STYLES_XML)
        for i, (name, rows, widths) in enumerate(SHEETS, start=1):
            xml = build_sheet_xml(rows)
            xml = xml.replace("__COLS__", build_cols_block(widths))
            # Merge cell A1 across all columns defined
            ncols = max(w[0] for w in widths) if widths else 5
            merge = '<mergeCells count="1"><mergeCell ref="A1:{0}1"/></mergeCells>'.format(col_letter(ncols))
            xml = xml.replace("__MERGE__", merge)
            z.writestr("xl/worksheets/sheet{0}.xml".format(i), xml)
    print("OK -> " + OUTPUT)


if __name__ == "__main__":
    main()
