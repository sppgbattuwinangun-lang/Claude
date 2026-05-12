#!/usr/bin/env python3
"""
Generator Berita Acara Penerimaan Insentif Sekolah
Menghasilkan 29 file .docx satu per sekolah berdasarkan:
  - Rekap Insentif Sekolah 27 April - 8 Mei 2026.xlsm (data)
  - Template BA Insentif.docx (template)

Tanpa library eksternal (pure stdlib) - membaca/menulis docx/xlsx sebagai ZIP.
"""
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xml_escape

# ============ PATH ============
BASE = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE, "Rekap Insentif Sekolah 27 April - 8 Mei 2026.xlsm")
TEMPLATE_PATH = os.path.join(BASE, "Template BA Insentif.docx")
OUTPUT_DIR = os.path.join(BASE, "Output_BA_Insentif")

# ============ KONFIGURASI ============
NAMA_PIHAK_PERTAMA_JABATAN = "PIC Sekolah"
NOMOR_SURAT_TEMPLATE = "{idx:03d}/BA/SPPG.BW/V/2026"

# ============ PARSE EXCEL ============
def parse_excel(xlsm_path):
    """Baca data dari file xlsm. Return list of dict."""
    NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    W = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

    with zipfile.ZipFile(xlsm_path, 'r') as z:
        with z.open('xl/sharedStrings.xml') as f:
            ss_xml = f.read()
        with z.open('xl/worksheets/sheet1.xml') as f:
            sheet_xml = f.read()

    # Parse shared strings
    root_ss = ET.fromstring(ss_xml)
    strings = []
    for si in root_ss.findall(f'{W}si'):
        texts = [t.text or '' for t in si.iter(f'{W}t')]
        strings.append(''.join(texts))

    # Parse sheet
    root = ET.fromstring(sheet_xml)

    def col_to_num(col):
        n = 0
        for c in col:
            n = n * 26 + (ord(c) - ord('A') + 1)
        return n

    rows_data = {}
    for row in root.iter(f'{W}row'):
        r = int(row.get('r'))
        cells = {}
        for c in row.findall(f'{W}c'):
            ref = c.get('r')
            col_letters = re.match(r'([A-Z]+)', ref).group(1)
            col_num = col_to_num(col_letters)
            t = c.get('t')
            v = c.find(f'{W}v', NS)
            if v is None:
                val = ''
            else:
                val = v.text
                if t == 's':
                    val = strings[int(val)]
            cells[col_num] = val
        rows_data[r] = cells

    # Header at row 3, data rows 4-32
    # Cols: A=1 No, B=2 Nama Sekolah, C=3 Jumlah Siswa, D=4 Keterangan,
    #       E=5 Insentif/Hari, F=6 Total Hari, G=7 Total Insentif,
    #       H=8 No Rek, I=9 Bank, J=10 Nama Penerima, K=11 Keterangan Tambahan
    schools = []
    for r in sorted(rows_data.keys()):
        row = rows_data[r]
        no = row.get(1, '')
        if not str(no).strip().isdigit():
            continue
        try:
            total_hari = int(float(row.get(6, 0) or 0))
        except Exception:
            total_hari = 0
        try:
            total_insentif = int(float(row.get(7, 0) or 0))
        except Exception:
            total_insentif = 0
        schools.append({
            'no': int(no),
            'nama_sekolah': str(row.get(2, '')).strip(),
            'nama_pic': str(row.get(10, '')).strip(),
            'total_hari': total_hari,
            'total_insentif': total_insentif,
            'bank': str(row.get(9, '')).strip(),
            'no_rek': str(row.get(8, '')).strip(),
        })
    return schools


# ============ FORMAT HELPERS ============
def fmt_rupiah(n):
    """Format angka ke rupiah Indonesia: 270000 -> 'Rp 270.000'"""
    s = f"{int(n):,}".replace(",", ".")
    return f"Rp {s}"

def sanitize_filename(s):
    """Hapus karakter terlarang untuk nama file."""
    s = s.replace('&', 'dan').replace('/', '-').replace('\\', '-')
    s = re.sub(r'[<>:"|?*]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


# ============ RENDER TEMPLATE ============
def render_template(template_path, out_path, data, nomor_surat):
    """Render satu BA untuk satu sekolah."""
    nama_sekolah = data['nama_sekolah']
    nama_pic = data['nama_pic']
    total_hari = data['total_hari']
    total_insentif = data['total_insentif']
    rekap_text = f"{total_hari} Hari | {fmt_rupiah(total_insentif)}"

    # Copy template jadi file output
    shutil.copy(template_path, out_path)

    # Modify document.xml inside zip
    with zipfile.ZipFile(template_path, 'r') as zin:
        doc_xml = zin.read('word/document.xml').decode('utf-8')

    xml = doc_xml

    # === Replacement 1: Nama Sekolah di tabel Ringkasan ===
    # Template: '<w:t>MTS &amp; MA UMMUL QURO</w:t>'
    xml = xml.replace(
        '<w:t>MTS &amp; MA UMMUL QURO</w:t>',
        f'<w:t>{xml_escape(nama_sekolah)}</w:t>',
        1
    )

    # === Replacement 2: Nomor Surat ===
    # Template: '001/BA/SPPG.BW/V/2026' (split in several w:t: '001/BA/SPPG.BW/', 'V', '/', '2026')
    # We replace the first segment and remove other segments by emptying them
    xml = xml.replace('<w:t>001/BA/SPPG.BW/</w:t>', f'<w:t>{nomor_surat}</w:t>', 1)
    # Remove V, /, 2026 yang terpisah (3 consecutive w:t after)
    # Hati-hati: ada 'V' tunggal dan '/' tunggal - cari berurutan
    # Cari 'V' dengan konteks unik sebelum
    xml = re.sub(
        r'(<w:t>%s</w:t></w:r>)(.*?)<w:t>V</w:t>' % re.escape(nomor_surat),
        r'\1\2<w:t></w:t>',
        xml,
        count=1,
        flags=re.DOTALL,
    )
    xml = re.sub(
        r'(<w:t>%s</w:t></w:r>.*?<w:t></w:t>.*?)<w:t>/</w:t>' % re.escape(nomor_surat),
        r'\1<w:t></w:t>',
        xml,
        count=1,
        flags=re.DOTALL,
    )
    xml = re.sub(
        r'(<w:t>%s</w:t></w:r>.*?<w:t></w:t>.*?<w:t></w:t>.*?)<w:t>2026</w:t>' % re.escape(nomor_surat),
        r'\1<w:t></w:t>',
        xml,
        count=1,
        flags=re.DOTALL,
    )

    # === Replacement 3: PIHAK KEDUA - Nama & Jabatan ===
    # Template ada 2x "<w:t>Nama :</w:t>" - pertama Pihak Pertama (Alfiansah), kedua Pihak Kedua (kosong)
    # Replace 2nd occurrence: ganti "Nama :" -> "Nama : {nama_pic}"
    parts = xml.split('<w:t>Nama :</w:t>', 2)
    if len(parts) == 3:
        xml = parts[0] + '<w:t>Nama :</w:t>' + parts[1] + \
              f'<w:t xml:space="preserve">Nama : {xml_escape(nama_pic)}</w:t>' + parts[2]

    # Sama untuk Jabatan
    parts = xml.split('<w:t>Jabatan :</w:t>', 2)
    if len(parts) == 3:
        xml = parts[0] + '<w:t>Jabatan :</w:t>' + parts[1] + \
              f'<w:t xml:space="preserve">Jabatan : {xml_escape(NAMA_PIHAK_PERTAMA_JABATAN)}</w:t>' + parts[2]

    # === Replacement 4: Nama PIC Sekolah (cell kosong di tabel Ringkasan) ===
    # paraId="31BDDB32" -> ini paragraph kosong dalam cell "Nama PIC Sekolah"
    # Inject <w:r>...<w:t>{nama_pic}</w:t></w:r> sebelum </w:p>
    run_injection = (
        '<w:r w:rsidRPr="001C6D40">'
        '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
        f'<w:t xml:space="preserve">{xml_escape(nama_pic)}</w:t></w:r>'
    )
    # Cari pattern paragraph kosong dengan paraId 31BDDB32
    pattern_pic_cell = re.compile(
        r'(<w:p w14:paraId="31BDDB32"[^>]*>'
        r'<w:pPr><w:rPr><w:rFonts[^/]*/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:pPr>)'
        r'(</w:p>)'
    )
    xml, n = pattern_pic_cell.subn(r'\1' + run_injection + r'\2', xml, count=1)
    # Fallback: kalau pattern tidak match (mungkin beda rsidR), pakai pattern lebih longgar
    if n == 0:
        pattern_pic_cell2 = re.compile(
            r'(<w:p w14:paraId="31BDDB32"[^>]*>.*?<w:pPr>.*?</w:pPr>)(</w:p>)',
            re.DOTALL
        )
        xml, n = pattern_pic_cell2.subn(r'\1' + run_injection + r'\2', xml, count=1)

    # === Replacement 5: Rekap Penerimaan ===
    # Cell berisi: <w:t>Jmlh Hari</w:t> ... <w:t xml:space="preserve"> | </w:t> ... <w:t>Total Insentif</w:t>
    xml = xml.replace(
        '<w:t>Jmlh Hari</w:t>',
        f'<w:t>{total_hari} Hari</w:t>',
        1
    )
    xml = xml.replace(
        '<w:t>Total Insentif</w:t>',
        f'<w:t>{xml_escape(fmt_rupiah(total_insentif))}</w:t>',
        1
    )

    # === Replacement 6: Nama Sekolah di blok TTD ===
    # Template: <w:t>(Nama Sekolah)</w:t>
    xml = xml.replace(
        '<w:t>(Nama Sekolah)</w:t>',
        f'<w:t>{xml_escape(nama_sekolah)}</w:t>',
        1
    )

    # === Replacement 7: Nama di TTD [……………………………….] ===
    # Template: <w:t>[</w:t> ... <w:t>……………………………….</w:t> ... <w:t>]</w:t>
    # Ganti middle dengan nama uppercase, dan bracket jadi empty
    nama_upper = nama_pic.upper()
    xml = xml.replace(
        '<w:t>……………………………….</w:t>',
        f'<w:t>{xml_escape(nama_upper)}</w:t>',
        1
    )
    # Hapus bracket dengan menjadikannya empty w:t (agar struktur XML aman)
    xml = xml.replace('<w:t>[</w:t>', '<w:t></w:t>', 1)
    xml = xml.replace('<w:t>]</w:t>', '<w:t></w:t>', 1)

    # === Tulis ke output zip ===
    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data_bytes = zin.read(item)
                if item == 'word/document.xml':
                    data_bytes = xml.encode('utf-8')
                zout.writestr(item, data_bytes)


# ============ MAIN ============
def main():
    print("Membaca data Excel...")
    schools = parse_excel(EXCEL_PATH)
    print(f"Ditemukan {len(schools)} sekolah.\n")

    # Print summary table
    print(f"{'No':<4} {'Sekolah':<30} {'PIC':<30} {'Hari':<6} {'Insentif':<15}")
    print("-" * 90)
    for s in schools:
        print(f"{s['no']:<4} {s['nama_sekolah']:<30} {s['nama_pic']:<30} {s['total_hari']:<6} {fmt_rupiah(s['total_insentif']):<15}")

    # Generate output
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    for f in os.listdir(OUTPUT_DIR):
        if f.endswith('.docx'):
            os.remove(os.path.join(OUTPUT_DIR, f))

    print(f"\nGenerating {len(schools)} Berita Acara ke {OUTPUT_DIR}...")
    for s in schools:
        fname = f"BA Insentif - {sanitize_filename(s['nama_sekolah'])}.docx"
        out_path = os.path.join(OUTPUT_DIR, fname)
        nomor_surat = NOMOR_SURAT_TEMPLATE.format(idx=s['no'])
        render_template(TEMPLATE_PATH, out_path, s, nomor_surat)
        print(f"  [OK] {fname}")

    total = sum(s['total_insentif'] for s in schools)
    print(f"\nTotal keseluruhan insentif: {fmt_rupiah(total)}")
    print(f"Total file dibuat: {len(schools)}")


if __name__ == '__main__':
    main()
