"""
Enhance 'Akumulasi Total.xlsx' -> menambah di Dashboard:
  1. Baris metric '% Surplus Total' (setelah Net Surplus/Defisit)
     Pakai IFERROR supaya tidak muncul #VALUE! atau #DIV/0!
  2. Helper data (rows 58-64) dengan label pendek: Terpakai/Sisa/Surplus/Defisit
  3. 2 Chart estetik profesional di sheet Dashboard:
     - Donut chart: Penyerapan Anggaran (Terpakai vs Sisa) - clean label
     - Column chart: Surplus vs Defisit dengan warna hijau/merah

Output: Akumulasi_Total_Enhanced.xlsx
"""

import zipfile
import shutil
import re
import os

SRC = "Akumulasi Total.xlsx"
DST = "Akumulasi_Total_Enhanced.xlsx"


# =========================================================================
# CHART XML (2 chart profesional untuk Dashboard)
# =========================================================================
#
# HELPER DATA di Dashboard rows 58-64 (hidden area):
#   A58: "Terpakai"        B58: =IFERROR(B30,0)
#   A59: "Sisa"            B59: =IFERROR(MAX(0,B31),0)    -- pastikan tidak negatif
#   A62: "Surplus"         B62: =IFERROR(B44,0)
#   A63: "Defisit"         B63: =IFERROR(ABS(B45),0)      -- ABS supaya positif, chart pakai warna merah
# =========================================================================

# Chart 1: Donut - Penyerapan Anggaran
# Label pendek + clean + title dengan subtitle
CHART_PENYERAPAN = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title>
      <c:tx><c:rich>
        <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" wrap="square" anchor="ctr" anchorCtr="1"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"><a:defRPr/></a:pPr>
          <a:r><a:rPr lang="id-ID" b="1" sz="1400"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>PENYERAPAN ANGGARAN</a:t></a:r>
        </a:p>
      </c:rich></c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:doughnutChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:v>Penyerapan</c:v></c:tx>
          <c:dPt>
            <c:idx val="0"/><c:bubble3D val="0"/>
            <c:spPr>
              <a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill>
              <a:ln w="25400"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
            </c:spPr>
          </c:dPt>
          <c:dPt>
            <c:idx val="1"/><c:bubble3D val="0"/>
            <c:spPr>
              <a:solidFill><a:srgbClr val="70AD47"/></a:solidFill>
              <a:ln w="25400"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
            </c:spPr>
          </c:dPt>
          <c:dLbls>
            <c:spPr>
              <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              <a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln>
            </c:spPr>
            <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1"><a:solidFill><a:srgbClr val="262626"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>
            <c:dLblPos val="outEnd"/>
            <c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/>
            <c:separator>: </c:separator>
          </c:dLbls>
          <c:cat><c:strRef>
            <c:f>Dashboard!$A$58:$A$59</c:f>
            <c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Terpakai</c:v></c:pt><c:pt idx="1"><c:v>Sisa</c:v></c:pt></c:strCache>
          </c:strRef></c:cat>
          <c:val><c:numRef>
            <c:f>Dashboard!$B$58:$B$59</c:f>
            <c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>0</c:v></c:pt></c:numCache>
          </c:numRef></c:val>
        </c:ser>
        <c:firstSliceAng val="0"/>
        <c:holeSize val="60"/>
      </c:doughnutChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
      <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1"/></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln>
  </c:spPr>
</c:chartSpace>
'''


# Chart 2: Column chart - Surplus vs Defisit
# Value positif keduanya (defisit pakai ABS), warnain beda via dPt
CHART_SURPLUS_DEFISIT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title>
      <c:tx><c:rich>
        <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" wrap="square" anchor="ctr" anchorCtr="1"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"><a:defRPr/></a:pPr>
          <a:r><a:rPr lang="id-ID" b="1" sz="1400"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>AKUMULASI SURPLUS vs DEFISIT</a:t></a:r>
        </a:p>
      </c:rich></c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:v>Nilai</c:v></c:tx>
          <c:spPr>
            <a:solidFill><a:srgbClr val="70AD47"/></a:solidFill>
            <a:ln><a:noFill/></a:ln>
          </c:spPr>
          <c:invertIfNegative val="0"/>
          <c:dPt>
            <c:idx val="0"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>
            <c:spPr>
              <a:solidFill><a:srgbClr val="70AD47"/></a:solidFill>
              <a:ln><a:noFill/></a:ln>
            </c:spPr>
          </c:dPt>
          <c:dPt>
            <c:idx val="1"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>
            <c:spPr>
              <a:solidFill><a:srgbClr val="C00000"/></a:solidFill>
              <a:ln><a:noFill/></a:ln>
            </c:spPr>
          </c:dPt>
          <c:dLbls>
            <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1"><a:solidFill><a:srgbClr val="262626"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>
            <c:dLblPos val="outEnd"/>
            <c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>
            <c:numFmt formatCode="&quot;Rp&quot;#,##0" sourceLinked="0"/>
          </c:dLbls>
          <c:cat><c:strRef>
            <c:f>Dashboard!$A$62:$A$63</c:f>
            <c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Surplus</c:v></c:pt><c:pt idx="1"><c:v>Defisit</c:v></c:pt></c:strCache>
          </c:strRef></c:cat>
          <c:val><c:numRef>
            <c:f>Dashboard!$B$62:$B$63</c:f>
            <c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>0</c:v></c:pt></c:numCache>
          </c:numRef></c:val>
        </c:ser>
        <c:gapWidth val="80"/>
        <c:axId val="901"/><c:axId val="902"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="901"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:majorTickMark val="none"/><c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:spPr><a:noFill/><a:ln><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>
        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"><a:solidFill><a:srgbClr val="262626"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>
        <c:crossAx val="902"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="902"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="1"/>
        <c:axPos val="l"/>
        <c:numFmt formatCode="&quot;Rp&quot;#,##0" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:tickLblPos val="none"/>
        <c:crossAx val="901"/>
      </c:valAx>
      <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="12700"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln>
  </c:spPr>
</c:chartSpace>
'''


# =========================================================================
# NEW drawing1.xml for Dashboard (dengan 2 chart)
# Posisi:
#   Chart A (Donut "Penyerapan Anggaran") = E5:J22 (kanan KPI cards)
#   Chart B (Bar "Surplus vs Defisit")    = E24:J42 (kanan section A-C)
# =========================================================================
DRAWING1_NEW = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>22</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="ChartPenyerapan"/>
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
    <xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>23</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>42</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="3" name="ChartSurplusDefisit"/>
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
</xdr:wsDr>
'''


# Drawing rels untuk Dashboard (menunjuk ke 2 chart baru)
# Nama file: chart6.xml & chart7.xml (biar nggak bentrok dengan 1-5 yang sudah ada di sheet Grafik)
DRAWING1_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart6.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart7.xml"/>
</Relationships>
'''


# =========================================================================
# MAIN
# =========================================================================

def main():
    if os.path.exists(DST):
        os.remove(DST)

    # Read shared strings to find/add new strings
    with zipfile.ZipFile(SRC, 'r') as zin:
        ss = zin.read('xl/sharedStrings.xml').decode()
        sheet1 = zin.read('xl/worksheets/sheet1.xml').decode()
        ct = zin.read('[Content_Types].xml').decode()

    # --- 1. Add new shared strings ---
    m = re.search(r'<sst[^>]*count="(\d+)"[^>]*uniqueCount="(\d+)"', ss)
    old_count = int(m.group(1))
    old_unique = int(m.group(2))

    # String indices (sequential)
    idx_surplus_pct = old_unique
    idx_header_viz  = old_unique + 1
    idx_terpakai    = old_unique + 2
    idx_sisa        = old_unique + 3
    idx_surplus     = old_unique + 4
    idx_defisit     = old_unique + 5
    num_new_strings = 6

    new_strings_xml = (
        '<si><t>% Surplus Total (dari Total Pagu)</t></si>'
        '<si><t>E. VISUALISASI</t></si>'
        '<si><t>Terpakai</t></si>'
        '<si><t>Sisa</t></si>'
        '<si><t>Surplus</t></si>'
        '<si><t>Defisit</t></si>'
    )

    ss_new = ss.replace(
        f'count="{old_count}" uniqueCount="{old_unique}"',
        f'count="{old_count + num_new_strings}" uniqueCount="{old_unique + num_new_strings}"'
    )
    ss_new = ss_new.replace('</sst>', new_strings_xml + '</sst>')

    # --- 2. Build new rows untuk Dashboard ---
    # Row 47: % Surplus Total (pakai IFERROR supaya tidak #VALUE!)
    # Formula: Net Surplus / Total Pagu
    #   Net Surplus = B46 (sudah ada: =B44+B45 = surplus + (-defisit))
    #   Total Pagu = SUM(Rekap Harian!K6:K36)
    #   IFERROR wrap: kalau total pagu 0 atau error -> 0
    row_47 = (
        '<row r="47" ht="22" customHeight="1">'
        f'<c r="A47" s="3" t="s"><v>{idx_surplus_pct}</v></c>'
        '<c r="B47" s="21">'
        '<f>IFERROR(B46/SUM(\'Rekap Harian\'!$K$6:$K$36),0)</f>'
        '</c>'
        '</row>'
    )

    # Helper data rows 58-63 (akan dipakai oleh chart)
    # Semua pakai IFERROR supaya nggak ada #VALUE!
    # Row 57: section header "E. VISUALISASI" (optional, agar user tahu)
    # Row 58: Terpakai | =IFERROR(B30,0)
    # Row 59: Sisa     | =IFERROR(MAX(0,B31),0)
    # (rows 60-61 spacer)
    # Row 62: Surplus  | =IFERROR(B44,0)
    # Row 63: Defisit  | =IFERROR(ABS(B45),0)
    helper_rows = (
        '<row r="57" ht="20" customHeight="1">'
        f'<c r="A57" s="17" t="s"><v>{idx_header_viz}</v></c>'
        '</row>'
        '<row r="58" ht="18" customHeight="1">'
        f'<c r="A58" s="10" t="s"><v>{idx_terpakai}</v></c>'
        '<c r="B58" s="24"><f>IFERROR(B30,0)</f></c>'
        '</row>'
        '<row r="59" ht="18" customHeight="1">'
        f'<c r="A59" s="10" t="s"><v>{idx_sisa}</v></c>'
        '<c r="B59" s="24"><f>IFERROR(MAX(0,B31),0)</f></c>'
        '</row>'
        '<row r="62" ht="18" customHeight="1">'
        f'<c r="A62" s="10" t="s"><v>{idx_surplus}</v></c>'
        '<c r="B62" s="24"><f>IFERROR(B44,0)</f></c>'
        '</row>'
        '<row r="63" ht="18" customHeight="1">'
        f'<c r="A63" s="10" t="s"><v>{idx_defisit}</v></c>'
        '<c r="B63" s="24"><f>IFERROR(ABS(B45),0)</f></c>'
        '</row>'
    )

    sheet1_new = sheet1

    # Replace empty row 47
    old_empty_47 = '<row r="47" ht="15.75" customHeight="1"/>'
    if old_empty_47 in sheet1_new:
        sheet1_new = sheet1_new.replace(old_empty_47, row_47)
    else:
        sheet1_new = re.sub(r'<row r="47"[^>]*/>', row_47, sheet1_new, count=1)

    # Replace empty rows 57-63 (mereka masing-masing empty self-closing)
    for rnum in (57, 58, 59, 62, 63):
        pattern = f'<row r="{rnum}" ht="15.75" customHeight="1"/>'
        sheet1_new = sheet1_new.replace(pattern, '')

    # Insert helper rows before </sheetData>
    # Tapi harus sebelum row 57 aslinya agar urutan row terjaga. Cari posisi
    # row 56 dulu, lalu inject setelahnya
    m_row56 = re.search(r'<row r="56"[^>]*/?>(?:</row>)?', sheet1_new)
    if m_row56:
        insert_pos = m_row56.end()
        sheet1_new = sheet1_new[:insert_pos] + helper_rows + sheet1_new[insert_pos:]
    else:
        # Fallback: insert before </sheetData>
        sheet1_new = sheet1_new.replace('</sheetData>', helper_rows + '</sheetData>')

    # --- 3. Update [Content_Types].xml ---
    chart_overrides = (
        '<Override ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml" PartName="/xl/charts/chart6.xml"/>'
        '<Override ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml" PartName="/xl/charts/chart7.xml"/>'
    )
    ct_new = ct.replace('</Types>', chart_overrides + '</Types>')

    # --- 4. Write new file ---
    with zipfile.ZipFile(SRC, 'r') as zin, zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'xl/sharedStrings.xml':
                data = ss_new.encode('utf-8')
            elif item.filename == 'xl/worksheets/sheet1.xml':
                data = sheet1_new.encode('utf-8')
            elif item.filename == '[Content_Types].xml':
                data = ct_new.encode('utf-8')
            elif item.filename == 'xl/drawings/drawing1.xml':
                data = DRAWING1_NEW.encode('utf-8')
            zout.writestr(item, data)

        # Add new files
        zout.writestr('xl/charts/chart6.xml', CHART_PENYERAPAN)
        zout.writestr('xl/charts/chart7.xml', CHART_SURPLUS_DEFISIT)
        zout.writestr('xl/drawings/_rels/drawing1.xml.rels', DRAWING1_RELS)

    print(f"OK -> {DST}")


if __name__ == "__main__":
    main()
