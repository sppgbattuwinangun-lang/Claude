"""
Enhance 'Akumulasi Total.xlsx' -> FOKUS HANYA DI SHEET DASHBOARD:
  1. Baris % Surplus Total (row 47) dengan IFERROR anti-error
  2. 2 Grafik estetik profesional di sheet Dashboard (kolom F-L):
     - Donut chart: Penyerapan Anggaran (Terpakai vs Sisa)
     - Column chart: Surplus vs Defisit
  3. Helper data tersembunyi di col N-O (tidak mengganggu tampilan)

Output: Akumulasi_Total_Enhanced.xlsx
"""

import zipfile
import re
import os

SRC = "Akumulasi Total.xlsx"
DST = "Akumulasi_Total_Enhanced.xlsx"


# =========================================================================
# CHART 1: DONUT - "PENYERAPAN ANGGARAN"
# Cat: Terpakai, Sisa  (hardcoded di strCache)
# Val: Dashboard!N5:N6  (helper data dengan IFERROR)
# =========================================================================
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
      <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
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
              <a:solidFill><a:srgbClr val="E67300"/></a:solidFill>
              <a:ln w="38100"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
            </c:spPr>
          </c:dPt>
          <c:dPt>
            <c:idx val="1"/><c:bubble3D val="0"/>
            <c:spPr>
              <a:solidFill><a:srgbClr val="548235"/></a:solidFill>
              <a:ln w="38100"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
            </c:spPr>
          </c:dPt>
          <c:dLbls>
            <c:spPr>
              <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              <a:ln w="12700"><a:solidFill><a:srgbClr val="A6A6A6"/></a:solidFill></a:ln>
            </c:spPr>
            <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"><a:solidFill><a:srgbClr val="262626"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>
            <c:dLblPos val="outEnd"/>
            <c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/>
            <c:separator>&#10;</c:separator>
          </c:dLbls>
          <c:cat><c:strRef>
            <c:f>Dashboard!$N$5:$N$6</c:f>
            <c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Terpakai</c:v></c:pt><c:pt idx="1"><c:v>Sisa</c:v></c:pt></c:strCache>
          </c:strRef></c:cat>
          <c:val><c:numRef>
            <c:f>Dashboard!$O$5:$O$6</c:f>
            <c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt></c:numCache>
          </c:numRef></c:val>
        </c:ser>
        <c:firstSliceAng val="0"/>
        <c:holeSize val="55"/>
      </c:doughnutChart>
      <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
      <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
      <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1"><a:latin typeface="Calibri"/></a:defRPr></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln>
  </c:spPr>
</c:chartSpace>
'''


# =========================================================================
# CHART 2: COLUMN - "AKUMULASI SURPLUS vs DEFISIT"
# Cat: Surplus, Defisit  (hardcoded di strCache)
# Val: Dashboard!O8:O9   (helper data dengan IFERROR, defisit sudah di-ABS)
# =========================================================================
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
      <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
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
            <a:solidFill><a:srgbClr val="548235"/></a:solidFill>
            <a:ln><a:noFill/></a:ln>
          </c:spPr>
          <c:invertIfNegative val="0"/>
          <c:dPt>
            <c:idx val="0"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>
            <c:spPr>
              <a:solidFill><a:srgbClr val="548235"/></a:solidFill>
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
            <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1"><a:solidFill><a:srgbClr val="262626"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>
            <c:dLblPos val="outEnd"/>
            <c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>
            <c:numFmt formatCode="&quot;Rp&quot;#,##0" sourceLinked="0"/>
          </c:dLbls>
          <c:cat><c:strRef>
            <c:f>Dashboard!$N$8:$N$9</c:f>
            <c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Surplus</c:v></c:pt><c:pt idx="1"><c:v>Defisit</c:v></c:pt></c:strCache>
          </c:strRef></c:cat>
          <c:val><c:numRef>
            <c:f>Dashboard!$O$8:$O$9</c:f>
            <c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>0</c:v></c:pt></c:numCache>
          </c:numRef></c:val>
        </c:ser>
        <c:gapWidth val="100"/>
        <c:axId val="901"/><c:axId val="902"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="901"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:majorTickMark val="none"/><c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:spPr><a:noFill/><a:ln w="9525"><a:solidFill><a:srgbClr val="A6A6A6"/></a:solidFill></a:ln></c:spPr>
        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"><a:solidFill><a:srgbClr val="262626"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>
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
    <a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln>
  </c:spPr>
</c:chartSpace>
'''


# =========================================================================
# DRAWING: 2 chart side-by-side di KANAN data utama (col F-L)
#   Chart 1 (Donut Penyerapan)         = col F-L (5..11), row 4-22
#   Chart 2 (Column Surplus vs Defisit)= col F-L (5..11), row 24-46
# Col index 0=A, 5=F, 11=L
# =========================================================================
DRAWING1_NEW = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>22</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
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
    <xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>23</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>45</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
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


DRAWING1_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart6.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart7.xml"/></Relationships>
'''


# =========================================================================
# MAIN
# =========================================================================
def main():
    if os.path.exists(DST):
        os.remove(DST)

    with zipfile.ZipFile(SRC, 'r') as zin:
        ss = zin.read('xl/sharedStrings.xml').decode()
        sheet1 = zin.read('xl/worksheets/sheet1.xml').decode()
        ct = zin.read('[Content_Types].xml').decode()

    # ----- 1. Tambah shared strings -----
    m = re.search(r'<sst[^>]*count="(\d+)"[^>]*uniqueCount="(\d+)"', ss)
    old_count = int(m.group(1))
    old_unique = int(m.group(2))

    idx_surplus_pct = old_unique
    idx_terpakai    = old_unique + 1
    idx_sisa        = old_unique + 2
    idx_surplus     = old_unique + 3
    idx_defisit     = old_unique + 4
    num_new = 5

    new_strings_xml = (
        '<si><t>% Surplus Total (dari Total Pagu)</t></si>'
        '<si><t>Terpakai</t></si>'
        '<si><t>Sisa</t></si>'
        '<si><t>Surplus</t></si>'
        '<si><t>Defisit</t></si>'
    )
    ss_new = ss.replace(
        f'count="{old_count}" uniqueCount="{old_unique}"',
        f'count="{old_count + num_new}" uniqueCount="{old_unique + num_new}"'
    )
    ss_new = ss_new.replace('</sst>', new_strings_xml + '</sst>')

    # ----- 2. Update sheet1 (Dashboard) -----
    sheet1_new = sheet1

    # A. Row 47: % Surplus Total (pakai IFERROR)
    # s="15" adalah format percent yang dipakai B32 (% Penyerapan)
    row_47 = (
        '<row r="47" ht="21" customHeight="1">'
        f'<c r="A47" s="3" t="s"><v>{idx_surplus_pct}</v></c>'
        '<c r="B47" s="15">'
        "<f>IFERROR(B46/SUM('Rekap Harian'!$K$6:$K$36),0)</f>"
        '</c>'
        '</row>'
    )
    old_row_47 = '<row r="47" ht="15.75" customHeight="1"/>'
    if old_row_47 in sheet1_new:
        sheet1_new = sheet1_new.replace(old_row_47, row_47)
    else:
        sheet1_new = re.sub(r'<row r="47"[^>]*/>', row_47, sheet1_new, count=1)

    # B. Helper data untuk chart di col N-O (area tersembunyi, tidak terlihat di layout utama)
    # Row 5: Terpakai | =IFERROR(B30,0)            [B30 = Total Penggunaan Dana]
    # Row 6: Sisa     | =IFERROR(MAX(0,B31),0)      [B31 = Total Sisa Dana]
    # Row 8: Surplus  | =IFERROR(B44,0)             [B44 = Total Akumulasi Surplus]
    # Row 9: Defisit  | =IFERROR(ABS(B45),0)        [B45 = Total Akumulasi Defisit (negatif)]
    # Ditambah ke row yang sudah ada (append cell ke row existing, bukan buat row baru)

    def inject_cells_into_row(xml, rnum, extra_cells_xml):
        """Tambah cells ke row existing (self-closing atau yang sudah punya isi)."""
        # Cari opening tag row dulu
        open_re = re.compile(rf'<row r="{rnum}"([^>]*?)(/>|>)')
        m = open_re.search(xml)
        if m is None:
            return xml
        attrs = m.group(1)
        close_char = m.group(2)
        start_idx = m.start()
        end_open = m.end()

        if close_char == '/>':
            # Self-closing: <row r="X" .../>  ->  <row r="X" ...>CELLS</row>
            replacement = f'<row r="{rnum}"{attrs}>{extra_cells_xml}</row>'
            return xml[:start_idx] + replacement + xml[end_open:]
        else:
            # Not self-closing: find matching </row>
            close_re = re.compile(r'</row>')
            mc = close_re.search(xml, end_open)
            if mc is None:
                return xml
            # Insert extra cells before </row>
            return xml[:mc.start()] + extra_cells_xml + xml[mc.start():]

    # Cells helper (text label + formula)
    helper_cells = {
        5: (
            f'<c r="N5" s="3" t="s"><v>{idx_terpakai}</v></c>'
            '<c r="O5" s="14"><f>IFERROR(B30,0)</f></c>'
        ),
        6: (
            f'<c r="N6" s="3" t="s"><v>{idx_sisa}</v></c>'
            '<c r="O6" s="14"><f>IFERROR(MAX(0,B31),0)</f></c>'
        ),
        8: (
            f'<c r="N8" s="3" t="s"><v>{idx_surplus}</v></c>'
            '<c r="O8" s="14"><f>IFERROR(B44,0)</f></c>'
        ),
        9: (
            f'<c r="N9" s="3" t="s"><v>{idx_defisit}</v></c>'
            '<c r="O9" s="14"><f>IFERROR(ABS(B45),0)</f></c>'
        ),
    }

    for rnum, cells in helper_cells.items():
        sheet1_new = inject_cells_into_row(sheet1_new, rnum, cells)

    # ----- 3. Update Content Types (register 2 chart baru) -----
    chart_overrides = (
        '<Override ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml" PartName="/xl/charts/chart6.xml"/>'
        '<Override ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml" PartName="/xl/charts/chart7.xml"/>'
    )
    ct_new = ct.replace('</Types>', chart_overrides + '</Types>')

    # ----- 4. Write output -----
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

        zout.writestr('xl/charts/chart6.xml', CHART_PENYERAPAN)
        zout.writestr('xl/charts/chart7.xml', CHART_SURPLUS_DEFISIT)
        zout.writestr('xl/drawings/_rels/drawing1.xml.rels', DRAWING1_RELS)

    print(f"OK -> {DST}")


if __name__ == "__main__":
    main()
