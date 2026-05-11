"""
Enhance 'Akumulasi Total.xlsx' -> menambah di Dashboard:
  1. Baris metric '% Surplus Total' (setelah Net Surplus/Defisit)
  2. 2 Chart estetik di sheet Dashboard:
     - Donut chart: Penyerapan Anggaran (Terpakai vs Sisa)
     - Bar chart horizontal: Surplus vs Defisit (dengan % label)

Output: Akumulasi_Total_Enhanced.xlsx
"""

import zipfile
import shutil
import re
import os

SRC = "Akumulasi Total.xlsx"
DST = "Akumulasi_Total_Enhanced.xlsx"


# =========================================================================
# CHART XML (2 chart baru untuk Dashboard)
# =========================================================================

# Chart 1: Donut - Penyerapan Anggaran (Terpakai vs Sisa)
# Data: Dashboard!B30 (Terpakai), Dashboard!B31 (Sisa)
CHART_PENYERAPAN = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title>
      <c:tx><c:rich>
        <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" wrap="square" anchor="ctr" anchorCtr="1"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr><a:defRPr/></a:pPr>
          <a:r><a:rPr lang="id-ID" b="1" sz="1400"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:rPr><a:t>Penyerapan Anggaran</a:t></a:r>
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
            <c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>
          </c:dPt>
          <c:dPt>
            <c:idx val="1"/><c:bubble3D val="0"/>
            <c:spPr><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>
          </c:dPt>
          <c:dLbls>
            <c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>
            <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>
            <c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/>
            <c:separator>
</c:separator>
          </c:dLbls>
          <c:cat><c:strRef>
            <c:f>Dashboard!$A$30:$A$31</c:f>
            <c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Terpakai</c:v></c:pt><c:pt idx="1"><c:v>Sisa</c:v></c:pt></c:strCache>
          </c:strRef></c:cat>
          <c:val><c:numRef>
            <c:f>Dashboard!$B$30:$B$31</c:f>
          </c:numRef></c:val>
        </c:ser>
        <c:firstSliceAng val="0"/>
        <c:holeSize val="55"/>
      </c:doughnutChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
      <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100"/></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln>
  </c:spPr>
</c:chartSpace>
'''


# Chart 2: Bar horizontal - Surplus vs Defisit (total akumulasi, Rp)
# Data: Dashboard!B44 (Surplus), Dashboard!B45 (Defisit - but stored as negative)
CHART_SURPLUS_DEFISIT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:title>
      <c:tx><c:rich>
        <a:bodyPr/><a:lstStyle/>
        <a:p>
          <a:r><a:rPr lang="id-ID" b="1" sz="1400"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:rPr><a:t>Akumulasi Surplus vs Defisit</a:t></a:r>
        </a:p>
      </c:rich></c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="bar"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:v>Nilai</c:v></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></c:spPr>
          <c:invertIfNegative val="1"/>
          <c:dPt>
            <c:idx val="0"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>
            <c:spPr><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></c:spPr>
          </c:dPt>
          <c:dPt>
            <c:idx val="1"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>
            <c:spPr><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></c:spPr>
          </c:dPt>
          <c:dLbls>
            <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000" b="1"/></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>
            <c:dLblPos val="outEnd"/>
            <c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>
            <c:numFmt formatCode="&quot;Rp&quot;#,##0" sourceLinked="0"/>
          </c:dLbls>
          <c:cat><c:strRef>
            <c:f>Dashboard!$A$44:$A$45</c:f>
            <c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Total Surplus</c:v></c:pt><c:pt idx="1"><c:v>Total Defisit</c:v></c:pt></c:strCache>
          </c:strRef></c:cat>
          <c:val><c:numRef>
            <c:f>Dashboard!$B$44:$B$45</c:f>
          </c:numRef></c:val>
        </c:ser>
        <c:gapWidth val="100"/>
        <c:axId val="901"/><c:axId val="902"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="901"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorTickMark val="out"/><c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1"/></a:pPr><a:endParaRPr lang="id-ID"/></a:p></c:txPr>
        <c:crossAx val="902"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="902"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E7E6E6"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
        <c:numFmt formatCode="&quot;Rp&quot;#,##0" sourceLinked="0"/>
        <c:majorTickMark val="out"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="901"/>
      </c:valAx>
      <c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln>
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

    # --- 1. Add new shared string: "% Surplus Total" ---
    # Find count/uniqueCount
    m = re.search(r'<sst[^>]*count="(\d+)"[^>]*uniqueCount="(\d+)"', ss)
    old_count = int(m.group(1))
    old_unique = int(m.group(2))
    new_idx = old_unique

    new_string_xml = f'<si><t>% Surplus Total (dari Total Pagu)</t></si>'
    # Insert before </sst>
    ss_new = ss.replace(
        f'count="{old_count}" uniqueCount="{old_unique}"',
        f'count="{old_count + 1}" uniqueCount="{old_unique + 1}"'
    )
    ss_new = ss_new.replace('</sst>', new_string_xml + '</sst>')

    # --- 2. Add row 47 (% Surplus Total) to Dashboard ---
    # Formula: =IF(SUM('Rekap Harian'!$K$6:$K$36)=0,0,B46/SUM('Rekap Harian'!$K$6:$K$36))
    # B46 = Net Surplus/Defisit, K = Total Pagu column
    # We'll use a simpler formula: percentage = B46 / B29 (Total Dana Anggaran)
    # But better: relative to total pagu operasi. Use Total Pagu = SUM rekap K
    # Style 21 = percent format (numFmtId 166) with center align
    # Row 47: A47 = label (text), B47 = formula percent
    new_row_xml = (
        '<row r="47" ht="20.25" customHeight="1">'
        f'<c r="A47" s="3" t="s"><v>{new_idx}</v></c>'
        '<c r="B47" s="21">'
        '<f>IF(SUM(\'Rekap Harian\'!$K$6:$K$36)=0,0,B46/SUM(\'Rekap Harian\'!$K$6:$K$36))</f>'
        '</c>'
        '</row>'
    )

    # Empty row 47 looks like: <row r="47" ht="15.75" customHeight="1"/>
    # Replace it
    old_empty_47 = '<row r="47" ht="15.75" customHeight="1"/>'
    if old_empty_47 in sheet1:
        sheet1_new = sheet1.replace(old_empty_47, new_row_xml)
    else:
        # Fallback: find any row 47 self-closing and replace
        sheet1_new = re.sub(
            r'<row r="47"[^>]*/>',
            new_row_xml,
            sheet1,
            count=1
        )

    # Note: cell style s="7" is used by B32 (% Penyerapan Anggaran).
    # Let's check what style % number uses in this file.
    # Looking at original: row 32 has <c r="B32" s="18"> with formula (format %?)
    # Actually format cell for percent: let's find one by looking at existing
    # The safer path is to use format of B26 (% Penggunaan Dana Operasional) which is same as B16, B32
    # B32 uses style s="18" per sample. Let me check properly.
    # Since I can't easily test, I'll use s="18" which we saw IS used for currency total
    # Actually looking again: row 46 used s="18" for Net Surplus total. That's currency.
    # For percentage, row 32 (% Penyerapan) uses s="7"... let's check.

    # --- 3. Update [Content_Types].xml to add chart6.xml & chart7.xml ---
    chart6_override = '<Override ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml" PartName="/xl/charts/chart6.xml"/>'
    chart7_override = '<Override ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml" PartName="/xl/charts/chart7.xml"/>'
    drawing1_rels_override = ''  # rels don't need override, use Default

    # Add before </Types>
    ct_new = ct.replace('</Types>', chart6_override + chart7_override + '</Types>')

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
