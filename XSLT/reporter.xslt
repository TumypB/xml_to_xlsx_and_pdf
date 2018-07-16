<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:ms="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="ms"
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt"
    xmlns:obj="urn:qrcode"
>
  <xsl:output method="xml" indent="yes"/>

  <xsl:variable name="styles" select="/book/styles/style"/>

  <xsl:variable name="abc">
    <xsl:text>ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:text>
  </xsl:variable>

  <xsl:template match="col">
    <xsl:variable name="v" select="substring-before(@width, 'pt')" />
    <!--http://webapp.docx4java.org/OnlineDemo/ecma376/SpreadsheetML/col.html-->
    <xsl:variable name="vv" select="$v div 2.83465 div 2" />
    <col min="{position()}" max="{position()}" width="{$vv}" customWidth="1" />
  </xsl:template>

  <xsl:template match="td" mode="mergeCell">
    <mergeCell>
      <xsl:variable name="c1" select="@colBegin"/>
      <xsl:variable name="c2" select="@colEnd"/>
      <xsl:attribute name="ref">
        <xsl:value-of select="substring($abc,$c1,1)" />
        <xsl:value-of select="@rowBegin" />
        <xsl:text>:</xsl:text>
        <xsl:value-of select="substring($abc,$c2,1)" />
        <xsl:value-of select="@rowEnd" />
      </xsl:attribute>
    </mergeCell>
  </xsl:template>

  <xsl:template match="td">
    <xsl:variable name="styleName" select="@style" />
    <xsl:variable name="sss" select="$styles[@name=$styleName]"/>
    <c r="{substring($abc,@colBegin,1)}{@rowBegin}" t="s" s="{count($styles[@name=$styleName]/preceding-sibling::style)}">
      <v>
        <xsl:value-of select="count(preceding::td)" />
      </v>
    </c>
  </xsl:template>

  <xsl:template match="td" mode="sst">
    <si>
      <xsl:for-each select="node()">
        <xsl:variable name="zz" select="."/>
        <xsl:choose>
          <xsl:when test="self::text()">
            <xsl:variable name="text" select="normalize-space(translate(self::text(), '&#9;&#10;', ''))" />
            <xsl:if test="string-length($text)>0">
              <r>
                <t>
                  <xsl:value-of select="$text"/>
                </t>
              </r>
            </xsl:if>
          </xsl:when>
          <xsl:when test="name() = 'br'">
            <r>
              <t xml:space="preserve">
</t>
            </r>
          </xsl:when>
          <xsl:when test="name() = 'span'">
            <r>
              <rPr>
                <b />
              </rPr>
              <t>
                <xsl:value-of select="text()"/>
              </t>
            </r>
          </xsl:when>
          <xsl:when test="name() = 'qrcode'">
          </xsl:when>
          <xsl:otherwise>
            <xsl:message terminate="yes">Неизвестный тэг в тексте ячейки</xsl:message>
          </xsl:otherwise>
        </xsl:choose>

      </xsl:for-each>
    </si>
  </xsl:template>

  <xsl:template match="tr">
    <xsl:variable name="h" select="substring-before(@height, 'pt')" />
    <row r="{@index}">
      <xsl:if test="$h > 0">
        <xsl:attribute name="ht">
          <xsl:value-of select="$h" />
        </xsl:attribute>
        <xsl:attribute name="customHeight">1</xsl:attribute>
      </xsl:if>
      <xsl:apply-templates select="td" />
    </row>
  </xsl:template>

  <xsl:template name="getParam">
    <xsl:param name="v"/>
    <xsl:param name="n"/>
    <xsl:choose>
      <xsl:when test="$n = 'left'">
        <xsl:value-of select="substring-before($v, ';')"/>
      </xsl:when>
      <xsl:when test="$n = 'top'">
        <xsl:value-of select="substring-before(substring-after($v, ';'), ';')"/>
      </xsl:when>
      <xsl:when test="$n = 'right'">
        <xsl:value-of select="substring-before(substring-after(substring-after($v, ';'), ';'), ';')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="substring-after(substring-after(substring-after($v, ';'), ';'), ';')"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="border">
    <xsl:param name="v"/>
    <xsl:param name="n"/>
    <xsl:variable name="zz">
      <xsl:call-template name="getParam">
        <xsl:with-param name="n" select="$n" />
        <xsl:with-param name="v" select="$v/@border" />
      </xsl:call-template>
    </xsl:variable>

    <xsl:element name="{$n}">
      <xsl:attribute name="style">
        <xsl:choose>
          <xsl:when test="$zz >= 1.5">thick</xsl:when>
          <xsl:when test="$zz >= 1">medium</xsl:when>
          <xsl:otherwise>thin</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <color indexed="64"/>
    </xsl:element>
  </xsl:template>

  <xsl:template match="styles">
    <!--http://officeopenxml.com/SSstyles.php-->
    <xsl:variable name="styleCount" select="count(style)" />

    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts>
        <font>
          <sz val="11" />
          <color theme="1" />
          <name val="Calibri" />
          <family val="2" />
          <charset val="204" />
          <scheme val="minor" />
        </font>
        <!--<xsl:for-each select="style">
          <xsl:variable name="pos" select="position()" />
          <font>
            
          </font>
        </xsl:for-each>-->
      </fonts>
      <fills>
        <fill>
          <patternFill patternType="none" />
        </fill>
      </fills>
      <borders>
        <border>
          <left />
          <right />
          <top />
          <bottom />
          <diagonal />
        </border>
        <xsl:for-each select="style">
          <border>
            <xsl:call-template name="border">
              <xsl:with-param name="v" select="."/>
              <xsl:with-param name="n" select="'left'"/>
            </xsl:call-template>
            <xsl:call-template name="border">
              <xsl:with-param name="v" select="."/>
              <xsl:with-param name="n" select="'right'"/>
            </xsl:call-template>
            <xsl:call-template name="border">
              <xsl:with-param name="v" select="."/>
              <xsl:with-param name="n" select="'top'"/>
            </xsl:call-template>
            <xsl:call-template name="border">
              <xsl:with-param name="v" select="."/>
              <xsl:with-param name="n" select="'bottom'"/>
            </xsl:call-template>
            <diagonal />
          </border>
        </xsl:for-each>

      </borders>

      <cellStyleXfs>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
      </cellStyleXfs>

      <cellXfs>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
        <xsl:for-each select="style">
          <xsl:variable name="pos" select="position()" />
          <xf numFmtId="0" fontId="0" fillId="0" borderId="{$pos}" xfId="0">
            <xsl:if test="@wrap=1">
              <xsl:attribute name="applyAlignment">1</xsl:attribute>
              <alignment wrapText="1"/>
            </xsl:if>
          </xf>
        </xsl:for-each>
      </cellXfs>

      <cellStyles>
        <cellStyle name="Обычный" xfId="0" builtinId="0" />
      </cellStyles>

      <dxfs />
      <tableStyles defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16" />
    </styleSheet>
  </xsl:template>

  <xsl:template match="/" >
    <xsl:variable name="withRowsCols">
      <xsl:apply-templates select="book/sheet[1]" />
    </xsl:variable>

    <xsl:variable name="table" select="msxsl:node-set($withRowsCols)"/>

    <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <pkg:part pkg:name="/xl/worksheets/sheet1.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml">
        <pkg:xmlData>
          <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <cols>
              <xsl:apply-templates select="book/sheet/colgroup/col" />
            </cols>
            <sheetData>
              <xsl:apply-templates select="$table/table/tr"/>
            </sheetData>
            <mergeCells>
              <xsl:apply-templates select="$table//td[@colBegin!=@colEnd or @rowBegin!=@rowEnd]" mode="mergeCell" />
            </mergeCells>
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />
            <drawing r:id="rId1" />
          </worksheet>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/xl/worksheets/_rels/sheet1.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml" />
          </Relationships>
        </pkg:xmlData>
      </pkg:part>

      <pkg:part pkg:name="/xl/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml">
        <pkg:xmlData>
          <xsl:apply-templates select="book/styles" />
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/xl/workbook.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml">
        <pkg:xmlData>
          <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheets>
              <sheet name="Лист1" sheetId="1" r:id="rId1" />
            </sheets>
          </workbook>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/xl/_rels/workbook.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
            <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
            <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
          </Relationships>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" />
          </Relationships>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/xl/sharedStrings.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml">
        <pkg:xmlData>
          <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <xsl:apply-templates select="$table/table/tr/td" mode="sst"/>
          </sst>
        </pkg:xmlData>
      </pkg:part>

      <pkg:part pkg:name="/xl/drawings/drawing1.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.drawing+xml">
        <pkg:xmlData>
          <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <xsl:for-each select="$table//td/qrcode">
              <xdr:twoCellAnchor editAs="twoCell">
                <xdr:from>
                  <xdr:col><xsl:value-of select="../@colBegin - 1"/></xdr:col>
                  <xdr:colOff>0.05cm</xdr:colOff>
                  <xdr:row><xsl:value-of select="../@rowBegin - 1"/></xdr:row>
                  <xdr:rowOff>0.05cm</xdr:rowOff>
                </xdr:from>
                <xdr:to>
                  <xdr:col><xsl:value-of select="../@colEnd"/></xdr:col>
                  <xdr:colOff>-0.05cm</xdr:colOff>
                  <xdr:row><xsl:value-of select="../@rowEnd"/></xdr:row>
                  <xdr:rowOff>-0.05cm</xdr:rowOff>
                </xdr:to>
                <xdr:pic>
                  <xdr:nvPicPr>
                    <xdr:cNvPr id="1" name="" />
                    <xdr:cNvPicPr>
                      <a:picLocks noChangeAspect="1" />
                    </xdr:cNvPicPr>
                  </xdr:nvPicPr>
                  <xdr:blipFill>
                    <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1">
                      <xsl:attribute name="r:embed">
                        <xsl:text>rId</xsl:text>
                        <xsl:value-of select="position()"/>
                      </xsl:attribute>
                    </a:blip>
                    <a:stretch>
                      <a:fillRect />
                    </a:stretch>
                  </xdr:blipFill>
                  <xdr:spPr>
                    <a:prstGeom prst="rect"/>
                  </xdr:spPr>
                </xdr:pic>
                <xdr:clientData />
              </xdr:twoCellAnchor>
            </xsl:for-each>
          </xdr:wsDr>
        </pkg:xmlData>
      </pkg:part>

      <pkg:part pkg:name="/xl/drawings/_rels/drawing1.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <xsl:for-each select="//td/qrcode">
              <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
                <xsl:attribute name="Id">
                  <xsl:text>rId</xsl:text>
                  <xsl:value-of select="position()"/>
                </xsl:attribute>
                <xsl:attribute name="Target">
                  <xsl:text>../media/image</xsl:text>
                  <xsl:value-of select="position()"/>
                  <xsl:text>.png</xsl:text>
                </xsl:attribute>
              </Relationship>
            </xsl:for-each>
          </Relationships>
        </pkg:xmlData>
      </pkg:part>

      <xsl:for-each select="//td/qrcode">
        <!--<pkg:part pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
          <xsl:attribute name="pkg:name">
            <xsl:text>/xl/drawings/_rels/drawing</xsl:text>
            <xsl:value-of select="position()"/>
            <xsl:text>.xml.rels</xsl:text>
          </xsl:attribute>
          <pkg:xmlData>
            <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
              <xsl:attribute name="rId">
                <xsl:text>rId</xsl:text>
                <xsl:value-of select="position()"/>
              </xsl:attribute>
              <xsl:attribute name="Target">
                <xsl:text>../media/image</xsl:text>
                <xsl:value-of select="position()"/>
                <xsl:text>.png</xsl:text>
              </xsl:attribute>
            </Relationship>
          </pkg:xmlData>
        </pkg:part>-->

        <pkg:part pkg:contentType="image/png" pkg:compression="store">
          <xsl:attribute name="pkg:name">
            <xsl:text>/xl/media/image</xsl:text>
            <xsl:value-of select="position()"/>
            <xsl:text>.png</xsl:text>
          </xsl:attribute>
          <pkg:binaryData>
            <xsl:value-of select="obj:GetQrCode(@data)"/>
          </pkg:binaryData>
        </pkg:part>
      </xsl:for-each>
    </pkg:package>
  </xsl:template>

  <xsl:template match="sheet">
    <table xmlns="">
      <xsl:variable name="rows">
        <xsl:call-template name="row">
          <xsl:with-param name="row" select="tr[1]"/>
          <xsl:with-param name="index" select="1"/>
        </xsl:call-template>
      </xsl:variable>

      <xsl:for-each select="msxsl:node-set($rows)/tr">
        <xsl:variable name="row" select="."/>
        <tr xmlns="" index="{@rowIndex}">
          <xsl:copy-of select="$row/@*"/>
          <xsl:for-each select="$row/td">
            <td xmlns="" colBegin="{@colBegin}" colEnd="{@colEnd}" rowBegin="{@rowBegin}" rowEnd="{@rowEnd}">
              <xsl:copy-of select="text() | * | @*"/>
            </td>
          </xsl:for-each>
        </tr>
      </xsl:for-each>
    </table>
  </xsl:template>

  <xsl:template name="row">
    <!-- Список предыдущих обработанных строк -->
    <xsl:param name="rows"/>
    <!-- Текущая строка -->
    <xsl:param name="row"/>
    <!-- Индекс текущей строки -->
    <xsl:param name="index"/>

    <!-- Генерация новой строки на основании текущей -->
    <xsl:variable name="new-row">
      <tr xmlns="" rowIndex="{$index}">
        <xsl:copy-of select="$row/@*"/>
        <!-- Заполняем ячейки -->
        <xsl:call-template name="td">
          <xsl:with-param name="rows" select="$rows"/>
          <xsl:with-param name="rowIndex" select="$index"/>
          <xsl:with-param name="td" select="$row/td[1]"/>
          <xsl:with-param name="index" select="1"/>
        </xsl:call-template>
      </tr>
    </xsl:variable>

    <xsl:variable name="new-row-nodes" select="msxsl:node-set($new-row)"/>

    <!-- Увеличиваем список строк -->
    <xsl:variable name="new-rows">
      <xsl:copy-of select="msxsl:node-set($rows)/tr"/>
      <xsl:copy-of select="$new-row-nodes/tr"/>
    </xsl:variable>

    <xsl:variable name="next" select="$row/following-sibling::tr[1]"/>
    <xsl:choose>
      <xsl:when test="$next">
        <xsl:call-template name="row">
          <xsl:with-param name="rows" select="$new-rows"/>
          <xsl:with-param name="row" select="$next"/>
          <xsl:with-param name="index" select="$index + 1"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <!-- Возвращаем итоговый результат -->
        <xsl:copy-of select="msxsl:node-set($new-rows)/tr"/>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="td">
    <!-- Список предыдущих обработанных строк -->
    <xsl:param name="rows"/>
    <!-- Индекс текущей строки -->
    <xsl:param name="rowIndex"/>
    <!-- Текущая колонка -->
    <xsl:param name="td"/>
    <!-- Индекс текущей колоки -->
    <xsl:param name="index"/>

    <!-- Находим ячейку, которая занимает аналогичную строку и колонку [@rowEnd >= $rowIndex and @colBegin <= $index and @colEnd >= $index] -->
    <xsl:variable name="max" select="msxsl:node-set($rows)/tr/td[@rowEnd >= $rowIndex and $index >= @colBegin and @colEnd >= $index]"/>

    <xsl:choose>
      <xsl:when test="$max">
        <!-- Генерирует новую фиктивную ячейку для rowSpan -->

        <xsl:call-template name="ghostCells">
          <xsl:with-param name="style" select="$max/@style"/>
          <xsl:with-param name="n" select="$max/@colEnd - $max/@colBegin + 1"/>
          <xsl:with-param name="rowIndex" select="$rowIndex"/>
          <xsl:with-param name="colIndex" select="$index"/>
        </xsl:call-template>

        <!-- Вызывает обработку этой же ячейки, но уже со сдвигом, для проверки на следующие объединные ячейки -->
        <xsl:call-template name="td">
          <xsl:with-param name="rows" select="$rows"/>
          <xsl:with-param name="rowIndex" select="$rowIndex"/>
          <xsl:with-param name="td" select="$td"/>
          <xsl:with-param name="index" select="$max/@colEnd + 1"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="$td">
          <xsl:variable name="colspan">
            <xsl:choose>
              <xsl:when test="$td/@colspan">
                <xsl:value-of select="$td/@colspan"/>
              </xsl:when>
              <xsl:otherwise>1</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>

          <xsl:variable name="rowspan">
            <xsl:choose>
              <xsl:when test="$td/@rowspan">
                <xsl:value-of select="$td/@rowspan"/>
              </xsl:when>
              <xsl:otherwise>1</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>


          <!-- Генерирует новую ячейку -->
          <td xmlns="" colBegin="{$index}" colEnd="{$index + $colspan - 1}" rowBegin="{$rowIndex}" rowEnd="{$rowIndex + $rowspan - 1}">
            <xsl:copy-of select="$td/text() | $td/* | $td/@*"/>
          </td>
          <xsl:call-template name="ghostCells">
            <xsl:with-param name="style" select="$td/@style"/>
            <xsl:with-param name="n" select="$colspan - 1"/>
            <xsl:with-param name="rowIndex" select="$rowIndex"/>
            <xsl:with-param name="colIndex" select="$index + 1"/>
          </xsl:call-template>

          <!-- И вызываем обработку следующей ячейки -->
          <xsl:variable name="next" select="$td/following-sibling::td[1]"/>
          <xsl:call-template name="td">
            <xsl:with-param name="rows" select="$rows"/>
            <xsl:with-param name="rowIndex" select="$rowIndex"/>
            <xsl:with-param name="td" select="$next"/>
            <xsl:with-param name="index" select="$index + $colspan"/>
          </xsl:call-template>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="ghostCells">
    <xsl:param name="style"/>
    <xsl:param name="n"/>
    <xsl:param name="rowIndex"/>
    <xsl:param name="colIndex"/>

    <xsl:if test="$n>0">
      <td xmlns="" colBegin="{$colIndex}" colEnd="{$colIndex}" rowBegin="{$rowIndex}" rowEnd="{$rowIndex}" style="{$style}"/>
      <xsl:call-template name="ghostCells">
        <xsl:with-param name="style" select="$style"/>
        <xsl:with-param name="n" select="$n - 1"/>
        <xsl:with-param name="rowIndex" select="$rowIndex"/>
        <xsl:with-param name="colIndex" select="$colIndex + 1"/>
      </xsl:call-template>
    </xsl:if>

  </xsl:template>

  <xsl:template match="@* | node()">
    <xsl:message terminate="yes">
      Hell
    </xsl:message>
  </xsl:template>
</xsl:stylesheet>