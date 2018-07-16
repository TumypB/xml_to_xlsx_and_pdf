<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:ms="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="ms"
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt">
  <xsl:output method="xml" indent="yes"/>

  <xsl:variable name="abc">
    <xsl:text>ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:text>
  </xsl:variable>

  <xsl:variable name="headrows"  select="3"/>
  <xsl:variable name="inst" select="Root/Inst"/>

  <xsl:template name="plus">
    <c t="inlineStr" s="6">
      <is>
        <t>+</t>
      </is>
    </c>
  </xsl:template>

  <xsl:template name="plusFilled">
    <c t="inlineStr" s="18">
      <is>
        <t>+</t>
      </is>
    </c>
  </xsl:template>

  <xsl:template name="plusFilled2">
    <c t="inlineStr" s="22">
      <is>
        <t>+</t>
      </is>
    </c>
  </xsl:template>

  <xsl:template match="Spec" >
    <xsl:variable name="cntEp" select="count(Edprog)"/>
    <c t="inlineStr" s="3">
      <is>
        <t>
          <xsl:value-of select="@code"/>
          <xsl:if test="string-length(@code)=6">
            <xsl:text>.</xsl:text>
            <xsl:value-of select="@codequal"/>
          </xsl:if>
        </t>
      </is>
    </c>
    <xsl:apply-templates select="Edprog[position()=1]"/>
  </xsl:template>

  <xsl:template name="itog">
    <xsl:param name="rown"/>
    <c t="inlineStr" s="11">
      <!--итоговый столбец СУММ ФБ-->
      <f>
        <xsl:text>SUM(E</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,H</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,K</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,N</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,Q</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,T</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>)</xsl:text>
      </f>
    </c>
    <c s="11">
      <f>
        <xsl:text>SUM(G</xsl:text>
        <!--итоговый столбец СУММ ПВЗ-->
        <xsl:value-of select="$rown"/>
        <xsl:text>,J</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,M</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,P</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,S</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,V</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>)</xsl:text>
      </f>
    </c>
    <c s="11">
      <v>
        <xsl:value-of select="sum(Course/@foreigner)"/>
      </v>
    </c>
    <c s="11">
      <!--итоговый столбец СУММ Всего-->
      <f>
        <xsl:text>SUM(W</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,X</xsl:text>
        <xsl:value-of select="$rown"/>
        <!--<xsl:text>,Y</xsl:text>
        <xsl:value-of select="$rown"/>-->
        <xsl:text>)</xsl:text>
      </f>
    </c>
  </xsl:template>

  <xsl:template name="itogPoSpec">
    <xsl:param name="rown"/>
    <c t="inlineStr" s="23">
      <!--итоговый столбец СУММ ФБ-->
      <f>
        <xsl:text>SUM(E</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,H</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,K</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,N</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,Q</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,T</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>)</xsl:text>
      </f>
    </c>
    <c s="23">
      <f>
        <xsl:text>SUM(G</xsl:text>
        <!--итоговый столбец СУММ ПВЗ-->
        <xsl:value-of select="$rown"/>
        <xsl:text>,J</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,M</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,P</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,S</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,V</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>)</xsl:text>
      </f>
    </c>
    <c s="23">
      <v>
        <xsl:value-of select="sum(Course/@foreigner)"/>
      </v>
    </c>
    <c s="23">
      <!--итоговый столбец СУММ Всего-->
      <f>
        <xsl:text>SUM(W</xsl:text>
        <xsl:value-of select="$rown"/>
        <xsl:text>,X</xsl:text>
        <xsl:value-of select="$rown"/>
        <!--<xsl:text>,Y</xsl:text>
        <xsl:value-of select="$rown"/>-->
        <xsl:text>)</xsl:text>
      </f>
    </c>
  </xsl:template>

  <xsl:template name="univer">
    <xsl:param name="col"/>
    <f>
      <xsl:for-each select="Root/Inst">
        <xsl:value-of select="$col"/>
        <xsl:value-of select="count(Spec/Edprog)
                                +count(Spec[count(Edprog)>1])
                                +count(preceding::Inst)
                                +count(preceding::Spec/Edprog)
                                +count(preceding::Spec[count(Edprog)>1])
                                +count(preceding::CodeQual)
                                +count(CodeQuals/CodeQual)
                                +$headrows+1"/>
        <xsl:text>+</xsl:text>
      </xsl:for-each>
      <xsl:text>0</xsl:text>
    </f>
  </xsl:template>

  <xsl:template match="Edprog" >
    <xsl:variable name="parent" select="count(../preceding::Spec[count(Edprog)>1])"/>
    <xsl:variable name="rown" select="$headrows+1
                  +count(preceding::Edprog)
                  +count(preceding::Inst)
                  +count(preceding::CodeQual)
                  +$parent"/>
    <c t="inlineStr">
      <xsl:attribute name="s">
        <xsl:choose>
          <xsl:when test="@studyForm='oz'">24</xsl:when>
          <xsl:otherwise>2</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <is>
        <t>
          <xsl:value-of select="@name"/>
          <xsl:if test="@studyForm='oz'">
            <xsl:text> (ОЗФО)</xsl:text>
          </xsl:if>
        </t>
      </is>
    </c>
    <xsl:choose>
      <xsl:when test="@base='spo'">
        <c t="inlineStr" s="7">
          <is>
            <t>СПО</t>
          </is>
        </c>
      </xsl:when>
      <xsl:when test="@base='vpo'">
        <c t="inlineStr" s="7">
          <is>
            <t>ВПО</t>
          </is>
        </c>
      </xsl:when>
      <xsl:otherwise>
        <c t="inlineStr" s="7"/>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:call-template name="Course">
      <xsl:with-param name="c" select="Course[@c=1]"/>
    </xsl:call-template>
    <xsl:call-template name="Course">
      <xsl:with-param name="c" select="Course[@c=2]"/>
    </xsl:call-template>
    <xsl:call-template name="Course">
      <xsl:with-param name="c" select="Course[@c=3]"/>
    </xsl:call-template>
    <xsl:call-template name="Course">
      <xsl:with-param name="c" select="Course[@c=4]"/>
    </xsl:call-template>
    <xsl:call-template name="Course">
      <xsl:with-param name="c" select="Course[@c=5]"/>
    </xsl:call-template>
    <xsl:call-template name="Course">
      <xsl:with-param name="c" select="Course[@c=6]"/>
    </xsl:call-template>
    <xsl:call-template name="itog">
      <xsl:with-param name="rown" select="$rown"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="Course" >
    <xsl:param name="c"/>
    <xsl:choose>
      <xsl:when test="$c">
        <c s="4">
          <v>
            <xsl:value-of select="$c/@b+$c/@bc"/>
          </v>
        </c>
        <xsl:call-template name="plus"/>
        <c s="5">
          <v>
            <xsl:value-of select="$c/@p"/>
          </v>
        </c>
      </xsl:when>
      <xsl:otherwise>
        <c s="4"/>
        <c t="inlineStr" s="6">
          <is>
            <t>-</t>
          </is>
        </c>
        <c s="5"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--<xsl:template name="csum" >
    <xsl:param name="inst"/>
    <xsl:param name="col"/>
    <xsl:variable name="specItog1" select="count($inst/Spec[count(Edprog)>1])"/>
    --><!--кол-во спец для текущ инст где больше 1 едпрога--><!--
    <xsl:variable name="precInst" select="$inst/preceding::Inst"/>
    <xsl:variable name="r1" select="$headrows+1+count($inst/preceding::Edprog)
                  +count($precInst)+count($precInst/Spec[count(Edprog)>1])
                  +count($precInst/CodeQuals/CodeQual)"/>
    <xsl:variable name="r2" select="$r1 -1+count($inst/Spec/Edprog)+$specItog1"/>

    <c s="17">
      <f>
        <xsl:text>SUM(</xsl:text>
        <xsl:value-of select="substring($abc,$col,1)"/>
        <xsl:value-of select="$r1"/>
        <xsl:text>:</xsl:text>
        <xsl:value-of select="substring($abc,$col,1)"/>
        <xsl:value-of select="$r2"/>
        <xsl:text>)</xsl:text>
        <xsl:for-each select="$inst/Spec[count(Edprog)>1]">
          <xsl:text>-</xsl:text>
          <xsl:value-of select="substring($abc,$col,1)"/>
          <xsl:variable name="v1" select="count(preceding::Spec/Edprog)+count(Edprog)"/>
          <xsl:variable name="v2" select="count(preceding::Spec[count(Edprog)>1])"/>
          <xsl:variable name="precedingQualSummary" select="count($inst/CodeQuals/CodeQual)+count($precInst/CodeQuals/CodeQual)"/>
          <xsl:value-of select="$headrows+1 + $v1 + $v2 + count($precInst) + ($precedingQualSummary - 2)"/>
        </xsl:for-each>
      </f>
    </c>
    <xsl:call-template name="plusFilled"/>
    <c s="19">
      <f>
        <xsl:text>SUM(</xsl:text>
        <xsl:value-of select="substring($abc,$col+2,1)"/>
        <xsl:value-of select="$r1"/>
        <xsl:text>:</xsl:text>
        <xsl:value-of select="substring($abc,$col+2,1)"/>
        <xsl:value-of select="$r2"/>
        <xsl:text>)</xsl:text>
        <xsl:for-each select="$inst/Spec[count(Edprog)>1]">
          <xsl:text>-</xsl:text>
          <xsl:value-of select="substring($abc,$col+2,1)"/>
          <xsl:variable name="v1" select="count(preceding::Spec/Edprog)+count(Edprog)"/>
          <xsl:variable name="v2" select="count(preceding::Spec[count(Edprog)>1])"/>
          <xsl:value-of select="$headrows+1 + $v1 + $v2 + count($precInst) + count($precInst/CodeQuals/CodeQual)"/>
        </xsl:for-each>
      </f>
    </c>
  </xsl:template>-->

  <xsl:template name="emptyC">
    <xsl:param name="cnt"/>
    <c s="6"></c>
    <xsl:if test="$cnt>1">
      <xsl:call-template name="emptyC">
        <xsl:with-param name="cnt" select="$cnt -1"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <xsl:template name="poSpec">
    <xsl:param name="course"/>
    <xsl:param name="specNode"/>

    <xsl:variable name="c1" select="sum($specNode/Edprog/Course[@c=$course]/@b)
                                + sum($specNode/Edprog/Course[@c=$course]/@bc)"/>
    <xsl:variable name="c2" select="sum($specNode/Edprog/Course[@c=$course]/@p)"/>
    <xsl:choose>
      <xsl:when test="$c1+$c2 != 0">
        <c s="20">
          <v>
            <xsl:value-of select="$c1"/>
          </v>
        </c>
        <xsl:call-template name="plusFilled2"/>
        <c s="21">
          <v>
            <xsl:value-of select="$c2"/>
          </v>
        </c>
      </xsl:when>
      <xsl:otherwise>
        <c s="20"/>
        <c t="inlineStr" s="22">
          <is>
            <t>-</t>
          </is>
        </c>
        <c s="21"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="Spec" mode="vsegoPoSpec">
    <xsl:variable name="rown">
      <xsl:value-of select="$headrows+1
                    +count(preceding::Spec[count(Edprog)>1])
                    +count(../preceding::Inst)
                    +count(preceding::Spec/Edprog)
                    +count(Edprog)"/>
    </xsl:variable>
    <row>
      <c s="8"/>
      <c s="8"/>
      <c t="inlineStr" s="20">
        <is>
          <xsl:choose>
            <xsl:when test="@codequal=62 or @codequal=68">
              <t>Всего по направлению</t>
            </xsl:when>
            <xsl:otherwise>
              <t>Всего по специальности</t>
            </xsl:otherwise>
          </xsl:choose>
        </is>
      </c>
      <c s="21"/>
      <xsl:call-template name="poSpec">
        <xsl:with-param name="course" select="1"/>
        <xsl:with-param name="specNode" select="."/>
      </xsl:call-template>
      <xsl:call-template name="poSpec">
        <xsl:with-param name="course" select="2"/>
        <xsl:with-param name="specNode" select="."/>
      </xsl:call-template>
      <xsl:call-template name="poSpec">
        <xsl:with-param name="course" select="3"/>
        <xsl:with-param name="specNode" select="."/>
      </xsl:call-template>
      <xsl:call-template name="poSpec">
        <xsl:with-param name="course" select="4"/>
        <xsl:with-param name="specNode" select="."/>
      </xsl:call-template>
      <xsl:call-template name="poSpec">
        <xsl:with-param name="course" select="5"/>
        <xsl:with-param name="specNode" select="."/>
      </xsl:call-template>
      <xsl:call-template name="poSpec">
        <xsl:with-param name="course" select="6"/>
        <xsl:with-param name="specNode" select="."/>
      </xsl:call-template>
      <xsl:call-template name="itogPoSpec">
        <xsl:with-param name="rown" select="$rown"/>
      </xsl:call-template>
    </row>
  </xsl:template>

  <xsl:template name="itogCodeQual">
    <xsl:param name="attribute" />
    <xsl:param name="style"/>
    <c>
      <xsl:attribute name="s">
        <xsl:value-of select="$style"/>
      </xsl:attribute>
      <v>
        <xsl:choose>
          <xsl:when test="string-length($attribute) > 0">
            <xsl:value-of select="$attribute"/>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </v>
    </c>
  </xsl:template>

  <xsl:template match="/" >
    <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <pkg:part pkg:name="/xl/worksheets/sheet1.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml">
        <pkg:xmlData>
          <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <cols>
              <col min="1" max="1" width="8.8" bestFit="1" customWidth="1"/>
              <col min="2" max="2" width="9.7" bestFit="1" customWidth="1"/>
              <col min="3" max="3" width="43" bestFit="1"  customWidth="1"/>
              <col min="4" max="4" width="5.1" bestFit="1" customWidth="1"/>
              <col min="5" max="5" width="4" bestFit="1" customWidth="1"/>
              <col min="6" max="6" width="1.7" bestFit="1" customWidth="1"/>
              <col min="7" max="7" width="4" bestFit="1" customWidth="1"/>
              <col min="8" max="8" width="4" bestFit="1" customWidth="1"/>
              <col min="9" max="9" width="1.7" bestFit="1" customWidth="1"/>
              <col min="10" max="11" width="4" bestFit="1" customWidth="1"/>
              <col min="12" max="12" width="1.7" bestFit="1" customWidth="1"/>
              <col min="13" max="13" width="4" bestFit="1" customWidth="1"/>
              <col min="14" max="14" width="4" bestFit="1" customWidth="1"/>
              <col min="15" max="15" width="1.7" bestFit="1" customWidth="1"/>
              <col min="16" max="17" width="4" bestFit="1" customWidth="1"/>
              <col min="18" max="18" width="1.7" bestFit="1" customWidth="1"/>
              <col min="19" max="20" width="4" bestFit="1" customWidth="1"/>
              <col min="21" max="21" width="1.7" bestFit="1" customWidth="1"/>
              <col min="22" max="22" width="4" bestFit="1" customWidth="1"/>
              <col min="23" max="23" width="4.9" bestFit="1" customWidth="1"/>
              <col min="24" max="24" width="4.9" bestFit="1" customWidth="1"/>
              <col min="25" max="25" width="5.7" bestFit="1" customWidth="1"/>
              <col min="26" max="26" width="5" bestFit="1" customWidth="1"/>
            </cols>
            <sheetData>
              <row>
                <c t="inlineStr" s="25">
                  <is>
                    <t>
                      Контингент <xsl:choose>
                        <xsl:when test="Root/Inst/Spec/Edprog/@studyForm='o'">
                          <xsl:text>очной формы обучения </xsl:text>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:text>заочной и очно-заочной формы обучения </xsl:text>
                        </xsl:otherwise>
                      </xsl:choose>
                      <xsl:text>на </xsl:text>
                      <xsl:value-of select="msxsl:format-date(Root/@date,'dd.MM.yyyy')"/>
                    </t>
                  </is>
                </c>
              </row>
              <!--Шапка таблицы-->
              <row>
                <c t="inlineStr" s="13">
                  <is>
                    <t>Институт</t>
                  </is>
                </c>
                <c t="inlineStr" s="13">
                  <is>
                    <t>Шифр</t>
                  </is>
                </c>
                <c t="inlineStr" s="13">
                  <is>
                    <t>Образовательная программа</t>
                  </is>
                </c>
                <c t="inlineStr" s="13">
                  <is>
                    <t>На базе</t>
                  </is>
                </c>
                <c t="inlineStr" s="15">
                  <is>
                    <t>Курсы</t>
                  </is>
                </c>
                <xsl:call-template name="emptyC">
                  <xsl:with-param name="cnt" select="17"/>
                </xsl:call-template>
                <c t="inlineStr" s="13">
                  <is>
                    <t>ФБ</t>
                  </is>
                </c>
                <c t="inlineStr" s="13">
                  <is>
                    <t>ПВЗ</t>
                  </is>
                </c>

                <c t="inlineStr" s="13">
                  <is>
                    <t>В т.ч. ин</t>
                  </is>
                </c>
                <c t="inlineStr" s="13">
                  <is>
                    <t>Все-го</t>
                  </is>
                </c>
              </row>
              <!--Номера курсов-->
              <row>
                <c s="8"  />
                <c s="8"  />
                <c/>
                <c s="8"/>
                <c s="4"/>
                <c s="14">
                  <v>1</v>
                </c>
                <c s="5"></c>
                <c s="4"/>
                <c s="14">
                  <v>2</v>
                </c>
                <c s="5"></c>
                <c s="4"/>
                <c s="14">
                  <v>3</v>
                </c>
                <c s="5"></c>
                <c s="4"/>
                <c s="14">
                  <v>4</v>
                </c>
                <c s="5"></c>
                <c s="4"/>
                <c s="14">
                  <v>5</v>
                </c>
                <c s="5"></c>
                <c s="4"/>
                <c s="14">
                  <v>6</v>
                </c>
                <c s="5"></c>
                <c/>
                <c s="8"  />
                <c s="8"  />
                <c s="8"  />
              </row>
              <xsl:for-each select="/Root/Inst">
                <row>
                  <c t="inlineStr" s="3">
                    <is>
                      <t>
                        <xsl:value-of select="@short"/>
                      </t>
                    </is>
                  </c>
                  <xsl:apply-templates select="Spec[position()=1]"/>
                </row>

                <xsl:for-each select="Spec[position()=1]">
                  <xsl:variable name="cntEp" select="count(Edprog)"/>
                  <!--кол-во едпрог в текущ спец-->
                  <xsl:for-each select="Edprog[position()>1]">
                    <row>
                      <c s="8"/>
                      <c s="8"/>
                      <xsl:apply-templates select="."/>
                    </row>
                    <xsl:if test="position()+1 =$cntEp">
                      <xsl:apply-templates mode="vsegoPoSpec" select=".."/>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:for-each>

                <xsl:for-each select="Spec[position()>1]">
                  <xsl:variable name="cntEp" select="count(Edprog)"/>
                  <!--кол-во едпрог в текущ спец-->
                  <row>
                    <c s="8"/>
                    <xsl:apply-templates select="."/>
                  </row>
                  <xsl:for-each select="Edprog[position()>1]">
                    <row>
                      <c s="8"/>
                      <c s="8"/>
                      <xsl:apply-templates select="."/>
                    </row>
                    <xsl:if test="position()+1 = $cntEp">
                      <xsl:apply-templates mode="vsegoPoSpec" select=".."/>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:for-each>
                <!--<CodeQuals>
              <CodeQual codeQual="62" c1b="174" c1p="4" c2b="118" c2p="4" c3b="108" c3p="4" c4b="86" c4p="8" foreigner="3" />
              <CodeQual codeQual="68" c1b="41" c1p="0" c2b="39" c2p="0" foreigner="1" />
            </CodeQuals>-->
                <xsl:for-each select="CodeQuals">
                  <xsl:for-each select="CodeQual">
                    <row>
                      <c s="8"/>
                      <c t="inlineStr" s="20">
                        <is>
                          <xsl:choose>
                            <xsl:when test="@codeQual=62">
                              <t>Всего по программам бакалавриата</t>
                            </xsl:when>
                            <xsl:when test="@codeQual=65">
                              <t>Всего по программам специалитета</t>
                            </xsl:when>
                            <xsl:otherwise>
                              <t>Всего по программам магистратуры</t>
                            </xsl:otherwise>
                          </xsl:choose>
                        </is>
                      </c>
                      <c s="20"/>
                      <c s="20"/>
                      <c s="20">
                        <v>
                          <xsl:value-of select="@c1b"/>
                        </v>
                      </c>
                      <xsl:call-template name="plusFilled2"/>
                      <c s="21">
                        <v>
                          <xsl:value-of select="@c1p"/>
                        </v>
                      </c>
                      <c s="20">
                        <v>
                          <xsl:value-of select="@c2b"/>
                        </v>
                      </c>
                      <xsl:call-template name="plusFilled2"/>
                      <c s="21">
                        <v>
                          <xsl:value-of select="@c2p"/>
                        </v>
                      </c>
                      <xsl:call-template name="itogCodeQual">
                        <xsl:with-param name="attribute" select="@c3b"/>
                        <xsl:with-param name="style" select="20"/>
                      </xsl:call-template>

                      <xsl:call-template name="plusFilled2"/>

                      <xsl:call-template name="itogCodeQual">
                        <xsl:with-param name="attribute" select="@c3p"/>
                        <xsl:with-param name="style" select="21"/>
                      </xsl:call-template>

                      <xsl:call-template name="itogCodeQual">
                        <xsl:with-param name="attribute" select="@c4b"/>
                        <xsl:with-param name="style" select="20"/>
                      </xsl:call-template>

                      <xsl:call-template name="plusFilled2"/>

                      <xsl:call-template name="itogCodeQual">
                        <xsl:with-param name="attribute" select="@c4p"/>
                        <xsl:with-param name="style" select="21"/>
                      </xsl:call-template>

                      <xsl:call-template name="itogCodeQual">
                        <xsl:with-param name="attribute" select="@c5b"/>
                        <xsl:with-param name="style" select="20"/>
                      </xsl:call-template>

                      <xsl:call-template name="plusFilled2"/>

                      <xsl:call-template name="itogCodeQual">
                        <xsl:with-param name="attribute" select="@c5p"/>
                        <xsl:with-param name="style" select="21"/>
                      </xsl:call-template>

                      <xsl:call-template name="itogCodeQual">
                        <xsl:with-param name="attribute" select="@c6b"/>
                        <xsl:with-param name="style" select="20"/>
                      </xsl:call-template>

                      <xsl:call-template name="plusFilled2"/>

                      <xsl:call-template name="itogCodeQual">
                        <xsl:with-param name="attribute" select="@c6p"/>
                        <xsl:with-param name="style" select="21"/>
                      </xsl:call-template>
                      <!--итоговые 4 столбца: всего ФБ; ПВЗ; в т.ч. ин; всего-->
                      <c s="23">
                        <v>
                          <xsl:value-of select="@c1b+@c2b+@c3b+@c4b+@c5b+@c6b"/>
                        </v>
                      </c>
                      <c s="23">
                        <v>
                          <xsl:value-of select="@c1p+@c2p+@c3p+@c4p+@c5p+@c6p"/>
                        </v>
                      </c>
                      <c s="23">
                        <v>
                          <xsl:value-of select="@foreigner"/>
                        </v>
                      </c>
                      <c s="23">
                        <v>
                          <xsl:value-of select="@c1b+@c2b+@c3b+@c4b+@c5b+@c6b+@c1p+@c2p+@c3p+@c4p+@c5p+@c6p"/>
                        </v>
                      </c>
                    </row>
                  </xsl:for-each>
                </xsl:for-each>
                <row>
                  <c s="10"/>
                  <c t="inlineStr" s="17">
                    <is>
                      <t>
                        <xsl:text>Итого по институту:</xsl:text>
                      </t>
                    </is>
                  </c>
                  <c s="18"/>
                  <c s="19"/>
                  <!--1 курс-->
                  <c s="17">
                    <f>
                      <xsl:for-each select="./CodeQuals/CodeQual">
                        <xsl:value-of select="@c1b"/>
                        <xsl:text>+</xsl:text>
                      </xsl:for-each>
                      <xsl:text>0</xsl:text>
                    </f>
                  </c>
                    <xsl:call-template name="plusFilled"/>
                  <c s="19">
                    <f>
                      <xsl:for-each select="./CodeQuals/CodeQual">
                        <xsl:value-of select="@c1p"/>
                        <xsl:text>+</xsl:text>
                      </xsl:for-each>
                      <xsl:text>0</xsl:text>
                    </f>
                  </c>
                  <!--2 курс-->
                  <c s="17">
                    <f>
                      <xsl:for-each select="./CodeQuals/CodeQual">
                        <xsl:value-of select="@c2b"/>
                        <xsl:text>+</xsl:text>
                      </xsl:for-each>
                      <xsl:text>0</xsl:text>
                    </f>
                  </c>
                  <xsl:call-template name="plusFilled"/>
                  <c s="19">
                    <f>
                      <xsl:for-each select="./CodeQuals/CodeQual">
                        <xsl:value-of select="@c2p"/>
                        <xsl:text>+</xsl:text>
                      </xsl:for-each>
                      <xsl:text>0</xsl:text>
                    </f>
                  </c>
                  <!--3 курс-->
                  <c s="17">
                    <f>
                      <xsl:for-each select="./CodeQuals/CodeQual">
                        <xsl:value-of select="@c3b"/>
                        <xsl:text>+</xsl:text>
                      </xsl:for-each>
                      <xsl:text>0</xsl:text>
                    </f>
                  </c>
                  <xsl:call-template name="plusFilled"/>
                  <c s="19">
                    <f>
                      <xsl:for-each select="./CodeQuals/CodeQual">
                        <xsl:value-of select="@c3p"/>
                        <xsl:text>+</xsl:text>
                      </xsl:for-each>
                      <xsl:text>0</xsl:text>
                    </f>
                  </c>
                  <!--4 курс-->
                  <c s="17">
                    <f>
                      <xsl:for-each select="./CodeQuals/CodeQual">
                        <xsl:value-of select="@c4b"/>
                        <xsl:text>+</xsl:text>
                      </xsl:for-each>
                      <xsl:text>0</xsl:text>
                    </f>
                  </c>
                  <xsl:call-template name="plusFilled"/>
                  <c s="19">
                    <f>
                      <xsl:for-each select="./CodeQuals/CodeQual">
                        <xsl:value-of select="@c4p"/>
                        <xsl:text>+</xsl:text>
                      </xsl:for-each>
                      <xsl:text>0</xsl:text>
                    </f>
                  </c>
                  <!--5 курс-->
                  <c s="17">
                    <f>
                      <xsl:for-each select="./CodeQuals/CodeQual">
                        <xsl:value-of select="@c5b"/>
                        <xsl:text>+</xsl:text>
                      </xsl:for-each>
                      <xsl:text>0</xsl:text>
                    </f>
                  </c>
                  <xsl:call-template name="plusFilled"/>
                  <c s="19">
                    <f>
                      <xsl:for-each select="./CodeQuals/CodeQual">
                        <xsl:value-of select="@c5p"/>
                        <xsl:text>+</xsl:text>
                      </xsl:for-each>
                      <xsl:text>0</xsl:text>
                    </f>
                  </c>
                  <!--6 курс-->
                  <c s="17">
                    <f>
                      <xsl:for-each select="./CodeQuals/CodeQual">
                        <xsl:value-of select="@c6b"/>
                        <xsl:text>+</xsl:text>
                      </xsl:for-each>
                      <xsl:text>0</xsl:text>
                    </f>
                  </c>
                  <xsl:call-template name="plusFilled"/>
                  <c s="19">
                    <f>
                      <xsl:for-each select="./CodeQuals/CodeQual">
                        <xsl:value-of select="@c6p"/>
                        <xsl:text>+</xsl:text>
                      </xsl:for-each>
                      <xsl:text>0</xsl:text>
                    </f>
                  </c>
                  <!--<xsl:call-template name="csum">
                    <xsl:with-param name="col" select="5"/>
                    <xsl:with-param name="inst" select="."/>
                  </xsl:call-template>
                  <xsl:call-template name="csum">
                    <xsl:with-param name="col" select="8"/>
                    <xsl:with-param name="inst" select="."/>
                  </xsl:call-template>
                  <xsl:call-template name="csum">
                    <xsl:with-param name="col" select="11"/>
                    <xsl:with-param name="inst" select="."/>
                  </xsl:call-template>
                  <xsl:call-template name="csum">
                    <xsl:with-param name="col" select="14"/>
                    <xsl:with-param name="inst" select="."/>
                  </xsl:call-template>
                  <xsl:call-template name="csum">
                    <xsl:with-param name="col" select="17"/>
                    <xsl:with-param name="inst" select="."/>
                  </xsl:call-template>
                  <xsl:call-template name="csum">
                    <xsl:with-param name="col" select="20"/>
                    <xsl:with-param name="inst" select="."/>
                  </xsl:call-template>-->
                  <xsl:variable name="specItog2" select="count(Spec[count(Edprog)>1])"/>
                  <xsl:variable name="r1" select="$headrows+1+count(preceding::Edprog)
                            +count(preceding::Inst)
                            +count(preceding::Inst/Spec[count(Edprog)>1])"/>
                  <xsl:variable name="r2" select="$r1 -1+count(Spec/Edprog)+$specItog2 +count(preceding::CodeQuals/CodeQual)-count(CodeQuals/CodeQual)"/>

                  <!--<xsl:variable name="r1" select="$headrows+1+count($inst/preceding::Edprog)
                  +count($precInst)+count($precInst/Spec[count(Edprog)>1])
                  +count($precInst/CodeQuals/CodeQual)"/>
    <xsl:variable name="r2" select="$r1 -1+count($inst/Spec/Edprog)+$specItog1"/>-->
                  <xsl:variable name="precedingCodeQuals" select="count(preceding::CodeQuals/CodeQual)"/>
                  <c s="16">
                    <f>
                      <xsl:text>SUM(W</xsl:text>
                      <xsl:value-of select="$r1+count(preceding::CodeQual)"/>
                      <xsl:text>:W</xsl:text>
                      <xsl:value-of select="$r2+count(CodeQuals/CodeQual)"/>
                      <xsl:text>)</xsl:text>
                      <xsl:for-each select="Spec[count(Edprog)>1]">
                        <xsl:text>-</xsl:text>
                        <xsl:value-of select="substring($abc,23,1)"/>
                        <xsl:variable name="v1" select="count(preceding::Spec/Edprog)+count(Edprog)"/>
                        <xsl:variable name="v2" select="count(preceding::Spec[count(Edprog)>1])+$precedingCodeQuals"/>
                        <xsl:variable name="precInst" select="../preceding::Inst"/>
                        <xsl:value-of select="$headrows+1 + $v1 + $v2 + count($precInst)"/>
                      </xsl:for-each>
                    </f>
                  </c>
                  <c s="16">
                    <f>
                      <xsl:text>SUM(X</xsl:text>
                      <xsl:value-of select="$r1+count(preceding::CodeQual)"/>
                      <xsl:text>:X</xsl:text>
                      <xsl:value-of select="$r2+count(CodeQuals/CodeQual)"/>
                      <xsl:text>)</xsl:text>
                      <xsl:for-each select="Spec[count(Edprog)>1]">
                        <xsl:text>-</xsl:text>
                        <xsl:value-of select="substring($abc,24,1)"/>
                        <xsl:variable name="v1" select="count(preceding::Spec/Edprog)+count(Edprog)"/>
                        <xsl:variable name="v2" select="count(preceding::Spec[count(Edprog)>1])+$precedingCodeQuals"/>
                        <xsl:variable name="precInst" select="../preceding::Inst"/>
                        <xsl:value-of select="$headrows+1 + $v1 + $v2 + count($precInst)+count(CodeQuals/CodeQual)"/>
                      </xsl:for-each>
                    </f>
                  </c>
                  <c s="16">
                    <f>
                      <xsl:text>SUM(Y</xsl:text>
                      <xsl:value-of select="$r1+count(preceding::CodeQual)"/>
                      <xsl:text>:Y</xsl:text>
                      <xsl:value-of select="$r2+count(CodeQuals/CodeQual)"/>
                      <xsl:text>)</xsl:text>
                      <xsl:for-each select="Spec[count(Edprog)>1]">
                        <xsl:text>-</xsl:text>
                        <xsl:value-of select="substring($abc,25,1)"/>
                        <xsl:variable name="v1" select="count(preceding::Spec/Edprog)+count(Edprog)"/>
                        <xsl:variable name="v2" select="count(preceding::Spec[count(Edprog)>1])+$precedingCodeQuals"/>
                        <xsl:variable name="precInst" select="../preceding::Inst"/>
                        <xsl:value-of select="$headrows+1 + $v1 + $v2 + count($precInst)"/>
                      </xsl:for-each>
                    </f>
                  </c>
                  <c s="16">
                    <f>
                      <xsl:text>SUM(Z</xsl:text>
                      <xsl:value-of select="$r1+count(preceding::CodeQual)"/>
                      <xsl:text>:Z</xsl:text>
                      <xsl:value-of select="$r2+count(CodeQuals/CodeQual)"/>
                      <xsl:text>)</xsl:text>
                      <xsl:for-each select="Spec[count(Edprog)>1]">
                        <xsl:text>-</xsl:text>
                        <xsl:value-of select="substring($abc,26,1)"/>
                        <xsl:variable name="v1" select="count(preceding::Spec/Edprog)+count(Edprog)"/>
                        <xsl:variable name="v2" select="count(preceding::Spec[count(Edprog)>1])+$precedingCodeQuals"/>
                        <xsl:variable name="precInst" select="../preceding::Inst"/>
                        <xsl:value-of select="$headrows+1 + $v1 + $v2 + count($precInst)"/>
                      </xsl:for-each>
                    </f>
                  </c>
                </row>
              </xsl:for-each>
              <!--Всего по университету-->
              <xsl:for-each select="/Root/CommonCodeQuals/CommonCodeQual">
                <row>
                  <c t="inlineStr" s="17">
                    <is>
                      <t>
                        <xsl:text>Итого по программам</xsl:text>
                        <xsl:choose>
                          <xsl:when test="@codeQual=62">
                            <xsl:text> бакалавриата</xsl:text>
                          </xsl:when>
                          <xsl:when test="@codeQual=65">
                            <xsl:text> специалитета</xsl:text>
                          </xsl:when>
                          <xsl:when test="@codeQual=68">
                            <xsl:text> магистратуры</xsl:text>
                          </xsl:when>
                        </xsl:choose>
                      </t>
                    </is>
                  </c>
                  <c s="18"/>
                  <c s="18"/>
                  <c s="19"/>
                  <c s="17">
                    <v>
                      <xsl:value-of select="@c1b"/>
                    </v>
                  </c>
                  <xsl:call-template name="plusFilled"/>
                  <c s="19">
                    <v>
                      <xsl:value-of select="@c1p"/>
                    </v>
                  </c>
                  <c s="17">
                    <v>
                      <xsl:value-of select="@c2b"/>
                    </v>
                  </c>
                  <xsl:call-template name="plusFilled"/>
                  <c s="19">
                    <v>
                      <xsl:value-of select="@c2p"/>
                    </v>
                  </c>
                  <c s="17">
                    <v>
                      <xsl:value-of select="@c3b"/>
                    </v>
                  </c>
                  <xsl:call-template name="plusFilled"/>
                  <c s="19">
                    <v>
                      <xsl:value-of select="@c3p"/>
                    </v>
                  </c>
                  <c s="17">
                    <v>
                      <xsl:value-of select="@c4b"/>
                    </v>
                  </c>
                  <xsl:call-template name="plusFilled"/>
                  <c s="19">
                    <v>
                      <xsl:value-of select="@c4p"/>
                    </v>
                  </c>
                  <c s="17">
                    <v>
                      <xsl:value-of select="@c5b"/>
                    </v>
                  </c>
                  <xsl:call-template name="plusFilled"/>
                  <c s="19">
                    <v>
                      <xsl:value-of select="@c5p"/>
                    </v>
                  </c>
                  <c s="17">
                    <v>
                      <xsl:value-of select="@c6b"/>
                    </v>
                  </c>
                  <xsl:call-template name="plusFilled"/>
                  <c s="19">
                    <v>
                      <xsl:value-of select="@c6p"/>
                    </v>
                  </c>
                  <c s="16">
                    <v>
                      <xsl:value-of select="@c1b+@c2b+@c3b+@c4b+@c5b+@c6b"/>
                    </v>
                  </c>
                  <c s="16">
                    <v>
                      <xsl:value-of select="@c1p+@c2p+@c3p+@c4p+@c5p+@c6p"/>
                    </v>
                  </c>
                  <c s="16">
                    <v>
                      <xsl:value-of select="@foreigner"/>
                    </v>
                  </c>
                  <c s="16">
                    <v>
                      <xsl:value-of select="@c1b+@c1p+@c2b+@c2p+@c3b+@c3p+@c4b+@c4p+@c5b+@c5p+@c6b+@c6p"/>
                    </v>
                  </c>
                </row>
              </xsl:for-each>
              <row>
                <xsl:variable name="r" select="3+count(Root/Inst/Spec/Edprog)+count(Root/Inst)"/>
                <c t="inlineStr" s="17">
                  <is>
                    <t>Итого по университету:</t>
                  </is>
                </c>
                <c s="18"/>
                <c s="18"/>
                <c s="19"/>

                <c s="17">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'E'"/>
                  </xsl:call-template>
                </c>
                <xsl:call-template name="plusFilled"/>
                <c s="19">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'G'"/>
                  </xsl:call-template>
                </c>
                <c s="17">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'H'"/>
                  </xsl:call-template>
                </c>
                <xsl:call-template name="plusFilled"/>
                <c s="19">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'J'"/>
                  </xsl:call-template>
                </c>
                <c s="17">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'K'"/>
                  </xsl:call-template>
                </c>
                <xsl:call-template name="plusFilled"/>
                <c s="19">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'M'"/>
                  </xsl:call-template>
                </c>
                <c s="17">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'N'"/>
                  </xsl:call-template>
                </c>
                <xsl:call-template name="plusFilled"/>
                <c s="19">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'P'"/>
                  </xsl:call-template>
                </c>
                <c s="17">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'Q'"/>
                  </xsl:call-template>
                </c>
                <xsl:call-template name="plusFilled"/>
                <c s="19">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'S'"/>
                  </xsl:call-template>
                </c>
                <c s="17">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'T'"/>
                  </xsl:call-template>
                </c>
                <xsl:call-template name="plusFilled"/>
                <c s="19">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'V'"/>
                  </xsl:call-template>
                </c>
                <c s="16">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'W'"/>
                  </xsl:call-template>
                </c>
                <c s="16">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'X'"/>
                  </xsl:call-template>
                </c>
                <c s="16">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'Y'"/>
                  </xsl:call-template>
                </c>
                <c s="16">
                  <xsl:call-template name="univer">
                    <xsl:with-param name="col" select="'Z'"/>
                  </xsl:call-template>
                </c>
              </row>
            </sheetData>

            <mergeCells>
              <mergeCell ref="A2:A3"/>
              <mergeCell ref="B2:B3"/>
              <mergeCell ref="C2:C3"/>
              <mergeCell ref="D2:D3"/>
              <mergeCell ref="E2:V2"/>
              <mergeCell ref="W2:W3"/>
              <mergeCell ref="X2:X3"/>
              <mergeCell ref="Y2:Y3"/>
              <mergeCell ref="Z2:Z3"/>
              <mergeCell ref="A1:Z1"/>
              <xsl:for-each select="Root/Inst">
                <xsl:variable name="r1" select="$headrows+1+count(preceding::Spec/Edprog)
                          +count(preceding::Spec[count(Edprog)>1])
                          +count(preceding::Inst)
                          +count(preceding::CodeQual)
                          "/>
                <xsl:variable name="r2" select="$r1 -1
                          +count(Spec/Edprog)
                          +count(Spec[count(Edprog)>1])
                          +count(CodeQuals/CodeQual)"/>
                <mergeCell>
                  <xsl:attribute name="ref">
                    <xsl:text>A</xsl:text>
                    <xsl:value-of select="$r1"/>
                    <xsl:text>:A</xsl:text>
                    <xsl:value-of select="1+$r2"/>
                  </xsl:attribute>
                </mergeCell>
              </xsl:for-each>
              <xsl:for-each select="Root/Inst">
                <!--считает кол-во спецев у которых эдпрогов больше 1-->
                <xsl:variable name="specItog" select="count(Spec[count(Edprog)>1])"/>
                <xsl:for-each select="Spec[count(Edprog)>1]">
                  <xsl:variable name="r1" select="$headrows+1
                            +count(preceding::Spec/Edprog)
                            +count(preceding::Spec[count(Edprog)>1])
                            +count(preceding::Inst)
                            +count(preceding::CodeQual)"/>
                  <xsl:variable name="r2" select="$r1 -1+count(Edprog)+count(CodeQuals/CodeQual)"/>
                  <mergeCell>
                    <xsl:attribute name="ref">
                      <xsl:text>B</xsl:text>
                      <xsl:value-of select="$r1"/>
                      <xsl:text>:B</xsl:text>
                      <xsl:value-of select="$r2+1"/>
                    </xsl:attribute>
                  </mergeCell>
                </xsl:for-each>
              </xsl:for-each>
            </mergeCells>
            <pageMargins left="0.2" right="0.2" top="0.2" bottom="0.2" header="0.2" footer="0.2"/>
            <pageSetup paperSize="9" orientation="landscape"/>
          </worksheet>
        </pkg:xmlData>
      </pkg:part>

      <pkg:part pkg:name="/xl/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml">
        <pkg:xmlData>
          <styleSheet >
            <fonts>
              <font>
                <!--0-->
                <sz val="10" />
                <name val="Arial" />
                <charset val="204" />
              </font>
              <font>
                <!--1-->
                <b />
                <sz val="10" />
                <name val="Arial" />
                <charset val="204" />
              </font>
            </fonts>
            <fills>
              <fill>
                <patternFill patternType="none" />
              </fill>
              <fill>
                <patternFill patternType="gray125" />
              </fill>
              <fill>
                <patternFill patternType="solid">
                  <!--2-->
                  <fgColor rgb="FFFFCCFF"/>
                  <bgColor indexed="64"/>
                </patternFill>
              </fill>
              <fill>
                <patternFill patternType="solid">
                  <!--3-->
                  <fgColor rgb="FFFF99FF"/>
                  <bgColor indexed="64"/>
                </patternFill>
              </fill>
              <fill>
                <patternFill patternType="solid">
                  <!--4-->
                  <fgColor rgb="FFFF66FF"/>
                  <bgColor indexed="64"/>
                </patternFill>
              </fill>
              <fill>
                <patternFill patternType="solid">
                  <!--5-->
                  <fgColor rgb="FFFFE5FF"/>
                  <bgColor indexed="64"/>
                </patternFill>
              </fill>
            </fills>
            <borders>
              <border>
                <!--0-->
                <left />
                <right />
                <top />
                <bottom />
                <diagonal />
              </border>
              <border>
                <!--1-->
                <left style="thin">
                  <color auto="1"/>
                </left>
                <right style="thin">
                  <color auto="1"/>
                </right>
                <top style="thin">
                  <color auto="1"/>
                </top>
                <bottom/>
                <diagonal/>
              </border>
              <border>
                <!--2-->
                <!--b-->
                <left style="thin">
                  <color auto="1"/>
                </left>
                <right/>
                <top style="thin">
                  <color auto="1"/>
                </top>
                <bottom style="thin">
                  <color auto="1"/>
                </bottom>
              </border>
              <border>
                <!--3-->
                <!--p-->
                <left/>
                <right style="thin">
                  <color auto="1"/>
                </right>
                <top style="thin">
                  <color auto="1"/>
                </top>
                <bottom style="thin">
                  <color auto="1"/>
                </bottom>
              </border>
              <border>
                <!--4-->
                <left/>
                <right/>
                <top style="thin">
                  <color auto="1"/>
                </top>
                <bottom style="thin">
                  <color auto="1"/>
                </bottom>
              </border>
              <border>
                <!--5-->
                <left style="thin">
                  <color auto="1"/>
                </left>
                <right style="thin">
                  <color auto="1"/>
                </right>
                <top style="thin">
                  <color auto="1"/>
                </top>
                <bottom style="thin">
                  <color auto="1"/>
                </bottom>
              </border>
              <border>
                <!--6-->
                <left style="thin">
                  <color auto="1"/>
                </left>
                <right style="thin">
                  <color auto="1"/>
                </right>
                <top/>
                <bottom/>
              </border>
              <border>
                <!--7-->
                <left style="thin">
                  <color auto="1"/>
                </left>
                <right style="thin">
                  <color auto="1"/>
                </right>
                <top/>
                <bottom style="thin">
                  <color auto="1"/>
                </bottom>
              </border>

            </borders>
            <cellStyleXfs count="1">
              <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
            </cellStyleXfs>
            <cellXfs>
              <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
              <!--0-->
              <xf numFmtId="0" fontId="1" fillId="0" borderId="2" xfId="0" applyFont="1" applyBorder="1"/>
              <!--1-->
              <xf numFmtId="0" fontId="0" fillId="0" borderId="5" xfId="0" applyAlignment="1">
                <!--2-->
                <alignment wrapText="1"/>
              </xf>
              <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1" applyBorder="1">
                <!--3-->
                <alignment horizontal="center" vertical="center"/>
              </xf>
              <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
              <!--4-->
              <xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0" applyBorder="1"/>
              <!--5-->
              <xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1"/>
              <!--6-->
              <xf numFmtId="0" fontId="0" fillId="0" borderId="5" xfId="0" applyBorder="1"/>
              <!--7-->
              <xf numFmtId="0" fontId="0" fillId="0" borderId="6" xfId="0" applyBorder="1"/>
              <!--8-->
              <xf numFmtId="0" fontId="1" fillId="0" borderId="3" xfId="0" applyFont="1" applyBorder="1"/>
              <!--9-->
              <xf numFmtId="0" fontId="0" fillId="0" borderId="7" xfId="0" applyBorder="1"/>
              <!--10-->
              <xf numFmtId="0" fontId="0" fillId="2" borderId="5" xfId="0" applyBorder="1"/>
              <!--11-->
              <xf numFmtId="0" fontId="1" fillId="3" borderId="5" xfId="0" applyBorder="1"/>
              <!--12-->
              <xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyAlignment="1" applyBorder="1">
                <!--13-->
                <alignment horizontal="center" vertical="center" wrapText="1"/>
              </xf>
              <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1">
                <!--14-->
                <alignment horizontal="center"/>
              </xf>
              <xf numFmtId="0" fontId="1" fillId="0" borderId="4" xfId="0" applyBorder="1" applyAlignment="1">
                <!--15-->
                <alignment horizontal="center"/>
              </xf>
              <xf numFmtId="0" fontId="1" fillId="4" borderId="5" xfId="0" applyBorder="1"/>
              <!--16-->
              <xf numFmtId="0" fontId="1" fillId="4" borderId="2" xfId="0" applyFont="1" applyBorder="1"/>
              <!--17-->
              <xf numFmtId="0" fontId="0" fillId="4" borderId="4" xfId="0" applyBorder="1"/>
              <!--18-->
              <xf numFmtId="0" fontId="1" fillId="4" borderId="3" xfId="0" applyFont="1" applyBorder="1"/>
              <!--19-->
              <xf numFmtId="0" fontId="0" fillId="3" borderId="2" xfId="0" applyBorder="1"/>
              <!--20-->
              <xf numFmtId="0" fontId="0" fillId="3" borderId="3" xfId="0" applyBorder="1"/>
              <!--21-->
              <xf numFmtId="0" fontId="0" fillId="3" borderId="4" xfId="0" applyBorder="1"/>
              <!--22-->
              <xf numFmtId="0" fontId="0" fillId="3" borderId="5" xfId="0" applyBorder="1"/>
              <!--23-->
              <xf numFmtId="0" fontId="0" fillId="5" borderId="5" xfId="0" applyAlignment="1" applyBorder="1" >
                <alignment wrapText="1"/>
              </xf>
              <!--24-->
              <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyBorder="1">
                <!--25-->
                <alignment horizontal="center" vertical="center" wrapText="1"/>
              </xf>
            </cellXfs>
            <cellStyles count="1">
              <cellStyle name="Normal" xfId="0" builtinId="0" />
            </cellStyles>
            <dxfs count="0" />
            <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16" />
          </styleSheet>
        </pkg:xmlData>
      </pkg:part>

      <pkg:part pkg:name="/xl/workbook.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml">
        <pkg:xmlData>
          <workbook  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheets>
              <sheet name="List1" sheetId="1" r:id="rId1" />
            </sheets>
          </workbook>
        </pkg:xmlData>
      </pkg:part>

      <pkg:part pkg:name="/xl/_rels/workbook.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" Id="rId1" />
            <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" Id="rId5" />
          </Relationships>
        </pkg:xmlData>
      </pkg:part>

      <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" Id="rId1" />
          </Relationships>
        </pkg:xmlData>
      </pkg:part>
    </pkg:package>
  </xsl:template>



  <xsl:template match="@* | node()">
    <xsl:message terminate="yes">
      Hell
    </xsl:message>
  </xsl:template>
</xsl:stylesheet>
