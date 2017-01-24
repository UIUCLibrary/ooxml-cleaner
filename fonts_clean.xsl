<?xml version="1.0"?>
<xsl:transform xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  version="1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <!-- Strips pointless font changes in an MS Word OOX document. -->

  <!-- Comma-separated list of font names. -->
  <xsl:param name="fontnames" select="'Cambria,Calibri'"/>

  <xsl:variable
    name="fontlist"
    select="concat(',', $fontnames, ',')"/>

  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="w:rFonts">
    <xsl:if
      test="not(contains($fontlist, concat(',', @w:ascii, ',')))">
      <xsl:copy>
        <xsl:apply-templates select="node()|@*"/>
      </xsl:copy>
    </xsl:if>
  </xsl:template>

  <xsl:template match="w:rPr">
    <xsl:if
      test="count(*) and
            not(count(*) = 1 and
                count(w:rFonts) = 1 and
                contains($fontlist,
                         concat(',', w:rFonts/@w:ascii, ',')))">
      <xsl:copy>
        <xsl:apply-templates select="node()|@*"/>
      </xsl:copy>
    </xsl:if>
  </xsl:template>

</xsl:transform>
