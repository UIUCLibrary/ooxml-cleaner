<?xml version="1.0"?>
<xsl:transform xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  version="1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <!-- Reports on styles present and used in an MS Word OOX document. -->

  <xsl:output method="text" encoding="UTF-8"/>

  <xsl:variable name="docbody" select="document('document.xml', /)"/>
  <xsl:variable name="endnotes" select="document('endnotes.xml', /)"/>
  <xsl:variable name="footnotes" select="document('footnotes.xml', /)"/>

  <xsl:template match="/">
    <xsl:text>PARAGRAPH STYLES&#xa;&#xa;</xsl:text>
    <xsl:apply-templates select="//w:style[@w:type='paragraph']">
      <xsl:sort select="@w:styleId"/>
    </xsl:apply-templates>
    <xsl:text>&#xa;CHARACTER STYLES&#xa;&#xa;</xsl:text>
    <xsl:apply-templates select="//w:style[@w:type='character']">
      <xsl:sort select="@w:styleId"/>
    </xsl:apply-templates>
    <xsl:text>&#xa;KEY&#xa;</xsl:text>
    <xsl:text>&#xa;+ style is used directly</xsl:text>
    <xsl:text>&#xa;d a derived style is used</xsl:text>
    <xsl:text>&#xa;l a linked style is used</xsl:text>
    <xsl:text>&#xa;• style is not used</xsl:text>
  </xsl:template>

  <xsl:template match="w:style">
    <xsl:variable name="styleid" select="@w:styleId"/>
    <xsl:variable name="referrers"
      select="//w:style[w:basedOn/@w:val=$styleid or
                        w:link/@w:val=$styleid]"/>
    <xsl:apply-templates select="." mode="usetest"/>
    <xsl:text> </xsl:text>
    <xsl:value-of select="$styleid"/>
    <xsl:if test="w:name/@w:val != $styleid">
      <xsl:text> (</xsl:text>
      <xsl:value-of select="w:name/@w:val"/>
      <xsl:text>)</xsl:text>
    </xsl:if>
    <xsl:text>&#xa;</xsl:text>
  </xsl:template>

  <xsl:template match="w:style" mode="usetest">
    <xsl:variable name="styleid" select="@w:styleId"/>
    <xsl:variable name="derived"
      select="//w:style[w:basedOn/@w:val=$styleid]"/>
    <xsl:variable name="links"
      select="//w:style[w:link/@w:val=$styleid]"/>
    <xsl:choose>
      <xsl:when
        test="$docbody//w:pPr/w:pStyle/@w:val=$styleid or
              $docbody//w:rPr/w:rStyle/@w:val=$styleid or
              $endnotes//w:pPr/w:pStyle/@w:val=$styleid or
              $endnotes//w:rPr/w:rStyle/@w:val=$styleid or
              $footnotes//w:pPr/w:pStyle/@w:val=$styleid or
              $footnotes//w:rPr/w:rStyle/@w:val=$styleid">
        <xsl:text>+</xsl:text>
      </xsl:when>
      <xsl:when
        test="$derived and
              $docbody//w:pPr/w:pStyle/@w:val=$derived/./@w:styleId or
              $docbody//w:rPr/w:rStyle/@w:val=$derived/./@w:styleId or
              $endnotes//w:pPr/w:pStyle/@w:val=$derived/./@w:styleId or
              $endnotes//w:rPr/w:rStyle/@w:val=$derived/./@w:styleId or
              $footnotes//w:pPr/w:pStyle/@w:val=$derived/./@w:styleId or
              $footnotes//w:rPr/w:rStyle/@w:val=$derived/./@w:styleId">
        <xsl:text>d</xsl:text>
      </xsl:when>
      <xsl:when
        test="@w:type='character' and
              $docbody//w:pPr/w:pStyle/@w:val=$links/./@w:styleId or
              $docbody//w:rPr/w:rStyle/@w:val=$links/./@w:styleId or
              $endnotes//w:pPr/w:pStyle/@w:val=$links/./@w:styleId or
              $endnotes//w:rPr/w:rStyle/@w:val=$links/./@w:styleId or
              $footnotes//w:pPr/w:pStyle/@w:val=$links/./@w:styleId or
              $footnotes//w:rPr/w:rStyle/@w:val=$links/./@w:styleId">
        <xsl:text>l</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:text>·</xsl:text>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

</xsl:transform>