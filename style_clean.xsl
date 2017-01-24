<?xml version="1.0"?>
<xsl:transform xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  version="1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <!-- Removes unused styles from an MS Word OOX document. -->

  <xsl:variable name="docbody" select="document('document.xml', /)"/>
  <xsl:variable name="endnotes" select="document('endnotes.xml', /)"/>
  <xsl:variable name="footnotes" select="document('footnotes.xml', /)"/>

  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="w:style">
    <xsl:variable name="styleid" select="@w:styleId"/>
    <xsl:variable name="derived"
      select="//w:style[w:basedOn/@w:val=$styleid]"/>
    <xsl:variable name="links"
      select="//w:style[w:link/@w:val=$styleid]"/>
    <xsl:if
      test="not(@w:type='paragraph' or @w:type='character') or
            $docbody//w:pPr/w:pStyle/@w:val=$styleid or
            $docbody//w:rPr/w:rStyle/@w:val=$styleid or
            $endnotes//w:pPr/w:pStyle/@w:val=$styleid or
            $endnotes//w:rPr/w:rStyle/@w:val=$styleid or
            $footnotes//w:pPr/w:pStyle/@w:val=$styleid or
            $footnotes//w:rPr/w:rStyle/@w:val=$styleid or
            $derived and
            ($docbody//w:pPr/w:pStyle/@w:val=$derived/./@w:styleId or
             $docbody//w:rPr/w:rStyle/@w:val=$derived/./@w:styleId or
             $endnotes//w:pPr/w:pStyle/@w:val=$derived/./@w:styleId or
             $endnotes//w:rPr/w:rStyle/@w:val=$derived/./@w:styleId or
             $footnotes//w:pPr/w:pStyle/@w:val=$derived/./@w:styleId or
             $footnotes//w:rPr/w:rStyle/@w:val=$derived/./@w:styleId) or
            @w:type='character' and
            ($endnotes//w:rPr/w:rStyle/@w:val=$links/./@w:styleId or
             $footnotes//w:pPr/w:pStyle/@w:val=$links/./@w:styleId or
             $footnotes//w:rPr/w:rStyle/@w:val=$links/./@w:styleId)">
      <xsl:copy>
        <xsl:apply-templates select="node()|@*"/>
      </xsl:copy>
    </xsl:if>
  </xsl:template>

</xsl:transform>
