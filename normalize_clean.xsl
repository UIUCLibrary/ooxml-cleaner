<?xml version="1.0"?>
<xsl:transform xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  version="1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <!-- Normalizes content in an MS Word OOX document. -->

  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="@w:rsidP | @w:rsidR | @w:rsidRDefault"/>

  <xsl:template
    match="w:r[(count(*) = 1 and count(w:t) = 1)]">
    <!-- If we are the same as our preceding sibling, we’ll be handled
         by it. -->
    <xsl:if
      test="not(preceding-sibling::*[1]
                                    [self::w:r]
                                    [count(*) = 1 and
                                     count(w:t) = 1])">
      <xsl:copy>
        <xsl:apply-templates select="@*"/>
        <w:t xml:space="preserve">
          <xsl:apply-templates
            select="."
            mode="combine-simple-content"/>
        </w:t>
      </xsl:copy>
    </xsl:if>
  </xsl:template>

  <xsl:template
    match="w:r[(count(*) = 2 and count(w:rPr) = 1 and
               count(w:t) = 1)]">
    <xsl:variable name="lang" select="w:rPr/w:lang/@w:val"/>
    <xsl:variable name="styleid" select="w:rPr/w:rStyle/@w:val"/>
    <!-- If we are the same as our preceding sibling, we’ll be handled
         by it. -->
    <xsl:if
      test="not(preceding-sibling::*[1]
                                    [self::w:r]
                                    [w:rPr/w:rStyle/@w:val = $styleid]
                                    [(not(w:rPr/w:lang) and
                                      string($lang) = '') or
                                     w:rPr/w:lang/@w:val = $lang]
                                    [count(*) = 2 and count(w:rPr) = 1
                                     and count(w:t) = 1])">
      <xsl:copy>
        <xsl:apply-templates select="@*"/>
        <xsl:apply-templates select="w:rPr"/>
        <w:t xml:space="preserve">
          <xsl:apply-templates select="w:t/@*"/>
          <xsl:apply-templates
            select="."
            mode="combine-styled-content"/>
        </w:t>
      </xsl:copy>
    </xsl:if>
  </xsl:template>

  <xsl:template match="w:r" mode="combine-simple-content">
    <xsl:if test="count(*) = 1 and count(w:t) = 1">
      <xsl:apply-templates select="w:t/node()"/>
      <xsl:apply-templates
        select="following-sibling::*[1]
                                    [self::w:r]
                                    [count(*) = 1 and count(w:t) = 1]"
        mode="combine-simple-content"/>
    </xsl:if>
  </xsl:template>

  <xsl:template match="w:r" mode="combine-styled-content">
    <xsl:variable name="lang" select="w:rPr/w:lang/@w:val"/>
    <xsl:variable name="styleid" select="w:rPr/w:rStyle/@w:val"/>
    <xsl:if
      test="count(*) = 2 and count(w:rPr) = 1 and count(w:t) = 1">
      <xsl:apply-templates select="w:t/node()"/>
      <xsl:apply-templates
        select="following-sibling::*[1]
                                    [self::w:r]
                                    [w:rPr/w:rStyle/@w:val = $styleid]
                                    [(not(w:rPr/w:lang) and
                                      string($lang) = '') or
                                     w:rPr/w:lang/@w:val = $lang]
                                    [count(*) = 2 and count(w:rPr) = 1
                                     and count(w:t) = 1]"
        mode="combine-styled-content"/>
    </xsl:if>
  </xsl:template>

</xsl:transform>
