<?xml version="1.0"?>
<xsl:transform xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  version="1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <!-- Strips out unwanted elements in an MS Word OOX document. -->

  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <!-- Drop bookmarks and proofing error markers. -->
  <xsl:template
    match="w:bookmarkEnd | w:bookmarkStart | w:lastRenderedPageBreak |
           w:proofErr"/>

</xsl:transform>
