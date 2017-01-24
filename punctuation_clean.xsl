<?xml version="1.0"?>
<xsl:transform xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  version="1.0"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <!-- Normalizes punctuation in an MS Word OOX document. -->

  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <!-- Normalize dashes. -->

  <!-- Normalize whitespace. -->

  <!-- What else? -->

</xsl:transform>
