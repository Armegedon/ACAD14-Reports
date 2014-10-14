<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">
	
<!--Description:Section data in comma delimited format, suitable for spreadsheet import.
This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
<!--CreatedBy:Autodesk Inc. -->
<!--DateCreated:09/06/2002 -->
<!--LastModifiedBy:Autodesk Inc. -->
<!--DateModified:09/11/2002 -->
<!--OutputExtension:csv -->

<xsl:output method="text" media-type="iso-8859-1" encoding="us-ascii"/>

<!-- ==== JavaScript Includes ==== -->
<xsl:include href="LandXMLUtils_JScript.xsl"/>
<!-- ============================= -->

<xsl:param name="SourceLinearUnit" select="//lx:Units/*/@linearUnit"/>

<xsl:template match="/">
<xsl:choose>
	<xsl:when test="$SourceLinearUnit='foot'">
		<xsl:text>"Imperial Units"&#xa;</xsl:text>
	</xsl:when>
	<xsl:when test="$SourceLinearUnit='meter'">
		<xsl:text>"Metric Units"&#xa;</xsl:text>
	</xsl:when>
</xsl:choose>
<xsl:for-each select="//lx:Alignment">
	<xsl:variable name="AlignmentName" select="@name"/>
	<xsl:for-each select="./lx:CrossSects">
		<xsl:text/>"Section: <xsl:value-of select="@name"/>"&#xa;<xsl:text/>
		<xsl:text/>"Station","area cut","centroid cut","volume cut","area fill","centroid fill","volume fill"&#xa;<xsl:text/>
		<xsl:for-each select="./lx:CrossSect">
			<xsl:text/>"<xsl:value-of select="@sta"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="@areaCut"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="@centroidCut"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="@volumeCut"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="@areaFill"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="@centroidFill"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="@volumeFill"/>"<xsl:text/>
			<xsl:text>&#xa;</xsl:text>
		</xsl:for-each>
	</xsl:for-each>
</xsl:for-each>

</xsl:template>

</xsl:stylesheet>