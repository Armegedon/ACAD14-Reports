<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">

<!--Description:Points in Engineering and Survey - Exchange (EAS-E) format.
This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
<!--CreatedBy:Autodesk Inc. -->
<!--DateCreated:06/15/2002 -->
<!--LastModifiedBy:Autodesk Inc. -->
<!--DateModified:09/27/2002 -->
<!--OutputExtension:eas -->
<xsl:output method="text" media-type="iso-8859-1" encoding="us-ascii"/>

<!-- ==== JavaScript Includes ==== -->
<xsl:include href="LandXMLUtils_JScript.xsl"/>
<xsl:include href="General_Formating_JScript.xsl"/>
<xsl:include href="Number_Formatting.xsl"/>
<xsl:include href="Cogo_Point.xsl"/>
<xsl:include href="header.xsl"/>
<!-- ============================= -->

<xsl:param name="SourceLinearUnit" select="//lx:Units/*/@linearUnit"/>

<xsl:template match="/">
<xsl:call-template name="SetGeneralFormatParameters"/>[EAS-E Version] = "1.0"
[Vendor/Software Version] = "<xsl:value-of select="//lx:Application/@manufacturer"/> - <xsl:value-of select="//lx:Application/@name"/> - <xsl:value-of select="//lx:Application/@version"/>"
[File Creation Date] = "<xsl:value-of select="$DateTime"/>"
[Project Name] = "<xsl:value-of select="//lx:Project/@name"/>"
[Units] = <xsl:text/>
	<xsl:choose>
		<xsl:when test="$Point.Coordinate.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceLinearUnit='foot'">
					<xsl:text/>"imperial"<xsl:text/>
				</xsl:when>
				<xsl:when test="$SourceLinearUnit='meter'">
					<xsl:text/>"metric"<xsl:text/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:text/>""<xsl:text/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Point.Coordinate.unit='foot'">
			<xsl:text/>"imperial"<xsl:text/>
		</xsl:when>
		<xsl:when test="$Point.Coordinate.unit='meter'">
			<xsl:text/>"metric"<xsl:text/>
		</xsl:when>
		<xsl:otherwise>
			<xsl:text/>""<xsl:text/>
		</xsl:otherwise>
	</xsl:choose>
[Coordinate System] = "<xsl:value-of select="//lx:CoordinateSystem/@name"/>"
[Scale Factor] = "1.0"
[Datum] = ""
[Comments] = "<xsl:value-of select="//lx:Project/@desc"/>"
#
[Type of Data] = "coordinate geometry"
[Comments] = "<xsl:value-of select="//lx:CgPoints/@name"/>"
<xsl:for-each select="//lx:CgPoint">
	<xsl:apply-templates select="." mode="set"/>
	<xsl:variable name="Northing" select="landUtils:GetCogoPointNorthing()"/>
	<xsl:variable name="Easting" select="landUtils:GetCogoPointEasting()"/>
	<xsl:variable name="Elevation" select="landUtils:GetCogoPointElevation()"/>
	<xsl:text/>[Point], <xsl:text/>
	<xsl:text/>"<xsl:value-of select="@name"/>", <xsl:text/>
	<xsl:text/><xsl:value-of select="landUtils:FormatNumber(string($Northing), string($SourceLinearUnit), string($Point.Coordinate.unit), string($Point.Coordinate.precision), string($Point.Coordinate.rounding))"/><xsl:text disable-output-escaping="yes"> </xsl:text>
	<xsl:text/><xsl:value-of select="landUtils:FormatNumber(string($Easting), string($SourceLinearUnit), string($Point.Coordinate.unit), string($Point.Coordinate.precision), string($Point.Coordinate.rounding))"/><xsl:text disable-output-escaping="yes"> </xsl:text>
	<!-- Elevation will use Coordinate unit setting, as mixing of metric and imperial is not allowed in this format -->
	<xsl:text/><xsl:value-of select="landUtils:FormatNumber(string($Elevation), string($SourceLinearUnit), string($Point.Coordinate.unit), string($Point.Elevation.precision), string($Point.Elevation.rounding))"/>, <xsl:text/>
	<xsl:text/>"<xsl:value-of select="@code"/>", <xsl:text/>
	<xsl:text/>"<xsl:value-of select="@desc"/>"<xsl:text/>
	<xsl:text>&#xa;</xsl:text>
</xsl:for-each>#
</xsl:template>
</xsl:stylesheet>