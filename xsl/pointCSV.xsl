<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">
	
<!--Description:Point data in comma delimited format, suitable for spreadsheet import.
This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
<!--CreatedBy:Autodesk Inc. -->
<!--DateCreated:06/15/2002 -->
<!--LastModifiedBy:Autodesk Inc. -->
<!--DateModified:09/27/2002 -->
<!--OutputExtension:csv -->
<xsl:output method="text" media-type="iso-8859-1" encoding="us-ascii"/>

<!-- ==== JavaScript Includes ==== -->
<xsl:include href="LandXMLUtils_JScript.xsl"/>
<xsl:include href="General_Formating_JScript.xsl"/>
<xsl:include href="Number_Formatting.xsl"/>
<xsl:include href="Cogo_Point.xsl"/>
<!-- ============================= -->
<xsl:param name="SourceLinearUnit" select="//lx:Units/*/@linearUnit"/>

<xsl:template match="/">
<xsl:text>"Point Name",</xsl:text>
<xsl:choose>
	<xsl:when test="$Point.Coordinate.unit='default'">
		<xsl:choose>
			<xsl:when test="$SourceLinearUnit='foot'">
				<xsl:text>"Northing (ft)","Easting (ft)",</xsl:text>
			</xsl:when>
			<xsl:when test="$SourceLinearUnit='meter'">
				<xsl:text>"Northing (m)","Easting (m)",</xsl:text>
			</xsl:when>
		</xsl:choose>
	</xsl:when>
	<xsl:when test="$Point.Coordinate.unit='foot'">
		<xsl:text>"Northing (ft)","Easting (ft)",</xsl:text>
	</xsl:when>
	<xsl:when test="$Point.Coordinate.unit='meter'">
		<xsl:text>"Northing (m)","Easting (m)",</xsl:text>
	</xsl:when>
</xsl:choose>
<xsl:choose>
	<xsl:when test="$Point.Elevation.unit='default'">
		<xsl:choose>
			<xsl:when test="$SourceLinearUnit='foot'">
				<xsl:text>"Elevation (ft)",</xsl:text>
			</xsl:when>
			<xsl:when test="$SourceLinearUnit='meter'">
				<xsl:text>"Elevation (m)",</xsl:text>
			</xsl:when>
		</xsl:choose>
	</xsl:when>
	<xsl:when test="$Point.Elevation.unit='foot'">
		<xsl:text>"Elevation (ft)",</xsl:text>
	</xsl:when>
	<xsl:when test="$Point.Elevation.unit='meter'">
		<xsl:text>"Elevation (m)",</xsl:text>
	</xsl:when>
</xsl:choose>
<xsl:text>"Description"&#xa;</xsl:text>
<xsl:call-template name="SetGeneralFormatParameters"/>

<xsl:for-each select="//lx:CgPoint[not(@pntRef)]">
	<xsl:apply-templates select="." mode="set"/>
	<xsl:variable name="Northing" select="landUtils:GetCogoPointNorthing()"/>
	<xsl:variable name="Easting" select="landUtils:GetCogoPointEasting()"/>
	<xsl:variable name="Elevation" select="landUtils:GetCogoPointElevation()"/>
	<xsl:text/>"<xsl:value-of select="@name"/>",<xsl:text/>
	<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($Northing), string($SourceLinearUnit), string($Point.Coordinate.unit), string($Point.Coordinate.precision), string($Point.Coordinate.rounding))"/>",<xsl:text/>
	<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($Easting), string($SourceLinearUnit), string($Point.Coordinate.unit), string($Point.Coordinate.precision), string($Point.Coordinate.rounding))"/>",<xsl:text/>
	<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($Elevation), string($SourceLinearUnit), string($Point.Elevation.unit), string($Point.Elevation.precision), string($Point.Elevation.rounding))"/>",<xsl:text/>
	<xsl:choose>
		<xsl:when test="@desc">
			<xsl:text/>"<xsl:value-of select="@desc"/>"<xsl:text/>
		</xsl:when>
		<xsl:when test="@code">
			<xsl:text/>"<xsl:value-of select="@code"/>"<xsl:text/>
		</xsl:when>
		<xsl:otherwise>
			<xsl:text/>"NA"<xsl:text/>
		</xsl:otherwise>
	</xsl:choose>
	<xsl:text>&#xa;</xsl:text>
</xsl:for-each>

</xsl:template>

</xsl:stylesheet>