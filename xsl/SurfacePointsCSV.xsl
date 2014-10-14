<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">
<!--Description:Surface Point data in comma delimited format, suitable for spreadsheet import.
This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
<!--CreatedBy:Autodesk Inc. -->
<!--DateCreated:06/15/2002 -->
<!--LastModifiedBy:Autodesk Inc. -->
<!--DateModified:10/03/2002 -->
<!--OutputExtension:csv -->
<xsl:output method="text" media-type="iso-8859-1" encoding="us-ascii"/>

<!-- ==== JavaScript Includes ==== -->
<xsl:include href="LandXMLUtils_JScript.xsl"/>
<xsl:include href="General_Formating_JScript.xsl"/>
<xsl:include href="Number_Formatting.xsl"/>
<xsl:include href="Surface_Layout.xsl"/>
<!-- ============================= -->

<xsl:param name="SourceLinearUnit" select="//lx:Units/*/@linearUnit"/>

<xsl:template match="/">
	<xsl:call-template name="SetGeneralFormatParameters"/>
	<xsl:for-each select="//lx:Surface">
		<xsl:text/>"Surface name: <xsl:value-of select="@name"/>"&#xa;<xsl:text/>
		<xsl:text>"Point ID",</xsl:text>
		<xsl:choose>
			<xsl:when test="$Surface.Point_Coordinates.unit='default'">
				<xsl:choose>
					<xsl:when test="$SourceLinearUnit='foot'">
						<xsl:text>"Northing (ft)","Easting (ft)",</xsl:text>
					</xsl:when>
					<xsl:when test="$SourceLinearUnit='meter'">
						<xsl:text>"Northing (m)","Easting (m)",</xsl:text>
					</xsl:when>
				</xsl:choose>
			</xsl:when>
			<xsl:when test="$Surface.Point_Coordinates.unit='foot'">
				<xsl:text>"Northing (ft)","Easting (ft)",</xsl:text>
			</xsl:when>
			<xsl:when test="$Surface.Point_Coordinates.unit='meter'">
				<xsl:text>"Northing (m)","Easting (m)",</xsl:text>
			</xsl:when>
		</xsl:choose>
		<xsl:choose>
			<xsl:when test="$Surface.Elevations.unit='default'">
				<xsl:choose>
					<xsl:when test="$SourceLinearUnit='foot'">
						<xsl:text>"Elevation (ft)"&#xa;</xsl:text>
					</xsl:when>
					<xsl:when test="$SourceLinearUnit='meter'">
						<xsl:text>"Elevation (m)"&#xa;</xsl:text>
					</xsl:when>
				</xsl:choose>
			</xsl:when>
			<xsl:when test="$Surface.Elevations.unit='foot'">
				<xsl:text>"Elevation (ft)"&#xa;</xsl:text>
			</xsl:when>
			<xsl:when test="$Surface.Elevations.unit='meter'">
				<xsl:text>"Elevation (m)"&#xa;</xsl:text>
			</xsl:when>
		</xsl:choose>
		<xsl:for-each select=".//lx:P">
			<xsl:variable name="Northing" select="landUtils:getPointNorthing(.)"/>
			<xsl:variable name="Easting" select="landUtils:getPointEasting(.)"/>
			<xsl:variable name="Elevation" select="landUtils:getPointElevation(.)"/>
			<xsl:text/>"<xsl:value-of select="@id"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($Northing), string($SourceLinearUnit), string($Surface.Point_Coordinates.unit), string($Surface.Point_Coordinates.precision), string($Surface.Point_Coordinates.rounding))"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($Easting), string($SourceLinearUnit), string($Surface.Point_Coordinates.unit), string($Surface.Point_Coordinates.precision), string($Surface.Point_Coordinates.rounding))"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($Elevation), string($SourceLinearUnit), string($Surface.Elevations.unit), string($Surface.Elevations.precision), string($Surface.Elevations.rounding))"/>"<xsl:text/>
			<xsl:text>&#xa;</xsl:text>
		</xsl:for-each>
	</xsl:for-each>
</xsl:template>

</xsl:stylesheet>