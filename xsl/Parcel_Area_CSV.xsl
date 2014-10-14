<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">

<!--Description:Parcel Area in comma delimited format, suitable for spreadsheet import.
This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
<!--CreatedBy:Autodesk Inc. -->
<!--DateCreated:06/15/2002 -->
<!--LastModifiedBy:Autodesk Inc. -->
<!--DateModified:09/17/2002 -->
<!--OutputExtension:csv -->
<xsl:output method="text" media-type="iso-8859-1" encoding="us-ascii"/>

<!-- ==== JavaScript Includes ==== -->
<xsl:include href="LandXMLUtils_JScript.xsl"/>
<xsl:include href="General_Formating_JScript.xsl"/>
<xsl:include href="Number_Formatting.xsl"/>
<xsl:include href="Text_Formatting.xsl"/>
<xsl:include href="Parcel_Layout.xsl"/>
<!-- ============================= -->
<xsl:param name="SourceLinearUnit" select="//lx:Units/*/@linearUnit"/>
<xsl:param name="SourceAreaUnit" select="//lx:Units/*/@areaUnit"/>

<xsl:template match="/">
<xsl:text>"Parcel Name",</xsl:text>
	<xsl:choose>
		<xsl:when test="$Parcel.2D_Area.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceAreaUnit='squareFoot'">
					<xsl:text>"Square Feet","Acres",</xsl:text>
				</xsl:when>
				<xsl:when test="$SourceAreaUnit='squareMeter'">
					<xsl:text>"Square Meters","Hectares",</xsl:text>
				</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Parcel.2D_Area.unit='squareFoot'">
			<xsl:text>"Square Feet","Acres",</xsl:text>
		</xsl:when>
		<xsl:when test="$Parcel.2D_Area.unit='squareMeter'">
			<xsl:text>"Square Meters","Hectares",</xsl:text>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="$Parcel.2D_Perimeter.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceLinearUnit='foot'">
					<xsl:text>"Perimeter (ft)"</xsl:text>
				</xsl:when>
				<xsl:when test="$SourceLinearUnit='meter'">
					<xsl:text>"Perimeter (m)"</xsl:text>
				</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Parcel.2D_Perimeter.unit='foot'">
			<xsl:text>"Perimeter (ft)"</xsl:text>
		</xsl:when>
		<xsl:when test="$Parcel.2D_Perimeter.unit='meter'">
			<xsl:text>"Perimeter (m)"</xsl:text>
		</xsl:when>
	</xsl:choose>
<xsl:text>&#xa;</xsl:text>
<xsl:for-each select="//lx:Parcel">
	<xsl:variable name="ParcelArea" select="@area"/>
	<xsl:text/>"<xsl:value-of select="@name"/>",<xsl:text/>
	<xsl:choose>
		<xsl:when test="$Parcel.2D_Area.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceAreaUnit='squareFoot'">
					<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareFoot', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/>",<xsl:text/>
					<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'acre', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/>",<xsl:text/>
				</xsl:when>
				<xsl:when test="$SourceAreaUnit='squareMeter'">
					<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareMeter', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/>",<xsl:text/>
					<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'hectare', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/>",<xsl:text/>
				</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Parcel.2D_Area.unit='squareFoot'">
			<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareFoot', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'acre', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/>",<xsl:text/>
		</xsl:when>
		<xsl:when test="$Parcel.2D_Area.unit='squareMeter'">
			<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareMeter', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'hectare', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/>",<xsl:text/>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
	<xsl:when test="not(./lx:CoordGeom/*[not(@length)])">
		<xsl:variable name="ParcelPerim" select="sum(./lx:CoordGeom/*/@length)"/>
		<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($ParcelPerim), string($SourceLinearUnit), string($Parcel.2D_Perimeter.unit), string($Parcel.2D_Perimeter.precision), string($Parcel.2D_Perimeter.rounding))"/>"<xsl:text/>
	</xsl:when>
	<xsl:otherwise>
		<xsl:text>""</xsl:text>
	</xsl:otherwise>
	</xsl:choose>
	<xsl:text>&#xa;</xsl:text>
</xsl:for-each>
  
</xsl:template> 

</xsl:stylesheet> 