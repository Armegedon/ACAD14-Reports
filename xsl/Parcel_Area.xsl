<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">

<!--Description:Area Report

This form provides area values for the selected parcel(s).  

This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
<!--CreatedBy:Autodesk Inc. -->
<!--DateCreated:06/15/2002 -->
<!--LastModifiedBy:Autodesk Inc. -->
<!--DateModified:09/18/2002 -->
<!--OutputExtension:html -->
<xsl:output method="html" encoding="UTF-8"/>

<!-- =========== JavaScript Includes ==== -->
<xsl:include href="LandXMLUtils_JScript.xsl"/>
<xsl:include href="General_Formating_JScript.xsl"/>
<xsl:include href="Plan_Comp_JScript.xsl"/>
<xsl:include href="Conversion_JScript.xsl"/>
<xsl:include href="Number_Formatting.xsl"/>
<xsl:include href="Parcel_Layout.xsl"/>
<xsl:include href="header.xsl"/>

<xsl:param name="SourceLinearUnit" select="//lx:Units/*/@linearUnit"/>
<xsl:param name="SourceAreaUnit" select="//lx:Units/*/@areaUnit"/>

<xsl:template match="/">
<html>
	<head>
		<title>Parcel Area Report for <xsl:value-of select="//lx:Project/@name"/></title>
	</head>
	<body>
	<div style="width:7in">
	<xsl:call-template name="AutoHeader">
		<xsl:with-param name="ReportTitle">Parcel Area Report</xsl:with-param>
		<xsl:with-param name="ReportDesc"><xsl:value-of select="//lx:Project/@name"/></xsl:with-param>
	</xsl:call-template>
	<br/>
	<table bordercolor="black" border="1" cellspacing="0" cellpadding="4" width="95%" align="center">
		<tr>
			<th>Parcel Name</th>
			<xsl:choose>
				<xsl:when test="$Parcel.2D_Area.unit='default'">
					<xsl:choose>
						<xsl:when test="$SourceAreaUnit='squareFoot'">
							<th>Square Feet</th>
							<th>Acres</th>
						</xsl:when>
						<xsl:when test="$SourceAreaUnit='squareMeter'">
							<th>Square Meters</th>
							<th>Hectares</th>
						</xsl:when>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="$Parcel.2D_Area.unit='squareFoot'">
					<th>Square Feet</th>
					<th>Acres</th>
				</xsl:when>
				<xsl:when test="$Parcel.2D_Area.unit='squareMeter'">
					<th>Square Meters</th>
					<th>Hectares</th>
				</xsl:when>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="$Parcel.2D_Perimeter.unit='default'">
					<xsl:choose>
						<xsl:when test="$SourceLinearUnit='foot'">
							<th>Perimeter (ft)</th>
						</xsl:when>
						<xsl:when test="$SourceLinearUnit='meter'">
							<th>Perimeter (m)</th>
						</xsl:when>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="$Parcel.2D_Perimeter.unit='foot'">
					<th>Perimeter (ft)</th>
				</xsl:when>
				<xsl:when test="$Parcel.2D_Perimeter.unit='meter'">
					<th>Perimeter (m)</th>
				</xsl:when>
			</xsl:choose>
		</tr>
		<xsl:for-each select="//lx:Parcel">
		<xsl:variable name="ParcelArea" select="@area"/>
		<tr>
			<td align="center"><xsl:value-of select="@name"/></td>
			<xsl:choose>
				<xsl:when test="$Parcel.2D_Area.unit='default'">
					<xsl:choose>
						<xsl:when test="$SourceAreaUnit='squareFoot'">
							<td align="center"><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareFoot', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/></td>
							<td align="center"><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'acre', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/></td>
						</xsl:when>
						<xsl:when test="$SourceAreaUnit='squareMeter'">
							<td align="center"><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareMeter', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/></td>
							<td align="center"><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'hectare', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/></td>
						</xsl:when>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="$Parcel.2D_Area.unit='squareFoot'">
					<td align="center"><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareFoot', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/></td>
					<td align="center"><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'acre', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/></td>
				</xsl:when>
				<xsl:when test="$Parcel.2D_Area.unit='squareMeter'">
					<td align="center"><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareMeter', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/></td>
					<td align="center"><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'hectare', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/></td>
				</xsl:when>
			</xsl:choose>
			<td align="center">
			<xsl:choose>
			<xsl:when test="not(./CoordGeom/*[not(@length)])">
				<xsl:variable name="ParcelPerim" select="sum(./lx:CoordGeom/*/@length)"/>
				<xsl:value-of select="landUtils:FormatNumber(string($ParcelPerim), string($SourceLinearUnit), string($Parcel.2D_Perimeter.unit), string($Parcel.2D_Perimeter.precision), string($Parcel.2D_Perimeter.rounding))"/>
			</xsl:when>
			<xsl:otherwise>&#160;</xsl:otherwise>
   			</xsl:choose>
			</td>
		</tr>
		</xsl:for-each>
	</table>
	</div>
	</body>
</html>
</xsl:template> 

</xsl:stylesheet> 