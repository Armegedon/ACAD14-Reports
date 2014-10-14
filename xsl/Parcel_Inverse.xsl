<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">

<!--Description:Inverse Report

This form lists the direction and distance or tangents or curve data computed from NE coordinates of the endpoints of the parcel segments.

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
	<xsl:call-template name="SetGeneralFormatParameters"/>
	<html>
		<head>
			<title>Parcel Inverse Report for <xsl:value-of select="//lx:Project/@name"/></title>
		</head>
		<body>
		<div style="width:7in">
		<xsl:call-template name="AutoHeader">
			<xsl:with-param name="ReportTitle">Parcel Inverse Report</xsl:with-param>
			<xsl:with-param name="ReportDesc"><xsl:value-of select="//lx:Project/@name"/></xsl:with-param>
		</xsl:call-template>
		<xsl:apply-templates select="//lx:Parcel"/>
		</div>
		</body>
	</html>
</xsl:template> 

<xsl:template match="lx:Parcel">
	<table border="2" width="100%">
		<tr>
			<td colspan="3">Parcel <xsl:value-of select="@name"/></td>
		</tr>
		<xsl:apply-templates select="./lx:CoordGeom"/>
		<tr>
			<td colspan="3" bgcolor="lightgrey">Area</td>
		</tr>
		<tr>
			<td width="33%">&#160;</td>
			<xsl:variable name="ParcelArea" select="@area"/>
			<xsl:choose>
				<xsl:when test="$Parcel.2D_Area.unit='default'">
					<xsl:choose>
						<xsl:when test="$SourceAreaUnit='squareFoot'">
							<td>Square feet</td>
							<td><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareFoot', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/></td>
						</xsl:when>
						<xsl:when test="$SourceAreaUnit='squareMeter'">
							<td>Square meters</td>
							<td><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareMeter', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/></td>
						</xsl:when>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="$Parcel.2D_Area.unit='squareFoot'">
					<td>Square feet</td>
					<td><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareFoot', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/></td>
				</xsl:when>
				<xsl:when test="$Parcel.2D_Area.unit='squareMeter'">
					<td>Square meters</td>
					<td><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareMeter', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/></td>
				</xsl:when>
			</xsl:choose>
		</tr>
	</table>
	<br/>
</xsl:template>

<xsl:template match="lx:CoordGeom">
	 <xsl:for-each select="./node()">
		<xsl:apply-templates select="."/>
	</xsl:for-each>
</xsl:template>

<xsl:template match="lx:Line">
	<xsl:apply-templates select="." mode="set"/>
	<xsl:variable name="angle" select="landUtils:GetLineAngle()"/>
<tr>
	<td colspan="3" bgcolor="lightgrey">
	<xsl:choose>
		<xsl:when test="./Start/@pntRef">
			Point Number:<xsl:value-of select="./Start/@pntRef"/>
		</xsl:when>
		<xsl:otherwise>
			<xsl:variable name="StartLineNorthing" select="landUtils:GetStartLineNorthing()"/>
			<xsl:variable name="StartLineEasting" select="landUtils:GetStartLineEasting()"/>
			Point whose Northing is <xsl:value-of select="landUtils:FormatNumber(string($StartLineNorthing), string($SourceLinearUnit), string($Parcel.Point_Coordinates.unit), string($Parcel.Point_Coordinates.precision), string($Parcel.Point_Coordinates.rounding))"/>
			and whose Easting is <xsl:value-of select="landUtils:FormatNumber(string($StartLineEasting), string($SourceLinearUnit), string($Parcel.Point_Coordinates.unit), string($Parcel.Point_Coordinates.precision), string($Parcel.Point_Coordinates.rounding))"/>
		</xsl:otherwise>
	</xsl:choose>
	</td>
</tr>
<tr>	
	<td width="33%">&#160;</td>
	<xsl:variable name="Bearing" select="landUtils:formatBearingDMS($angle)"/>
	<xsl:variable name="LineLength" select="landUtils:GetLineLength()"/>
	<td>Bearing: <xsl:value-of select="$Bearing"/></td>
	<td>Length: <xsl:value-of select="landUtils:FormatNumber(string($LineLength), string($SourceLinearUnit), string($Parcel.Line_Segment_Length.unit), string($Parcel.Line_Segment_Length.precision), string($Parcel.Line_Segment_Length.rounding))"/></td>
</tr>
</xsl:template>

<xsl:template match="lx:Curve">
	<xsl:apply-templates select="." mode="set"/>
	<xsl:variable name="angle" select="landUtils:GetCurveAngle()"/>
	<tr>
		<td colspan="3" bgcolor="lightgrey">
		<xsl:choose>
			<xsl:when test="./Start/@pntRef">
				<xsl:value-of select="./Start/@pntRef"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="StartNorthing" select="landUtils:GetStartingNorthing()"/>
				<xsl:variable name="StartEasting" select="landUtils:GetStartingEasting()"/>
				Point whose Northing is <xsl:value-of select="landUtils:FormatNumber(string($StartNorthing), string($SourceLinearUnit), string($Parcel.Point_Coordinates.unit), string($Parcel.Point_Coordinates.precision), string($Parcel.Point_Coordinates.rounding))"/>
				and whose Easting is <xsl:value-of select="landUtils:FormatNumber(string($StartEasting), string($SourceLinearUnit), string($Parcel.Point_Coordinates.unit), string($Parcel.Point_Coordinates.precision), string($Parcel.Point_Coordinates.rounding))"/>
			</xsl:otherwise>
		</xsl:choose>
		</td>
	</tr>
	<tr>
		<td width="33%">&#160;</td>
		<td colspan="2" bgcolor="lightgrey">Curve</td>
	</tr>
	<tr>
		<td width="33%">&#160;</td>
		<td bgcolor="lightgrey" style="font:bold">Direction P.C. to Radius:</td>
		<xsl:variable name="PCtoRadius" select="landUtils:CalculateBearingDMS(./lx:Start, ./lx:Center)"/>
		<td><xsl:value-of select="$PCtoRadius"/></td>
	</tr>
	<tr>
		<td width="33%">&#160;</td>
		<td bgcolor="lightgrey" style="font:bold">Radius Length:</td>
		<xsl:variable name="RadiusLength" select="landUtils:GetCurveRadius()"/>
		<td><xsl:value-of select="landUtils:FormatNumber(string($RadiusLength), string($SourceLinearUnit), string($Parcel.Curve_Segment_Radius.unit), string($Parcel.Curve_Segment_Radius.precision), string($Parcel.Curve_Segment_Radius.rounding))"/></td>
	</tr>
	<tr>
		<td width="33%">&#160;</td>
		<td bgcolor="lightgrey" style="font:bold">Delta:</td>
		<!-- <xsl:param name="Parcel.Curve_Segment_Delta_Angle.format"/> -->
		<td><xsl:value-of select="landUtils:FormatAngle(string($angle), 'degrees', 'degrees', string($Parcel.Curve_Segment_Delta_Angle.format), string($Parcel.Curve_Segment_Delta_Angle.precision), string($Parcel.Curve_Segment_Delta_Angle.rounding))"/></td>
		<!-- <td><xsl:value-of select="landUtils:FormatNumber(string($angle), string($SourceLinearUnit), 'default', string($Parcel.Curve_Segment_Delta_Angle.precision), string($Parcel.Curve_Segment_Delta_Angle.rounding))"/></td> -->
	</tr>
	<tr>
		<td width="33%">&#160;</td>
		<td bgcolor="lightgrey" style="font:bold">Curve Length:</td>
		<xsl:variable name="CurveLength" select="landUtils:GetCurveLength()"/>
		<td><xsl:value-of select="landUtils:FormatNumber(string($CurveLength), string($SourceLinearUnit), string($Parcel.Curve_Segment_Length.unit), string($Parcel.Curve_Segment_Length.precision), string($Parcel.Curve_Segment_Length.rounding))"/></td>
	</tr>
	<tr>
		<td width="33%">&#160;</td>
		<td bgcolor="lightgrey" style="font:bold">Chord Length:</td>
		<xsl:variable name="ChordLength" select="landUtils:GetChordLength()"/>
		<td><xsl:value-of select="landUtils:FormatNumber(string($ChordLength), string($SourceLinearUnit), string($Parcel.Curve_Segment_Chord_Length.unit), string($Parcel.Curve_Segment_Chord_Length.precision), string($Parcel.Curve_Segment_Chord_Length.rounding))"/></td>
	</tr>
	<tr>
		<td width="33%">&#160;</td>
		<td bgcolor="lightgrey" style="font:bold">Chord Direction:</td>
		<xsl:variable name="ChordDir" select="landUtils:CalculateBearingDMS(./lx:Start, ./lx:End)"/>
		<td><xsl:value-of select="$ChordDir"/></td>
	</tr>
	<tr>
		<td width="33%">&#160;</td>
		<td bgcolor="lightgrey" style="font:bold">Direction Radius to P.T.:</td>
		<xsl:variable name="RadiustoPT" select="landUtils:CalculateBearingDMS(./lx:Center, ./lx:End)"/>
		<td><xsl:value-of select="$RadiustoPT"/></td>
	</tr>
</xsl:template>

</xsl:stylesheet>
