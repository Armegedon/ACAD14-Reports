<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">

<!--Description:Surveyor's Certificate.  This report creates a survey description of a parcel.  Choose only one parcel from the Data Summary window to create the report from (if multiple parcels are selected only the first parcel will be used).
This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
<!--CreatedBy:Autodesk Inc. -->
<!--DateCreated:09/05/2002 -->
<!--LastModifiedBy:Autodesk Inc. -->
<!--DateModified:3/14/2003 -->
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
		<title>Surveyor's Certificate for Parcel <xsl:value-of select="//lx:Parcel/@name"/></title>
	</head>
	<body>
	<div style="width:7in">
	<xsl:for-each select="//lx:Parcel">
	<xsl:apply-templates select="." mode="set"/>
	<xsl:if test="position()=1">
	<h2>Parcel <xsl:value-of select="@name"/></h2>
	<font size="+1">SURVEYOR'S CERTIFICATE</font><br/><br/>
	I, <font style="text-decoration:underline"><xsl:value-of select="$Owner.Preparer.name"/></font> Registered Land Surveyor,
	do hereby certify that I have surveyed, divided, and mapped<br/><br/>
	<hr color="black" size="1"/><br/><hr color="black" size="1"/>
	more particularly described as:<br/><br/>
	<table><tr><td>&#160;&#160;&#160;&#160;&#160;</td><td>
		<xsl:variable name="ComNorthing" select="landUtils:GetCommencePointNorthing()"/>
		<xsl:variable name="ComEasting" select="landUtils:GetCommencePointEasting()"/>
		Commencing at a point of Northing <xsl:value-of select="landUtils:FormatNumber(string($ComNorthing), string($SourceLinearUnit), string($Parcel.Point_Coordinates.unit), string($Parcel.Point_Coordinates.precision), string($Parcel.Point_Coordinates.rounding))"/>
		and Easting <xsl:value-of select="landUtils:FormatNumber(string($ComEasting), string($SourceLinearUnit), string($Parcel.Point_Coordinates.unit), string($Parcel.Point_Coordinates.precision), string($Parcel.Point_Coordinates.rounding))"/>
		<xsl:for-each select="./lx:CoordGeom">
			<xsl:apply-templates select="."/>
		</xsl:for-each>
		to the point of beginning.<br/><br/>
	</td></tr></table>
	<xsl:variable name="ParcelArea" select="@area"/>
	Said described parcel contains
	<xsl:choose>
		<xsl:when test="$Parcel.2D_Area.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceAreaUnit='squareFoot'">
					<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareFoot', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/> square feet
					(<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'acre', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/> acres),
				</xsl:when>
				<xsl:when test="$SourceAreaUnit='squareMeter'">
					<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareMeter', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/> square meters
					(<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'hectare', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/> hectares),
				</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Parcel.2D_Area.unit='squareFoot'">
					<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareFoot', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/> square feet
					(<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'acre', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/> acres),
		</xsl:when>
		<xsl:when test="$Parcel.2D_Area.unit='squareMeter'">
					<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareMeter', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/> square meters
					(<xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'hectare', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/> hectares),
		</xsl:when>
	</xsl:choose>
	more or less, subject to any and all easements, reservations, restrictions and conveyances of record.<br/>
	</xsl:if>
	</xsl:for-each>
	</div>
	</body>
	</html>
</xsl:template>

<xsl:template match="lx:Line">
	<xsl:apply-templates select="." mode="set"/>
	<xsl:variable name="angle" select="landUtils:GetLineAngle()"/>
	<xsl:variable name="LineLength" select="landUtils:GetLineLength()"/>
	;<br/>thence bearing <xsl:value-of select="landUtils:formatBearingDMS($angle)"/>
	a distance of <xsl:value-of select="landUtils:FormatNumber(string($LineLength), string($SourceLinearUnit), string($Parcel.Line_Segment_Length.unit), string($Parcel.Line_Segment_Length.precision), string($Parcel.Line_Segment_Length.rounding))"/>&#160;
	<xsl:choose>
		<xsl:when test="$Parcel.Line_Segment_Length.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceLinearUnit='foot'">feet</xsl:when>
				<xsl:when test="$SourceLinearUnit='meter'">meters</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Parcel.Line_Segment_Length.unit='foot'">feet</xsl:when>
		<xsl:when test="$Parcel.Line_Segment_Length.unit='meter'">meters</xsl:when>
	</xsl:choose>
</xsl:template>

<xsl:template match="lx:Curve">
	<xsl:apply-templates select="." mode="set"/>
	<xsl:variable name="CurveAngle" select="landUtils:GetCurveAngle()"/>
	<xsl:variable name="CurveRadius" select="landUtils:GetCurveRadius()"/>
	<xsl:variable name="ChordDir" select="landUtils:CalculateBearingDMS(./lx:Start,./lx:End)"/>
	<xsl:variable name="ChordLength" select="landUtils:GetChordLength()"/>
	;<br/>thence along a curve to the <xsl:value-of select="landUtils:GetCurveDirection()"/>,
	having a radius of <xsl:value-of select="landUtils:FormatNumber(string($CurveRadius), string($SourceLinearUnit), string($Parcel.Curve_Segment_Radius.unit), string($Parcel.Curve_Segment_Radius.precision), string($Parcel.Curve_Segment_Radius.rounding))"/>&#160;
	<xsl:choose>
		<xsl:when test="$Parcel.Curve_Segment_Radius.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceLinearUnit='foot'">feet,</xsl:when>
				<xsl:when test="$SourceLinearUnit='meter'">meters,</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Parcel.Curve_Segment_Radius.unit='foot'">feet,</xsl:when>
		<xsl:when test="$Parcel.Curve_Segment_Radius.unit='meter'">meters,</xsl:when>
	</xsl:choose>
    a delta angle of <xsl:value-of select="landUtils:FormatAngle(string($CurveAngle), string('default'), string($Parcel.Curve_Segment_Radius.unit), string($Parcel.Curve_Segment_Delta_Angle.format), string($Parcel.Curve_Segment_Delta_Angle.precision), string($Parcel.Curve_Segment_Delta_Angle.rounding))"/>,
	and whose long chord bears <xsl:value-of select="$ChordDir"/>
	a distance of <xsl:value-of select="landUtils:FormatNumber(string($ChordLength), string($SourceLinearUnit), string($Parcel.Curve_Segment_Chord_Length.unit), string($Parcel.Curve_Segment_Chord_Length.precision), string($Parcel.Curve_Segment_Chord_Length.rounding))"/>&#160;
	<xsl:choose>
		<xsl:when test="$Parcel.Curve_Segment_Chord_Length.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceLinearUnit='foot'">feet</xsl:when>
				<xsl:when test="$SourceLinearUnit='meter'">meters</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Parcel.Curve_Segment_Chord_Length.unit='foot'">feet</xsl:when>
		<xsl:when test="$Parcel.Curve_Segment_Chord_Length.unit='meter'">meters</xsl:when>
	</xsl:choose>
</xsl:template>

</xsl:stylesheet>
