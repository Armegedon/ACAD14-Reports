<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">

<!--Description:Metes and Bounds

This form lists the direction and distance or tangents or curve data computed from a begin point expressed as NE.  

This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
<!--CreatedBy:Autodesk Inc. -->
<!--DateCreated:06/15/2002 -->
<!--LastModifiedBy:Autodesk Inc. -->
<!--DateModified:10/9/2002 -->
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

<xsl:template match="/">
	<html>
	<head>
		<title>Metes and Bounds Report for Project <xsl:value-of select="//lx:Project/@name"/></title>
	</head>
	<body>
	<div style="width:7in">
	<xsl:call-template name="AutoHeader">
		<xsl:with-param name="ReportTitle">Metes and Bounds Report</xsl:with-param>
		<xsl:with-param name="ReportDesc"><xsl:value-of select="//lx:Project/@name"/></xsl:with-param>
	</xsl:call-template>
	<xsl:apply-templates select="//lx:Parcel"/>
	</div>
	</body>
	</html>
</xsl:template>

<xsl:template match="lx:Parcel">
	<xsl:apply-templates select="." mode="set"/>
	<h2>Metes and Bounds description of parcel <xsl:value-of select="@name"/></h2>
	<xsl:apply-templates select="./lx:CoordGeom"/>
	to the point of beginning.<br/><br/><hr/>
</xsl:template>

<xsl:template match="lx:CoordGeom">
	<xsl:variable name="ComNorthing" select="landUtils:GetCommencePointNorthing()"/>
	<xsl:variable name="ComEasting" select="landUtils:GetCommencePointEasting()"/>
	Beginning at a point whose Northing is <xsl:value-of select="landUtils:FormatNumber(string($ComNorthing), string($SourceLinearUnit), string($Parcel.Point_Coordinates.unit), string($Parcel.Point_Coordinates.precision), string($Parcel.Point_Coordinates.rounding))"/>
	and whose Easting is <xsl:value-of select="landUtils:FormatNumber(string($ComEasting), string($SourceLinearUnit), string($Parcel.Point_Coordinates.unit), string($Parcel.Point_Coordinates.precision), string($Parcel.Point_Coordinates.rounding))"/>
	<xsl:for-each select="./node()">
		<xsl:apply-templates select="."/>
	</xsl:for-each>
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