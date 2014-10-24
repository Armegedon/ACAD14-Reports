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
<xsl:param name="SourceAreaUnit" select="//lx:Units/*/@areaUnit"/>

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
	, TO THE POINT OF BEGINNING,<br/><br/>CONTAINING A CALCULATED AREA OF 

	
	<xsl:variable name="ParcelArea" select="@area"/>
	<xsl:choose>
		<xsl:when test="$Parcel.2D_Area.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceAreaUnit='squareFoot'">
					<xsl:text/><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareFoot', string('0'), string($Parcel.2D_Area.rounding))"/><xsl:text/>
					
				</xsl:when>
				<xsl:when test="$SourceAreaUnit='squareMeter'">
					<xsl:text/><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareMeter', string('0'), string($Parcel.2D_Area.rounding))"/><xsl:text/>
					
				</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Parcel.2D_Area.unit='squareFoot'">
			<xsl:text/><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareFoot', string('0'), string($Parcel.2D_Area.rounding))"/><xsl:text/>
			
		</xsl:when>
		<xsl:when test="$Parcel.2D_Area.unit='squareMeter'">
			<xsl:text/><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'squareMeter', string('0'), string($Parcel.2D_Area.rounding))"/><xsl:text/>
			
		</xsl:when>
	</xsl:choose>
	
	 SQUARE FEET OR

	<xsl:choose>
		<xsl:when test="$Parcel.2D_Area.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceAreaUnit='squareFoot'">
					
					<xsl:text/><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'acre', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/><xsl:text/>
				</xsl:when>
				<xsl:when test="$SourceAreaUnit='squareMeter'">
					
					<xsl:text/><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'hectare', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/><xsl:text/>
				</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Parcel.2D_Area.unit='squareFoot'">
			
			<xsl:text/><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'acre', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/><xsl:text/>
		</xsl:when>
		<xsl:when test="$Parcel.2D_Area.unit='squareMeter'">
			
			<xsl:text/><xsl:value-of select="landUtils:FormatNumber(string($ParcelArea), string($SourceAreaUnit), 'hectare', string($Parcel.2D_Area.precision), string($Parcel.2D_Area.rounding))"/><xsl:text/>
		</xsl:when>
	</xsl:choose>

	 ACRES, MORE OR LESS.<br/><br/><hr/>


</xsl:template>

<xsl:template match="lx:CoordGeom">
	<xsl:variable name="ComNorthing" select="landUtils:GetCommencePointNorthing()"/>
	<xsl:variable name="ComEasting" select="landUtils:GetCommencePointEasting()"/>
	BEGINNING AT A POINT WHOSE NORTHING IS <xsl:value-of select="landUtils:FormatNumber(string($ComNorthing), string($SourceLinearUnit), string($Parcel.Point_Coordinates.unit), string($Parcel.Point_Coordinates.precision), string($Parcel.Point_Coordinates.rounding))"/>
	AND WHOSE EASTING IS <xsl:value-of select="landUtils:FormatNumber(string($ComEasting), string($SourceLinearUnit), string($Parcel.Point_Coordinates.unit), string($Parcel.Point_Coordinates.precision), string($Parcel.Point_Coordinates.rounding))"/>
	<xsl:for-each select="./node()">
		<xsl:apply-templates select="."/>
	</xsl:for-each>
</xsl:template>

<xsl:template match="lx:Line">
	<xsl:apply-templates select="." mode="set"/>
	<xsl:variable name="angle" select="landUtils:GetLineAngle()"/>
	<xsl:variable name="LineLength" select="landUtils:GetLineLength()"/>
	;<br/>THENCE <xsl:value-of select="landUtils:formatBearingDMS($angle)"/>,
	A DISTANCE OF <xsl:value-of select="landUtils:FormatNumber(string($LineLength), string($SourceLinearUnit), string($Parcel.Line_Segment_Length.unit), string('0.00'), string($Parcel.Line_Segment_Length.rounding))"/><!-- &#160;-->
	<xsl:choose>
		<xsl:when test="$Parcel.Line_Segment_Length.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceLinearUnit='USSurveyFoot'"> FEET</xsl:when>
				<xsl:when test="$SourceLinearUnit='foot'"> FEET</xsl:when>
				<xsl:when test="$SourceLinearUnit='meter'"> METERS</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Parcel.Line_Segment_Length.unit='USSurveyFoot'"> FEET</xsl:when>
		<xsl:when test="$Parcel.Line_Segment_Length.unit='foot'"> FEET</xsl:when>
		<xsl:when test="$Parcel.Line_Segment_Length.unit='meter'"> METERS</xsl:when>
	</xsl:choose>
</xsl:template>

<xsl:template match="lx:Curve">
	<xsl:apply-templates select="." mode="set"/>
	<xsl:variable name="CurveLength" select="landUtils:GetCurveLength()"/>
	<xsl:variable name="CurveAngle" select="landUtils:GetCurveAngle()"/>
	<xsl:variable name="CurveRadius" select="landUtils:GetCurveRadius()"/>
	<xsl:variable name="ChordDir" select="landUtils:CalculateBearingDMS(./lx:Start,./lx:End)"/>
	<xsl:variable name="ChordLength" select="landUtils:GetChordLength()"/>
	, TO A POINT ON A CURVE;<br/>THENCE ALONG THE ARC OF A CURVE TO THE <xsl:value-of select="landUtils:GetCurveDirection()"/> HAVING A RADIUS OF <xsl:value-of select="landUtils:FormatNumber(string($CurveRadius), string($SourceLinearUnit), string($Parcel.Curve_Segment_Radius.unit), string('0.00'), string($Parcel.Curve_Segment_Radius.rounding))"/>&#160;
	<xsl:choose>
		<xsl:when test="$Parcel.Curve_Segment_Radius.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceLinearUnit='foot'">FEET,</xsl:when>
				<xsl:when test="$SourceLinearUnit='USSurveyFoot'">FEET,</xsl:when>
				<xsl:when test="$SourceLinearUnit='meter'">METERS,</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Parcel.Curve_Segment_Radius.unit='foot'">FEET,</xsl:when>
		<xsl:when test="$Parcel.Curve_Segment_Radius.unit='USSurveyFoot'">FEET,</xsl:when>
		<xsl:when test="$Parcel.Curve_Segment_Radius.unit='meter'">METERS,</xsl:when>
	</xsl:choose>
	A CENTRAL ANGLE OF <xsl:value-of select="landUtils:FormatAngle(string($CurveAngle), string('default'), string($Parcel.Curve_Segment_Radius.unit), string($Parcel.Curve_Segment_Delta_Angle.format), string($Parcel.Curve_Segment_Delta_Angle.precision), string($Parcel.Curve_Segment_Delta_Angle.rounding))"/>,
	AN ARC LENGTH OF <xsl:value-of select="landUtils:FormatNumber(string($CurveLength), string($SourceLinearUnit), string($Parcel.Curve_Segment_Length.unit), string('0.00'), string($Parcel.Curve_Segment_Length.rounding))"/> FEET,
	THE CHORD OF WHICH BEARS <xsl:value-of select="$ChordDir"/>, 
	<xsl:value-of select="landUtils:FormatNumber(string($ChordLength), string($SourceLinearUnit), string($Parcel.Curve_Segment_Chord_Length.unit), string('0.00'), string($Parcel.Curve_Segment_Chord_Length.rounding))"/>&#160;
	<xsl:choose>
		<xsl:when test="$Parcel.Curve_Segment_Chord_Length.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceLinearUnit='foot'">FEET</xsl:when>
				<xsl:when test="$SourceLinearUnit='USSurveyFoot'">FEET</xsl:when>
				<xsl:when test="$SourceLinearUnit='meter'">METERS</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Parcel.Curve_Segment_Chord_Length.unit='foot'">FEET</xsl:when>
		<xsl:when test="$Parcel.Curve_Segment_Chord_Length.unit='USSurveyFoot'">FEET</xsl:when>
		<xsl:when test="$Parcel.Curve_Segment_Chord_Length.unit='meter'">METERS</xsl:when>
	</xsl:choose>
</xsl:template>

</xsl:stylesheet>