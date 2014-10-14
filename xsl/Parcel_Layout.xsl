<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2002 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">
				
<!-- =========== Parcel  Parameter definitions ============ -->
<xsl:param name="Parcel.2D_Area.unit">default</xsl:param>
<xsl:param name="Parcel.2D_Area.precision">0.00</xsl:param>
<xsl:param name="Parcel.2D_Area.rounding">normal</xsl:param>
<xsl:param name="Parcel.2D_Perimeter.unit">default</xsl:param>
<xsl:param name="Parcel.2D_Perimeter.precision">0.00</xsl:param>
<xsl:param name="Parcel.2D_Perimeter.rounding">normal</xsl:param>
<xsl:param name="Parcel.Line_Segment_Direction.unit">default</xsl:param>
<xsl:param name="Parcel.Line_Segment_Direction.precision">0.00</xsl:param>
<xsl:param name="Parcel.Line_Segment_Direction.rounding">normal</xsl:param>
<xsl:param name="Parcel.Line_Segment_Length.unit">default</xsl:param>
<xsl:param name="Parcel.Line_Segment_Length.precision">0.00</xsl:param>
<xsl:param name="Parcel.Line_Segment_Length.rounding">normal</xsl:param>
<xsl:param name="Parcel.Point_Coordinates.unit">default</xsl:param>
<xsl:param name="Parcel.Point_Coordinates.precision">0.00</xsl:param>
<xsl:param name="Parcel.Point_Coordinates.rounding">normal</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Length.unit">default</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Length.precision">0.00</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Length.rounding">normal</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Radius.unit">default</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Radius.precision">0.00</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Radius.rounding">normal</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Chord_Length.unit">default</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Chord_Length.precision">0.00</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Chord_Length.rounding">normal</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Delta_Angle.format">decimal</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Delta_Angle.precision">0.000</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Delta_Angle.rounding">normal</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Tangent_Length.unit">default</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Tangent_Length.precision">0.00</xsl:param>
<xsl:param name="Parcel.Curve_Segment_Tangent_Length.rounding">normal</xsl:param>

<!-- =========== JavaScript Includes ==== -->
<xsl:include href="H_Curve_Calculations.xsl"/>
<xsl:include href="H_Line_Calculations.xsl"/>
<xsl:include href="Cogo_Point.xsl"/>

<!-- =========== Begin Processing ======= -->
<xsl:template match="lx:Parcel" mode="set">
	<xsl:variable name="parName" select="landUtils:SetParcelName(@name)"/>
	<xsl:variable name="parArea" select="landUtils:SetParcelArea(string(@area))"/>
	<xsl:variable name="parCentN" select="landUtils:SetParcelCenter(string(./Center))"/>
	<xsl:apply-templates select="./lx:CoordGeom/node()/lx:Start" mode="setCommence"/>
</xsl:template>

<xsl:template match="lx:Start" mode="setCommence">
	<xsl:if test="position()=1">
	<xsl:choose >
		<xsl:when test="@PntRef">
			<xsl:variable name="nref" select="landUtils:SetCommencePoint(string(//lx:CgPoints/CgPoint[@name=@pntRef]))"/>
		</xsl:when>
		<xsl:otherwise>
			<xsl:variable name="nref" select="landUtils:SetCommencePoint(string(.))"/>
		</xsl:otherwise>
	</xsl:choose>
	</xsl:if>
</xsl:template>

<msxsl:script language="JScript" implements-prefix="landUtils"><![CDATA[
var parcelName;
var parcelArea = 0;
var parcelCenterNorthing = 0;
var parcelCenterEasting = 0;

var commencePointNorthing = 0;
var commencePointEasting = 0;

// --------------------------------------------------------------------
// Parcel attribute property functions
// --------------------------------------------------------------------
function SetParcelName(name)
{
	parcelName = name;
	return name;
}
function GetParcelName()
{
	return parcelName;
}

function SetParcelArea(area)
{
	
	parcelArea = new Number(parseFloat(area));
	return "" + parcelArea;
}

function GetParcelArea()
{
	return "" + area;
}

// --------------------------------------------------------------------
// Center Point property functions
// --------------------------------------------------------------------
function SetParcelCenter(pointStr)
{
	strPoint = pointStr.split(" ");
	parcelCenterNorthing = new Number(parseFloat(strPoint[0]));
	parcelCenterEasting = new Number(parseFloat(strPoint[1]));
	return pointStr;
}
function GetParcelCenterNorthing()
{
	return "" + parcelCenterNorthing;
}
function GetParcelCenterEasting()
{
	return "" + parcelCenterEasting;
}
// --------------------------------------------------------------------
// Commencement Point property functions
// --------------------------------------------------------------------
function SetCommencePoint(pointStr)
{
	strPoint = pointStr.split(" ");
	commencePointNorthing = new Number(parseFloat(strPoint[0]));
	commencePointEasting = new Number(parseFloat(strPoint[1]));
	
	return "done";
}

function SetCommencePointNorthing(pointStr)
{
	commencePointNorthing = new Number(parseFloat(pointStr));
	return pointStr;
}
function SetCommencePointEasting(pointStr)
{
	commencePointEasting = new Number(parseFloat(pointStr));
	return pointStr;
}
function GetCommencePointNorthing()
{
	return Number(commencePointNorthing);
}
function GetCommencePointEasting()
{
	return Number(commencePointEasting);
}


]]></msxsl:script>
</xsl:stylesheet>