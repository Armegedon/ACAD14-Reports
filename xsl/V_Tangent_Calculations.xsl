<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet 	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
                		xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                 		xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
                		xmlns:lxml="urn:lx_utils"
                		version="1.0">
<msxsl:script language="JScript" implements-prefix="landUtils"> 
<![CDATA[
var pviOneStation = 0;
var pviOneElevation = 0;

var pviTwoStation = 0;
var pviTwoElevation = 0;

var vTangentHDiff = 0;
var vTangentVDiff = 0;
var vTangentLength = 0;
var vTangentGrade = 0;
// -----------------------------------------------------------------
// PVI Point One property functions
// -----------------------------------------------------------------
function SetPVIPointOne(pointStr)
{
	var strPoint = pointStr.split(" ");
	pviOneStation = new Number(parseFloat(strPoint[0]));
	pviOneElevation = new Number(parseFloat(strPoint[1]));

	return pointStr;
}
function SetPVIPointOneStation(station)
{
	pviOneStation = new Number(parseFloat(station));
	return offset;
}
function SetPVIPointOneElevatin(elevation)
{
	pviOneElevation = new Number(parseFloat(elevation));
	return elevation;
}
function GetPVIPointOneStation()
{
	return "" + pviOneStation;
}
function GetPVIPointOneElevation()
{
	return "" + pviOneElevation ;
}
// -----------------------------------------------------------------
// PVI Point Two property functions
// -----------------------------------------------------------------
function SetPVIPointTwo(pointStr)
{
	var strPoint = pointStr.split(" ");
	pviTwoStation = new Number(parseFloat(strPoint[0]));
	pviTwoElevation = new Number(parseFloat(strPoint[1]));

	return pointStr;
}
function SetPVIPointTwoStation(station)
{
	pviTwoStation = new Number(parseFloat(station));
	return offset;
}
function SetPVIPointTwoElevatin(elevation)
{
	pviTwoElevation = new Number(parseFloat(elevation));
	return elevation;
}
function GetPVIPointTwoStation()
{
	return "" + pviTwoStation;
}
function GetPVIPointTwoElevation()
{
	return "" + pviTwoElevation ;
}
// -----------------------------------------------------------------
// Calculation functions
// -----------------------------------------------------------------
function CalculationTangentValues()
{
	vTangentHDiff = pviTwoStation - pviOneStation;
	vTangentVDiff = pviTwoElevation - pviOneElevation;
	vTangentLength = Math.sqrt(Math.pow(vTangentHDiff, 2) + Math.pow(vTangentVDiff, 2));
	vTangentGrade = (vTangentVDiff / vTangentHDiff) * 100;
	return "done";
}
// -----------------------------------------------------------------
// Calculated property functions
// -----------------------------------------------------------------
function GetHorzDistance()
{
	return formatLinearNumber(vTangentHDiff);
}
function GetVertDistance()
{
	return formatLinearNumber(vTangentVDiff) ;
}
function GetVTangentLength()
{
	return formatLinearNumber(vTangentLength);
}
function GetVTangentGrade()
{
	return formatLinearNumber(vTangentGrade);
}
]]>
</msxsl:script>

</xsl:stylesheet>
