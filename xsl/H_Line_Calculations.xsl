<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet 	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
				xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
				xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit" 
				version="1.0">

<xsl:template match="lx:Line" mode="set">
	
	<xsl:choose >
		<xsl:when test="./lx:Start/@PntRef">
			<xsl:variable name="sp" select="./lx:Start/@pntRef"></xsl:variable>
			<xsl:variable name="nref" select="landUtils:SetLineStart(string(//lx:CgPoints/CgPoint[@name=$sp]))"></xsl:variable>
		</xsl:when>
		<xsl:otherwise>
			<xsl:variable name="nref" select="landUtils:SetLineStart(string(./lx:Start))"></xsl:variable>
		</xsl:otherwise>
	</xsl:choose>
	
	<xsl:choose >
		<xsl:when test="./End/@PntRef">
			<xsl:variable name="ep" select="./lx:End/@pntRef"></xsl:variable>
			<xsl:variable name="enref" select="landUtils:SetEndLinePoint(string(//lx:CgPoints/CgPoint[@name=$ep]))"></xsl:variable>
		</xsl:when>
		<xsl:otherwise>
			<xsl:variable name="enref" select="landUtils:SetEndLinePoint(string(./lx:End))"></xsl:variable>
		</xsl:otherwise>
	</xsl:choose>
	<xsl:variable name="calcline" select="landUtils:CalculateLineValues()"></xsl:variable>
</xsl:template>

<msxsl:script language="JScript" implements-prefix="landUtils"><![CDATA[
// Base units set to one for imperial and 2 for metric
var baseUnits = 1; 

// Calculation Rounding method
// 1 for normal
// 2 for Up
// 3 for down
// 4 for truncate
var rounding = 1;

var startLineNorthing = 0;
var startLineEasting = 0;
var endLineNorthing = 0;
var endLineEasting = 0;

var lineAngle = 0;
var lineLength = 0;

// -----------------------------------------------------------------
// Set/Get Calculation rounding functions
// -----------------------------------------------------------------
function SetRounding(roundingString)
{
	if(roundingString == "Normal")
	{
		rounding = 1;
	}
	else if(roundingString == "Rounding Up")
	{
		rounding = 2;
	}
	else if(roundingString == "Rounding Down")
	{
		rounding = 3;
	}
	else
	{
		rounding = 4;
	}
	return roundingString;
}
function GetRounding()
{
	if(rounding == 1)
	{
		return "Normal";
	}
	else if(rounding == 2)
	{
		return "Rounding Up";
	}
	else if(rounding == 3)
	{
		return "Rounding Down";
	}
	else
	{
		return "Truncate";
	}
}

// -----------------------------------------------------------------
// Set/Get Base Units functions
// -----------------------------------------------------------------
function SetBaseUnits(unitString)
{
	if(unitString=="imperial")
	{
		baseUnit = 1;
	}
	else
	{
		baseUnit = 2;
	}
	return unitString;
}

function GetBaseUnits()
{
	if(baseUnit == 1)
	{
		return "imperial";;
	}
	else
	{
		return "metric";
	}
}
// -----------------------------------------------------------------
// Starting Point (PC) property functions
// -----------------------------------------------------------------
function SetLineStart(pointStr)
{
	var strPoint = pointStr.split(" ");
	startLineNorthing = new Number(parseFloat(strPoint[0]));
	startLineEasting = new Number(parseFloat(strPoint[1]));
	return pointStr;
}
function SetStartLineNorthing(northingText)
{
	startLineNorthing = new Number(parseFloat(northingText));
	return northingText;
}
function SetStartLineEasting(eastingText)
{
	startLineEasting = new Number(parseFloat(eastingText));
	return eastingText;
}
function GetStartLineNorthing()
{
//	return formatCoordNumber(startLineNorthing );
	return "" + startLineNorthing;
}
function GetStartLineEasting()
{
//	return formatCoordNumber(startLineEasting );
	return "" + startLineEasting;
}

// -----------------------------------------------------------------
// Ending Point (PT) property functions
// -----------------------------------------------------------------
function SetEndLinePoint(pointStr)
{
	var strPoint = pointStr.split(" ");
	endLineNorthing = new Number(parseFloat(strPoint[0]));
	endLineEasting = new Number(parseFloat(strPoint[1]));
	return pointStr;
}
function SetEndLineNorthing(northingText)
{
	endLineNorthing = new Number(parseFloat(northingText));
	return northingText;
}
function SetEndEasting(eastingText)
{
	endLineEasting = new Number(parseFloat(eastingText));
	return eastingText;
}
function GetEndLineNorthing()
{
	//return formatCoordNumber(endLineNorthing );
	return "" + endLineNorthing;
}
function GetEndLineEasting()
{
	//return formatCoordNumber(endLineEasting );
	return "" + endLineEasting;
}
// -----------------------------------------------------------------
// Line calculation functions
// -----------------------------------------------------------------
function CalculateLineValues()
{
	lineAngle = 0;
	lineLength = 0;
	var xDiff = startLineEasting  -  endLineEasting ;
   	var yDiff = startLineNorthing - endLineNorthing  ;

	lineLength = Math.sqrt(Math.pow(xDiff, 2) + Math.pow(yDiff, 2));

 	var tanA = yDiff / xDiff;
   	var angle = (Math.atan(tanA)* 180)/Math.PI;

	// lineAngle will be 0-360 degrees, E=0 deg, counter-clockwise

	if (angle > 0)
	{
		if(startLineNorthing > endLineNorthing)  // SW
		{
			lineAngle = 180. + angle;
		}
		else    // NE
		{
			lineAngle = angle;
		}
	}

	else if (angle < 0)
	{
		if(startLineNorthing > endLineNorthing)  // SE
		{
			lineAngle = 360. + angle;
		}
		else	// NW
		{
			lineAngle = 180. + angle;
		}
	}

	else
	{
		if(startLineEasting > endLineEasting)
		{
			lineAngle = 180.;
		}
		else
   		{
			lineAngle = 0.;
		}
	}
	return "done";
}

function GetLineAngle()
{
	return lineAngle;
//	return formatAngleNumber(lineAngle);
}
function GetLineLength()
{
//	return formatLinearNumber(lineLength);
	return lineLength;
}
function GetHoriz() //also known as departure
{
	return startLineEasting - endLineEasting;
}
function GetVert() //also known as latitude
{
	return startLineNorthing - endLineNorthing;
}

function GetLineLength()
{
	var startN = GetStartLineNorthing();
	var n1 = new Number (startN);
	var startE = GetStartLineEasting();
	var e1 = new Number (startE);
	var endN = GetEndLineNorthing();
	var n2 = new Number (endN);
	var endE = GetEndLineEasting();
	var e2 = new Number (endE);
	
	var xDiff = e1 - e2 ;
   	var yDiff = n1 - n2 ;

	var lineLength = Math.sqrt(Math.pow(xDiff, 2) + Math.pow(yDiff, 2));
	return "" + lineLength;
}

]]></msxsl:script>
</xsl:stylesheet>