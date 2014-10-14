<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet 	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
				xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
				xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit" 
				version="1.0">
<!-- =========== JavaScript Include==== -->
<xsl:include href="H_Curve_Calculations.xsl"></xsl:include>
<xsl:include href="H_Spiral_Calculations.xsl"></xsl:include>
<xsl:include href="H_Line_Calculations.xsl"></xsl:include>
<xsl:include href="Cogo_Point.xsl"></xsl:include>
<!-- ================================= -->

<!-- =========== Formatting parameters==== -->
<xsl:param name="Alignment.Elevation.unit">default</xsl:param>
<xsl:param name="Alignment.Elevation.precision">0.00</xsl:param>
<xsl:param name="Alignment.Elevation.rounding">normal</xsl:param>
<xsl:param name="Alignment.Coordinate.unit">default</xsl:param>
<xsl:param name="Alignment.Coordinate.precision">0.00</xsl:param>
<xsl:param name="Alignment.Coordinate.rounding">normal</xsl:param>
<xsl:param name="Alignment.Linear.unit">default</xsl:param>
<xsl:param name="Alignment.Linear.precision">0.00</xsl:param>
<xsl:param name="Alignment.Linear.rounding">normal</xsl:param>
<xsl:param name="Alignment.Angular.unit">default</xsl:param>
<xsl:param name="Alignment.Angular.precision">0.00</xsl:param>
<xsl:param name="Alignment.Angular.rounding">normal</xsl:param>
<xsl:param name="Alignment.Station.Display">##+##</xsl:param>
<xsl:param name="Alignment.Station.unit">default</xsl:param>
<xsl:param name="Alignment.Station.precision">0.00</xsl:param>
<xsl:param name="Alignment.Station.rounding">normal</xsl:param>
<!-- ================================= -->

<xsl:template match="lx:Alignment" mode="set">
	<xsl:variable name="ststa" select="landUtils:SetStartingStation(string(@staStart))"></xsl:variable>
</xsl:template>

<xsl:template match="lx:CoordGeom" mode="set">
	<xsl:variable name="len" select="landUtils:SetArrayLength(count(./node()))"></xsl:variable>
	 <xsl:for-each select="./node()">
		<xsl:apply-templates select="."></xsl:apply-templates>
		<xsl:choose >
			<xsl:when test="position() = last()">
			<xsl:apply-templates select="." mode="end"></xsl:apply-templates>
			</xsl:when>
		</xsl:choose>
	</xsl:for-each>
	
</xsl:template>

<xsl:template match="lx:Line">
	<xsl:apply-templates select="." mode="set"></xsl:apply-templates>
	<xsl:variable name="len" select="landUtils:GetLineLength()"></xsl:variable>
	<xsl:variable name="delta" select="landUtils:GetLineAngle()"/>
	<xsl:variable name="north" select="landUtils:GetStartLineNorthing()"></xsl:variable>
	<xsl:variable name="east" select="landUtils:GetStartLineEasting()"></xsl:variable>
	<xsl:variable name="endNorth" select="landUtils:GetEndLineNorthing()"/>
	<xsl:variable name="endEast" select="landUtils:GetEndLineEasting()"/>
	
	<xsl:variable name="setSN" select="landUtils:SetHStartPointNorthing($north)"/>
	<xsl:variable name="setSE" select="landUtils:SetHPointStartPointEasting($east)"/>
	<xsl:variable name="setEN" select="landUtils:SetHEndPointNorthing($endNorth)"/>
	<xsl:variable name="setEE" select="landUtils:SetHEndPointEasting($endEast)"/>

	<xsl:variable name="setDelta" select="landUtils:AddHDelta($delta)"/>
	<xsl:variable name="addl" select="landUtils:AddPIPoint($len, $north, $east)"></xsl:variable>
</xsl:template>

<xsl:template match="lx:Curve">
	<xsl:apply-templates select="." mode="set"></xsl:apply-templates>
	
	<xsl:variable name="curSN" select="landUtils:GetStartingNorthing()"/>
	<xsl:variable name="curSE" select="landUtils:GetStartingEasting()"/>
	<xsl:variable name="curEN" select="landUtils:GetCenterNorthing()"/>
	<xsl:variable name="curEE" select="landUtils:GetCenterEasting()"/>
	<xsl:variable name="endNorth" select="landUtils:GetCurveEndNorthing()"/>
	<xsl:variable name="endEast" select="landUtils:GetCurveEndEasting()"/>

	<xsl:variable name="piN" select="landUtils:GetCurvePINorthing()"/>
	<xsl:variable name="piE" select="landUtils:GetCurvePIEasting()"/>
	<xsl:variable name="tLen" select="landUtils:GetCurveTangent()"/>
<!--		<xsl:variable name="curLen" select="@length"/>-->
	<xsl:variable name="curLen" select="landUtils:GetCurveLength()"/>
	<xsl:variable name="curDelta" select="landUtils:GetCurveAngle()"/>
	<xsl:variable name="curRadius" select="landUtils:GetCurveRadius()"/>
	<xsl:variable name="rotValue" select="landUtils:GetRotatonValue()"/>
	
	
	<xsl:variable name="setRad" select="landUtils:AddHRadius($curRadius)"/>
	<xsl:variable name="setSN" select="landUtils:SetHStartPointNorthing($curSN)"/>
	<xsl:variable name="setSE" select="landUtils:SetHPointStartPointEasting($curSE)"/>
	<xsl:variable name="setCN" select="landUtils:SetHCenterPointNorthing($curEN)"/>
	<xsl:variable name="seCtE" select="landUtils:SetHCenterPointEasting($curEE)"/>
	<xsl:variable name="setEN" select="landUtils:SetHEndPointNorthing($endNorth)"/>
	<xsl:variable name="setEE" select="landUtils:SetHEndPointEasting($endEast)"/>
	<xsl:variable name="setDelta" select="landUtils:AddHDelta($curDelta)"/>
	
<!--	<xsl:variable name="addl" select="landUtils:AddPIPointCurve($curLen, $tLen, $piN, $piE)"></xsl:variable>-->
	<xsl:variable name="addl" select="landUtils:AddPIPointCurve($curLen, $tLen, $piN, $piE, $rotValue)"></xsl:variable>

</xsl:template>

<xsl:template match="lx:Line" mode="end">
	<xsl:apply-templates select="." mode="set"></xsl:apply-templates>
	<xsl:variable name="delta" select="landUtils:GetLineAngle()"/>
	<xsl:variable name="len" select="landUtils:GetLineLength()"></xsl:variable>
	<xsl:variable name="sNorth" select="landUtils:GetStartLineNorthing()"/>
	<xsl:variable name="sEast" select="landUtils:GetStartLineEasting()"/>
	
	<xsl:variable name="north" select="landUtils:GetEndLineNorthing()"></xsl:variable>
	<xsl:variable name="east" select="landUtils:GetEndLineEasting()"></xsl:variable>
	
	<xsl:variable name="setSN" select="landUtils:SetHStartPointNorthing($north)"/>
	<xsl:variable name="setSE" select="landUtils:SetHPointStartPointEasting($east)"/>
	
	<xsl:variable name="setEN" select="landUtils:SetHEndPointNorthing($north)"/>
	<xsl:variable name="setEE" select="landUtils:SetHEndPointEasting($east)"/>

	<xsl:variable name="setDelta" select="landUtils:AddHDelta($delta)"/>
	<xsl:variable name="addl" select="landUtils:AddPIPoint($len, $north, $east)"></xsl:variable>
</xsl:template>

<xsl:template match="lx:Curve" mode="end">
	<xsl:apply-templates select="." mode="set"></xsl:apply-templates>
	<xsl:variable name="curSN" select="landUtils:GetStartingNorthing()"/>
	<xsl:variable name="curSE" select="landUtils:GetStartingEasting()"/>
	<xsl:variable name="curEN" select="landUtils:GetCenterNorthing()"/>
	<xsl:variable name="curEE" select="landUtils:GetCenterEasting()"/>
	<xsl:variable name="endNorth" select="landUtils:GetCurveEndNorthing()"/>
	<xsl:variable name="endEast" select="landUtils:GetCurveEndEasting()"/>
	
<!--	 	<xsl:variable name="piN" select="landUtils:GetCurvePINorthing()" /> Seems like you could do this. The next 2 lines are how it was originally
 	 	<xsl:variable name="piE" select="landUtils:GetCurvePIEasting()" />-->
	<xsl:variable name="piN" select="landUtils:GetEndLineNorthing()"/>
	<xsl:variable name="piE" select="landUtils:GetEndLineEasting()"/> 
	
	<xsl:variable name="tLen" select="landUtils:GetCurveTangent()"/>
<!--		<xsl:variable name="curLen" select="@length"/>-->
	<xsl:variable name="curLen" select="landUtils:GetCurveLength()"/>
	<xsl:variable name="curDelta" select="landUtils:GetCurveAngle()"/>
	<xsl:variable name="curRadius" select="landUtils:GetCurveRadius()"/>
	<xsl:variable name="rotValue" select="landUtils:GetRotatonValue()"/>

	<xsl:variable name="setRad" select="landUtils:AddHRadius($curRadius)"/>
	<xsl:variable name="setSN" select="landUtils:SetHStartPointNorthing($curSN)"/>
	<xsl:variable name="setSE" select="landUtils:SetHPointStartPointEasting($curSE)"/>
	<xsl:variable name="setCN" select="landUtils:SetHCenterPointNorthing($curEN)"/>
	<xsl:variable name="setCE" select="landUtils:SetHCenterPointEasting($curEE)"/>
	<xsl:variable name="setEN" select="landUtils:SetHEndPointNorthing($endNorth)"/>
	<xsl:variable name="setEE" select="landUtils:SetHEndPointEasting($endEast)"/>
	<xsl:variable name="setDelta" select="landUtils:AddHDelta($curDelta)"/>

		<xsl:variable name="addl" select="landUtils:AddPIPointCurve($curLen, $tLen, $piN, $piE, $rotValue)"></xsl:variable>
</xsl:template>

<xsl:template match="lx:Spiral">
	<xsl:apply-templates select="." mode="set"></xsl:apply-templates>
	<xsl:variable name="spiSN" select="landUtils:GetSpiralStartNorthing()"/>
	<xsl:variable name="spiSE" select="landUtils:GetSpiralStartEasting()"/>
	<xsl:variable name="spiEN" select="landUtils:GetSpiralEndNorthing()"/>
	<xsl:variable name="spiEE" select="landUtils:GetSpiralEndEasting()"/>

	<xsl:variable name="spiStartRadius" select="landUtils:GetSpiralStartRadius()"/>
	<xsl:variable name="spiEndRadius" select="landUtils:GetSpiralEndRadius()"/>
	<xsl:variable name="setRad" select="landUtils:AddHRadius($spiStartRadius)"/>

	<xsl:variable name="piN" select="landUtils:GetSpiralPINorthing()"/>
	<xsl:variable name="piE" select="landUtils:GetSpiralPIEasting()"/>
	<xsl:variable name="ltLen" select="landUtils:GetSpiralLongTangent()"/>
	<xsl:variable name="stLen" select="landUtils:GetSpiralShortTangent()"/>
	<xsl:variable name="spiLen" select="landUtils:GetSpiralLength()"/>
	<xsl:variable name="spiTheta" select="landUtils:GetSpiralTheta_Radians()"/>
	<xsl:variable name="rotValue" select="landUtils:GetSpiralRotation()"/>

	
	<xsl:variable name="setSN" select="landUtils:SetHStartPointNorthing($spiSN)"/>
	<xsl:variable name="setSE" select="landUtils:SetHPointStartPointEasting($spiSE)"/>
	<xsl:variable name="setEN" select="landUtils:SetHEndPointNorthing($spiEN)"/>
	<xsl:variable name="setEE" select="landUtils:SetHEndPointEasting($spiEE)"/>

<xsl:variable name="addl" select="landUtils:AddPIPointSpiral($spiLen, $ltLen, $piN, $piE, $rotValue)"/>

</xsl:template>

<xsl:template match="lx:Spiral" mode="end">
	<xsl:apply-templates select="." mode="set"></xsl:apply-templates>

	<xsl:variable name="spiSN" select="landUtils:GetSpiralStartNorthing()"/>
	<xsl:variable name="spiSE" select="landUtils:GetSpiralStartEasting()"/>
	<xsl:variable name="spiEN" select="landUtils:GetSpiralEndNorthing()"/>
	<xsl:variable name="spiEE" select="landUtils:GetSpiralEndEasting()"/>
	
	<xsl:variable name="spiStartRadius" select="landUtils:GetSpiralStartRadius()"/>
	<xsl:variable name="spiEndRadius" select="landUtils:GetSpiralEndRadius()"/>
	<xsl:variable name="setRad" select="landUtils:AddHRadius($spiStartRadius)"/>

	<xsl:variable name="piN" select="landUtils:GetSpiralPINorthing()"/>
	<xsl:variable name="piE" select="landUtils:GetSpiralPIEasting()"/>
	<xsl:variable name="ltLen" select="landUtils:GetSpiralLongTangent()"/>
	<xsl:variable name="stLen" select="landUtils:GetSpiralShortTangent()"/>
	<xsl:variable name="spiLen" select="landUtils:GetSpiralLength()"/>
	<xsl:variable name="spiTheta" select="landUtils:GetSpiralTheta_Radians()"/>
	<xsl:variable name="rotValue" select="landUtils:GetSpiralRotation()"/>

	<xsl:variable name="setSN" select="landUtils:SetHStartPointNorthing($spiSN)"/>
	<xsl:variable name="setSE" select="landUtils:SetHPointStartPointEasting($spiSE)"/>
	<xsl:variable name="setEN" select="landUtils:SetHEndPointNorthing($spiEN)"/>
	<xsl:variable name="setEE" select="landUtils:SetHEndPointEasting($spiEE)"/>
	
	<xsl:variable name="addl" select="landUtils:AddPIPointSpiral($spiLen, $ltLen, $piN, $piE, $rotValue)"/>

</xsl:template>

<msxsl:script language="JScript" implements-prefix="landUtils"><![CDATA[
// -----------------------------------------------------------------
// Global properties
// -----------------------------------------------------------------
var startSta = 0;
var incr = 25;
var trackStation = 0;
var length = 4;
var ndx = 1;
var nLen = 0;
var curveRotation = -1;

var hElementArray;
var hStartPointNorthingArray;
var hStartPointEastingArray;
var hCenterPointNorthingArray;
var hCenterPointEastingArray;
var hEndPointNorthingArray;
var hEndPointEastingArray;
var hStationArray;
var hLengthArray;
var hCurveDirectionArray;

var hDeltaArray;
var hRadiusArray;

var PIstationArray;
var PInorthingArray;
var PIeastingArray;

// -----------------------------------------------------------------
// Station Increment property functions
// -----------------------------------------------------------------
function SetStationIncrement(increment)
{
	incr = new Number(increment);
	return increment;
}
function GetStationIncrement()
{
	return "" + incr;
}

function Check(index)
{
	var nx = new Number(index);
	
	var end = hStationArray[nLen];
	
	if((end + incr) > trackStation)
	{
		return "Ok";
	}
	return "Quit";
}

function IncrementStation(station)
{
	var sta = new Number(station);
	trackStation = trackStation + incr;
	return "" + (sta + incr);
}

function DecrementStation(station)
{
	var sta = new Number(station);
	trackStation = trackStation - incr;
	return "" + (sta - incr);
}
// -----------------------------------------------------------------
// Starting Station property functions
// -----------------------------------------------------------------
function SetStartingStation(station)
{
	startSta = new Number(parseFloat(station));
	CheckForEnd = startSta;
	trackStation = startSta;
	length = startSta;
	return station;
}
function GetStartStation()
{
	return "" + startSta;
}
// -----------------------------------------------------------------
// Array functions
// -----------------------------------------------------------------
function SetArrayLength(len)
{
	nLen = new Number(len);
	hElementArray = new Array(len + 1);
	hStartPointNorthingArray = new Array(len + 1);
	hStartPointEastingArray = new Array(len + 1);
	hCenterPointNorthingArray = new Array(len + 1);
	hCenterPointEastingArray = new Array(len + 1);
	hEndPointNorthingArray = new Array(len + 1);
	hEndPointEastingArray = new Array(len + 1);
	hLengthArray = new Array(len + 1);
	hStationArray = new Array(len + 1);
	hDeltaArray = new Array(len + 1);
	hRadiusArray = new Array(len + 1);
	hCurveDirectionArray = new Array(len + 1);

	PIstationArray = new Array(len + 1);
	PInorthingArray = new Array(len + 1);
	PIeastingArray = new Array(len + 1);
	ndx = 1;
	
	hStationArray[0] = 0;
	return len;
}
function GetArrayLength()
{
	return "" + nLen;
}
// -----------------------------------------------------------------
// Starting Point array property functions
// -----------------------------------------------------------------
function SetHStartPointNorthing(northing)
{
	var temp = new Number(northing);
	hStartPointNorthingArray[ndx] = temp;
	return northing;
}
function GetHStartPointNorthing(index)
{
	var indx = new Number(index);
	return "" + hStartPointNorthingArray[indx+ 1];
}
function SetHPointStartPointEasting(easting)
{
	var temp = new Number(easting);
	hStartPointEastingArray[ndx] = temp;
	return easting;
}
function GetHStartPointEasting(index)
{
	var indx = new Number(index);
	return hStartPointEastingArray[indx];
}
// -----------------------------------------------------------------
// Center Point array property functions
// -----------------------------------------------------------------
function SetHCenterPointNorthing(northing)
{
	var temp = new Number(northing);
	hCenterPointNorthingArray[ndx] = temp;
	return northing;
}
function GetHCenterPointNorthing(index)
{
	var indx = new Number(index);
	return "" + hCenterPointNorthingArray[indx+ 1];
}
function SetHCenterPointEasting(easting)
{
	var temp = new Number(easting);
	hCenterPointEastingArray[ndx] = temp;
	return easting;
}
function GetHCenterPointEasting(index)
{
	var indx = new Number(index);
	return hCenterPointEastingArray[indx];
}
// -----------------------------------------------------------------
// Ending Point array property functions
// -----------------------------------------------------------------
function SetHEndPointNorthing(northing)   //function SetHEndPointNorthing(index, northing)
{
	var temp = new Number(northing);
//	var indx = new Number(index);
//	hEndPointNorthingArray[indx] = temp;
	hEndPointNorthingArray[ndx] = temp;
	return "" + northing;
}
function SetHEndPointEasting(easting)  //function SetHEndPointEasting(index, easting)
{
	var temp = new Number(easting);
//	var indx = new Number(index);
//	hEndPointEastingArray[indx] = temp;
	hEndPointEastingArray[ndx] = temp;
	return "" + easting;
}
function AddHDelta(delta)
{
	var del = new Number(delta);
	hDeltaArray[ndx] = del;
	return delta;
}

function AddHRadius(radius)
{
	var rad = new Number(radius);
	hRadiusArray[ndx] = rad;
	
	return "" + rad;
}

function AddPIPoint(len, northing, easting)
{
	var l = new Number(parseFloat(len));
	var n = new Number(parseFloat(northing));
	var e = new Number(parseFloat(easting));

	PInorthingArray[ndx] = n;
	PIeastingArray[ndx] = e;

	PIstationArray[ndx] = length;

	hLengthArray[ndx] = l;
	hElementArray[ndx] = 1;
	hCurveDirectionArray[ndx] = 0;
	
	if (ndx > 1)
		hStationArray[ndx] = hStationArray[ndx-1] + hLengthArray[ndx-1];
	else
		hStationArray[ndx] = 0 + length;
	
	length = length + l;
	
	ndx = ndx + 1;
	return ndx;
}

function AddPIPointCurve(len, tlen, northing, easting, rotation)
{
	var l = new Number(parseFloat(len));
	var tl = new Number(parseFloat(tlen));
	var n = new Number(parseFloat(northing));
	var e = new Number(parseFloat(easting));
	PInorthingArray[ndx] = n;
	PIeastingArray[ndx] = e;
	hLengthArray[ndx] = l;
	hElementArray[ndx] = 2;
	
	if (ndx > 1)
		hStationArray[ndx] = hStationArray[ndx-1] + hLengthArray[ndx-1];
	//	hStationArray[ndx] = 21;//hStationArray[ndx-1] + hLengthArray[ndx];
	else
		hStationArray[ndx] = 0 + length;

	
	length = length + tl;
	PIstationArray[ndx] = length;  //PI is current start station (length) + the tangent length
	length = length - tl;                //Move length back to the start station
	length = length + l;               //Add the length of the curve 
	
	hCurveDirectionArray[ndx] = rotation;
	ndx = ndx + 1;
	return ndx;
}

function AddPIPointSpiral(len, tlen, northing, easting, rotation)
{
	var l = new Number(parseFloat(len));
	var tl = new Number(parseFloat(tlen));
	var n = new Number(parseFloat(northing));
	var e = new Number(parseFloat(easting));
	PInorthingArray[ndx] = n;
	PIeastingArray[ndx] = e;
	
	hLengthArray[ndx] = l; //PI is current start station (length) + the long tangent length
	hElementArray[ndx] = 3;
	
	if (ndx > 1)
		//hStationArray[ndx] = 0 + hLengthArray[ndx];
		hStationArray[ndx] = hStationArray[ndx-1] + hLengthArray[ndx-1];
	else
		hStationArray[ndx] = 0 + length;

	length = length + tl;
	PIstationArray[ndx] = length;  //PI is current start station (length) + the tangent length
	length = length - tl;                //Move ‘length’ back to the start station
	length = length + l;               //Add the length of the spiral 

	hCurveDirectionArray[ndx] = rotation;
	ndx = ndx + 1;
	return ndx;
}

function GetDistanceBetweenPIs(index)
{
	return (PIstationArray[index+1] - PIstationArray[index]);
}
function SetPntAtIndex(index, station, northing, easting)
{
	return index;
}
function GetStationAtIndex(index)
{
	var indx = new Number(index);
	return hStationArray[indx];  -1
}
function GetStationAtPI(index)
{
	tempPI = PIstationArray [index];
	return "" + tempPI;
}
function GetNorthingAtPI(index)
{
	tempPI = PInorthingArray [index];
	return "" + tempPI;
}
function GetEastingAtPI(index)
{
	tempPI = PIeastingArray [index];
	return "" + tempPI;
}

function GetEndNorthing(index)
{
	var north = hEndPointNorthingArray [index];
	return "" + north;
}
function GetEndEasting(index)
{
	var east = hEndPointEastingArray [index];
	return "" + east;
}

function GetElementTypeAtIndex(index)
{
	var et = hElementArray[index];
	return "" + et;
}
function GetPreviousElementType(index)
{
	if (index == 1)
	{
		var et = 1;
	}
	else
	{
		et = hElementArray[index-1];
	}
	return "" + et;
}
function GetNextElementType(index)
{
	if (index >= nLen)
	{
		var et = hElementArray[index];
	}
	
	else
	{
		et = hElementArray[index+1];
	}
	return "" + et;
}

function ComparePreviousCurveDirection(index)
{
	if (index == 1)
	{
		return "0";
	}
	else
	{
		var LastRot = hCurveDirectionArray[index-1];
		var ThisRot = hCurveDirectionArray[index];
		if (LastRot == ThisRot)
		{
			return "Compound";
		}
		else
		{
			return "Reverse";
		}
		
	}

}

function CompareNextCurveDirection(index)
{
	if (index >= nLen)
	{
		return "0";
	}
	else
	{
		var NextRot = hCurveDirectionArray[index+1];
		var ThisRot = hCurveDirectionArray[index];
		if (NextRot == ThisRot)
		{
			return "Compound";
		}
		else
		{
			return "Reverse";
		}
		
	}
}

function GetNextStation(index)
{
	if (index >= nLen)
	{
		var et = hStationArray[index];
//		var et = PIstationArray[index];
	}
	
	else
	{
//		et = PIstationArray[index+1];
		et = hStationArray[index+1];
	}
	return "" + et;
}

function GetLengthAtIndex(index)
{
	return "" + hLengthArray[index];
	//if(PInorthingArray[index] != null)
	//{
		//var cN = new Number(PInorthingArray[index]);
		//var cE = new Number(PIeastingArray[index]);
		//var xN = new Number(PInorthingArray[index+ 1]);
		//var xE = new Number(PIeastingArray[index+ 1]);
		
		//var nDiff = xN - cN;
		//var eDiff = xE - cE;
		
		//return Math.sqrt(Math.pow(eDiff, 2) + Math.pow(nDiff, 2));
		////return index;
	//}
	//else
	//{
		//return "dd";
	//}
}
function GetDirectionAtIndex_Old(index)
{
	lineAngle = 0;
	var indx = new Number(index);
	var cN = new Number(PInorthingArray[indx ]);
	var cE = new Number(PIeastingArray[indx ]);
	var xN = new Number(PInorthingArray[indx + 1]);
	var xE = new Number(PIeastingArray[indx + 1]);
		
	var yDiff = xN - cN;
	var xDiff = xE - cE;
   	
    	var tanA = yDiff / xDiff;
    	var tempAngle = Math.atan(tanA);
    	if(tempAngle < 0)
    	{
    		if(cN > xN)
    		{
    			lineAngle = 360 + ( (tempAngle * 180) / Math.PI);    		
   		}
    		else
    		{
    			lineAngle = 180 + ( (tempAngle * 180) / Math.PI);
    		}
    	}
    	else
    	{
    		if(cN > xN)
    		{
    			lineAngle =180 + ((tempAngle * 180) / Math.PI);
    		}
    		else
    		{
     			lineAngle = ((tempAngle * 180) / Math.PI);
    		}
    	}
	
	return lineAngle; //"N: " + xN + " E: " + xE; //
}


function GetDirectionAtIndex(index)
{
	lineAngle = 0;
	var indx = new Number(index);
	var cN = new Number(PInorthingArray[indx ]);
	var cE = new Number(PIeastingArray[indx ]);
	var xN = new Number(PInorthingArray[indx + 1]);
	var xE = new Number(PIeastingArray[indx + 1]);
		
	var yDiff = xN - cN;
	var xDiff = xE - cE;
   	
    	var tanA = yDiff / xDiff;
    	var tempAngle = Math.atan(tanA);
    	if(tempAngle > 0)
	{
    		if(cN > xN)  // SW
    		{
     		lineAngle = 270.0 - tempAngle; 
   		}
    		else    // NE
    		{
    			lineAngle = tempAngle; 
    		}
    }
    
    else if (tempAngle < 0)
    {
        	if (cN > xN)    //SE
    		{
     		lineAngle = 90 - tempAngle;
   		}
    		else      //NW
    		{
    			lineAngle = 270 - tempAngle;
    		}
    }
    
    else
    {
    		if (cN > xN)
    		{
     		lineAngle = 180.00;
   		}
    		else
    		{
    			lineAngle = 0.00;
		}
	}

	return lineAngle;
}

// -----------------------------------------------------------------
// Calulation functions
// -----------------------------------------------------------------

function OffsetOfPoint(northing, easting)
{
}
function StationOfPoint(northing, easting)
{
	var lenIncr = startSta;
	
}
function GetStationForIndex(index)
{
	return "" + hStationArray[index];
}
function GetIndexNearestStation(station)
{
	var idx = 0;
	var stat = startSta;
	var st =  new Number(station);
	
	for(var i = 1; i <= nLen; i++)
	{
		var staCheck = hStationArray[i];
		
		if(st < staCheck)
		{
			return "" + i;
		}
	}
	return "fall through " + idx + " Station: " + station;
}
function NorthingOfIncrStation()
{
	var north = 0;
	var sta = new Number(trackStation );
	var eleNdx = GetIndexNearestStation(trackStation );
	
	var spN = GetHStartPointNorthing(eleNdx);
	var spE = GetHStartPointEasting(eleNdx);
	var spSta = hStationArray[eleNdx - 1];
	var dLen = sta - spSta ;
	var delta = hDeltaArray[eleNdx];
	var length = hLengthArray[eleNdx];
	
	var etype = GetElementTypeAtIndex(eleNdx);
	if(etype == 1)
	{
		north = spN - (dLen * Math.cos((delta * Math.PI) / 180 )) ;
	}
	else if(etype == 2)
	{
		var theta = delta * (dLen / length);
		var rad = hRadiusArray[eleNdx];
		var cpN = hCenterPointNorthingArray[eleNdx];
		north = cpN - (rad * Math.cos((theta * Math.PI) / 180) );
	}
	
	else if(etype == 3)
	{
		var sLen = GetSpiralLength();
		var sStartRad = GetSpiralStartRadius();
		var sEndRad = GetSpiralEndRadius();
		var sRad = sEndRad;
		if (sRad == "NaN")
		{
			sRad = sStartRad;
		}
		
		var Temp1 = Math.pow(dLen, 3) / (6 * sRad * sLen);
		var Temp2 = Math.pow(dLen, 4) / (56 * Math.pow(sRad, 2) * Math.pow(sLen, 2));
		var Temp3 = Math.pow(dLen, 8) / (7040 * Math.pow(sRad, 4) * Math.pow(sLen, 4));
		var y = spE - (Temp1 * (1 - Temp2 + Temp3));
		north = 0;

	}

	return "" + north;
}
function GetTrackedStation()
{
	return "" + trackStation;
}
function EastingOfIncrStation()
{
	var east = 0;
	var sta = new Number(trackStation );
	var eleNdx = GetIndexNearestStation(trackStation);
	
	var spN = GetHStartPointNorthing(eleNdx);
	var spE = GetHStartPointEasting(eleNdx);
	var spSta = hStationArray[eleNdx - 1];
	var dLen = sta - spSta ; 
	var delta = hDeltaArray[eleNdx];
	
	var etype = GetElementTypeAtIndex(eleNdx);
	if(etype == 1)
	{
		east = spE - (dLen * Math.sin((delta * Math.PI) / 180 )) ;
	}
	else if(etype == 2)
	{
	
		var theta = delta * (dLen / length);
		var rad = hRadiusArray[eleNdx];
		var cpE = hCenterPointEastingArray[eleNdx];
		east = cpE - (rad * Math.sin((theta * Math.PI) / 180) );
	}
	
	else if(etype == 3)
	{
		var sLen = dLen; //GetSpiralLength();
		var sStartRad = GetSpiralStartRadius();
		var sEndRad = GetSpiralEndRadius();
		var sRad = sEndRad;
		if (sRad == "NaN")
		{
			sRad = sStartRad;
		}
		
		var Temp1 = Math.pow(dLen, 4) / (40 * Math.pow(sRad, 2) * Math.pow(sLen, 2));
		var Temp2 = Math.pow(dLen, 8) / (3456 * Math.pow(sRad, 4) * Math.pow(sLen, 4));
		var x =  (dLen * (1 - Temp1 + Temp2));

		east = 0;

	}

	return "" + east;
}
function NorthingOfStation(station)
{
	var north = 0;
	var sta = new Number(station);
	var eleNdx = GetIndexNearestStation(station);
	
	var spN = GetHStartPointNorthing(eleNdx);
	var spE = GetHStartPointEasting(eleNdx);
	var spSta = hStationArray[eleNdx - 1];
	var dLen = sta - spSta ;
	var delta = hDeltaArray[eleNdx];
	var length = hLengthArray[eleNdx];
	
	var etype = GetElementTypeAtIndex(eleNdx);
	if(etype == 1)
	{
		north = spN - (dLen * Math.cos((delta * Math.PI) / 180 )) ;
	}
	else if(etype == 2)
	{
		var theta = delta * (dLen / length);
		var rad = hRadiusArray[eleNdx];
		var cpN = hCenterPointNorthingArray[eleNdx];
		north = cpN - (rad * Math.cos((theta * Math.PI) / 180) );
	}
	
	return "" + north;
}
function EastingOfStation(station)
{
	var east = 0;
	var sta = new Number(station);
	var eleNdx = GetIndexNearestStation(station);
	
	var spN = GetHStartPointNorthing(eleNdx);
	var spE = GetHStartPointEasting(eleNdx);
	var spSta = hStationArray[eleNdx - 1];
	var dLen = sta - spSta ;
	var delta = hDeltaArray[eleNdx];
	
	var etype = GetElementTypeAtIndex(eleNdx);
	if(etype == 1)
	{
		east = spE - (dLen * Math.sin((delta * Math.PI) / 180 )) ;
	}
	else if(etype == 2)
	{
		var theta = delta * (dLen / length);
		var rad = hRadiusArray[eleNdx];
		var cpE = hCenterPointEastingArray[eleNdx];
		east = cpE - (rad * Math.sin((theta * Math.PI) / 180) );
	}
	
	return "" + east;
}
function GetDirectionAtStation(station)
{
	var sta = new Number(station);
	var eleNdx = GetIndexNearestStation(station);
	return "" + GetDirectionAtIndex(eleNdx);
	//return "" + eleNdx;
}

function getLastStation(startStation, length)
{
	var sta = new Number(startStation);
	var len = new Number (length);
	var end = sta + len;
	return "" + end;
}






]]></msxsl:script>

</xsl:stylesheet>
