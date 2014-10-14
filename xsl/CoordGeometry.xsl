<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet version="1.0"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">
	
<xsl:include href="NodeConversion_JScript.xsl"/>
<xsl:include href="Computations_JScript.xsl"/>

<xsl:template name="ApplyStationEquations">
<xsl:param name="StartStation" select="0"></xsl:param>
	<xsl:variable name="applyStaEq" select="landUtils:ApplyStationEquations($StartStation)"/>
</xsl:template>

<xsl:template match="lx:CoordGeom" mode="set">
	<xsl:variable name="clear" select="landUtils:Clear()"/>
	<xsl:variable name="len" select="landUtils:SetGeomArrayLength(count(./*))"></xsl:variable>
	 <xsl:for-each select="./*">
	 	<xsl:variable name="pos" select="position()"/>
		<xsl:choose >
			<!-- Adding a line call -->
			<xsl:when test="name() = 'Line'">
				<xsl:variable name="start" select="./lx:Start"/>
				<xsl:variable name="end" select="./lx:End"/>
				<xsl:variable name="add" select="landUtils:AddLineElement($pos, $start, $end)"/>
			</xsl:when>

			<!-- Adding a irregular line call -->
			<xsl:when test="name() = 'IrregularLine'">
				<xsl:variable name="start" select="./lx:Start"/>
				<xsl:variable name="end" select="./lx:End"/>
				<xsl:choose >
					<xsl:when test="./PntList2D">
						<xsl:variable name="add" select="landUtils:AddIrregularLineElement2d($pos, $start, $end, string(./PntList2D))"/>
					</xsl:when>
					<xsl:when test="./PntList3D">
						<xsl:variable name="add" select="landUtils:AddIrregularLineElement3d($pos, $start, $end, string(./PntList3D))"/>
					</xsl:when>
				</xsl:choose>
			</xsl:when>
			
			<!-- Adding a curve call -->
			<xsl:when test="name() = 'Curve'">
				<xsl:variable name="start" select="./lx:Start"/>
				<xsl:variable name="center" select="./lx:Center"/>
				<xsl:variable name="end" select="./lx:End"/>
				<xsl:variable name="rot" select="@rot"/>
				<xsl:variable name="add" select="landUtils:AddCurveElement($pos, $start, $center, $end, string($rot))"/>
			</xsl:when>
			
			<!-- Adding a spiral call -->
			<xsl:when test="name() = 'Spiral'">
				<xsl:variable name="splen" select="@length"/>
				<xsl:variable name="rs" select="@radiusStart"/>
				<xsl:variable name="re" select="@radiusEnd"/>
				<xsl:variable name="rot" select="@rot"/>
				<xsl:variable name="spt" select="@spiType"/>
				<xsl:variable name="start" select="./lx:Start"/>
				<xsl:variable name="pi" select="./lx:PI"/>
				<xsl:variable name="end" select="./lx:End"/>
				<xsl:variable name="add" select="landUtils:AddSpiralElement($pos, $splen, $rs, $re, $rot, $spt, $start, $pi, $end)"/>
			</xsl:when>
			
		</xsl:choose>
		<xsl:choose >
			<xsl:when test="position() = last()">
			</xsl:when>
		</xsl:choose>
	</xsl:for-each>
	
	<!-- <xsl:variable name="trim" select="landUtils:TrimGeomArray()"/> -->
	
</xsl:template>

<msxsl:script language="JScript" implements-prefix="landUtils"> 
<![CDATA[
// -----------------------------------------------------------------
// Global properties
// -----------------------------------------------------------------
var GeomArray;        					// Geometry Element Array
var StationEqArray = new Array(1);   	// Station Equation Array
var PIElementArray;   					// Array of PI elements (VPI and ParaCurve)
var staCurrent;							// used when incrementing stations

// -----------------------------------------------------------------
// Helper functions
// -----------------------------------------------------------------
function Clear()
{
	GeomArray = new Array(1);
	StationEqArray = new Array(1);
	PIElementArray = new Array(1);
	staCurrent = 0;
	return "Done";
}

function iterationTracker(index, range)
{
	var ndx = Number(index);
	var rng = Number(range);
	
	if(ndx <= rng)
	{
		return ndx + 1;
	}
	else
	{
		return String(rng) ;
	}
}

function TrimGeomArray()
{
	var i;
	var stop;
	for(i = 2; i < GeomArray.length; i++)
	{
		var check = GeomArray[i];
		if(check.Northing)
		{
		}
		else
		{
			stop = i;
		}
	}
	var copiedArray = GeomArray;
	GeomArray = new Array(stop);
	for(var c = 0; c < stop; c++)
	{
		GeomArray[c] = copiedArray[c];
	}
	return stop;
}

function SetStationCurrent(station)
{
	staCurrent = parseFloat(station);
	return station;
}

function IncrementStation(increment)
{
	var incr = parseFloat(increment);
	staCurrent = staCurrent + incr;
	
	return "" + staCurrent;
}

function GetIndexOfStation(station)
{
	try
	{
		var fSta = parseFloat(station);
		var i;
		i = 1;
		for(; i < GeomArray.length; i++)
		{
			if( fSta >= GetGeomStartingStation(i))
			{
				return i ;
			}
		}
		return -1;
	}
	catch(e)
	{
		return -1;
	}
}

function GetNorthingAtEqStation(station)
{
	try
	{
		var north = 0;
		var fSta = parseFloat(station);
		
		var eleIndex;
		eleIndex = 0;
		
		var staIndx;
		staIndx = GetIndexOfStation(station);
		
		if(staIndx >= 0)
		{
			eleIndex = staIndx;
		}
		else
		{
			return "Error";
		}
		
		var theEle;
		 theEle = GeomArray[eleIndex];
	
		var dLen = fSta - theEle.StartStation;
		
		if(theEle.type == "Line")
		{
			var spN = GetLineStartNorthing(eleIndex);
			var delta = GetLineDirection(eleIndex);
			north = spN - (dLen * Math.sin((delta * Math.PI) / 180 )) ;
		}
		else if(theEle.type == "Curve")
		{
			var delta = GetCurveDelta(eleIndex);
			var length = GetCurveLength(eleIndex);
			var theta = delta * (dLen / length);
			var rad = GetCurveRadius(eleIndex);
			var cpN = GetCurveCenterNorthing(eleIndex);
			north = cpN - (rad * Math.sin((theta * Math.PI) / 180) );
		}
		
		else if(theEle.type == "Spiral")
		{
			var sLen = GetSpiralLength(eleIndex);
			var sStartRad = GetSpiralStartRadius(eleIndex);
			var sEndRad = GetSpiralEndRadius(eleIndex);
			var sRad = sEndRad;
			if (sRad == "NaN")
			{
				sRad = sStartRad;
			}
			
			var Temp1 = Math.pow(dLen, 3) / (6 * sRad * sLen);
			var Temp2 = Math.pow(dLen, 4) / (56 * Math.pow(sRad, 2) * Math.pow(sLen, 2));
			var Temp3 = Math.pow(dLen, 8) / (7040 * Math.pow(sRad, 4) * Math.pow(sLen, 4));
			var y = spE - (Temp1 * (1 - Temp2 + Temp3));
			north = y;
		}
		return north;
	}
	catch(e)
	{
		return "Error";
	}
}

function GetEastingAtEqStation(station)
{
	try
	{
		var east = 0;
		var fSta = parseFloat(station);
		
		var eleIndex;
		eleIndex = 0;
		
		var staIndx;
		staIndx = GetIndexOfStation(station);
		
		if(staIndx >= 0)
		{
			eleIndex = staIndx;
		}
		else
		{
			return "Error";
		}
		
		var theEle;
		 theEle = GeomArray[eleIndex];
	
		var dLen = fSta - theEle.StartStation;
		
		if(theEle.type == "Line")
		{
			var spE = GetLineStartEasting(eleIndex);
			var delta = GetLineDirection(eleIndex);
			east = spE - (dLen * Math.cos((delta * Math.PI) / 180 )) ;
		}
		else if(theEle.type == "Curve")
		{
			var delta = GetCurveDelta(eleIndex);
			var length = GetCurveLength(eleIndex);
			var theta = delta * (dLen / length);
			var rad = GetCurveRadius(eleIndex);
			var cpE = GetCurveCenterEasting(eleIndex);
			
			east = cpE - (rad * Math.cos((theta * Math.PI) / 180) );
		}
		
		else if(theEle.type == "Spiral")
		{
			var sLen = dLen;
			var sStartRad = GetSpiralStartRadius(eleIndex);
			var sEndRad = GetSpiralEndRadius(eleIndex);
			var sRad = sEndRad;
			if (sRad == "NaN")
			{
				sRad = sStartRad;
			}
			
			var Temp1 = Math.pow(dLen, 4) / (40 * Math.pow(sRad, 2) * Math.pow(sLen, 2));
			var Temp2 = Math.pow(dLen, 8) / (3456 * Math.pow(sRad, 4) * Math.pow(sLen, 4));
			var x =  (dLen * (1 - Temp1 + Temp2));
	
			east = x;
		}
		
		return east;
	}
	catch(e)
	{
		return "Error";
	}
	
}

function GetNorthingAtInternalStation(station)
{
		var north = 0;
		var fSta = parseFloat(station);
		var i;
		var eleIndex;
		eleIndex = 0;

		var staIndx;
		staIndx = GetIndexOfStation(station);
		
		if(staIndx >= 0)
		{
			eleIndex = staIndx;
		}
		else
		{
			return "Error";
		}

		var theEle = GeomArray[eleIndex];
		var dLen = fSta - theEle.InternalStartStation;
		
		if(theEle.type == "Line")
		{
			var spN = GetLineStartNorthing(eleIndex);
			var delta = GetLineDirection(eleIndex);
			north = spN - (dLen * Math.sin((delta * Math.PI) / 180 )) ;
		}
		else if(theEle.type == "Curve")
		{
			var delta = GetCurveDelta(eleIndex);
			var length = GetCurveLength(eleIndex);
			var theta = delta * (dLen / length);
			var rad = GetCurveRadius(eleIndex);
			var cpN = GetCurveCenterNorthing(eleIndex);
			north = cpN - (rad * Math.sin((theta * Math.PI) / 180) );
		}
		
		else if(theEle.type == "Spiral")
		{
			var sLen = GetSpiralLength(eleIndex);
			var sStartRad = GetSpiralStartRadius(eleIndex);
			var sEndRad = GetSpiralEndRadius(eleIndex);
			var sRad = sEndRad;
			if (sRad == "NaN")
			{
				sRad = sStartRad;
			}
			
			var Temp1 = Math.pow(dLen, 3) / (6 * sRad * sLen);
			var Temp2 = Math.pow(dLen, 4) / (56 * Math.pow(sRad, 2) * Math.pow(sLen, 2));
			var Temp3 = Math.pow(dLen, 8) / (7040 * Math.pow(sRad, 4) * Math.pow(sLen, 4));
			var y = spE - (Temp1 * (1 - Temp2 + Temp3));
			north = y;
		}
		return north;
}

function GetEastingAtInternalStation(station)
{
		var east = 0;
		var fSta = parseFloat(station);
		var i;
		var eleIndex;
		eleIndex = 0;
		
		var staIndx;
		staIndx = GetIndexOfStation(station);
		
		if(staIndx >= 0)
		{
			eleIndex = staIndx;
		}
		else
		{
			return "Error";
		}

		var theEle = GeomArray[eleIndex];
		var dLen = fSta - theEle.InternalStartStation;
		
		if(theEle.type == "Line")
		{
			var spE = GetLineStartEasting(eleIndex);
			var delta = GetLineDirection(eleIndex);
			east = spE - (dLen * Math.cos((delta * Math.PI) / 180 )) ;
		}
		else if(theEle.type == "Curve")
		{
			var delta = GetCurveDelta(eleIndex);
			var length = GetCurveLength(eleIndex);
			var theta = delta * (dLen / length);
			var rad = GetCurveRadius(eleIndex);
			var cpE = GetCurveCenterEasting(eleIndex);
			
			east = cpE - (rad * Math.cos((theta * Math.PI) / 180) );
		}
		
		else if(theEle.type == "Spiral")
		{
			var sLen = dLen;
			var sStartRad = GetSpiralStartRadius(eleIndex);
			var sEndRad = GetSpiralEndRadius(eleIndex);
			var sRad = sEndRad;
			if (sRad == "NaN")
			{
				sRad = sStartRad;
			}
			
			var Temp1 = Math.pow(dLen, 4) / (40 * Math.pow(sRad, 2) * Math.pow(sLen, 2));
			var Temp2 = Math.pow(dLen, 8) / (3456 * Math.pow(sRad, 4) * Math.pow(sLen, 4));
			var x =  (dLen * (1 - Temp1 + Temp2));
	
			east = x;
		}
		
		return east;
}

function GetStartRegion(index)
{
	var geomElement = GeomArray[index];
	return "" + geomElement.StartRegion;
}

function GetEndRegion(index)
{
	var geomElement = GeomArray[index];
	return "" + geomElement.EndRegion;
}

// -----------------------------------------------------------------
// PI Array functions
// -----------------------------------------------------------------

function ProcessPIPoints()
{
	PIElementArray = new Array(1);
	
	// Add the start of the PI array (start of the first element)
	var firstElement = GeomArray[1];
	
	var firstN = firstElement.StartN;
	var firstE = firstElement.StartE;
	var firstSta = firstElement.StartStation;
	var firstPI = new PIElement(firstN, firstE, firstSta);

	PIElementArray.push(firstPI);
	
	// variables for SCS combination
	var startSpiral;
	var spiralCurve;
	var endSpiral;
	
	// Iterate the Geometry elements to store PI's (always looking ahead)
	var i;
	var isCombo = 0;
	var sumLen = 0;

	for( i = 1; i < GeomArray.length; i++)
	{
		var gele = GeomArray[i];
		var type = gele.type;
		
		if(type == "Line")
		{
			if(i < GeomArray.length - 2)
			{
				isCombo = 0;
				var nele = GeomArray[i + 1];
				if(nele.type == "Line")
				{
					var piLine = new PIElement(gele.EndN, gele.EndE, gele.InternalEndStation);
					sumLen = sumLen + GetLineLength(i);

					PIElementArray.push(piLine);
				}
			}
		}
		else if(type == "Curve")
		{
			if(isCombo == 1)
			{
				spiralCurve = gele;
			}
			else
			{
				var pN = GetCurvePINorthing(i);
				var pE = GetCurvePIEasting(i);
				var curLen = GetCurveLength(i);
				
				var pSta = gele.InternalStartStation + (curLen / 2);
				
				var cpElement = new PIElement(pN, pE, pSta);
				PIElementArray.push(cpElement);
				sumLen = sumLen + GetCurveLength(i);
			}
		}
		else if(type == "Spiral")
		{
			if(i < GeomArray - 2)
			{
				if(isCombo == 1)
				{
					endSpiral = gele;
					var stN = startSpiral.StartN;
					var stE = startSpiral.StartE;
					var stDir = CalculateDirection(startSpiral.StartN, startSpiral.StartE, startSpiral.piN, startSpiral.piE);
					
					var enN = startSpiral.EndN;
					var enE = startSpiral.EndE;
					var enDir = CalculateDirection(startSpiral.EndN, startSpiral.EndE, startSpiral.piN, startSpiral.piE);
					
					var pn = GetIntersectionNorthing();
					var pe = GetIntersectionEasting();
					
					var staAdd = startSpiral.length + GetCurveLength(i - 1) + endSpiral.length;
					
					var sta = sumLen + (staAdd / 2);
					
					var spelement = new PIElement(pn, pe, sta);
					PIElementArray.push(spelement);
					
					isCombo = 0;
				}
				else
				{
					var nextElement = GeomArray[i + 1];
					if(nextElement.type == "Curve")
					{
						startSpiral = gele;
						isCombo = 1;
					}
					else
					{
						var pN = GetSpiralPINorthing(i);
						var pE = GetSpiralPIEasting(i);
						var pSta = GetSpiralLength(i);
				
						var pElement = new PIElement(pN, pE, pSta);
						PIElementArray.push(pElement);
						sumLen = sumLen + GetSpiralLength(i);
					}
				}
			}
			else
			{
				var pN = GetSpiralPINorthing(i);
				var pE = GetSpiralPIEasting(i);
				var pSta = gele.InternalEndStation;
				
				var pElement = new PIElement(pN, pE, pSta);
				PIElementArray.push(pElement);
			}
			
		}
			if(i == GeomArray.length - 1)
			{
				var en = gele.EndN;
				var ee = gele.EndE;
				var eSta = gele.InternalEndStation;
				
				var endPI = new PIElement(en, ee, eSta);
				PIElementArray.push(endPI);
			}
	}
	
	// Might as well return the length of the new array for debugging purposes
	return PIElementArray.length;
}

function GetPINorthingAt(index)
{
	var piNdx = Number(index);
	var pElement = PIElementArray[piNdx];
	if(pElement)
	{
		return String(pElement.Northing);
	}
	else
	{
		return String(pElement);
	}
	return index;
}
function GetPIEastingAt(index)
{
	var ndx = Number(index);
	var piElement = PIElementArray[index];
	if(piElement)
	{
		return piElement.Easting.toString();
	}
	else
	{
		return 0;
	}
}
function GetPIStationAt(index)
{
	var ndx = Number(index);
	var piElement = PIElementArray[index];
	if(piElement)
	{
		return piElement.AheadStation.toString();
	}
	else
	{
		return 0;
	}
}

function GetPILength(index)
{
	var ndx = Number(index);
	
	if(ndx < PIElementArray.length - 1)
	{
		var thisPI = PIElementArray[ndx];
		var nextPI = PIElementArray[ndx + 1];
		
		return CalculateLength(thisPI.Northing,  thisPI.Easting, nextPI.Northing, nextPI.Easting);
	}
	else
	{
		return 0;
	}
}

function GetPIDirection(index)
{
	var ndx = Number(index);
	
	if(ndx > 0 && ndx < PIElementArray.length - 1)
	{
		var thisPI = PIElementArray[ndx];
		var nextPI = PIElementArray[ndx + 1];
		
		return CalculateDirection(thisPI.Northing,  thisPI.Easting, nextPI.Northing, nextPI.Easting);
	}
	else
	{
		return 0;
	}
}


function extendPIArray()
{
	var copiedArray = PIElementArray;
	var piArrayLength = copiedArray.length;
	
	PIElementArray = new Array(1);
	for(var i = 1; i < piArrayLength; i++)
	{
		PIElementArray.push(copiedArray[i]);
	}
	return PIElementArray.length - 1;
}

function GetPIArraySize()
{
	var piArrSize = PIElementArray.length;
	return PIElementArray.length - 1;
}
// -----------------------------------------------------------------
// Station Equation Array functions
// -----------------------------------------------------------------
function GetStationArrayLength()
{
	return StationEqArray.length - 1;
}

function AddStationEquation(internalStation, backStation, aheadStation)
{
	var is = internalStation;
	var bs = backStation;
	var as = aheadStation;
	var staEq = new StationEquation(is, bs, as);
	StationEqArray.push(staEq);
	return StationEqArray.length;
}

function GetEquationInnerStation(index)
{
	var equation = StationEqArray[index];
	if(equation == null)
	{
		return "Not a station equation";
	}
	return equation.InternalStation;
}

function GetEquationBackStation(index)
{
	var equation = StationEqArray[index];
	if(equation == null)
	{
		return "Not a station equation";
	}
	return equation.BackStation;
}

function GetEquationAheadStation(index)
{
	var equation = StationEqArray[index];
	if(equation == null)
	{
		return "Not a station equation";
	}
	return equation.AheadStation;
}

function ApplyStationing(startStation)
{
	StationEqArray = new Array(1);
	PIElementArray = new Array(1);
	var staStartStr = NodeToText(startStation);
	var accumLength = new Number(parseFloat(staStartStr));
	for(var e = 1; e < GeomArray.length; e++)
	{
		var geomElement = GeomArray[e];
		
		GeomArray[e].StartRegion = 1;
		GeomArray[e].EndRegion = 1;
		
		if(geomElement.type == "Line")
		{
			var lineLenStr = GetLineLength(e);
			var lineLen = new Number(parseFloat(lineLenStr));
			
			geomElement.StartStation = accumLength;
			geomElement.InternalStartStation = accumLength;
			geomElement.BackStartStation = accumLength;
			
			geomElement.PIStation = accumLength;
			geomElement.InternalPIStation = accumLength;
			geomElement.BackPIStation = accumLength;
			
			accumLength = accumLength + lineLen;
			
			geomElement.EndStation = accumLength;
			geomElement.InternalEndStation = accumLength;
			geomElement.BackEndStation = accumLength;
		}
		else if(geomElement.type == "Curve")
		{
			var curLenStr = GetCurveLength(e);
			var curLength = new Number(parseFloat(curLenStr));
			
			geomElement.StartStation = accumLength;
			geomElement.InternalStartStation = accumLength;
			geomElement.BackStartStation = accumLength;

			geomElement.PIStation = accumLength + curLength / 2;
			geomElement.InternalPIStation = accumLength + curLength / 2;
			geomElement.BackPIStation = accumLength + curLength / 2;

			accumLength = accumLength + curLength ;
			
			geomElement.EndStation = accumLength;
			geomElement.InternalEndStation = accumLength;
			geomElement.BackEndStation = accumLength;
		}
		else
		{
			var spiLengthStr = GetSpiralLength(e);
			var spiLength = new Number(parseFloat(spiLengthStr ));
			
			geomElement.StartStation = accumLength;
			geomElement.InternalStartStation = accumLength;
			geomElement.BackStartStation = accumLength;
			
			geomElement.PIStation = accumLength + spiLength / 2;
			geomElement.InternalPIStation = accumLength + spiLength / 2;
			geomElement.BackPIStation = accumLength + spiLength / 2;

			accumLength = accumLength + spiLength ;
			
			geomElement.EndStation = accumLength;
			geomElement.InternalEndStation = accumLength;
			geomElement.BackEndStation = accumLength;
		}
	}		
	ProcessPIPoints();
	return "Done";
}

function ApplyStationEquations()
{
	var i;
	var eqDiff = 0;
	for(i = 1; i < StationEqArray.length; i++)
	{
		eqDiff = 0;
		var equation = StationEqArray[i];
		if(equation ==  null)
		{
			return "No Equations";
		}
		var eqInternal = equation.InternalStation;
		var eqBackStation = equation.BackStation;
		var eqAheadStation = equation.AheadStation;
		eqDiff = eqInternal - eqAheadStation;
		
		var e;
		for(e = 1; e < GeomArray.length; e++)
		{
			var geomElement = GeomArray[e];
			var InternalStartStation = geomElement.InternalStartStation;
			var InternalPIStation = geomElement.InternalPIStation;
			var InternalEndStation = geomElement.InternalEndStation;
			
			var geomStartStation = geomElement.StartStation;
			var geomPIStation  = geomElement.PIStation;
			var geomEndStation = geomElement.EndStation;
			
			if(InternalStartStation >= eqInternal)
			{
				geomElement.StartStation = geomElement.InternalStartStation - eqDiff;
				GeomArray[e].StartRegion = i + 1;
			}
			if(InternalPIStation >= eqInternal)
			{
				geomElement.PIStation = geomElement.InternalPIStation - eqDiff;
			}
			if(InternalEndStation >= eqInternal)
			{
				geomElement.EndStation = geomElement.InternalEndStation - eqDiff;
				GeomArray[e].EndRegion = i + 1;
			}
		}		
	}
	ApplyStaEquationsToPI();
	return eqDiff.toString();
}
function ApplyStaEquationsToPI()
{
	for(var i = 1; i < StationEqArray.length; i++)
	{
		var staEq = StationEqArray[i];
		var inSta = staEq.InternalStation;
		var ahSta = staEq.AheadStation;
		var eqDiff = ahSta - inSta;
		
		for(var p = 1; p < PIElementArray.length; p++)
		{
			var piEle = PIElementArray[p];
			var piSta = piEle.Station;
			
			if(piSta > inSta)
			{
				piEle.AheadStation = piEle.Station + eqDiff;
			}
		}
	}
}

function GetGeomStartingStation(index)
{
	var geom = GeomArray[index];
	if(geom.StartStation)
	{
		return geom.StartStation.toString();
	}
	else
	{
		return 0;
	}
}
function GetGeomPIStation(index)
{
	var geom = GeomArray[index];
	return geom.PIStation.toString();
}

function GetGeomEndStation(index)
{
	var geom = GeomArray[index];
	return geom.EndStation.toString();
}

function GetGeomStartNorthing(index)
{
	var ndx = Number(index);
	var geom = GeomArray[ndx];
	var geoType = geom.type;
	
	if(geoType == "Line")
	{
		var line = GeomArray[ndx];
		return Number(line.StartN);
	}
	else if(geoType == "Curve")
	{
		var curve = GeomArray[ndx];
		return Number(curve.StartN);
	}
	else(geoType == "Spiral")
	{
		var spiral = GeomArray[ndx];
		return Number(spiral.StartN);
	}
}
function GetGeomStartEasting(index)
{
	var ndx = Number(index);
	var geom = GeomArray[ndx];
	var geoType = geom.type;
	
	if(geoType == "Line")
	{
		var line = GeomArray[ndx];
		return Number(line.StartE);
	}
	else if(geoType == "Curve")
	{
		var curve = GeomArray[ndx];
		return Number(curve.StartE);
	}
	else(geoType == "Spiral")
	{
		var spiral = GeomArray[ndx];
		return Number(spiral.StartE);
	}
}
// PI point coordinates
function GetGeomPINorthing(index)
{
	var ndx = Number(index);
	var geom = GeomArray[ndx];
	var geoType = geom.type;
	
	if(geoType == "Line")
	{
		var line = GeomArray[ndx];
		return Number(line.piN);
	}
	else if(geoType == "Curve")
	{
		var curve = GeomArray[ndx];
		return Number(curve.piN);
	}
	else(geoType == "Spiral")
	{
		var spiral = GeomArray[ndx];
		return Number(spiral.piN);
	}
}
function GetGeomPIEasting(index)
{
	var ndx = Number(index);
	var geom = GeomArray[ndx];
	var geoType = geom.type;
	
	if(geoType == "Line")
	{
		var line = GeomArray[ndx];
		return Number(line.piE);
	}
	else if(geoType == "Curve")
	{
		var curve = GeomArray[ndx];
		return Number(curve.piE);
	}
	else(geoType == "Spiral")
	{
		var spiral = GeomArray[ndx];
		return Number(spiral.piE);
	}
}

// End point coordinates
function GetGeomEndNorthing(index)
{
	var ndx = Number(index);
	var geom = GeomArray[ndx];
	var geoType = geom.type;
	
	if(geoType == "Line")
	{
		var line = GeomArray[ndx];
		return Number(line.EndN);
	}
	else if(geoType == "Curve")
	{
		var curve = GeomArray[ndx];
		return Number(curve.EndN);
	}
	else(geoType == "Spiral")
	{
		var spiral = GeomArray[ndx];
		return Number(spiral.EndN);
	}
}
function GetGeomEndEasting(index)
{
	var ndx = Number(index);
	var geom = GeomArray[ndx];
	var geoType = geom.type;
	
	if(geoType == "Line")
	{
		var line = GeomArray[ndx];
		return Number(line.EndE);
	}
	else if(geoType == "Curve")
	{
		var curve = GeomArray[ndx];
		return Number(curve.EndE);
	}
	else(geoType == "Spiral")
	{
		var spiral = GeomArray[ndx];
		return Number(spiral.EndE);
	}
}

// -----------------------------------------------------------------
// Geometry Array functions
// -----------------------------------------------------------------
function SetGeomArrayLength(arrLength)
{
	var nLen = Number(arrLength);
	GeomArray = new Array(1);
	return arrLength;
}

function GetGeomArrayLength()
{
	return GeomArray.length - 1;
}

// -----------------------------------------------------------------
// Add Elements to the array functions
// -----------------------------------------------------------------

function AddLineElement(index, start, end)
{
	var ndx = Number(index);
	
	var sText = NodeToText(start);
	var eText = NodeToText(end);
	
	var strCoords = sText.split(" ");
	var sNorth = strCoords [0];
	var sEast = strCoords [1];
	
	var strECoords = eText.split(" ");
	var eNorth = strECoords [0];
	var eEast = strECoords [1];

	var line = new LineElement(sNorth, sEast, eNorth, eEast);
	GeomArray.push(line);
	return sNorth ;
}

function AddIrregularLineElement2d(start, end, pntlist)
{
	var sText = NodeToText(start);
	var eText = NodeToText(end);
	
	var strCoords = sText.split(" ");
	var sNorth = strCoords [0];
	var sEast = strCoords [1];
	
	var strECoords = eText.split(" ");
	var eNorth = strECoords [0];
	var eEast = strECoords [1];
	
	var irrLine = new IrregularLine2d(sNorth, sEast, eNorth, eEast);
	GeomArrary.push(irrLine);
	return sNorth;
}

function AddIrregularLineElement3d(start, end, pntlist)
{
	var sText = NodeToText(start);
	var eText = NodeToText(end);
	
	var strCoords = sText.split(" ");
	var sNorth = strCoords [0];
	var sEast = strCoords [1];
	
	var strECoords = eText.split(" ");
	var eNorth = strECoords [0];
	var eEast = strECoords [1];
	
	var irrLine = new IrregularLine3d(sNorth, sEast, eNorth, eEast);
	GeomArrary.push(irrLine);
	return sNorth;
}

function AddCurveElement(index, start, center, end, rot)
{
	var ndxStr = NodeToText(index);
	var ndx = new Number(parseInt(ndxStr));
	
	var sText = NodeToText(start);
	var eText = NodeToText(end);
	var cText = NodeToText(center);
	
	var stCoords = sText.split(" ");
	var cenCoords = cText.split(" ");
	var endCoords = eText.split(" ");
	
	var curve  = new CurveElement(stCoords[0], stCoords[1], cenCoords[0], cenCoords[1], endCoords[0], endCoords[1], rot);
	GeomArray.push(curve);
	
	return index;
}

function AddSpiralElement(index, length, radiusStart, radiusEnd, rotation, spiType, start, pi, end)
{
	var ndxStr = NodeToText(index);
	var lenStr = NodeToText(length);
	var radStStr = NodeToText(radiusStart);
	var redEnStr = NodeToText(radiusEnd);
	var rotStr = NodeToText(rotation);
	var spiStr = NodeToText(spiType);
	var ndx = new Number(parseInt(ndxStr));
	
	var sText = NodeToText(start);
	var pText = NodeToText(pi);
	var eText = NodeToText(end);
	
	var stCoords = sText.split(" ");
	var pCoords = pText.split(" ");
	var endCoords = eText.split(" ");
	
	
	var spiral = new SpiralElement(lenStr, radStStr , redEnStr , rotStr , spiStr , stCoords[0], stCoords[1], pCoords[0], pCoords[1], endCoords[0], endCoords[1]);
	GeomArray.push(spiral);

	return index;
}

// -----------------------------------------------------------------
// Retrieve from Geometry Array functions
// -----------------------------------------------------------------

function GetGeomElementType(index)
{
	var ndx = new Number(parseInt(index));
	var Num_of_Elements = GetGeomArrayLength();
	
	if (ndx < 1)
		return "Line";

	if (ndx > Num_of_Elements)
		return "Line";

	var geom = GeomArray[ndx];
	
	return geom.type ;
}

function GetGeomElementLength(index)
{
	var ndx = new Number(parseInt(index));
	var geom = GeomArray[index];
	
	if(geom.type == "Line")
	{
		return GetLineLength(index);
	}
	else if(geom.type == "Curve")
	{
		return GetCurveLength(index);
	}
	else
	{
		return GetSpiralLength(index);
	}
}

// -----------------------------------------------------------------
// Line element functions
// -----------------------------------------------------------------

// Starting point coordinates
function GetLineStartNorthing(index)
{
	var line = GeomArray[index];
	if(line.type != "Line")
	{
		return "Not a Line";
	}
	return line.StartN.toString();
}
function GetLineStartEasting(index)
{
	var line = GeomArray[index];
	if(line.type != "Line")
	{
		return "Not a Line";
	}
	return line.StartE.toString();
}

// End point coordinates
function GetLineEndNorthing(index)
{
	var line = GeomArray[index];
	if(line.type != "Line")
	{
		return "Not a Line";
	}
	return line.EndN.toString();
}
function GetLineEndEasting(index)
{
	var line = GeomArray[index];
	if(line.type != "Line")
	{
		return "Not a Line";
	}
	return line.EndE.toString();
}

function GetLineLength(index)
{
	var lineLength = 0;
	var line = GeomArray[index];
	
	var xDiff = line.EndE -  line.StartE ;
   	var yDiff = line.EndN - line.StartN ;
    	
	lineLength = Math.sqrt(Math.pow(xDiff, 2) + Math.pow(yDiff, 2));
	return lineLength;
}

function GetLineDirection(index)
{
	var line = GeomArray[index];
	if(line.type != "Line")
	{
		return "Not a Line";
	}
	var calcDir = CalculateDirection(line.StartN, line.StartE, line.EndN, line.EndE);
	return calcDir;
}


// -----------------------------------------------------------------
// Curve element functions
// -----------------------------------------------------------------
function GetCurveRotation(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	return curve.Rot;
}

// Starting point coordinates
function GetCurveStartNorthing(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	return curve.StartN.toString();
}
function GetCurveStartEasting(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	return curve.StartE.toString();
}

// Center point coordinates
function GetCurveCenterNorthing(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	
	return curve.CenterN.toString();
}
function GetCurveCenterEasting(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	
	return curve.CenterE.toString();
}

// End point coordinates
function GetCurveEndNorthing(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	return curve.EndN.toString();
}
function GetCurveEndEasting(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	return curve.EndE.toString();
}

function CheckCurve(index) //Sometimes curve data is reversed in the XML report
{
	var geomCount = GetGeomArrayLength();
	if (geomCount <= 1)
	{
		return 1;
	}

	var curve = GeomArray[index];
	
	var sn = curve.StartN;
	var se = curve.StartE;
	var en = curve.EndN;
	var ee = curve.EndE;
	if (index > 1)
	{
		var PreviousElement = GeomArray[index - 1];
		var LastN = PreviousElement.EndN;
		var LastE = PreviousElement.EndE;
		if ((sn != LastN) || (se != LastE)) //Start of this element and end of previous one should be the same
			return 1;
		else
			return 0;
	}
	else
	{
		var NextElement = GeomArray[index + 1];
		var NextN = NextElement.StartN;
		var NextE = NextElement.StartE;
		if ((en != NextN) || (ee != NextE)) //Start of next element and end of this one should be the same
			return 1;
		else
			return 0;
	}
}

// PI point coordinates (calculated)
function GetCurvePINorthing(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	
	var Is_Curve_Data_OK = CheckCurve(index);
	if (Is_Curve_Data_OK == 0) //If data was reversed in the XML file
	{
		var tempN = curve.StartN;
		var tempE = curve.StartE;
		curve.StartN = curve.EndN;
		curve.StartE = curve.EndE;
		curve.EndN = tempN;
		curve.EndE = tempE;
	}
	
	var tanLen = GetCurveTangent(index);
	
	//var ang = GetCurveAngle(index);
	//curvePINorthing =curve.StartN - (tanLen * Math.cos((ang * Math.PI) / 180 ));
	//return curvePINorthing.toString();
	
	var ang = 	CalculateDirection(curve.CenterN, curve.CenterE, curve.StartN, curve.StartE);
	var mAngle = 0;
	if(curve.Rot == "cw")
	{
		mAngle = Math.sin(((ang + 90) * Math.PI) / 180 );

	}
	else
	{
		mAngle = Math.sin(((ang - 90) * Math.PI) / 180 );

	}
	return curve.StartN - (tanLen * mAngle);
}
function GetCurvePIEasting(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	
	var Is_Curve_Data_OK = CheckCurve(index);
	if (Is_Curve_Data_OK == 0) //If data was reversed in the XML file
	{
		var tempN = curve.StartN;
		var tempE = curve.StartE;
		curve.StartN = curve.EndN;
		curve.StartE = curve.EndE;
		curve.EndN = tempN;
		curve.EndE = tempE;
	}

	var tanLen = GetCurveTangent(index);
	
	//var ang = GetCurveAngle(index)
	//curvePINorthing =curve.StartN - (tanLen * Math.sin((ang * Math.PI) / 180 ));
	//return curvePINorthing.toString();
	
	var ang = 	CalculateDirection(curve.CenterN, curve.CenterE, curve.StartN, curve.StartE);
	var mAngle = 0;
	if(curve.Rot == "cw")
	{
		mAngle = Math.cos(((ang + 90) * Math.PI) / 180 );
	}
	else
	{
		mAngle = Math.cos(((ang - 90) * Math.PI) / 180 );
	}
	
	return curve.StartE - (tanLen * mAngle);
}

function SetCurvePIStation(station)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	curve.PIStation = station;
}

function GetCurvePIStation(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	return curve.PIStation.toString();
}

function GetCurveAngle(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	var rotation = curve.Rot;
	
	var curveAngle = 0;
	var startDir = 0;
	var endDir = 0;
	
	var startXDiff = curve.CenterE - curve.StartE;
	var startYDiff = curve.CenterN - curve.StartN;
	
	var endXDiff = curve.CenterE - curve.EndE;
	var endYDiff = curve.CenterN - curve.EndN;
		
	var startTan =startXDiff / startYDiff  ;
	var endTan = endXDiff / endYDiff;
	
	var startDirTmp = Math.atan(startTan);
	var endDirTmp = Math.atan(endTan);
	
	    	if(curve.CenterN >= curve.StartN)
    		{
     		startDir = 270 - ( (startDirTmp * 180) / Math.PI); 
   		}
    		else
    		{
    			startDir =90 - ( (startDirTmp * 180) / Math.PI) ;
    		}
    		
    		
    		
    		if(curve.CenterN >= curve.EndN)
    		{
    			endDir= 270 -  ( (endDirTmp * 180) / Math.PI);    		
  		}
    		else
    		{
    			endDir= 90 - ( (endDirTmp * 180) / Math.PI);
 		}
    	
    		
    	if(rotation == "cw")
		{
		
			if(startDir<endDir)
			{
				curveAngle = 360.0 + startDir - endDir ;
			}
			else
			{
				curveAngle = (startDir - endDir);
			}
		}
		
	else
		{
			if(startDir > endDir)
			{
				curveAngle = 360 - (startDir - endDir) ;
			}
			else
			{
				curveAngle = endDir - startDir;
			}
		}
	return (curveAngle ); 
}
function GetCurveRadius(index)
{
	var curve = GeomArray[index];
	var curveRadius = 0;
	
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	
	var yDiff = curve.StartN - curve.CenterN;
	var xDiff = curve.StartE - curve.CenterE;
	
	curveRadius = Math.sqrt(Math.pow(xDiff, 2) + Math.pow(yDiff, 2));
	return curveRadius;
}
function GetChordLength(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}

	var cRadius = GetCurveRadius(index);
	var cAngle = GetCurveAngle(index);
	return 2 * cRadius * Math.sin(.017453293 * (cAngle / 2));
}
function GetChordDirection(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}

	var angle = CalculateDirection(curve.StartN, curve.StartE, curve.EndN, curve.EndE);
	return angle;
}
function GetCurveTangent(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	var cRadius = GetCurveRadius(index);
	var cAngle = GetCurveAngle(index);

	return Math.abs(cRadius * (Math.tan(.017453293 * (cAngle / 2))));
}

function GetCurveLength(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return 0;
	}
	var cRadius = GetCurveRadius(index);
	var cAngle = GetCurveAngle(index);

	return (Math.PI * cRadius * cAngle) / 180;
}

function GetCurveStartDirection(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return 0;
	}
	var dir = CalculateDirection(curve.CenterN, curve.CenterE, curve.StartN, curve.StartE);
	return dir;
}

function GetCurveEndDirection(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return 0;
	}
	var dir = CalculateDirection(curve.CenterN, curve.CenterE, curve.EndN, curve.EndE);
	return dir;
}

function GetCurveRotation(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return 0;
	}
	return curve.Rot.toString();
}

function GetCurveExternal(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	var cRadius = GetCurveRadius(index);
	var cAngle = GetCurveAngle(index);
	
	var tan = cRadius * (Math.tan(.017453293 * (cAngle / 2)));
	var ext = tan * Math.tan(.017453293 * (cAngle / 4));
	return Math.abs(ext);
}

function GetCurveMiddle(index)
{
	var curve = GeomArray[index];

	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	var cAngle = GetCurveAngle(index);
	var cRadius = GetCurveRadius(index);
	
	var mid = cRadius *  (1 - (Math.cos(.017453293 * (cAngle / 2))));
	return mid;
}
function GetCurveDOC(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	var cRadius = GetCurveRadius(index);
	var doc = 5729.578 / cRadius;
	return doc;
}

function ComparePreviousCurveDirection(index)
{
	if (index < 2)
	{
		return "0";
	}
	
	var curve = GeomArray[index];

	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}

	var LastRot = GetCurveRotation[index-1];
	var ThisRot = GetCurveRotation[index];
	if (LastRot == ThisRot)
	{
		return "Compound";
	}
	else
	{
		return "Reverse";
	}
}

function CompareNextCurveDirection(index)
{
	var Number_of_Elements = GetGeomArrayLength();

	if (index >= Number_of_Elements)
	{
		return "0";
	}
	
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}

	var NextRot = GetCurveRotation[index+1];
	var ThisRot = GetCurveRotation[index];
	if (NextRot == ThisRot)
	{
		return "Compound";
	}
	else
	{
		return "Reverse";
	}
}


function GetCurveES(index)
{
	var curve = GeomArray[index];
	if(curve.type != "Curve")
	{
		return "Not a Curve";
	}
	var cRadius = GetCurveRadius(index);
	var delta = GetCurveAngle(index);

	var Temp1 = Math.cos(delta / 2);
	var Secant = 1 / Temp1;
	var Es = cRadius * (Secant - 1);
	return Es;
}

function GetCurveBackDirection(index)
{
	var startDir = GetCurveStartDirection(index);
	var adjDir = 0;
	
	var rot = GetCurveRotation(index);
	if(rot == "cw")
	{
		adjDir = startDir + 90;
	}
	else
	{
		adjDir = startDir - 90;
	}
	
	if(adjDir < 0)
	{
		adjDir = 360 + adjDir;
	}
	return adjDir;
}

function GetCurveAheadDirection(index)
{
	var endDir = GetCurveEndDirection(index);
	var adjDir = 0;
	
	var rot = GetCurveRotation(index);
	if(rot == "cw")
	{
		adjDir = endDir - 90;
	}
	else
	{
		adjDir = endDir + 90;
	}
	if(adjDir < 0)
	{
		adjDir = 360 + adjDir;
	}
	return adjDir;
}

function GetDirectionOfConcavity(index)
{
	var chordDir = GetChordDirection(index);
	var rot = GetCurveRotation(index);
	
	if(chordDir < 45)
	{
		if(rot == "cw")
		{
			return "Northerly";
		}
		else if(rot == "ccw")
		{
			return "Southerly";
		}
		else
		{
			return "Rotation undetermined";
		}
	}
	else if(chordDir >= 45 & chordDir < 135)
	{
		if(rot == "cw")
		{
			return "Easterly";
		}
		else if(rot == "ccw")
		{
			return "Westerly"
		}
		else
		{
			return "Rotation undetermined";
		}
	} 
	else if(chordDir >= 135 & chordDir < 225)
	{
		if(rot == "cw")
		{
			return "Westerly";
		}
		else if(rot == "ccw")
		{
			return "Easterly"
		}
		else
		{
			return "Rotation undetermined";
		}
	}
	else
	{
		if(rot == "cw")
		{
			return "Southerly";
		}
		else if(rot == "ccw")
		{
			return "Northerly";
		}
		else
		{
			return "Rotation undetermined";
		}
	}
	return "ConcavityDir";
}

function GetDirectionOfCurveRun(index)
{
	var runDir = GetCurveBackDirection(index);
	var rot = GetCurveRotation(index);
	
	if(runDir < 45)
	{
		if(rot == "cw")
		{
			return "Northerly";
		}
		else if(rot == "ccw")
		{
			return "Southerly";
		}
		else
		{
			return "Rotation undetermined";
		}
	}
	else if(runDir >= 45 & runDir < 135)
	{
		if(rot == "cw")
		{
			return "Easterly";
		}
		else if(rot == "ccw")
		{
			return "Westerly"
		}
		else
		{
			return "Rotation undetermined";
		}
	} 
	else if(runDir >= 135 & runDir < 225)
	{
		if(rot == "cw")
		{
			return "Westerly";
		}
		else if(rot == "ccw")
		{
			return "Easterly"
		}
		else
		{
			return "Rotation undetermined";
		}
	}
	else
	{
		if(rot == "cw")
		{
			return "Southerly";
		}
		else if(rot == "ccw")
		{
			return "Northerly";
		}
		else
		{
			return "Rotation undetermined";
		}
	}
	return "CurveRunDir";
}

// -----------------------------------------------------------------
// Spiral element functions
// -----------------------------------------------------------------
function GetSpiralLength(index)
{
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return 0;
	}
	return spiral.length.toString();
}

function GetSpiralDirection(index)
{
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return 0;
	}
	
	var angle = CalculateDirection(spiral.StartN, spiral.StartE, spiral.EndN, spiral.EndE);
	return angle;
}
function GetSpiralStartRadius(index)
{
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return "Not a spiral";
	}
	return spiral.radiusStart;
}
function GetSpiralEndRadius(index)
{
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return "Not a spiral";
	}
	return spiral.radiusEnd;
}
function GetSpiralRadius(index) //Returns the non-infinite radius, or the smaller of the two
{
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return "Not a spiral";
	}
	
	var Start = GetSpiralStartRadius(index);
	var End = GetSpiralEndRadius(index);
	if (Start !='INF' && End !='INF') {
	   Start=Number(GetSpiralStartRadius(index));
	   End=Number(GetSpiralEndRadius(index));
  	 if (Start < End)
     		return Start;
      else
      	return End;
	}else if (Start =='INF') {
  	return End;
  }else {
    return Start
  }
		
}
function GetSpiralRotation(index)
{
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return "Not a spiral";
	}
	return spiral.rotation;
}
function GetSpiralType(index)
{
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return "Not a spiral";
	}
	return spiral.spiType;
}
function GetSpiralStartNorthing(index)
{
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return "Not a spiral";
	}
	return spiral.StartN.toString();
}
function GetSpiralStartEasting(index)
{
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return "Not a spiral";
	}
	return spiral.StartE.toString();
}
function GetSpiralPINorthing(index)
{
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return "Not a spiral";
	}
	return spiral.piN.toString();
}
function GetSpiralPIEasting(index)
{
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return "Not a spiral";
	}
	return spiral.piE.toString();
}
function GetSpiralEndNorthing(index)
{
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return "Not a spiral";
	}
	return spiral.EndN.toString();
}
function GetSpiralEndEasting(index)
{
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return "Not a spiral";
	}
	return spiral.EndE.toString();
}

function GetSpiralChordLength(index)
{
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return "Not a spiral";
	}

   	var xDiff = spiral.StartE - spiral.EndE;
   	var yDiff = spiral.StartN - spiral.EndN;
   	return Math.sqrt(Math.pow(xDiff, 2) + Math.pow(yDiff, 2));
   	var tanA = yDiff / xDiff;

	return (2 * curveRadius * Math.sin(.017453293 * (curveAngle / 2)));
}

function Adjust_Length_For_Compound_Spiral_If_Necessary(index, Length)
{
	var StartRad = GetSpiralStartRadius(index);
	var EndRad = GetSpiralEndRadius(index);
	if ((StartRad != 'INF') && (EndRad != 'INF')) //if neither the start nor end radius is infinite, the length will be adjusted
	{
  	  StartRad=Number(GetSpiralStartRadius(index));
    	EndRad=Number(GetSpiralEndRadius(index));
      if (StartRad > EndRad)	{
      		var R1 = new Number(StartRad);
          var R2 = new Number(EndRad);
       }
	else
	{
  		R1 = new Number(EndRad);
      R2 = new Number(StartRad);
	}

	var Len1 = (Length * R2) / (R1 - R2);
	var Len2 = new Number(Length);
	return (Len1 + Len2);
	}
	

	return (Length);
}

function GetSpiralA(index)
{
	var Len = GetSpiralLength(index);
	Len = Adjust_Length_For_Compound_Spiral_If_Necessary(index, Len);
	var Radius = GetSpiralRadius(index);
	
	var A = Math.sqrt(Len * Radius);
	return (A);
}

function GetSpiralTotalX(index) 
{
	var Len = GetSpiralLength(index);
	var Radius = GetSpiralRadius(index);
	var FullLen = Adjust_Length_For_Compound_Spiral_If_Necessary(index, Len);
		
	if (Len != FullLen) //If neither the Start Radius nor the End Radius is infinite
	{
		var TotalX = Get_Adjusted_X_or_Y(index, 1);
	}
	else
	{
		var Temp1 = (Math.pow(Len, 2)) / (40 * Math.pow(Radius, 2));
		var Temp2 = (Math.pow(Len, 4)) / (3456 * Math.pow(Radius, 4));
		TotalX = Len * (1 - Temp1 + Temp2);
	}
	return (TotalX);
}

function GetSpiralTotalY(index)
{
	var Len = GetSpiralLength(index);
	var FullLen = Adjust_Length_For_Compound_Spiral_If_Necessary(index, Len);
	var Radius = GetSpiralRadius(index);
	
	if (Len != FullLen) //If neither the Start Radius nor the End Radius is infinite
	{
		var TotalY = Get_Adjusted_X_or_Y(index, 0);
	}
	else
	{
		var Temp1 = (Math.pow(Len, 2)) / (Radius * 6);
		var Temp2 = (Math.pow(Len, 2)) / (56 * Math.pow(Radius, 2));
		var Temp3 = (Math.pow(Len, 4)) / (7040 * Math.pow(Radius, 4));
		TotalY = Temp1 * (1 - Temp2 + Temp3);
	}
	return (TotalY);
}

function Get_Adjusted_X_or_Y(index, XorY)
{
	var spiral = GeomArray[index];

	var StartRad = GetSpiralStartRadius(index);
	var EndRad = GetSpiralEndRadius(index);
	
	var X2 = spiral.piE;
	var Y2 = spiral.piN;

	if (StartRad > EndRad)
	{
		var X1 = spiral.StartE;
		var Y1 = spiral.StartN;
		var X3 = spiral.EndE;
		var Y3 = spiral.EndN;
	}
	else
	{
		X3 = spiral.StartE;
		Y3 = spiral.StartN;
		X1 = spiral.EndE;
		Y1 = spiral.EndN;
	}
	
	var M = (Y2 - Y1) / (X2 - X1);
	var C1 = Y1 - (M * X1);
	var C2 = Y3 + (X3 / M);
	var X = (M * (C2 - C1) / (Math.pow(M, 2) + 1));
	var Y = ((Math.pow(M, 2) * C2) + C1) / (Math.pow(M, 2) + 1);
	
	var TotalX = Math.sqrt(Math.pow((Y - Y1), 2) + (Math.pow((X - X1), 2)));
	var TotalY = Math.sqrt(Math.pow((Y - Y3), 2) + (Math.pow((X - X3), 2)));
	
	if (XorY == 1)
		return TotalX;
	else
		return TotalY;
}



function GetSpiralX(index) 
{
    var dTotalX = 0;
    var dTotalY = 0;
    var dRadiusInOut = GetSpiralRadius(index);
    var m_dRadiusStart=GetSpiralStartRadius(index);
    var m_dRadiusEnd=GetSpiralEndRadius(index);
    var Len = GetSpiralLength(index);
	  var FullLen = Adjust_Length_For_Compound_Spiral_If_Necessary(index, Len);

    // if spiral is compound then total X is calculated as from the point
    // where large radius starts and the reference X axis is the tangent at
    // point of large radius
    if( m_dRadiusStart !='INF' && m_dRadiusEnd!='INF' )
    {
        var dDeltaX=0, dDeltaY=0;
        dTotalX=getOffsetPtX(FullLen, FullLen, dRadiusInOut);
        dTotalY=getOffsetPtY(FullLen, FullLen, dRadiusInOut);
        
        dDeltaX=getOffsetPtX(FullLen, (FullLen-Len), dRadiusInOut);
        dDeltaY=getOffsetPtY(FullLen, (FullLen-Len), dRadiusInOut);
    
        var dDeltaAngle=FullLen / (2 * dRadiusInOut)- GetSpiralTheta_Radians(index);
 
        
        if( (dDeltaAngle > 0) || (dDeltaAngle < 0) )
        {
            dTotalX = (dTotalX-dDeltaX)*Math.cos(-dDeltaAngle) - (dTotalY-dDeltaY)*Math.sin(-dDeltaAngle);
        }
    }
    else
    {
        dTotalX=getOffsetPtX(FullLen, Len, dRadiusInOut);
    }
    return dTotalX;
}

function GetSpiralY(index) 
{
    var dTotalX = 0, dTotalY = 0;
    var dRadiusInOut = GetSpiralRadius(index);
    var m_dRadiusStart=GetSpiralStartRadius(index);
    var m_dRadiusEnd=GetSpiralEndRadius(index);
    var Len = GetSpiralLength(index);
	  var FullLen = Adjust_Length_For_Compound_Spiral_If_Necessary(index, Len);
    // if spiral is compound then total X is calculated as from the point
    // where large radius starts and the reference X axis is the tangent at
    // point of large radius
    if( m_dRadiusStart !='INF' && m_dRadiusEnd  !='INF' )
    {
        var dDeltaX=0, dDeltaY=0;
        dTotalX=getOffsetPtX(FullLen, FullLen, dRadiusInOut);
        dTotalY=getOffsetPtY(FullLen, FullLen, dRadiusInOut);
        
        dDeltaX=getOffsetPtX(FullLen, (FullLen-Len), dRadiusInOut);
        dDeltaY=getOffsetPtY(FullLen, (FullLen-Len), dRadiusInOut);
        
        var dDeltaAngle=FullLen / (2 * dRadiusInOut)- GetSpiralTheta_Radians(index);
        if( dDeltaAngle > 0 || dDeltaAngle < 0 )
        {
             dTotalY = (dTotalX-dDeltaX)*Math.sin(-dDeltaAngle) + (dTotalY-dDeltaY)*Math.cos(-dDeltaAngle);
        }
    }
    else
    {
        dTotalY=getOffsetPtY(FullLen, Len, dRadiusInOut);
    }
    return dTotalY;
}

function getOffsetPtX(dTotalLength, dLength, dRadiusInOut)
{
    

    var theta = new Array(18);
    var theta_test = 0;
    var dAngle = 0;

    dAngle = (dLength*dLength*0.5)/(dTotalLength*dRadiusInOut);

    theta[0] = 1.0;
    theta_test = Math.log(Math.abs(dAngle))/Math.LN10;
    theta_test = -300 - theta_test;

    for(var i=1; i<=17; i++)
    {
        if( Math.log(Math.abs(theta[i-1]))/Math.LN10 < theta_test )
        {
            break;
        }
        theta[i] = theta[i-1]*(dAngle);
    }

    return  dLength*(1.0 - (theta[2]/10.0) + (theta[4]/216.0) - 
        (theta[6]/9360.0) + (theta[8]/685440.0) - 
        (theta[10]/76204800.0) + 
        (theta[12]/11975040000.0) - (theta[14]/2528170444800.0) +
        (theta[16]/690452066304000.0) );

}

function getOffsetPtY(dTotalLength, dLength, dRadiusInOut)
{
    var theta=new Array(18);
    var theta_test = 0;
    var dAngle = 0;

    dAngle = (dLength*dLength*0.5)/(dTotalLength*dRadiusInOut);

    theta[0] = 1.0;
    theta_test = Math.log(Math.abs(dAngle))/Math.LN10;
    theta_test = -300 - theta_test;

    for(var i=1; i<=17; i++)
    {
        if( Math.log(Math.abs(theta[i-1]))/Math.LN10 < theta_test )
        {
            break;
        }
        theta[i] = theta[i-1]*(dAngle);
    }

    return dLength*( (theta[1]/3.0) - (theta[3]/42.0) + 
        (theta[5]/1320.0) - (theta[7]/75600.0) + 
        (theta[9]/6894720.0) - 
        (theta[11]/918086400.0) + 
        (theta[13]/168129561600.0) - (theta[15]/40537905408000.0) +
        (theta[17]/1.244905998336e16) );
}

function GetSpiralLongTangent(index)
{	
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return "Not a spiral";
	}
  var Theta=GetSpiralTheta_Radians(index);
  return GetSpiralX(index) - GetSpiralY(index) / Math.tan(Theta);

}

function GetSpiralShortTangent(index)
{	
	var spiral = GeomArray[index];
	if(spiral.type != "Spiral")
	{
		return "Not a spiral";
	}
	var Theta = GetSpiralTheta_Radians(index);
	var TotalY = GetSpiralY(index);
	return TotalY / Math.sin(Theta);	
}

function GetSpiralTheta_Radians(index)
{
	var Len = GetSpiralLength(index);
	var FullLen = Adjust_Length_For_Compound_Spiral_If_Necessary(index, Len);
	var Radius = GetSpiralRadius(index);
	var Theta_In_Radians = FullLen / (2 * Radius);	

	if (Len != FullLen) //If neither the Start Radius nor the End Radius is infinite
	{
		var StartRad = parseFloat(GetSpiralStartRadius(index));
		var EndRad = parseFloat(GetSpiralEndRadius(index));
		if (StartRad > EndRad)
		{
			var R1 = new Number(EndRad);
		}
		else
		{
			R1 = new Number(StartRad);
		}
		var Length_to_Tangent = FullLen - Len;
		var Theta1 = Length_to_Tangent*Length_to_Tangent/ (2 * R1*FullLen);
		Theta_In_Radians = Theta_In_Radians - Theta1;
	}
	 
	
	return Theta_In_Radians;
}

function GetSpiralTheta_Degrees(index)
{
	return  GetSpiralTheta_Radians(index) * 57.295779513082320876798154814105;
}

function GetSpiralP(index)
{
	var Radius = GetSpiralRadius(index);
	var Len = GetSpiralLength(index);
	Len = Adjust_Length_For_Compound_Spiral_If_Necessary(index, Len);
	var Y = GetSpiralY(index);
	var Theta = GetSpiralTheta_Radians(index);
	
	var Temp2 = Math.cos(Theta); //Here Theta is converted to radians
	var P = Y - (Radius * (1-Temp2));
	return(P);
}

function GetSpiralK(index)
{
	var Radius = GetSpiralRadius(index);
	var Len = GetSpiralLength(index);
	Len = Adjust_Length_For_Compound_Spiral_If_Necessary(index, Len);
	var X = GetSpiralX(index);
	var Theta = GetSpiralTheta_Radians(index);
	
	var K = X - (Radius * Math.sin(Theta)); 
	return (K);
}


// -----------------------------------------------------------------
// Object Definitions
// -----------------------------------------------------------------
function StationEquation(InternalStation, BackStation, AheadStation)
{
	this.name="hello"
	this.InternalStation = Number(InternalStation);
	this.BackStation= Number(BackStation);
	this.AheadStation= Number(AheadStation);
}

// Geometery element definitions
function PIElement(Northing, Easting, Station)
{
	this.Northing = Number(Northing);
	this.Easting = Number(Easting);
	this.Station = Number(Station);
	this.AheadStation = Number(Station);
}

function CoordPoint(pointN, pointE, pointZ)
{
	this.PointN = pointN;
	this.PointE = pointE;
	this.PointZ = pointZ;
}

function GeometryElement(type)
{
	this.type = "GeomElement";
	
	this.StartN = 0;
	this.StartE= 0;
	
	this.piN = 0;
	this.piE = 0;
	
	this.EndN= 0;
	this.EndE= 0;
	
	this.StartStation = 0;
	this.InternalStartStation = 0;
	this.BackStartStation = 0;
	this.StartRegion;
	
	this.PIStation = 0;
	this.InternalPIStation = 0;
	this.BackPIStation = 0;
	
	this.EndStation = 0;
	this.InternalEndStation = 0;
	this.BackEndStation = 0;
	this.EndRegion;
}

function LineElement(StartN, StartE, EndN, EndE)
{
	this.type="Line";
	this.StartN = new Number(parseFloat(StartN));
	this.StartE= new Number(parseFloat(StartE));
	this.EndN= new Number(parseFloat(EndN));
	this.EndE= new Number(parseFloat(EndE));
}
LineElement.prototype = new GeometryElement;

function CurveElement(StartN, StartE, CenterN, CenterE, EndN, EndE, Rot)
{
	this.type = "Curve";
	this.StartN = new Number(parseFloat(StartN));
	this.StartE= new Number(parseFloat(StartE));
	this.CenterN = new Number(parseFloat(CenterN));
	this.CenterE = new Number(parseFloat(CenterE));
	this.EndN= new Number(parseFloat(EndN));
	this.EndE= new Number(parseFloat(EndE));
	this.Rot =new String(Rot);
}
CurveElement.prototype = new GeometryElement;

function SpiralElement(length, radiusStart, radiusEnd, rotation, spiType, StartN, StartE, piN, piE, EndN, EndE)
{
	this.type = "Spiral";
	this.length= new Number(parseFloat(length));
	this.radiusStart= radiusStart;
	this.radiusEnd= radiusEnd;
	this.rotation= rotation;
	this.spiType= spiType;
	this.StartN = new Number(parseFloat(StartN));
	this.StartE= new Number(parseFloat(StartE));
	this.piN= new Number(parseFloat(piN));
	this.piE= new Number(parseFloat(piE));
	this.EndN= new Number(parseFloat(EndN));
	this.EndE= new Number(parseFloat(EndE));
}
SpiralElement.prototype = new GeometryElement;

function IrregularLineElement2d(StartN, StartE, EndN, EndE)
{
	this.type="IrregularLine2d";
	this.StartN = new Number(parseFloat(StartN));
	this.StartE= new Number(parseFloat(StartE));
	this.EndN= new Number(parseFloat(EndN));
	this.EndE= new Number(parseFloat(EndE));
	this.PntArray = new Array(1);
}
IrregularLineElement2d.prototype = new GeometryElement;

function IrregularLineElement3d(StartN, StartE, EndN, EndE)
{
	this.type="IrregularLine3d";
	this.StartN = new Number(parseFloat(StartN));
	this.StartE= new Number(parseFloat(StartE));
	this.EndN= new Number(parseFloat(EndN));
	this.EndE= new Number(parseFloat(EndE));
	this.PntArray = new Array(1);
}
IrregularLineElement3d.prototype = new GeometryElement;

]]>
</msxsl:script>

</xsl:stylesheet>
