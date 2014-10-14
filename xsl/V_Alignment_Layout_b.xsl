<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet version="1.0"
    xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt"
    xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
    xmlns:lxml="urn:lx_utils">
    
	<!-- =========== JavaScript Include==== -->
	<xsl:include href="NodeConversion_JScript.xsl" />
	<xsl:include href="Computations_JScript.xsl" />
	<!-- ================================= -->
	<!-- ==========Formatting parameters===== -->
	<xsl:param name="Profile.Elevation.unit">default</xsl:param>
	<xsl:param name="Profile.Elevation.precision">0.00</xsl:param>
	<xsl:param name="Profile.Elevation.rounding">normal</xsl:param>
	<xsl:param name="Profile.Coordinate.unit">default</xsl:param>
	<xsl:param name="Profile.Coordinate.precision">0.00</xsl:param>
	<xsl:param name="Profile.Coordinate.rounding">normal</xsl:param>
	<xsl:param name="Profile.Linear.unit">default</xsl:param>
	<xsl:param name="Profile.Linear.precision">0.00</xsl:param>
	<xsl:param name="Profile.Linear.rounding">normal</xsl:param>
	<xsl:param name="Profile.Angular.precision">0.00</xsl:param>
	<xsl:param name="Profile.Angular.rounding">normal</xsl:param>
	<xsl:param name="Profile.Station.Increment">20</xsl:param>
	<xsl:param name="Profile.Station.UseDefaultStartStation">yes</xsl:param>
	<xsl:param name="Profile.Station.StartStation">0</xsl:param>
	<xsl:param name="Profile.Station.Display">XX+XX</xsl:param>
	<xsl:param name="Profile.Station.unit">default</xsl:param>
	<xsl:param name="Profile.Station.precision">0.00</xsl:param>
	<xsl:param name="Profile.Station.rounding">normal</xsl:param>
	<!-- ================================= -->
	<xsl:template match="lx:ProfAlign" mode="set">
		<xsl:variable name="arrSet" select="landUtils:ResetVElementArraySize()" />
		<xsl:for-each select="./*">
			<xsl:variable name="tmp" select="landUtils:SetStartingStationToDefault(position(), string(.))" />
			<xsl:choose>
				<xsl:when test="name() = 'PVI'">
					<xsl:variable name="addPvi" select="landUtils:AddVPI(string(.))" />
				</xsl:when>
				<xsl:when test="name() = 'ParaCurve' ">
					<xsl:variable name="len" select="@length" />
					<xsl:variable name="addPvi" select="landUtils:AddVCurve(string(.), string($len))" />
				</xsl:when>
				<xsl:when test="name() = 'CircCurve' ">
					<xsl:variable name="len" select="@length" />
					<xsl:variable name="radius" select="@radius" />
					<!-- note that radius value is not recorded -->
					<xsl:variable name="addPvi" select="landUtils:AddVCurve(string(.), string($len))" />
				</xsl:when>
				<xsl:when test="name() = 'UnsymParaCurve' ">
					<xsl:variable name="len" select="@lengthIn" />
					<xsl:variable name="lenOut" select="@lengthOut" />
					<!-- note that lengthIn is used & lengthOut is ignored for now -->
					<xsl:variable name="addPvi" select="landUtils:AddVCurve(string(.), string($len))" />
				</xsl:when>
			</xsl:choose>
		</xsl:for-each>
		<xsl:variable name="sortedByStation" select="landUtils:SortByStation()" />
		<xsl:variable name="valid" select="landUtils:ValidateVertical()" />
	</xsl:template>
	<msxsl:script language="JScript" implements-prefix="landUtils"> 
<![CDATA[



// -----------------------------------------------------------------
// Global Arrays (element at 0 is always null)
// -----------------------------------------------------------------
var VElementArray;
var VStationEqArray;
var incr = 25;
var trackStation = 25;
var useDefault = "yes";
var secondStation = "true";
var stationDisplay = "##+##";

// -----------------------------------------------------------------
// InOut station functions
// -----------------------------------------------------------------
function InternalToStation(station)
{
	var n = Number(station);
	
	for(var i = 0; i < VStationEqArray.length; i++)
	{
		var eq = VStationEqArray[i];
		
		if(n > eq.InternalStation)
		{
			return n + (eq.Station - eq.InternalStation);
		}
	}
	return n;	
}
function StationToInternal(station)
{
	var n = Number(station);
	for(var i = 0; i < VStationEqArray.length; i++)
	{
		var eq = VStationEqArray[i];
		
		if(n > eq.AheadStation)
		{
			return n - (eq.Station - eq.InternalStation);
		}
	}
	return n;	
}
// -----------------------------------------------------------------
// Helper functions
// -----------------------------------------------------------------
function IncrementStation_Jeff(station, increment)
{
	var sta = Number(station);
	var incr = Number(increment);
	
	return sta + incr;
}

function ResetSecondStationVariable()
{
	//The 'secondStation' variable needs to be reset to "true" at the start of each profile
	secondStation = "true";
	return '';
}

function IncrementStation(station, increment)
{
	var sta = Number(station);
	var incr = Number(increment);

	if (secondStation == "true")
	{
		trackStation = KeepStationsIncremental(incr, sta);
		secondStation = "false";
		return "" + trackStation;
	}
	trackStation = trackStation + incr;
	return sta + incr;
}

function KeepStationsIncremental(increment, startStation)
{
	incr = new Number(increment);
	var startSta = new Number(parseInt(startStation));
	
	var count = startSta + 1;
	var next = startSta + incr;
	while(count < next)
	{
		if ((count % incr) == 0)
		{
			return count;
		}
		count += 1; 
	}
	
	return next;
}

function SetUseDefaultStation(YesOrNo)
{
	useDefault = YesOrNo;
	return "" + useDefault;
}

function SetStartingStation(station)
{
	var startSta = new Number(parseFloat(station));
	trackStation = startSta;
	return station;
}
function SetStartingStationToDefault(index, pviStr)
{
	var strPoint = pviStr.split(" ");
	var startSta = new Number(parseFloat(strPoint[0]));
	
	if (index == 1)
	{
		trackStation = startSta;
	}
	
	return "" + startSta;
}

function SetStationIncrement(increment)
{
	incr = new Number(increment);
	return increment;
}
function GetStationIncrement()
{
	return "" + incr;
}
function GetTrackedStation()
{
	return "" + trackStation;
}

function Check(index)
{
	var nx = new Number(index);
	
	var end = GetEndStation(); //vStations[vCount];
	
	if((end + incr) > trackStation) 
	{
		if (end < trackStation)  
		{
			return "LastOne";
		}
		return "Ok";
	}
	return "Quit";
}



// -----------------------------------------------------------------
// Data Retrieval functions
// -----------------------------------------------------------------
function GetVInternalStation(index)
{
	var ndx = Number(index);
	var ele = VElementArray[ndx];
	
	return ele.InternalStation ;
}

function GetVStation(index)
{
	var ndx = Number(index);
	var ele = VElementArray[ndx];
	
	return ele.Station ;
}
function GetVElevation(index)
{
	var ndx = Number(index);
	var ele = VElementArray[ndx];
	
	return ele.Elevation;
}

function GetGradeIn(index)
{
	var ndx = Number(index);
	var ele = VElementArray[ndx];
	
	return ele.GradeIn;
}

function GetGradeOut(index)
{
	var ndx = Number(index);
	var ele = VElementArray[ndx];
	
	return ele.GradeOut;
}

function GetGradeIn(index)
{
	var ndx = Number(index);
	var ele = VElementArray[ndx];
	
	return ele.GradeIn;
}

function GetVCurveLength(index)
{
	var ndx = Number(index);
	var ele = VElementArray[ndx];
	
	return ele.CurveLength;
}
// -----------------------------------------------------------------
// Element calculation functions
// -----------------------------------------------------------------
function GetIDofVPI(index)
{
	var ndx = Number(index);
	var ele = VElementArray[ndx];
	return ele.id;
}

function GetPVCInternalStation(index)
{
	var ndx = Number(index);
	var ele = VElementArray[ndx];
	var curv = ele.CurveLength;
	
	return ele.InternalStation - (curv / 2);
}

function GetPVCStation(index)
{
	var ndx = Number(index);
	var ele = VElementArray[ndx];
	var curv = ele.CurveLength;
	
	return ele.Station - (curv / 2);
}

function GetPVTInternalStation(index)
{
	var ndx = Number(index);
	var ele = VElementArray[ndx];
	var curv = ele.CurveLength;
	
	return ele.InternalStation + (curv / 2);
}

function GetPVTStation(index)
{
	var ndx = Number(index);
	var ele = VElementArray[ndx];
	var curv = ele.CurveLength;
	
	return ele.Station + (curv / 2);
}
function GetGradeDelta(index)
{
	var ndx = Number(index);
	var ele = VElementArray[ndx];
	
	return Math.abs(ele.GradeIn - ele.GradeOut);
}
function GetKValue(index)
{
	var ndx = Number(index);

	if(ndx > 0 && ndx < VElementArray.length)
	{
		var curLen = GetVCurveLength(index);
		var gradeDelta = GetGradeDelta(index);
		
		return curLen / gradeDelta;
	}
	else
	{
		return 0;
	}
}

function GetVCurveType(index)
{
	var gradeIn = GetGradeIn(index);
	var gradeOut = GetGradeOut(index);
	
	var inward = new Number(gradeIn);
	var outward = new Number(gradeOut);
	
  	if  (inward > outward)
	{
     		return "crest";
     	}
     	 
 	else if (inward < outward)
 	{
	 	return "sag";
	}

  return "none";
}

function GetEndStation()
{
	var last =  VElementArray[VElementArray.length - 1];
	return last.Station;
}

function GetEndAlignmentStation(start, length)
{
	var s = new Number(start);
	var e = new Number(length);
	var end = (s+e);
	return end.toString();
}

function GetHighPointStation(index)
{
	var gradeIn = GetGradeIn(index);
	var gradeOut = GetGradeOut(index);
	
	//If both grades are positive or if both are negative, don't report this at all
	if ((gradeIn < 0 && gradeOut < 0))
	{
		return "NotReporting";
	}
	
	if ((gradeIn > 0 && gradeOut > 0))
	{
			return "NotReporting";
	}
	
	var curveLen = GetVCurveLength(index);
	
	var pvcStation = GetPVCStation(index);
	var pvtStation = GetPVTStation(index);
	
	var thisStation = GetVStation(index);
	
	if(gradeIn >= 0)
	{
		if(gradeOut >= 0)
		{
			return pvtStation;
		}
		else
		{
       		var A = GetGradeDelta(index);
       		var distToHighPoint =  ( (gradeIn) * (curveLen) ) / A ;
       		return  pvcStation + distToHighPoint;
		}
	}
	else
	{
		if(gradeOut >= 0)
		{
       		var A = GetGradeDelta(index);
       		var distToHighPoint =  ( (gradeIn) * (curveLen) ) / A ;
       		return  thisStation + distToHighPoint;
		}
		else
		{
			return pvcStation;
		}
	}
     	return 0;
}

function GetLowPointStation(index)
{
	var gradeIn = GetGradeIn(index);
	var gradeOut = GetGradeOut(index);
	
	//If both grades are positive or if both are negative, don't report this at all
	if ((gradeIn < 0 && gradeOut < 0))
	{
		return "NotReporting";
	}
	
	if ((gradeIn > 0 && gradeOut > 0))
	{
			return "NotReporting";
	}

	var curveLen = GetVCurveLength(index);
	
	var pvcStation = GetPVCStation(index);
	var pvtStation = GetPVTStation(index);
	
	var thisStation = GetVStation(index);

	var R = (gradeOut - gradeIn) / curveLen;

	var distToLowPoint =  (0 - gradeIn) / R;

	if(gradeIn >= 0)
	{
		if(gradeOut >= 0)
		{
			return pvcStation;
		}
		else
		{
       		return  pvcStation + distToLowPoint;
		}
	}
	else
	{
		if(gradeOut >= 0)
		{
       		return pvcStation + distToLowPoint;
		}
		else
		{
			return pvtStation;
		}
	}
	
  	/*if ( (gradeIn) < 0.0 && (gradeOut) > 0.0 )  
  	{
     		if ( Math.abs(gradeIn) == Math.abs(gradeOut) )
     		{
      			 return t_Station;
     		}
     		else
     		{
       		return "" +( pvcStation + distToLowPoint);
     		}
   	}
   	else
     		if ( Math.abs(gradeIn) == Math.abs(gradeOut) )
     		{
      			 return t_Station;
     		}
     		else
     		{
       		//return "" +( thisStation + distToLowPoint); //Original code
       		return  pvcStation + distToLowPoint;   //CG code
     		}*/
	return 0;
}


function GetVElevationAtStation(aStation)
{		
	var retValue = 0;
	var station = new Number(aStation);

	for(var i = 1; i < VElementArray.length; i++)
	{
		var stringPVT = GetPVTStation(i);
		var numPVT = new Number(stringPVT);
		var eSta = VElementArray[i].Station;

		var ele = VElementArray[i];

		if(station < numPVT)
		{
			if (i < 2)
				return "NaN"; //Avoid 'NULL or not an object' error
				
			var nEle = VElementArray[i - 1];

			var ele2 = VElementArray[i];
			var pvcSta = ele.InternalStation - ele.CurveLength/2.0;

			if((ele.CurveLength > 0) && (station > pvcSta))
				return GetElevationOnVCurve(i, station);
			else
				return interpolate(ele.InternalStation, ele.Elevation, nEle.InternalStation, nEle.Elevation, station, 0.00);
			
		}
		else
		{

			if(i == VElementArray.length - 1)
			{
				var nEle = VElementArray[i - 1];
				if(ele.type == "vpi")
				{
					return interpolate(nEle.InternalStation, nEle.Elevation, ele.InternalStation, ele.Elevation, station, 0.00);
				}
				else if(ele.type = "vcurve")
				{
					return GetElevationOnVCurve(i, station);
				}
			}
		}
	}
	return retValue;
}

function GetElevationOnVCurve(index, aStation)
{
	var ndx = new Number(index);
	var station = new Number(aStation);
	
	var prevE = VElementArray[ndx - 1];
	var ele = VElementArray[ndx];
	var nextE = VElementArray[ndx + 1];
  	
  	var curveLen = ele.CurveLength;
   	var pviStation = ele.InternalStation;
   	var pviElevation = ele.Elevation;
   	
   	var startStation = prevE.InternalStation;
   	var startElevation = prevE.Elevation;
   	
   	var endStation = pviStation; //nextE.InternalStation;
  	var endElevation = pviElevation; //nextE.Elevation;
   	
   	var gradeIn = GetGradeIn(ndx);
   	var gradeOut = GetGradeOut(ndx);

   	var elev, elev_bvc;
   	var x_dist, a_val;

  	 x_dist = (station - (pviStation - curveLen/2)) / 100.0; 

   	var temp_ptX = pviStation - ( curveLen/2 );

   	elev_bvc = interpolate(endStation, endElevation, startStation, startElevation, temp_ptX, 0.0);

   	// do the final calculation - y = (r/2)*x*x + g1*x + elev_bvc

   	var curveLen2 = curveLen / 100.0;

   	a_val = (gradeOut - gradeIn) / (2.0 * curveLen2); 
   		
   	elev = (a_val * x_dist * x_dist) + (gradeIn * x_dist) + elev_bvc;


   	return elev;
}

function interpolate(x1, y1, x2, y2, x3, y3)
{
	var slope, elevation;

  	slope = (y1 - y2) / (x1 - x2);

  	if ( y3 == 0.0 )
      		elevation = y1 - ((x1 - x3) * slope);
  	else
      		elevation = x1 - ((y1 - y3) * slope);

  	return elevation; 
}

// -----------------------------------------------------------------
// Element to element calculation functions
// -----------------------------------------------------------------
function GetHorzDistance(startIndex, endIndex)
{
	var sndx = Number(startIndex);
	var endx = Number(endIndex);
	
	var stEle = VElementArray[sndx];
	var enEle = VElementArray[endx];
	
	return enEle.InternalStation - stEle.InternalStation ;
}
function GetVertDistance(startIndex, endIndex)
{
	var sndx = Number(startIndex);
	var endx = Number(endIndex);
	
	var stEle = VElementArray[sndx];
	var enEle = VElementArray[endx];
	
	return enEle.Elevation - stEle.Elevation;
}
function GetVLength(startIndex, endIndex)
{
	var sndx = Number(startIndex);
	var endx = Number(endIndex);
	
	var stEle = VElementArray[sndx];
	var enEle = VElementArray[endx];
	
	var eDiff = stEle.Elevation - enEle.Elevation;
	var sDiff = stEle.InternalStation - enEle.InternalStation;
	
	var tanlength = Math.sqrt(Math.pow(sDiff, 2) + Math.pow(eDiff, 2));
	
	return tanlength;
}
function GetVGrade(startIndex, endIndex)
{
	var sndx = Number(startIndex);
	var endx = Number(endIndex);
	
	var stEle = VElementArray[sndx];
	var enEle = VElementArray[endx];
	
	var eDiff = stEle.Elevation - enEle.Elevation;
	var sDiff = stEle.InternalStation - enEle.InternalStation;
	
	//vTangentLength = Math.sqrt(Math.pow(sDiff, 2) + Math.pow(eDiff, 2));
	var grade = (eDiff / sDiff) * 100;

	return grade;
}

// -----------------------------------------------------------------
// VStationEquation Array Functions
// -----------------------------------------------------------------
function AddVStationEquation(internalSta, backSta, aheadSta)
{
	var inStr = NodeToText(internalSta);
	var bStr = NodeToText(backSta);
	var aStr = NodeToText(aheadSta);
	
	var nStaEq = new VStationEquation(inStr, bStr, aStr);
	
	VStationEqArray.push(nStaEq);
	return "ok";	
}

function ApplyVStationEquations()
{
	// Defect fix - VStationEqArray.length can't be called if VStationEqArray hasn't been initialized
	// A more accurate fix would be to discover why this array is not initialized before this method is called
	if(typeof(VStationEqArray) != "undefined")
	{
		var len = VStationEqArray.length;

		for(var se = 1; se < len; se++)
		{
			var staEq = VStationEqArray[se]; 
			var staEqInternal = staEq.InternalStation;
			var eqDiff = staEq.InternalStation - staEq.AheadStation;
			
			for(var i = 1; i < VElementArray.length; i++)
			{
				var vEle = VElementArray[i];
				if(staEqInternal < vEle.InternalStation)
				{
					vEle.Station = vEle.InternalStation - eqDiff;
				}
			}
		}
	}	
	return "done";
}

// -----------------------------------------------------------------
// VElement Array Functions
// -----------------------------------------------------------------

function funcSortByStationAscending(first, second)
{
    if ( first.InternalStation < second.InternalStation )
    {
        return -1;
    } 
    else if (first.InternalStation > second.InternalStation )
    {
        return 1;
    } 
    else
    {
        return 0;
    }
}


// LandXML 7.0 sorts by element type rather than by station value,
// so we'll sort by station here
function SortByStation()
{
    // Remove extra element at start of array before sort:
    VElementArray.shift();
    //
    VElementArray.sort( funcSortByStationAscending );
    //
    // reinsert a dummy element at start of array
    var dummy = new VPI(0.0, 0.0);
    VElementArray.unshift(dummy);  
    return 1;
}



// ValidateVertical is really calculating and storing gradeIn and
// gradeOut...function name is a misnomer
function ValidateVertical()
{
	var i;
	for(i = 1; i < VElementArray.length - 1; i++)
	{
		var ele = VElementArray[i];
		var nele = VElementArray[i + 1];
		
		var gradeOut = GetVGrade(i, i + 1);
		
		ele.GradeOut = gradeOut;
		nele.GradeIn = gradeOut;
	}

	return 1;
}
function GetVEndStation()
{
	var endElems = VElementArray[VElementArray.length - 1];
	return endElems.Station;
}

function AddVPI(nodeStr)
{
	var insureString = String(nodeStr);	
	var coords = insureString.split(" ");	
	var vpi = new VPI(coords[0], coords[1]);	
	vpi.id = VElementArray.length;
	VElementArray.push(vpi);
	
	return coords[0];
}

// May be called for a CircCurve, ParaCurve or UnsymParaCurve,
// but only the PVI & length parameters are accepted because
// that's what we report.  In the future we will deal with
// radius for CircCurve and lengthIn/lengthOut for UnsymParaCurve
// when we have a form that supports that data.
function AddVCurve(nodeStr, lengthStr)
{
	var insureString = String(nodeStr);
	var coords = insureString.split(" ");
	var vcurve = new VCurve(coords[0], coords[1], lengthStr);
	vcurve.id = VElementArray.length;
	VElementArray.push(vcurve);
	
	return VElementArray.length;
}

function ResetVElementArraySize()
{
	VElementArray = new Array(1);
	VStationEqArray = new Array(1);
	
	return "ok";
}

function GetVElementArraySize()
{
	return VElementArray.length;
}

// -----------------------------------------------------------------
// Object Definitions
// -----------------------------------------------------------------
function VStationEquation(InternalStation, BackStation, AheadStation)
{
	this.type="StationEquation";
	this.InternalStation= new Number(parseFloat(InternalStation));
	this.BackStation= new Number(parseFloat(BackStation));
	this.AheadStation= new Number(parseFloat(AheadStation));
}

function VGeomElement()
{
	this.type = "VGeom";
	this.id = 0;
	this.InternalStation = 0;
	this.Station = 0;
	this.Elevation = 0;
	this.GradeIn = 0;
	this.GradeOut = 0;
	this.CurveLength = 0;
}

function VPI(station, elevation)
{
	this.type = "vpi";
	this.Station = Number(station);
	this.InternalStation = Number(station);
	this.Elevation = Number(elevation);
}
VPI.prototype = new VGeomElement;

function VCurve(station, elevation, curveLength)
{
	this.type = "vcurve";
	this.Station = Number(station);
	this.InternalStation = Number(station);
	this.Elevation = Number(elevation);
	this.CurveLength = Number(curveLength);
}
VCurve.prototype = new VGeomElement;

// -----------------------------------------------------------------
// VStation Equation Array Functions
// -----------------------------------------------------------------

]]>
</msxsl:script>
</xsl:stylesheet>
