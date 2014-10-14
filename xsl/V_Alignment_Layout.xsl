<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet 	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
                		xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                 		xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
                		xmlns:lxml="urn:lx_utils"
                		version="1.0">
<!-- =========== JavaScript Include==== -->
<xsl:include href="V_Tangent_Calculations.xsl"></xsl:include>
<xsl:include href="V_Curve_Calculations.xsl"></xsl:include>
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
<xsl:param name="Profile.Station.Increment">0</xsl:param>
<xsl:param name="Profile.Station.UseDefaultStartStation">yes</xsl:param>
<xsl:param name="Profile.Station.StartStation">0</xsl:param>
<xsl:param name="Profile.Station.Display">XX+XX</xsl:param>
<xsl:param name="Profile.Station.unit">default</xsl:param>
<xsl:param name="Profile.Station.precision">0.00</xsl:param>
<xsl:param name="Profile.Station.rounding">normal</xsl:param>
<!-- ================================= -->

<xsl:template match="lx:ProfAlign" mode="set">
	<xsl:variable name="cnt" select="landUtils:SetArraySizes(count(./node()))"></xsl:variable>
	<xsl:for-each select="./node()">
		<xsl:variable name="tmp" select="landUtils:ProcessPVI(position(), string(.))"/>
		<xsl:choose >
			<xsl:when test="@length">
				<xsl:variable name="clen" select="landUtils:SetVLength(position(), string(@length))"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="clen" select="landUtils:SetVLength(position(), 0)"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:for-each>
</xsl:template>


<msxsl:script language="JScript" implements-prefix="landUtils"> 
<![CDATA[
var vCount = 0;
var vStations = new Array(1);
var vElevations = new Array(1);
var vGrades = new Array(1);
var vLengths = new Array(1);
var incr = 25;
var trackStation = 25;
var useDefault = "yes";
var secondStation = "true";
var stationDisplay = "##+##";

function SetStationDisplay(displayOption)
{
	stationDisplay = displayOption;
	return "" + stationDisplay;
}


// -----------------------------------------------------------------
// Split functions
// -----------------------------------------------------------------
function ProcessPVI(index, pviStr)
{
	var ndx = new Number(index);
	
	
	var strPoint = pviStr.split(" ");
	var tmpStation = new Number(parseFloat(strPoint[0]));
	var tmpElevation = new Number(parseFloat(strPoint[1]));
	
	SetVStation(ndx, tmpStation);
	SetVElevation(ndx, tmpElevation);
	
	if (ndx == 1)
	{
		SetStartingStation(tmpStation);
	}
	
	return pviStr;
}

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

function Check(index)
{
	var nx = new Number(index);
	
	var end = vStations[vCount];
	
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

function IncrementStation(station)
{
	var sta = new Number(station);
	if (secondStation == "true")
	{
		trackStation = KeepStationsIncremental(incr, sta);
		secondStation = "false";
		return "" + trackStation;
	}
	trackStation = trackStation + incr;
	return "" + (sta + incr);
}

function DecrementStation(station)
{
	var sta = new Number(station);
	trackStation = trackStation - incr;
	return "" + (sta - incr);
}

function SetStartingStation(station)
{
	startSta = new Number(parseFloat(station));
	CheckForEnd = startSta;
	trackStation = startSta;
	length = startSta;
	return station;
}
function SetUseDefaultStation(YesOrNo)
{
	useDefault = YesOrNo;
	return "" + useDefault;
}

function GetStartStation()
{
	return "" + startSta;
}
function GetTrackedStation()
{
	return "" + trackStation;
}

function GetLastStation()
{
	var end = vStations[vCount];
	return "" + end;
}

function GetLastElevation()
{
	var elev = vElevations[vCount];
	return "" + elev;
}

// -----------------------------------------------------------------
// Array Set functions
// -----------------------------------------------------------------
function SetArraySizes(count)
{
	var cnt = new Number(count);
	vCount = cnt;
	vStations = new Array(cnt + 1);
	vElevations =  new Array(cnt + 1);
	vGrades = new Array(cnt + 1);
	vLengths = new Array(cnt + 1);
	
	return "" + cnt;
}

function SetVStation(index, station)
{
	vStations[index] = station;
	return station;
}
function SetVElevation(index, elevation)
{
	vElevations[index] = elevation;
	return elevation;
}
function SetVGrade(index, grade)
{
	vGrades[index] = grade;
	return grade;
}


function SetVLength(index, length)
{
	var ndx = new Number(index);
	vLengths[ndx] = length;

	return length;
}

// -----------------------------------------------------------------
// Array Get functions
// -----------------------------------------------------------------
function GetVStation(index)
{
	ndx = new Number(index);
	return "" + vStations[ndx] ;
}
function GetPVCStation(index)
{
	var ndx = new Number(index);
	if(vLengths[ndx] != 0 || vLengths[ndx] != null)
	{
		var curSta = vStations[ndx];
		var curLen = vLengths[ndx];
		
		return ""  + (curSta - (curLen / 2));
	}
	else
	{
		return "" + 0;
	}
}
function GetPVCElevation(index)
{
}
function GetPVTStation(index)
{
	var ndx = new Number(index);
	if(vLengths[ndx] != 0 || vLengths[ndx] != null)
	{
		var curSta = vStations[ndx];
		var curLen = vLengths[ndx];
		
		return ""  + (curSta + (curLen / 2));
	}
	else
	{
		return "" + 0;
	}
}
function GetPVTElevation(index)
{
}
function GetVElevation(index)
{
	ndx = new Number(index);
	
	if(vElevations[ndx] != null)
	{
		return "" + vElevations[ndx] ;
	}
	else
	{
		return "No Elevation at Index";
	}
	
}

function GetGradeIn(index)
{
	var ndx = new Number(index);
	if(ndx > 0)
	{
		var hdiff = vStations[ndx] - vStations[ndx - 1];
		var vdiff = vElevations[ndx] - vElevations[ndx - 1];
		
		return "" + ((vdiff / hdiff) * 100);
	}
	else
	{
		return "" + 0;
	}	
}
function GetGradeOut(index)
{
	var ndx = new Number(index);
	if(ndx < vCount )
	{
		var hdiff = vStations[ndx + 1] - vStations[ndx];
		var vdiff = vElevations[ndx + 1] - vElevations[ndx];
		
		return "" + ((vdiff / hdiff) * 100);
	}
	else
	{
		return "" + 0;
	}	
}
function GetGradeDelta(index)
{
	var ndx = new Number(index);
	if(ndx > 0 && ndx < vCount)
	{
		var gIn = GetGradeIn(index);
		var gOut = GetGradeOut(index);
		
		var gi = new Number(gIn);
		var go = new Number(gOut);
		
		return "" + Math.abs(gi - go);
	}
	else
	{
		return "" + 0;
	}
}
function GetVCurveLength(index)
{
	var ndx = new Number(index);
	
	return "" + vLengths[ndx];
}
function GetKValue(index)
{
	var ndx = new Number(index);
	if(ndx > 0 && ndx < vCount)
	{
		var curLen = GetVCurveLength(index);
		var gradeDelta = GetGradeDelta(index);
		
		return "" + curLen / gradeDelta;
	}
	else
	{
		return "" + 0;
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

function GetHighPointStation(index)
{
	var grade_In = GetGradeIn(index);
	var grade_Out = GetGradeOut(index);
	
	//If both grades are positive or if both are negative, don't report this at all
	if ((grade_In < 0 && grade_Out < 0))
	{
		return "NotReporting";
	}
	
	if ((grade_In > 0 && grade_Out > 0))
	{
			return "NotReporting";
	}
	
	var curve_Len = GetVCurveLength(index);
	
	var gradeIn = new Number(grade_In);
	var gradeOut = new Number(grade_Out);
	var curveLen = new Number(curve_Len);
	
	var t_Station = GetVStation(index);
	var pc_Station = GetPVCStation(index);
	var pt_Station = GetPVTStation(index);
	
	var pvcStation = new Number(pc_Station);
	var pvtStation = new Number(pt_Station);
	
	var thisStation = new Number(t_Station);
	
  	if ( (gradeIn) > 0.0 && (gradeOut) < 0.0 )  
  	{
     		if ( Math.abs(gradeIn) == Math.abs(gradeOut) )
     		{
      			 return t_Station;
     		}
     		else
     		{
       		var A;
       		var dA = GetGradeDelta(index) ;
       		var bA = new Number(dA );
       		var A = bA;
       		var distToHighPoint =  ( (gradeIn) * (curveLen) ) / A ;
       		return "" +( pvcStation + distToHighPoint);
     		}
   	}
   	else
     		if ( Math.abs(gradeIn) == Math.abs(gradeOut) )
     		{
      			 return t_Station;
     		}
     		else
     		{
       		var A;
       		var dA = GetGradeDelta(index) ;
       		var bA = new Number(dA );
       		var A = bA;
       		var distToHighPoint =  ( (gradeOut) * (curveLen) ) / A ;
       		return "" +( thisStation + distToHighPoint);
     		}
}

function GetLowPointStation(index)
{
	var grade_In = GetGradeIn(index);
	var grade_Out = GetGradeOut(index);
	
	//If both grades are positive or if both are negative, don't report this at all
	if ((grade_In < 0 && grade_Out < 0))
	{
		return "NotReporting";
	}
	
	if ((grade_In > 0 && grade_Out > 0))
	{
			return "NotReporting";
	}

	
	var curve_Len = GetVCurveLength(index);
	
	var gradeIn = new Number(grade_In);
	var gradeOut = new Number(grade_Out);
	var curveLen = new Number(curve_Len);
	
	var t_Station = GetVStation(index);
	var pc_Station = GetPVCStation(index);
	var pt_Station = GetPVTStation(index);
	
	var pvcStation = new Number(pc_Station);
	var pvtStation = new Number(pt_Station);
	
	var thisStation = new Number(t_Station);
	var R = (gradeOut - gradeIn) / curveLen;

	var distToLowPoint =  (0 - gradeIn) / R;

	
  	if ( (gradeIn) < 0.0 && (gradeOut) > 0.0 )  
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
       		return "" +( pvcStation + distToLowPoint);   //CG code
     		}
}


function GetVElevationAtStation(aStation)
{
	var endStation, endElevation, startStation, startElevation;
	var station = new Number(aStation);

	for(var icount = 1; icount <= vCount; icount++)
	{
		var counter = new Number(vStations[icount]);
		var stringPVT = GetPVTStation(icount);
		var numPVT = new Number(stringPVT);
		if(station < numPVT)
		{
			//return GetElevationOnVCurve(icount, station);
			
			var pvcSta = vStations[icount] - vLengths[icount]/2.0; //Added by CG
			
			//Edited by CG
			if((vLengths[icount] > 0) && (station > pvcSta))
				return GetElevationOnVCurve(icount, station);
				
			//if(vLengths[icount] > 0) 
			
			//End of edit
			
			else
			{
				endStation = vStations[icount];
				endElevation = vElevations[icount];
					
				startStation = vStations[icount - 1];				
				startElevation = vElevations[icount - 1];
				
				return "" + (interpolate(endStation, endElevation, startStation, startElevation, station, 0.00));
			}
		}
	}
	return "" + 0;
}

function GetElevationOnVCurve(index, aStation)
{
	var ndx = new Number(index);
	if(vLengths[ndx] > 0){
		var station = new Number(aStation)
  		pviStation = vStations[ndx];
  	
  		var curveLen = vLengths[ndx];
   		var pviStation = vStations[ndx];
   		var pviElevation = vElevations[ndx];
   		var startStation = vStations[ndx - 1];
   		var startElevation = vElevations[ndx - 1];
   	
   		var grd_in = GetGradeIn(ndx);
   		var grd_out = GetGradeOut(ndx);
   		var gradeIn = new Number(grd_in);
   		var gradeOut = new Number(grd_out);

   		var elev, elev_bvc;
   		var x_dist, a_val;

  	 	x_dist = (station - (pviStation - curveLen/2)) / 100.0; 

   		var endStation     = pviStation;
   		var endElevation   = pviElevation;

   		var temp_ptX = pviStation - ( curveLen/2 );

   		elev_bvc = interpolate(endStation, endElevation, startStation, startElevation, temp_ptX, 0.0);		

   		// do the final calculation - y = (r/2)*x*x + g1*x + elev_bvc

   		var curveLen2 = curveLen / 100.0;

   		a_val = (gradeOut - gradeIn) / (2.0 * curveLen2); 
   		
   		elev = (a_val * x_dist * x_dist) + (gradeIn * x_dist) + elev_bvc;

   		return "" + elev;
   }
   else
   {
   	return "";
   }
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

function CalculateSightDistance(index, eyeHeight, objHeight)
{
	var cur_len = GetVCurveLength(index);
	var grd_chg = GetGradeDelta(index);
	var curveLength = new Number(cur_len);
	var gradeChange = new Number(grd_chg);
	var eye = new Number(eyeHeight);
	var obj = new Number(objHeight);
	var calcOne = 100*(curveLength)*Math.pow(Math.sqrt(2*eye) + Math.sqrt(2*obj), 2);
	var calcTwo = Math.sqrt(calcOne / gradeChange);
	
	if(calcTwo < curveLength)
	{
		return "" + calcTwo;
	}
	else
	{
		return "" + ((curveLength / 2) + calcTwo);
	}
}

function GetEndStation(start, length)
{
	var s = new Number(start);
	var e = new Number(length);
	var end = (s+e);
	return end.toString();
}


]]>
</msxsl:script>

</xsl:stylesheet>
