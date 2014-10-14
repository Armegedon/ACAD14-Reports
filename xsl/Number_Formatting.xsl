<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2002 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet version="1.0"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">

<xsl:include href="NodeConversion_JScript.xsl"/>

<xsl:template name="FormatNumber">
	<xsl:param name="number"/>
	<xsl:param name="fromUnit"/>
	<xsl:param name="toUnit"/>
	<xsl:param name="precision"/>
	<xsl:param name="rounding"/>
	<xsl:choose>
		<xsl:when test="$number">
			<xsl:value-of select="landUtils:FormatNumber(string($number), string($fromUnit), string($toUnit), string($precision), string($rounding))"/>
		</xsl:when>
		<xsl:otherwise>
			No number entered
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<xsl:template name="FormatStation">
	<xsl:param name="station"/>
	<xsl:param name="pattern"/>
	<xsl:param name="precision"/>
	<xsl:param name="rounding"/>
	<xsl:choose>
		<xsl:when test="$station">
			<xsl:value-of select="landUtils:FormatStation($station, $pattern, $precision, $rounding)"/>
		</xsl:when>
		<xsl:otherwise>
			No station entered
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<xsl:template name="FormatDirection">
	<xsl:param name="direction"/>
	<xsl:param name="fromFormat"/>
	<xsl:param name="toFormat"/>
	<xsl:param name="precision"/>
	<xsl:param name="rounding"/>
	<xsl:param name="unit"/>
	<xsl:choose>
		<xsl:when test="$direction">
			<xsl:value-of select="landUtils:FormatDirection(string($direction), string($fromFormat), string($toFormat), string($precision), string($rounding), string($unit))"></xsl:value-of>
		</xsl:when>
		<xsl:otherwise>
			No direction entered
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<msxsl:script language="JScript" implements-prefix="landUtils"> 
<![CDATA[
/////////////////////////////////////////////////////////////////////////
// DMS Labels

var angleDegreeLabel = "&deg;";
var angleMinuteLabel = "'";
var angleSecondLabel = "\"";

function SetAngleDegreeLabel(label)
{
	angleDegreeLabel = label;
	return label;
}

function SetAngleMinuteLabel(label)
{
	angleMinuteLabel = label;
	return label;
}

function SetAngleSecondLabel(label)
{
	angleSecondLabel = label;
	return label;
}


var dirDegreeLabel = "&deg;";
var dirMinuteLabel = "'";
var dirSecondLabel = "\"";

function SetDirDegreeLabel(label)
{
	dirDegreeLabel = label;
	return label;
}

function SetDirMinuteLabel(label)
{
	dirMinuteLabel = label;
	return label;
}

function SetDirSecondLabel(label)
{
	dirSecondLabel = label;
	return label;
}

///////////////////////////////////////////////////////////////////////
// Direction labels
var northLabel = "N";
var southLabel = "S";
var eastLabel = "E";
var westLabel = "W";

function SetNorthLabel(label)
{
	northLabel = label;
	return label;
}

function SetSouthLabel(label)
{
	southLabel = label;
	return label;
}

function SetEastLabel(label)
{
	eastLabel = label;
	return label;
}

function SetWestLabel(label)
{
	westLabel = label;
	return label;
}
/////////////////////////////////////////////////////////////////////////////////////////////////////////
// Number formatting
function FormatNumber(numberStr, fromStr, toStr, precisionStr, roundingStr)
{
	var numStr = numberStr;
	var fStr = fromStr;
	var tStr = toStr;
	var pStr = precisionStr;
	var rStr = roundingStr;
	
	var origNumber= new Number(parseFloat(numStr));
	
	var convNumber = 0;
	var precDepth;
	var fixedDec;
	var rNumber;
	
	// Do conversion if necessary
	if(tStr == "default")
	{
		convNumber = origNumber;
	}
	else
	{
		convNumber = ConvertNumber(origNumber, fStr, tStr);
	}
	
	rNumber = FormatPrecision(convNumber, pStr, rStr);
	
	return rNumber.toString();
	//return pStr;
}

function ConvertNumber(num, fromStr, toStr)
{
	var convNumber = new Number(parseFloat(num));
	var numer = new Number(parseFloat(num));
	
	var cnvFromStr = fromStr.toLowerCase();
	var cnvToStr = toStr.toLowerCase();
	
	if(cnvFromStr == "foot")
	{
		if(cnvToStr == "meter")
		{
			convNumber = FeetToMeters(numer)
		}
		else if(cnvToStr == "ussurveyfoot")
		{
			convNumber = FeetToUSSurveyFoot(numer);
		}
		else if(cnvToStr == "inch")
		{
			convNumber = numer*12;
		}
		else if(cnvToStr == "milimeter")
		{
			convNumber = FeetToMeters(numer)*1000;
		}
	}
	else if(cnvFromStr == "meter")
	{
		if(cnvToStr == "foot")
		{
			convNumber = MToFeet(numer);
		}
		else if(cnvToStr == "ussurveyfoot")
		{
			convNumber = MetersToUSSurveyFoot(numer);
		}
		else if(cnvToStr == "inch")
		{
			convNumber = MToFeet(numer)*12;
		}
		else if(cnvToStr == "milimeter")
		{
			convNumber = numer*1000;
		}
	}
	else if(cnvFromStr == "millimeter")
	{
		if(cnvToStr == "meter")
		{
			convNumber = numer/1000.;
		}
		else if(cnvToStr == "foot")
		{
			numer = numer/1000.;
			convNumber = MToFeet(numer);
		}
	}
	else if(cnvFromStr == "inch")
	{
		if(cnvToStr == "foot")
		{
			convNumber = numer/12.;
		}
		else if(cnvToStr == "meter")
		{
			numer = numer/12.;
			convNumber = FeetToMeters(numer);
		}
		
	}
	else if(cnvFromStr == "squarefoot")
	{
		if(cnvToStr == "squaremeter")
		{
			convNumber = SqFeetToSqMeters(numer);
		}
		if(cnvToStr == "acre")
		{
			convNumber = SqFtToAcres(numer);
		}
		if(cnvToStr == "hectare")
		{
			convNumber = SqFeetToSqMeters(numer) / 10000.;
		}
	}
	else if(cnvFromStr == "squaremeter")
	{
		if(cnvToStr == "squarefoot")
		{
			convNumber = SqMetersToSqFeet(numer);
		}
		if(cnvToStr == "acre")
		{
			convNumber = SqMetersToAcres(numer);
		}
		if(cnvToStr == "hectare")
		{
			convNumber = numer / 10000.;
		}
	}
	else
	{
		convNumber = numer;
	}
	
	return convNumber;
}

function ParsePrecision(precStr)
{
	var decPrec;
	var per = ".";
	var s = precStr.indexOf(".");
	if(s <= 0)
	{
		decPrec= 0;
	}
	else
	{
		var d = precStr.length - s - 1;
		decPrec= d;
	}	
	return decPrec;
}
/////////////////////////////////////////////////////////////////////////////////////////////////////////
// Direction formatting
function FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit)
{
   if(unit == "Grads")
      return ConvertFromDegree2Grad(adjDir , precisionStr, roundingStr) + "g";
   else if(unit == "Radian")
      return ConvertFromDegree2Radian(adjDir , precisionStr, roundingStr) + "rad";
   else
      return FormatDMS(adjDir , precisionStr, roundingStr, false);
}

function FormatDirection(directionStr, fromStr, toStr, precisionStr, roundingStr, unit)
{
	var rDir = directionStr;
	var dirNum = new Number(parseFloat(directionStr));
	var convNumber = 0;
	
	if(fromStr == "Conventional")
	{
		if(toStr == "Bearing")
		{
			if(dirNum <= 90 && dirNum >= 0)
			{
				var adjDir = 90 - dirNum;
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return northLabel +" " + dirStr + " " + eastLabel;
			}
			else if(dirNum <= 180 && dirNum > 90)
			{
				var adjDir =  dirNum - 90;
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return northLabel + " " + dirStr + " " + westLabel;
			}
			else if(dirNum <= 270 && dirNum > 180)
			{
				var adjDir = 270 - dirNum;
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return southLabel + " " + dirStr + " " + westLabel;
			}
			else if(dirNum < 360 && dirNum > 270)
			{
				var adjDir = dirNum - 270;
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return southLabel + " " + dirStr + " " + eastLabel;
			}
		}
		else if(toStr == "North Azimuth")
		{
			if(dirNum <= 90 && dirNum >= 0)
			{
				var adjDir = 90 - dirNum;
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return dirStr;
			}
			else if(dirNum <= 180 && dirNum > 90)
			{
				var adjDir =  360 - (dirNum - 90);
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return dirStr;
			}
			else if(dirNum <= 270 && dirNum > 180)
			{
				var adjDir =  360 - (dirNum - 90);
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return dirStr;
			}
			else if(dirNum < 360 && dirNum > 270)
			{
				var adjDir = 360 - dirNum + 90;
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return dirStr;
			}
		}
		else if(toStr == "South Azimuth")
		{
			if(dirNum <= 90 && dirNum >= 0)
			{
				var adjDir = 360 - (dirNum + 90);
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return dirStr;
			}
			else if(dirNum <= 180 && dirNum > 90)
			{
				var adjDir =  180 - (dirNum - 90);
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return dirStr;
			}
			else if(dirNum <= 270 && dirNum > 180)
			{
				var adjDir =  270 - dirNum;
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return dirStr;
			}
			else if(dirNum < 360 && dirNum > 270)
			{
				var adjDir = dirNum;
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return dirStr;
			}
		}
	}
	else if(fromStr == "North Azimuth")
	{
		if(toStr == "Bearing")
		{
			if(dirNum <= 90 && dirNum >= 0)
			{
				var adjDir = dirNum;
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return northLabel + " " + dirStr + " " + eastLabel;
			}
			else if(dirNum <= 180 && dirNum > 90)
			{
				var adjDir = 180 - dirNum;
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return southLabel + " " + dirStr + " " + eastLabel;
			}
			else if(dirNum <= 270 && dirNum > 180)
			{
				var adjDir = dirNum - 180;
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return southLabel + " " + dirStr + " " + westLabel;
			}
			else if(dirNum < 360 && dirNum > 270)
			{
				var adjDir = 360 - dirNum;
				var dirStr = FormatDirectionWithUnit(adjDir , precisionStr, roundingStr, unit);
				return northLabel + " " + dirStr + " " + westLabel;
			}
		}
	}
	else if(fromStr == "Bearing")
	{
		if(toStr == "North Azimuth")
		{
		}
	}
	
	return rDir;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////
// Angle formatting

function FormatAngle(angleStr, fromStr, toStr, formatStr, precisionStr, roundingStr)
{
	var retStr = angleStr;
	
	var angNumber = new Number(parseFloat(angleStr));
	var convNumber = angNumber;
	var rNumber;
	
	// Convert Number
	if(fromStr == "Degrees")
	{
		if(toStr == "Radians")
		{
			convNumber = (angNumber * Math.PI) / 180;
		}
	}
	else if(fromStr == "Radians")
	{
		if(toStr == "Degrees")
		{
			convNumber = (angNumber * 180) / Math.PI;
		}
	}
	
	if(formatStr == "DMS")
	{
		rNumber = FormatDMS(convNumber, precisionStr, roundingStr, true);
	}
	else
	{
		rNumber = FormatPrecision(convNumber, precisionStr, roundingStr);
	}
	
	return rNumber.toString();
}


function ConvertFromDegree2Grad(angle, precisionStr, roundingStr)
{
  return FormatPrecision(angle * 10 / 9, precisionStr, roundingStr);
}

function ConvertFromDegree2Radian(angle, precisionStr, roundingStr)
{
  return FormatPrecision(angle * Math.PI / 180, precisionStr, roundingStr);
}

function FormatDMS(angle, precisionStr, roundingStr, isAngle)
{
	var degrees = Math.floor(angle);

	var dMin = 60. * (angle - degrees);
	var minutes = Math.floor(dMin);

	var dSec = 60. * (dMin - minutes);
	var seconds = FormatPrecision(dSec,  precisionStr, roundingStr);
	
	//fix Fenway defect #655976	
	//if format seconds make seconds equal 60, then carry it to minutes.
	if(seconds>=60.)
	{
		seconds-=60.;
		seconds = FormatPrecision(seconds,  precisionStr, roundingStr);
		minutes++;
	}
	if(minutes>=60.)
	{
		minutes-=60.;		
		degrees++;
	}	
  var Sdeg, Smin, Ssec;
	if(degrees < 10) { Sdeg = "0" + String(degrees); } else { Sdeg = String(degrees); }
	if(minutes < 10) { Smin = "0" + String(minutes); } else { Smin = String(minutes); }
	if(seconds < 10) { Ssec = "0" + String(seconds); } else { Ssec = String(seconds); }	
	
	if(isAngle)
	  return Sdeg + angleDegreeLabel + Smin + angleMinuteLabel  + Ssec + angleSecondLabel;
  else
    return Sdeg + dirDegreeLabel + Smin + dirMinuteLabel  + Ssec + dirSecondLabel;
}

function FormatPrecision(numStr, precisionStr, roundingStr)
{
	var convNumber = new Number(parseFloat(numStr));
	var rNumber = numStr;
	
	// Set Precision value
	var precDepth = ParsePrecision(precisionStr);
	
	// Do rounding
	if(roundingStr == "normal")
	{
		rNumber = convNumber.toFixed(precDepth);
		//Defect #780823 It's a broken implementation of toFixed(), 0.95.toFixed(0)=1, 0.94.toFixed(0)=0, so we apply a workaround

    var depthMulti = 1.0;
    var i;
		for(i = 0; i < precDepth; i++)
		{
			depthMulti *= 10;
		}
    var cNumber = convNumber * depthMulti + 0.5; 
    var floorNumber = Math.floor(cNumber); 
    var rcNumber = floorNumber / depthMulti; 
    rNumber = rcNumber.toFixed(precDepth);
	}
	else if(roundingStr == "round up")
	{
		var depthMulti = 1;
		var i;
		for(i = 0; i < precDepth; i++)
		{
			depthMulti *= 10;
		}
		var cNumber = convNumber * depthMulti;
		var ceilNumber = Math.ceil(cNumber);
		var rcNumber = ceilNumber / depthMulti;
		rNumber = rcNumber.toFixed(precDepth);
	}
	else if(roundingStr == "round down")
	{
		var depthMulti = 1;
		var i;
		for(i = 0; i < precDepth; i++)
		{
			depthMulti *= 10;
		}
		var cNumber = convNumber * depthMulti;
		var floorNumber = Math.floor(cNumber);
		var rcNumber = floorNumber / depthMulti;
		rNumber = rcNumber.toFixed(precDepth);
	}
	else if(roundingStr == "truncate")
	{
		var depthMulti = 1;
		var i;
		for(i = 0; i < precDepth; i++)
		{
			depthMulti *= 10;
		}
		var cNumber = convNumber * depthMulti;
		var roundNumber = Math.floor(cNumber);
		var rcNumber = roundNumber / depthMulti;
		rNumber = rcNumber.toFixed(precDepth);
	}	
	else
	{
		rNumber = convNumber;
	}
	return rNumber;
} 

/////////////////////////////////////////////////////////////////////////////////////////////////////////
// Station formatting

function FormatStation(stationStr, patterStr, precisionStr, roundingStr)
{
	var convNumber = new Number(parseFloat(stationStr));
	var precDepth;
	var rNumber;

	rNumber = FormatPrecision(stationStr,  precisionStr, roundingStr);
	
	// Do Station Pattern
	var numStr = rNumber.toString();
	
	var plusIndex;
	var dotIndexPattern;
	var dotIndexNumber;
	var insertIndex;
	
	plusIndex = FindPlusIndex(patterStr);
	dotIndexPattern = FindDotIndex(patterStr);
	dotIndexNumber = FindDotIndex(numStr);
	
	insertIndex = dotIndexNumber - (dotIndexPattern - plusIndex - 1)
	
	var rString = "";
	var startIndex = 0;
	
	if (convNumber < 0)
	{
		insertIndex = insertIndex - 1;
	}
	//defect #1043195 & #1043205
	if (insertIndex <=0) {
	  rString = rString +"0+";
	  for (var i=0;i <Math.abs(insertIndex);i++) {
       rString= rString+"0"
    }
  }
	for(var i = 0; i < numStr.length; i++)
	{
		if(i == insertIndex && i !=0 && plusIndex > 0)
		{
			rString = rString + '+';
		}
		rString = rString + numStr.charAt(i);
	}
	
	return rString;
}
function FindPlusIndex(stationPattern)
{
	var decPrec;
	var s = stationPattern.indexOf("+");
	if(s < 0)
	{
		decPrec= -1; //stationPattern.length;
	}
	else
	{
		var d = s;
		decPrec= d;
	}	
	return decPrec;
}

function FindDotIndex(dotStr)
{
	var decPrec;
	var s = dotStr.indexOf(".");
	if(s < 0)
	{
		decPrec=dotStr.length;
	}
	else
	{
		var d = s;
		decPrec= d;
	}	
	return decPrec;

}
/////////////////////////////////////////////////////////////////////////////////////////////////////////
// Unit Conversion Functions
//--------------------------------------------------------------
// English to English - Linear
//--------------------------------------------------------------
function FeetToUSSurveyFoot(feet)
{
	var ft = new Number(parseFloat(feet));
	return ft * 1.000002000004
}

function MetersToUSSurveyFoot(meters)
{
	var met = new Number(parseFloat(meters));
	return (met / .3048) * 1.000002000004
}

function FeetToMiles(feet)
{
	var ft = new Number(parseFloat(feet));
	return ft / 5280;
}

function MilesToFeet(miles)
{
	var mls = new Number(parseFloat(miles));
	return 5280 * mls;
}
//--------------------------------------------------------------
// English to English - Area
//--------------------------------------------------------------
function SqFtToAcres(sqFt)
{
	var sq = new Number(parseFloat(sqFt));
	return sq / 43560;
}
function AcresToSqFeet(acres)
{
	var acrs = new Number(parseFloat(acres));
	return acrs * 43560;
}

function AcresToSqMiles(acres)
{
	var acrs = new Number(parseFloat(acres));
	return acrs / 640;
}
//--------------------------------------------------------------
// Metric to Metric - Length
//--------------------------------------------------------------
function CmToMeter(d)
{
	return d/100;
}
function CmToKm(d)
{
	return d/100000;
}
function CmToMm(d)
{
	return d*10;
}
function MmToCm(d)
{
	return d/10;
}

//--------------------------------------------------------------
// Metric to Metric - Area
//--------------------------------------------------------------


//--------------------------------------------------------------
// Metric to English - Length
//--------------------------------------------------------------
function MmToInches(d)
{
	return d/.254;
}
function CmToFeet(d)
{
	return d / 30.48;
}
function MToFeet(d)
{
	return d/.3048;
}

function CmToInch(d)
{
	return d/2.54;
}
function CmToYards(d)
{
	return d/91.44;
}
function CmToMiles(d)
{
	return d/160934.4;
}
function KmToMiles(d)
{
	return d / 1.609344;
}
//--------------------------------------------------------------
// English to Metric - Length
//--------------------------------------------------------------

function InchesToMm(d)
{
	return d * 25.4;
}

function InchesToCm(d)
{
	return d * 2.54;
}
function FeetToMm(d)
{
	return d * 304.8;
}
function FeetToCm(d)
{
	return d * 30.48;
}

function FeetToMeters(d)
{
	return d * .3048;
}
function FeetToKm(d)
{
	return d * .003048;
}

function YardsToMeters(d)
{
	return d * .9144;
}

function MilesToKm(d)
{
	return d * 1.609344;
}

//--------------------------------------------------------------
// English to Metric - Area
//--------------------------------------------------------------
function SqInchesToSqCm(d)
{
	return d * 6.4516;
}

function SqFeetToSqMeters(d)
{
	return d * 0.09290304;
}

function SqYardsToSqMeters(d)
{
	return d * 0.83612736;
}

function SqMilesToSqKm(d)
{
	return d * 2.589988110336;
}

function AcresToHectares(d)
{
	return d * 0.40468564224;
}

//--------------------------------------------------------------
// Metric to English - Area
//--------------------------------------------------------------
function SqCmToSqInches(d)
{
	return d / 6.4516;
}
function SqMetersToSqFeet(d)
{
	return d / 0.09290304;
}
function SqMetersToSqYards(d)
{
	return d / 0.83612736;
}
function SqKmToSqMiles(d)
{
	return d / 2.589988110336;
}
function HectaresToAcres(d)
{
	return d / 0.40468564224;
}
function SqMetersToAcres(d)
{
	return d / 4046.8564224;
}

]]>
</msxsl:script>

</xsl:stylesheet>