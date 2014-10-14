<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet 	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
				xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
				xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit" 
				version="1.0">
<xsl:template match="lx:Curve" mode="set">
	<xsl:variable name="rot" select="landUtils:SetRotation(string(@rot))"></xsl:variable>
	<xsl:variable name="cLength" select="landUtils:SetCurveLength(string(@length))"></xsl:variable>


	<xsl:choose >
		<xsl:when test="./lx:Start/@PntRef">
			<xsl:variable name="sp" select="./lx:Start/@pntRef"></xsl:variable>
			<xsl:variable name="nref" select="landUtils:SetStartingPoint(string(//lx:CgPoints/CgPoint[@name=$sp]))"></xsl:variable>
		</xsl:when>
		<xsl:otherwise>
			<xsl:variable name="nref" select="landUtils:SetStartingPoint(string(./lx:Start))"></xsl:variable>
		</xsl:otherwise>
	</xsl:choose>

	<xsl:choose >
		<xsl:when test="./lx:Center/@PntRef">
			<xsl:variable name="cp" select="./lx:Center/@pntRef"></xsl:variable>
			<xsl:variable name="enref" select="landUtils:SetCenterPoint(string(//lx:CgPoints/CgPoint[@name=$cp]))"></xsl:variable>
		</xsl:when>
		<xsl:otherwise>
			<xsl:variable name="enref" select="landUtils:SetCenterPoint(string(./lx:Center))"></xsl:variable>
		</xsl:otherwise>
	</xsl:choose>

	
	<xsl:choose >
		<xsl:when test="./lx:End/@PntRef">
			<xsl:variable name="ep" select="./End/@pntRef"></xsl:variable>
			<xsl:variable name="enref" select="landUtils:SetEndPoint(string(//lx:CgPoints/CgPoint[@name=$ep]))"></xsl:variable>
		</xsl:when>
		<xsl:otherwise>
			<xsl:variable name="enref" select="landUtils:SetEndPoint(string(./lx:End))"></xsl:variable>
		</xsl:otherwise>
	</xsl:choose>


	<!-- <xsl:variable name="cen" select="landUtils:SetCenterPoint(string(./lx:Center))"></xsl:variable> -->
	<xsl:variable name="calc" select="landUtils:CalculateCurveValues()"></xsl:variable>
</xsl:template>
<msxsl:script language="JScript" implements-prefix="landUtils"><![CDATA[
var rotation = 0;
var startNorthing = 0;
var startEasting = 0;
var centerNorthing = 0;
var centerEasting = 0;
var endNorthing = 0;
var endEasting = 0;
var curvePINorthing = 0;
var curvePIEasting = 0;
var curveLen = 0;

var curveAngle = 0;
var curveRadius = 0;
var startDir = 0;
var endDir = 0;

// -----------------------------------------------------------------
// Rotation property functions
// -----------------------------------------------------------------
function SetRotation(rotText)
{
	var rt = rotText;
	var cwt = "cw";
	if(rt == cwt )
	{
		rotation = 1;
	}
	else
	{
		rotation = -1;
	}
	
	return rotText;
}
function GetRotatonValue()
{
	return rotation;
}
function GetRotationText()
{
	if(rotation == 1)
	{
		return "CW";
	}
	else
	{
		return "CCW";
	}
}
function GetCurveDirection()
{
	if(rotation == 1)
	{
		return "RIGHT";
	}
	else
	{
		return "LEFT";
	}
}
// -----------------------------------------------------------------
// Starting Point (PC) property functions
// -----------------------------------------------------------------
function SetStartingPoint(pointStr)
{
	var strPoint = pointStr.split(" ");
	startNorthing = new Number(parseFloat(strPoint[0]));
	startEasting = new Number(parseFloat(strPoint[1]));
	return pointStr;
}
function SetStartingNorthing(northingText)
{
	startNorthing = new Number(parseFloat(northingText));
	return northingText;
}
function SetStartingEasting(eastingText)
{
	startEasting = new Number(parseFloat(eastingText));
	return eastingText;
}
function GetStartingNorthing()
{
	return (0 + startNorthing);
}
function GetStartingEasting()
{
	return (0 + startEasting);
}
// -----------------------------------------------------------------
// Center Point (RP) property functions
// -----------------------------------------------------------------
function SetCenterPoint(pointStr)
{
	var strPoint = pointStr.split(" ");
	centerNorthing = new Number(parseFloat(strPoint[0]));
	centerEasting = new Number(parseFloat(strPoint[1]));
	return pointStr;
}
function SetCenterNorthing(northingText)
{
	centerNorthing = new Number(parseFloat(northingText));
	return northingText;
}
function SetCenterEasting(eastingText)
{
	centerEasting = new Number(parseFloat(eastingText));
	return eastingText;
}
function GetCenterNorthing()
{
	//return formatCoordNumber(centerNorthing);
	return (0 + centerNorthing);
}
function GetCenterEasting()
{
//	return formatCoordNumber(centerEasting);
	return (0 + centerEasting);

}
// -----------------------------------------------------------------
// Ending Point (PT) property functions
// -----------------------------------------------------------------
function SetEndPoint(pointStr)
{
	var strPoint = pointStr.split(" ");
	endNorthing = new Number(parseFloat(strPoint[0]));
	endEasting = new Number(parseFloat(strPoint[1]));
	return pointStr;
}
function SetCurveEndNorthing(northingText)
{
	endNorthing = new Number(parseFloat(northingText));
	return northingText;
}
function SetCurveEndEasting(eastingText)
{
	endEasting = new Number(parseFloat(eastingText));
	return eastingText;
}
function GetCurveEndNorthing()
{
	return "" + endNorthing;
	//return formatCoordNumber(endNorthing);
}
function GetCurveEndEasting()
{
	return "" + endEasting;
//	return formatCoordNumber(endEasting);
}
// -----------------------------------------------------------------
// PI property functions
// -----------------------------------------------------------------
function GetCurvePINorthing()
{
//	return formatCoordNumber(curvePINorthing );
	return "" + curvePINorthing;
}
function GetCurvePIEasting()
{
	//return formatCoordNumber(curvePIEasting );
	return "" + curvePIEasting;
}
// -----------------------------------------------------------------
// Tangent property functions
// -----------------------------------------------------------------
function GetStartDirection()
{
//	return formatAngleNumber(startDir );
	return startDir;
}
function GetEndDirection()
{
//	return formatAngleNumber(endDir );
	return endDir;
}
// -----------------------------------------------------------------
// Curve calculation functions
// -----------------------------------------------------------------
function CalculateCurveValues()
{
	var ca = GetCurveAngle();
	var cr = GetCurveRadius();
	CalculatePI();
	return "done";	
}
function CalculatePI()
{
	var tanLen = GetCurveTangent();
	var bearAngle = CalculateBearingAngle(startDir - 90);
	//var bearAngle = startDir - 90;
	
	curvePINorthing =startNorthing - (tanLen * Math.cos((bearAngle * Math.PI) / 180 ));
	//curvePINorthing = bearAngle;
	//curvePIEasting = (tanLen * Math.sin(bearAngle));
	curvePIEasting = startEasting + (tanLen * Math.sin((bearAngle * Math.PI) / 180 ));
}
function CalculateBearingAngle(angle)
{
	if(angle <= 90 && angle > 0)
	{
		return angle;
	}
	else if(angle <= 180 && angle > 90)
	{
		return angle - 90;
	}
	else if(angle <= 270 && angle > 180)
	{
		return 270 - angle;
	}
	else
	{
		return angle - 270;
	}
}

function GetCurveAngle()
{
	curveAngle = 0;
	startDir = 0;
	endDir = 0;
	var startXDiff = centerEasting - startEasting;
	var startYDiff = centerNorthing - startNorthing;
				
				
	var endXDiff = centerEasting - endEasting;
	var endYDiff = centerNorthing - endNorthing;

	var startTan =startXDiff / startYDiff  ;
	var endTan = endXDiff / endYDiff;


	var startDirTmp = Math.atan(startTan); 
	var endDirTmp = Math.atan(endTan);
	
	//var R = GetCurveRadius();
	//var L = GetCurveLength();
	//var Delta = (L / R);
	//return Delta;
	
	//One attempt
	//var totalXDiff = endEasting - startEasting;
	//var totalYDiff = endNorthing - startNorthing;
	//var Chord = Math.sqrt(Math.pow(totalXDiff, 2) + Math.pow(totalYDiff, 2));
	//var Temp1 = Chord / (2 * R);
	//var Angle = 2 * (Math.asin(Temp1));
	//return Angle;
	
	//Another attempt
	//var StartToCenter = Math.sqrt(Math.pow(startXDiff, 2) + Math.pow(startYDiff, 2));
	//var Temp1 = (StartToCenter) / StartToEnd;
	//var Temp2 = Math.asin(Temp1);
	//var Temp3 = Math.cos(Temp2 / 2);
	//var angle = Math.asin(Temp3);
	//return 0 + Temp1;
	//return startDirTmp;
	
	
    		if(centerNorthing > startNorthing)
    		{
     		startDir = 270 - ( (startDirTmp * 180) / Math.PI); 
   		}
    		else
    		{
    			startDir =90 - ( (startDirTmp * 180) / Math.PI) ;
    		}
    		
    		
    		
    		if(centerNorthing > endNorthing)
    		{
    			endDir= 270 -  ( (endDirTmp * 180) / Math.PI);    		
  		}
    		else
    		{
    			endDir= 90 - ( (endDirTmp * 180) / Math.PI);
 		}
    	
    		
    	if(rotation == 1)
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
		
	
	return (curveAngle );  //return formatAngleNumber(curveAngle );
}


function GetCurveRadius()
{
	var yDiff = startNorthing - centerNorthing;
	var xDiff = startEasting - centerEasting;
	
	curveRadius = Math.sqrt(Math.pow(xDiff, 2) + Math.pow(yDiff, 2));
	return curveRadius;
}
function GetChordLength()
{
//	return formatLinearNumber(2 * curveRadius * Math.sin(.017453293 * (curveAngle / 2)));
	return (2 * curveRadius * Math.sin(.017453293 * (curveAngle / 2)));
}
function GetChordDirection_Old()
{
	var xDiff = startEasting - endEasting;
   	var yDiff = startNorthing - endNorthing;
   	var tanA = yDiff / xDiff;
    	
    	var angle = (Math.atan(tanA)* 180)/Math.PI;

	return 180 + angle;
	//return angle;
	//return Math.abs(360 - angle);
}

function GetChordDirection()
{
	var xDiff = startEasting - endEasting;
   	var yDiff = startNorthing - endNorthing;
   	var tanA = yDiff / xDiff;
    	
    	var angle = (Math.atan(tanA)* 180)/Math.PI;

	if (angle > 0)
	{
    		if(startNorthing > endNorthing)  // SW
    		{
     		angle = 270.0 - angle; 
   		}
    		else    // NE
    		{
    			return angle;
    		}
    }
    
    else if (angle < 0)
    {
        	if(startNorthing > endNorthing)   //SE
    		{
     		angle = 90 - angle;
   		}
    		else      //NW
    		{
    			angle = 270 - angle;
    			return angle;
    		}
    }
    
    else
    {
    		if(startEasting > endEasting)
    		{
     		angle = 180.00;
   		}
    		else
    		{
    			angle = 0.00;
		}
	}

	return angle;
}




function GetCurveTangent()
{
//	return formatLinearNumber(curveRadius * (Math.tan(.017453293 * (curveAngle / 2))));
	return Math.abs(curveRadius * (Math.tan(.017453293 * (curveAngle / 2))));
}

function SetCurveLength(length)
{
	curveLen = length; //new Number(parseFloat(length));
	return length;
}
function GetCurveLength()
{
	return "" + curveLen;
}

function GetCurveLength_NoGood()
{
	 var Rad = GetCurveRadius();
	 var Angle = GetCurveAngle();
	return Angle;
//	return ((Math.PI * Rad * Angle) / 180);
	
	//return ((Math.PI * curveRadius * curveAngle) / 180);
	
	//	return formatLinearNumber((Math.PI * curveRadius * curveAngle) / 180);
}

function GetCurveExternal()
{
	var tan = curveRadius * (Math.tan(.017453293 * (curveAngle / 2)));
	var ext = tan * Math.tan(.017453293 * (curveAngle / 4));
//	return formatLinearNumber(ext);
	return ext;
}

function GetCurveMiddle()
{
	var R = GetCurveRadius();
	var L = GetCurveLength();

	var A = GetCurveAngle();
	var Delta = (L / R);

	
	
	var mid = R *  (1 - (Math.cos(.017453293 * (A / 2))));
//	var mid = curveRadius *  (1 - (Math.cos(.017453293 * (curveAngle / 2))));
	return mid;
//	return formatLinearNumber(mid);
}
function GetCurveDOC()
{
	var doc = 5729.578 / curveRadius;
	return doc;
//	return formatAngleToDMS(doc);
}

]]></msxsl:script>
</xsl:stylesheet>
