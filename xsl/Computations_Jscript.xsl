<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet version="1.0"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">
	
<xsl:include href="NodeConversion_JScript.xsl"/>

<msxsl:script language="JScript" implements-prefix="landUtils"> 
<![CDATA[
function CalculateLength(startNorthing, startEasting, endNorthing, endEasting)
{
	var retLength = 0;
	var sn = Number(startNorthing);
	var se = Number(startEasting);
	var en = Number(endNorthing);
	var ee = Number(endEasting);
	
	var nDiff = en - sn;
	var eDiff = ee - se;
	
	retLength = Math.sqrt(Math.pow(nDiff, 2) + Math.pow(eDiff, 2));
	
	return retLength;
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
function CalculateDirection(startNorthing, startEasting, endNorthing, endEasting)
{
	var retAngle = 0;
	
	var nOne = Number(startNorthing);
	var eOne = Number(startEasting);
	
	var nTwo = Number(endNorthing);
	var eTwo = Number(endEasting);
	
	var xDiff = eOne  -  eTwo;;
   	var yDiff = nOne - nTwo;

 	var tanA = yDiff / xDiff;
   	var angle = (Math.atan(tanA)* 180)/Math.PI;

	// retAngle will be 0-360 degrees, E=0 deg, counter-clockwise

	if (angle > 0)
	{
		if(nOne > nTwo)  // SW
		{
			retAngle = 180. + angle;
		}
		else    // NE
		{
			retAngle = angle;
		}
	}

	else if (angle < 0)
	{
		if(nOne > nTwo)  // SE
		{
			retAngle = 360. + angle;
		}
		else	// NW
		{
			retAngle = 180. + angle;
		}
	}

	else
	{
		if(eOne > eTwo)
		{
			retAngle = 180.;
		}
		else
   		{
			retAngle = 0.;
		}
	}
	return retAngle ;
}

function CalculateNorthAzimuth(startNorthing, startEasting, endNorthing, endEasting)
{
	var ang = CalculateDirection(startNorthing, startEasting, endNorthing, endEasting);
	var angNum = new Number(ang);
	var azimuth =  angNum - 90;
	
	return azimuth;	
}

function CalculateAngle(startNorthing, startEasting, centerNorthing, centerEasting, endNorthing, endEasting, rotation)
{
	var rot = rotation.text;
	
	var startDir = CalculateDirection(centerNorthing, centerEasting, startNorthing, startEasting);
	var endDir = CalculateDirection(centerNorthing, centerEasting, endNorthing, endEasting);
	
	if(rot == "cw"){
		var curveAngle = 180 - (endDir-startDir)  ;
		return curveAngle;
	}
	else
	{
		var curveAngle  = 180 - endDir-startDir ;
		return curveAngle;
	}
}
function GetIntersectionNorthing(startN, startE, dirOne, dirTwo, endN, endE)
{
	var dOne = Number(dirOne);
	var dTwo = Number(dirTwo);
	
	var sn = Number(startN);
	var se = Number(startE);
	var en = Number(endN);
	var ee = Number(endE);	
	
	var m1 = Math.tan(dOne);
	var m2 = Math.tan(dTwo);
	
	var c1 = sn - (m1 * se);
	var c2 = en - (m2 * ee);
	
	var ret =  ((m1*c2) - (m2*c1)) / (m1 - m2);
	
	return ret;
}

function GetIntersectionEasting(startN, startE, dirOne, dirTwo, endN, endE)
{
	var dOne = Number(dirOne);
	var dTwo = Number(dirTwo);
	
	var sn = Number(startN);
	var se = Number(startE);
	var en = Number(endN);
	var ee = Number(endE);	
	
	var m1 = Math.tan(dOne);
	var m2 = Math.tan(dTwo);
	
	var c1 = sn - (m1 * se);
	var c2 = en - (m2 * ee);
	
	var ret =  (c2 - c1) / (m1 - m2);
	
	return ret;
}

]]>
</msxsl:script>

</xsl:stylesheet>
