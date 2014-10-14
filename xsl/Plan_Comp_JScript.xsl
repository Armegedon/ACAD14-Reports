<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2002 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet 	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
					xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
               		xmlns:msxsl="urn:schemas-microsoft-com:xslt"
               		xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
               		xmlns:lxml="urn:lx_utils"
               		version="1.0">
<msxsl:script language="JScript" implements-prefix="landUtils"><![CDATA[

function NodeToText(node)
{
	var str;
	if(node.length > 0)
	{
		return node.item(0).text;
	}
	else
	{
		return "" + node;
	}
}
function GetNodeType(node)
{
	return node;
}

function CalculateLength(pointOne, pointTwo)
{
	var oneStr, twoStr;	

	oneStr = NodeToText(pointOne);
	twoStr = NodeToText(pointTwo);

	var oneStrs = oneStr.split(" ");
	var twoStrs= twoStr.split(" ");
	
	var nOne = new Number(oneStrs[0]);
	var eOne = new Number(oneStrs[1]);
	
	var nTwo = new Number(twoStrs[0]);
	var eTwo = new Number(twoStrs[1]);
	
	var yDiff = nTwo - nOne;
	var xDiff = eTwo - eOne;
	
	var len = Math.sqrt(Math.pow(xDiff, 2) + Math.pow(yDiff, 2));

	return formatLinearNumber(len);
}

function CalculateLengthText(oneStr, twoStr)
{
	var oneStrs = oneStr.split(" ");
	var twoStrs= twoStr.split(" ");
	
	var nOne = new Number(oneStrs[0]);
	var eOne = new Number(oneStrs[1]);
	
	var nTwo = new Number(twoStrs[0]);
	var eTwo = new Number(twoStrs[1]);
	
	var yDiff = nTwo - nOne;
	var xDiff = eTwo - eOne;
	
	var len = Math.sqrt(Math.pow(xDiff, 2) + Math.pow(yDiff, 2));

	return formatLinearNumber(len);
}

function CalculateBearingDecimal(pointOne, pointTwo)
{
	var ang = CalculateDirection(pointOne, pointTwo);
	var angNum;
	
	if(ang >= 0 && ang <= 90)
	{
		angNum = formatAngleNumber(90. - ang);
		return "N " + angNum + " E";
	}
	else if(ang > 90 && ang <= 180)
	{
		angNum = formatAngleNumber(ang - 90.);
		return "N " + angNum + " W";
	}
	else if(ang > 180 && ang < 270)
	{
		angNum = formatAngleNumber(270. - ang);
		return "S " + angNum + " W";
	}
	else
	{
		angNum = formatAngleNumber(ang - 270.);
		return "S " + angNum + " E";
	}	
}

function CalculateBearingDMS(pointOne, pointTwo)
{

	var ang = CalculateDirection(pointOne, pointTwo);
	var angNum;
	
	if(ang >= 0 && ang <= 90)
	{
		angNum = 90. - ang;
		var bearing = formatAngleToDMS(angNum);
		return "NORTH " + bearing + " EAST";
	}
	else if(ang > 90 && ang <= 180)
	{
		angNum = ang - 90.;
		var bearing = formatAngleToDMS(angNum);
		return "NORTH " + bearing + " WEST";
	}
	else if(ang > 180 && ang < 270)
	{
		angNum = 270. - ang;
		var bearing = formatAngleToDMS(angNum);
		return "SOUTH " + bearing + " WEST";
	}
	else
	{
		angNum = ang - 270.;
		var bearing = formatAngleToDMS(angNum);
		return "SOUTH " + bearing + " EAST";
	}	
}

function CalculateBearingDMSText(oneText, twoText)
{

	var ang = CalculateDirectionText(oneText, twoText);
	var angNum;
	
	if(ang > 0 && ang < 90)
	{
		angNum = 90 - ang;
		var bearing = formatAngleToDMS(angNum);
		return "N " + bearing + " E";
	}
	else if(ang > 90 && ang < 180)
	{
		angNum = ang - 90;
		var bearing = formatAngleToDMS(angNum);
		return "N " + bearing + " W";
	}
	else if(ang > 180 && ang < 270)
	{
		angNum = 270 - ang;
		var bearing = formatAngleToDMS(angNum);
		return "S " + bearing + " W";
	}
	else
	{
		angNum = ang - 270;
		var bearing = formatAngleToDMS(angNum);
		return "S " + bearing + " E";
	}	
}


function CalculateNorthAzimuth(pointOne, pointTwo)
{
	var ang = CalculateDirection(pointOne, pointTwo);
	var angNum = new Number(ang);
	var azimuth =  angNum - 90;
	
	return azimuth;	
}

function CalculateNorthAzimuthDecimal(pointOne, pointTwo)
{
	var ang = CalculateDirection(pointOne, pointTwo);
	var angNum = new Number(ang);
	var azimuth =  formatAngleNumber(angNum - 90);
	
	return azimuth;	
}

function CalculateNorthAzimuthDMS(pointOne, pointTwo)
{
	var ang = CalculateDirection(pointOne, pointTwo);
	var angNum = new Number(ang);
	
	return formatAngleToDMS(angNum - 90);	
}

function CalculateDirection(pointOne, pointTwo)
{
	oneStr = NodeToText(pointOne);
	twoStr = NodeToText(pointTwo);
	
	var oneStrs = oneStr.split(" ");
	var twoStrs= twoStr.split(" ");
	
	var nOne = new Number(oneStrs[0]);
	var eOne = new Number(oneStrs[1]);
	
	var nTwo = new Number(twoStrs[0]);
	var eTwo = new Number(twoStrs[1]);
	
    var xDiff = eOne - eTwo;
   	var yDiff = nOne - nTwo;
   	
    	var tanA = yDiff / xDiff;
    	var tempAngle = Math.atan(tanA);
    	var angle = 0;
    	
    	if(tempAngle < 0)
    	{
    		if(nOne > nTwo )
    		{
    			angle= 360 + ( (tempAngle * 180) / Math.PI);    		
   		}
    		else
    		{
    			angle= 180 + ( (tempAngle * 180) / Math.PI);
    		}
    	}
    	else
    	{
    		if(nOne > nTwo )
    		{
    			angle=180 + ((tempAngle * 180) / Math.PI);
    		}
    		else
    		{
     			angle= ((tempAngle * 180) / Math.PI);
    		}
    	}
	return angle;
}

function CalculateDirectionText(textOne, textTwo)
{
	var oneStrs = textOne.split(" ");
	var twoStrs= textTwo.split(" ");
	
	var nOne = new Number(oneStrs[0]);
	var eOne = new Number(oneStrs[1]);
	
	var nTwo = new Number(twoStrs[0]);
	var eTwo = new Number(twoStrs[1]);
	
    	var xDiff = eOne - eTwo;
   	var yDiff = nOne - nTwo;
   	
    	var tanA = yDiff / xDiff;
    	var tempAngle = Math.atan(tanA);
    	var angle = 0;
    	
    	if(tempAngle < 0)
    	{
    		if(nOne > nTwo )
    		{
    			angle= 360 + ( (tempAngle * 180) / Math.PI);    		
   		}
    		else
    		{
    			angle= 180 + ( (tempAngle * 180) / Math.PI);
    		}
    	}
    	else
    	{
    		if(nOne > nTwo )
    		{
    			angle=180 + ((tempAngle * 180) / Math.PI);
    		}
    		else
    		{
     			angle= ((tempAngle * 180) / Math.PI);
    		}
    	}
	return angle;
}

function CalculateDirectionDecimal(pointOne, pointTwo)
{
	var angle = CalculateDirection(pointOne, pointTwo);
	return formatAngleNumber(angle);
}

function CalculateDirectionDMS(pointOne, pointTwo)
{
	var angle = CalculateDirection(pointOne, pointTwo);
	return formatAngleToDMS(angle);
}

function CalculateAngle(startPoint, centerPoint, endPoint, rotation)
{
	var rot = rotation.text;
	
	var startDir = CalculateDirection(centerPoint, startPoint);
	var endDir = CalculateDirection(centerPoint, endPoint);
	
	if(rot == "cw"){
		var curveAngle = 180 - (endDir-startDir)  ;
		return formatAngleNumber(curveAngle);
	}
	else
	{
		var curveAngle  = 180 - endDir-startDir ;
		return formatAngleNumber(curveAngle);
	}
}
]]></msxsl:script>
</xsl:stylesheet>