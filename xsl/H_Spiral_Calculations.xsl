<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet 	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
				xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
				xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit" 
				version="1.0">
                		
<xsl:template match="lx:Spiral" mode="set">
 	<xsl:variable name="spitype"  select="landUtils:SetSpiralType(@spiType)"></xsl:variable>
	<xsl:variable name="spiRot" select="landUtils:SetSpiralRotation(@rot)"></xsl:variable>
	<xsl:variable name="spiRadStart" select="landUtils:SetSpiralStartRadius(string(@radiusStart))"></xsl:variable>
	<xsl:variable name="spiRadEnd" select="landUtils:SetSpiralEndRadius(string(@radiusEnd))"></xsl:variable>
	<xsl:variable name="spilen" select="landUtils:SetSpiralLength(string(@length))"></xsl:variable>
	<xsl:variable name="spitamlong" select="landUtils:SetSpiralLongTangent(string(@tanLong))"></xsl:variable>    
	<xsl:variable name="spitanshort" select="landUtils:SetSpiralShortTangent(string(@tanShort))"></xsl:variable> 
	<xsl:variable name="spitheta" select="landUtils:SetSpiralTheta(string(@theta))"></xsl:variable>            

	<xsl:choose >
		<xsl:when test="./lx:Start/@PntRef">
			<xsl:variable name="sp" select="./lx:Start/@pntRef"></xsl:variable>
			<xsl:variable name="nref" select="landUtils:SetSpiralStartingPoint(string(//lx:CgPoints/CgPoint[@name=$sp]))"></xsl:variable>
		</xsl:when>
		<xsl:otherwise>
			<xsl:variable name="nref" select="landUtils:SetSpiralStartingPoint(string(./lx:Start))"></xsl:variable>
		</xsl:otherwise>
	</xsl:choose>

	<xsl:choose >
		<xsl:when test="./lx:PI/@PntRef">
			<xsl:variable name="cp" select="./lx:PI/@pntRef"></xsl:variable>
			<xsl:variable name="enref" select="landUtils:SetSpiralPIPoint(string(//lx:CgPoints/CgPoint[@name=$cp]))"></xsl:variable>
		</xsl:when>
		<xsl:otherwise>
			<xsl:variable name="enref" select="landUtils:SetSpiralPIPoint(string(./lx:PI))"></xsl:variable>
		</xsl:otherwise>
	</xsl:choose>

	
	<xsl:choose >
		<xsl:when test="./lx:End/@PntRef">
			<xsl:variable name="ep" select="./End/@pntRef"></xsl:variable>
			<xsl:variable name="enref" select="landUtils:SetSpiralEndPoint(string(//lx:CgPoints/CgPoint[@name=$ep]))"></xsl:variable>
		</xsl:when>
		<xsl:otherwise>
			<xsl:variable name="enref" select="landUtils:SetSpiralEndPoint(string(./lx:End))"></xsl:variable>
		</xsl:otherwise>
	</xsl:choose>


</xsl:template>
<msxsl:script language="JScript" implements-prefix="landUtils"><![CDATA[

var spiType = "CLOTHOID";
var spiRotation = "CCW"
var spiLength = 0;
var spiLongTanLength = 0;
var spiShortTanLength = 0;
var spiStartRadius = 0;
var spiEndRadius = 0;
var spiRadius = 0;
var spiTheta = 0;

var spiStartNorthing = 0;
var spiStartEasting = 0;
var spiPINorthing = 0;
var spiPIEasting = 0;
var spiEndNorthing = 0;
var spiEndEasting = 0;

// -----------------------------------------------------------------
// Setting Spiral default values
// -----------------------------------------------------------------
function SetSpiralType(spitype)
{
	spiType = spitype;
	return spiType;
}
function GetSpiralType()
{
	return spiType;
}
// -----------------------------------------------------------------
function SetSpiralRotation(rot)
{
	spiRotation = rot;
	return rot;
}
function GetSpiralRotation()
{
	return spiRotation;
}
// -----------------------------------------------------------------
function SetSpiralLength(length)
{
	spiLength = new Number(parseFloat(length));
	return length;
}
function GetSpiralLength()
{
	return "" + spiLength;
}
// -----------------------------------------------------------------
function SetSpiralStartRadius(radius)
{
	spiStartRadius = new Number(parseFloat(radius));
	return radius;
}
function GetSpiralStartRadius()
{
	return "" + spiStartRadius;
}
// -----------------------------------------------------------------
function SetSpiralEndRadius(radius)
{
	spiEndRadius = new Number(parseFloat(radius));
	return radius;
}
function GetSpiralEndRadius()
{
	return "" + spiEndRadius;
}
// -----------------------------------------------------------------

function GetSpiralRadius()
{
	var rStart = GetSpiralStartRadius();
	var r1 = new Number(rStart);
	var rEnd = GetSpiralEndRadius();
	var r2 = new Number(rEnd);
	if (rStart < rEnd) //NaN is greater that any number. Since one of the slopes will be infinite, I want the other one.
		return rStart;
	else
		return rEnd;
}
// -----------------------------------------------------------------
// Spiral Starting Point property functions
// -----------------------------------------------------------------
function SetSpiralStartingPoint(pointStr)
{
	var strPoint = pointStr.split(" ");
	spiStartNorthing = new Number(parseFloat(strPoint[0]));
	spiStartEasting = new Number(parseFloat(strPoint[1]));
	return pointStr;
}
function SetSpiralStartNorthing(northingText)
{
	spiStartNorthing = new Number(parseFloat(northingText));
	return northingText;
}
function SetSpiralStartEasting(eastingText)
{
	spiStartEasting = new Number(parseFloat(eastingText));
	return eastingText;
}
function GetSpiralStartNorthing()
{
//	return formatCoordNumber(spiStartNorthing);
	return (0 + spiStartNorthing);
}
function GetSpiralStartEasting()
{
//	return formatCoordNumber(spiStartEasting);
	return (0 + spiStartEasting);
}
// -----------------------------------------------------------------
// Spiral PI Point property functions
// -----------------------------------------------------------------
function SetSpiralPIPoint(pointStr)
{
	var strPoint = pointStr.split(" ");
	spiPINorthing = new Number(parseFloat(strPoint[0]));
	spiPIEasting = new Number(parseFloat(strPoint[1]));
	return pointStr;
}
function SetSpiralPINorthing(northingText)
{
	spiPINorthing = new Number(parseFloat(northingText));
	return northingText;
}
function SetSpiralPIEasting(eastingText)
{
	spiPIEasting = new Number(parseFloat(eastingText));
	return eastingText;
}
function GetSpiralPINorthing()
{
	return (0 + spiPINorthing);
//	return formatCoordNumber(spiPINorthing);
}
function GetSpiralPIEasting()
{
	return (0 + spiPIEasting);
//	return formatCoordNumber(spiPIEasting);
}
// -----------------------------------------------------------------
// Ending Point (PT) property functions
// -----------------------------------------------------------------
function SetSpiralEndPoint(pointStr)
{
	var strPoint = pointStr.split(" ");
	spiEndNorthing = new Number(parseFloat(strPoint[0]));
	spiEndEasting = new Number(parseFloat(strPoint[1]));
	return pointStr;
}
function SetSpiralEndNorthing(northingText)
{
	spiEndNorthing = new Number(parseFloat(northingText));
	return northingText;
}
function SetSpiralEndEasting(eastingText)
{
	spiEndEasting = new Number(parseFloat(eastingText));
	return eastingText;
}
function GetSpiralEndNorthing()
{
	return (0 + spiEndNorthing);
//	return formatCoordNumber(spiEndNorthing);
}
function GetSpiralEndEasting()
{
	return (0 + spiEndEasting);
//	return formatCoordNumber(spiEndEasting);
}

// -----------------------------------------------------------------
// Formula calculation functions
// -----------------------------------------------------------------

// -----------------------------------------------------------------

function GetSpiralChordDirection()
{
	var xDiff = spiStartEasting - spiEndEasting;
   	var yDiff = spiStartNorthing - spiEndNorthing;
   	var tanA = yDiff / xDiff;
    	
    	var angle = (Math.atan(tanA)* 180)/Math.PI;

	if (angle > 0)
	{
    		if(spiStartNorthing > spiEndNorthing)  // SW
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
        	if(spiStartNorthing > spiEndNorthing)   //SE
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
    		if(spiStartEasting > spiEndEasting)
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

function GetSpiralChordLength()
{

	var xDiff = spiStartEasting - spiEndEasting;
   	var yDiff = spiStartNorthing - spiEndNorthing;
   	return Math.sqrt(Math.pow(xDiff, 2) + Math.pow(yDiff, 2));
   	var tanA = yDiff / xDiff;

	return (2 * curveRadius * Math.sin(.017453293 * (curveAngle / 2)));
}

function GetSpiralTS()
{
	var R = GetSpiralRadius();
	var P = GetSpiralP(R);
	var K = GetSpiralK(R);
	
}


function SetSpiralTheta(theta)
{
	spiTheta = new Number(parseFloat(theta));
	return theta;
}
function GetSpiralTheta_NoGood()
{
	return "" + spiTheta;
}

function SetSpiralLongTangent(tanlength)
{
	spiLongTanLength = new Number(parseFloat(tanlength));
	return tanlength;
}
function GetSpiralLongTangent_NoGood()
{
	return "" + spiLongTanLength;
}

function SetSpiralShortTangent(tanlength)
{
	spiShortTanLength = new Number(parseFloat(tanlength));
	return tanlength;
}
function GetSpiralShortTangent_NoGood()
{
	return "" + spiShortTanLength;
}

function GetSpiralA(Radius)
{
	var Len = GetSpiralLength();
	var A = Math.sqrt(Len * Radius);
	//return formatLinearNumber(A);
	return (A);
}

function GetSpiralTotalX()
{
	var Len = GetSpiralLength();
	var Radius = GetSpiralRadius();

	var Temp1 = (Math.pow(Len, 2)) / (40 * Math.pow(Radius, 2));
	var Temp2 = (Math.pow(Len, 4)) / (3456 * Math.pow(Radius, 4));
	var TotalX = Len * (1 - Temp1 + Temp2);
	
	return (TotalX);
}

function GetSpiralTotalY()
{
	var Len = GetSpiralLength();
	var Radius = GetSpiralRadius();

	var Temp1 = (Math.pow(Len, 2)) / (Radius * 6);
	var Temp2 = (Math.pow(Len, 2)) / (56 * Math.pow(Radius, 2));
	var Temp3 = (Math.pow(Len, 4)) / (7040 * Math.pow(Radius, 4));
	var TotalY = Temp1 * (1 - Temp2 + Temp3);

	return (TotalY);
}

function GetSpiralX(Radius)
{
	var Len = GetSpiralLength();
	
	var K = Math.sqrt(Len * Radius * 2);
	var Temp1 = (Math.pow(Len, 4)) / (10 * Math.pow(K, 4));
	var Temp2 = (Math.pow(Len, 8)) / (216 * Math.pow(K, 8));
	var X = Len * (1 - Temp1 + Temp2);
	
	return (X);
}

function GetSpiralY(Radius)
{
	var Len = GetSpiralLength();
	var Temp1 = (Math.pow(Len, 3)) / (Radius * Len * 6);
	var Temp2 = (Math.pow(Len, 4)) / (56 * Math.pow(Radius, 2) * Math.pow(Len, 2));
	var Temp3 = (Math.pow(Len, 8)) / (7040 * Math.pow(Radius, 4) * Math.pow(Len, 4));
	var Y = Temp1 * (1 - Temp2 + Temp3);

	return (Y);
}

function GetSpiralLongTangent()
{	
	var Theta = GetSpiralTheta_Radians();
	var TotalX = GetSpiralTotalX();
	var TotalY = GetSpiralTotalY();
	var Cotan = 1 / Math.tan(Theta);
	
	var LongTan = TotalX - (TotalY * Cotan);	
	return (LongTan);
}

function GetSpiralShortTangent()
{
	var Theta = GetSpiralTheta_Radians();
	var TotalY = GetSpiralTotalY();
	var Cosecant = 1 / Math.sin(Theta);
	
	var ShortTan = TotalY * Cosecant;	
	return (ShortTan);
}

function GetSpiralTheta_Radians() //Used in calculations for Short and Long Tangents
{
	var Len = GetSpiralLength();
	var Radius = GetSpiralRadius();

	var Theta = Len / (2 * Radius);	
	return (Theta);
}

function GetSpiralTheta_Degrees() //This is the reported theta value
{
	var Len = GetSpiralLength();
	var Radius = GetSpiralRadius();

	var Theta_In_Radians = Len / (2 * Radius);	
	
	var Theta = (Theta_In_Radians * 180) / Math.PI;
	return Theta;
}


function GetSpiralP(Radius)
{
	var Len = GetSpiralLength();
	var Y = GetSpiralY(Radius);
	var Temp1 = Len / (Radius * 2);
	var Temp2 = Math.cos(Temp1);
	var P = Y - (Radius * (1-Temp2));
	return(P);
}

function GetSpiralK(Radius)
{
	var Len = GetSpiralLength();
	var P = GetSpiralP(Radius);
	var K = (Len / 2) * (1 - (P / (5 * Radius)));
	return (K);
}


]]></msxsl:script>
</xsl:stylesheet>
