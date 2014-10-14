<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet 	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
				xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
				xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit" 
				version="1.0">

<xsl:param name="linearPrec"/>
<xsl:param name="coordPrec"/>
<xsl:param name="anglePrec"/>

<msxsl:script language="JScript" implements-prefix="landUtils"><![CDATA[
var linearPrec = 2;
var coordPrec = 6;
var anglePrec = 3;

var linearUnit="ft";
var areaUnit="sq ft";
var volumeUnit="cu ft";
var tempUnit="f";
var pressUnit="HG";

var unitType="Imperial";

function setUnitType(unitType)
{
	unitType=unitType;
	return unitType;
}
function getUnitType()
{
	return unitType;
}

function setLinearUnit(Unit)
{
	linearUnit=Unit;
	return Unit;
}
function getLinearUnit()
{
	return linearUnit;
}

function setAreaUnit(Unit)
{
	areaUnit=Unit;
	return Unit;
}
function getAreaUnit()
{
	return areaUnit;
}

function setVolumeUnit(Unit)
{
	volumeUnit =Unit;
	return Unit;
}
function getVolumeUnit()
{
	return volumeUnit ;
}

function setTemperatureUnit(Unit)
{
	tempUnit=Unit;
	return Unit;
}
function getTemperatureUnit()
{
	return tempUnit;
}

function setPressureUnit(Unit)
{
	pressUnit=Unit;
	return Unit;
}
function getPressureUnit()
{
	return pressUnit;
}

//Precision Settings
function setLinearPrec(prec)
{
	var s = prec.indexOf(".");
	if(s < 0)
	{
		linearPrec = 0;
	}
	else
	{
		var d = prec.length - s - 1;
		linearPrec = d;
	}	
	return prec;
}

function setCoordPrec(prec)
{
	var s = prec.indexOf(".");
	if(s < 0)
	{
		coordPrec = 0;
	}
	else
	{
		var d = prec.length - s - 1;
		coordPrec = d;
	}	
	return prec;
}

function setAnglePrec(prec)
{
	var s = prec.indexOf(".");
	if(s < 0)
	{
		anglePrec = 0;
	}
	else
	{
		var d = prec.length - s - 1;
		anglePrec = d;
	}	
	return prec;
}

function formatLinearNumber(number)
{
	var strFormatted;
	strFormatted = number.toFixed(linearPrec);
	return strFormatted;
}
function formatLinearText(text)
{
	var strFormatted;

	var strNum = new Number(parseFloat(text));
	
	if(isNaN(strNum))
	{
		return numStr;
	}
	else
	{
		strFormatted = strNum.toFixed(linearPrec);
		return strFormatted;
	}
}

function formatLinearString(numStr)
{
	var strFormatted;
	var str = numStr.nextNode().text;
	var strNum;

	strNum = new Number(parseFloat(str));
	
	if(isNaN(strNum))
	{
		return numStr;
	}
	else
	{
		strFormatted = strNum.toFixed(linearPrec);
		return strFormatted;
	}
}

function formatLinearAttribute(numAttr)
{
	var strFormatted;
	var numStr = numAttr.nextNode().text;
	var strNum;
	strNum = new Number(numStr);
	strFormatted = strNum.toFixed(linearPrec);
	return strFormatted;
}

function formatCoordNumber(number)
{
	var strFormatted;
	strFormatted = number.toFixed(coordPrec);
	return strFormatted;
}
function formatCoordString(numStr)
{
	var strFormatted;
	var strNum;
	strNum = new Number(numStr);
	strFormatted = strNum.toFixed(coordPrec);
	return strFormatted;
}

function formatAngleNumber(number)
{
	var strFormatted;
	strFormatted = number.toFixed(anglePrec);
	return strFormatted;
}

function formatAngleToDMS(angle)
{
	var degrees = Math.floor(angle);

	var dMin = 60. * (angle - degrees);
	var minutes = Math.floor(dMin);

	var dSec = 60. * (dMin - minutes);
	var seconds = formatAngleNumber(dSec);

	return degrees + "-" + minutes + "-" + seconds;
}

function formatBearingDMS(angle)
{
	// decimal degrees are E=0, N=90, W=180, S=270 (counter-clockwise)
	var angNum;

	if(angle >= 0 && angle <= 90)
	{
		angNum = 90. - angle;
		var bearing = formatAngleToDMS(angNum);
		return "N " + bearing + " E";
	}
	else if(angle > 90 && angle <= 180)
	{
		angNum = angle - 90.;
		var bearing = formatAngleToDMS(angNum);
		return "N " + bearing + " W";
	}
	else if(angle > 180 && angle < 270)
	{
		angNum = 270. - angle;
		var bearing = formatAngleToDMS(angNum);
		return "S " + bearing + " W";
	}
	else
	{
		angNum = angle - 270.;
		var bearing = formatAngleToDMS(angNum);
		return "S " + bearing + " E";
	}	
}
]]></msxsl:script>

<xsl:template name="SetGeneralFormatParameters">
		<xsl:if test="$linearPrec">
			<xsl:variable name="vardef" select="landUtils:setLinearPrec($linearPrec)"/>
		</xsl:if>
		
		<xsl:if test="$coordPrec">
			<xsl:variable name="vardef" select="landUtils:setCoordPrec($coordPrec)"/>
		</xsl:if>

		<xsl:if test="$anglePrec">
			<xsl:variable name="vardef" select="landUtils:setAnglePrec($anglePrec)"/>
		</xsl:if>
</xsl:template>

<xsl:template match="lx:Units">
	<xsl:for-each select="./node()">	
		<xsl:choose>
			<xsl:when test="name()='Metric'">
				<xsl:variable name="unitDef" select="landUtils:setUnitType('Metric')"/>
				<xsl:variable name="vardef1" select="landUtils:setLinearUnit(' M')"/>
				<xsl:variable name="vardef2" select="landUtils:setAreaUnit(@areaUnit)"/>
				<xsl:variable name="vardef3" select="landUtils:setVolumeUnit(@volumeUnit)"/>
				<xsl:variable name="vardef4" select="landUtils:setTemperatureUnit(@temperatureUnit)"/>
				<xsl:variable name="vardef5" select="landUtils:setPressureUnit(@pressureUnit)"/>
			</xsl:when>
			<xsl:when test="name()='Imperial'">
				<xsl:variable name="unitDef" select="landUtils:setUnitType('Imperial')"/>
				<xsl:variable name="vardef1" select="landUtils:setLinearUnit(' ft')"/>
				<xsl:variable name="vardef2" select="landUtils:setAreaUnit(@areaUnit)"/>
				<xsl:variable name="vardef3" select="landUtils:setVolumeUnit(@volumeUnit)"/>
				<xsl:variable name="vardef4" select="landUtils:setTemperatureUnit(@temperatureUnit)"/>
				<xsl:variable name="vardef5" select="landUtils:setPressureUnit(@pressureUnit)"/>
			</xsl:when>
		<xsl:otherwise >
			<xsl:variable name="vardef1" select="landUtils:setLinearUnit(@linearUnit)"/>
			<xsl:variable name="vardef2" select="landUtils:setAreaUnit(@areaUnit)"/>
			<xsl:variable name="vardef3" select="landUtils:setVolumeUnit(@volumeUnit)"/>
			<xsl:variable name="vardef4" select="landUtils:setTemperatureUnit(@temperatureUnit)"/>
			<xsl:variable name="vardef5" select="landUtils:setPressureUnit(@pressureUnit)"/>
		</xsl:otherwise>
		</xsl:choose>
	</xsl:for-each>

</xsl:template>


</xsl:stylesheet>