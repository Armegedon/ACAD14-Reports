<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet 	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
				xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
				xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit" 
				version="1.0">
<msxsl:script language="JScript" implements-prefix="landUtils"><![CDATA[
//--------------------------------------------------------------
// English to English - Linear
//--------------------------------------------------------------
function FeetToMiles(feet)
{
	var ft = new Number(parseFloat(feet));
	return formatLinearNumber(ft / 5280);
}

function MilesToFeet(miles)
{
	var mls = new Number(parseFloat(miles));
	return formatLinearNumber(5280 * mls);
}
//--------------------------------------------------------------
// English to English - Area
//--------------------------------------------------------------
function SqFtToAcres(sqFt)
{
	var sq = new Number(parseFloat(sqFt));
	return formatLinearNumber(sq / 43560);
}
function AcresToSqFeet(acres)
{
	var acrs = new Number(parseFloat(acres));
	return formatLinearNumber(acrs * 43560);
}

function AcresToSqMiles(acres)
{
	var acrs = new Number(parseFloat(acres));
	return formatLinearNumber(acrs / 640);
}
//--------------------------------------------------------------
// Metric to Metric - Length
//--------------------------------------------------------------
function CmToMeter(d)
{
	return formatLinearNumber(d/100);
}
function CmToKm(d)
{
	return formatLinearNumber(d/100000);
}
function CmToMm(d)
{
	return formatLinearNumber(d*10);
}
function MmToCm(d)
{
	return formatLinearNumber(d/10);
}

//--------------------------------------------------------------
// Metric to Metric - Area
//--------------------------------------------------------------


//--------------------------------------------------------------
// Metric to English - Length
//--------------------------------------------------------------
function CmToFeet(d)
{
	return formatLinearNumber(d / 30.48);
}
function CmToInch(d)
{
	return formatLinearNumber(d/2.54);
}
function CmToYards(d)
{
	return formatLinearNumber(d/91.44);
}
function CmToMiles(d)
{
	return formatLinearNumber(d/160934.4);
}
function KmToMiles(d)
{
	return formatLinearNumber(d / 1.609344);
}
//--------------------------------------------------------------
// English to Metric - Length
//--------------------------------------------------------------

function InchesToMm(d)
{
	return formatLinearNumber(d * 25.4);
}

function InchesToCm(d)
{
	return formatLinearNumber(d * 2.54);
}

function FeetToCm(d)
{
	return formatLinearNumber(d * 30.48);
}

function FeetToMeters(d)
{
	return formatLinearNumber(d * .3048);
}

function YardsToMeters(d)
{
	return formatLineatNumber(d * .9144);
}

function MilesToKm(d)
{
	return formatLinearNumber(d * 1.609344);
}

//--------------------------------------------------------------
// English to Metric - Area
//--------------------------------------------------------------
function SqInchesToSqCm(d)
{
	return formatlinearNumber(d * 6.4516);
}

function SqFeetToSqMeters(d)
{
	return formatLinearNumber(d * 0.09290304);
}

function SqYardsToSqMeters(d)
{
	return formatLinearNumber( d * 0.83612736);
}

function SqMilesToSqKm(d)
{
	return formatLinearNumber( d * 2.589988110336);
}

function AcresToHectares(d)
{
	return formatLinearNumber(d * 0.40468564224);
}

//--------------------------------------------------------------
// Metric to English - Area
//--------------------------------------------------------------
function SqCmToSqInches(d)
{
	return formatLinearNumber(d / 6.4516);
}
function SqMetersToSqFeet(d)
{
	return formatLinearNumber(d / 0.09290304);
}
function SqMetersToSqYards(d)
{
	return formatLinearNumber(d / 0.83612736);
}
function SqKmToSqMiles(d)
{
	return formatLinearNumber(d / 2.589988110336);
}
function HectaresToAcres(d)
{
	return formatLinearNumber(d / 0.40468564224);
}
]]></msxsl:script>
</xsl:stylesheet>
