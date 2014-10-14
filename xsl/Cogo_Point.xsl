<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit">
				
<xsl:param name="Point.Elevation.unit">default</xsl:param>
<xsl:param name="Point.Elevation.precision">0.00</xsl:param>
<xsl:param name="Point.Elevation.rounding">normal</xsl:param>
<xsl:param name="Point.Coordinate.unit">default</xsl:param>
<xsl:param name="Point.Coordinate.precision">0.00</xsl:param>
<xsl:param name="Point.Coordinate.rounding">normal</xsl:param>

<xsl:template match="lx:CgPoint" mode="set">
	<xsl:variable name="st" select="landUtils:SetCogoPoint(string(.))"/>
</xsl:template>
<msxsl:script language="JScript" implements-prefix="landUtils"><![CDATA[
var cogoPointNorthing = 0;
var cogoPointEasting = 0;
var cogoPointElevation = 0;

// -----------------------------------------------------------------
// Cogo Point property functions
// -----------------------------------------------------------------
function SetCogoPoint(pointStr)
{
	strPoint = pointStr.split(" ");
	cogoPointNorthing = new Number(parseFloat(strPoint[0]));
	cogoPointEasting = new Number(parseFloat(strPoint[1]));
	cogoPointElevation = 0;
	if(strPoint.length > 2)
	{
		cogoPointElevation = new Number(parseFloat(strPoint[2]));
	}
	
	return "done";
}

function SetCogoPointNorthing(pointStr)
{
	cogoPointNorthing = new Number(parseFloat(pointStr));
	return pointStr;
}
function SetCogoPointEasting(pointStr)
{
	cogoPointEasting = new Number(parseFloat(pointStr));
	return pointStr;
}
function SetCogoPointElevation(pointStr)
{
	cogoPointElevation = new Number(parseFloat(pointStr));
	return pointStr;
}
function GetCogoPointNorthing()
{
	return Number(cogoPointNorthing);
}
function GetCogoPointEasting()
{
	return Number(cogoPointEasting);
}
function GetCogoPointElevation()
{
	return Number(cogoPointElevation);
}

]]></msxsl:script>
</xsl:stylesheet>