<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2004 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">
	
<xsl:include href="Number_Formatting.xsl" />
<xsl:include href="LandXMLUtils_JScript.xsl" />				

<!-- =========== Pipe  Parameter definitions ============ -->
<xsl:param name="Pipe_Reports.Pipe_Length.unit">default</xsl:param>
<xsl:param name="Pipe_Reports.Pipe_Length.precision">0.00</xsl:param>
<xsl:param name="Pipe_Reports.Pipe_Length.rounding">normal</xsl:param>
<xsl:param name="Pipe_Reports.Pipe_Length.type">2-D</xsl:param>

<xsl:param name="Pipe_Reports.Coordinate.unit">default</xsl:param>
<xsl:param name="Pipe_Reports.Coordinate.precision">0.00</xsl:param>
<xsl:param name="Pipe_Reports.Coordinate.rounding">normal</xsl:param>

<xsl:param name="Pipe_Reports.Elevation.unit">foot</xsl:param>
<xsl:param name="Pipe_Reports.Elevation.precision">0.00</xsl:param>
<xsl:param name="Pipe_Reports.Elevation.rounding">normal</xsl:param>

<xsl:param name="Pipe_Reports.Linear.unit">default</xsl:param>
<xsl:param name="Pipe_Reports.Linear.precision">0.00</xsl:param>
<xsl:param name="Pipe_Reports.Linear.rounding">normal</xsl:param>

<xsl:param name="Pipe_Reports.Station.Display">##+##</xsl:param>
<xsl:param name="Pipe_Reports.Station.precision">0.00</xsl:param>
<xsl:param name="Pipe_Reports.Station.rounding">normal</xsl:param>

<xsl:param name="Pipe_Reports.Angular.unit">default</xsl:param>
<xsl:param name="Pipe_Reports.Angular.precision">0.00</xsl:param>
<xsl:param name="Pipe_Reports.Angular.rounding">normal</xsl:param>



<msxsl:script language="JScript" implements-prefix="landUtils"> 

<![CDATA[
var delta = 0;
var counter = 0;

function ResetInternalCounterAndGetCurNodeCounter(node)
{
	counter = node.length;
	return counter;
	
}
function DecreCounter()
{
	counter = Number(counter) - 1;
	return counter;
}
function getPipeLengthUnit(SourceLinearUnit, SettingLinearUnit)
{
    var unit = "";
	var cnvSourceLinearUnit = SourceLinearUnit.toLowerCase();
	var cnvSettingLinearUnit = SettingLinearUnit.toLowerCase();
	
	if(cnvSettingLinearUnit == "default")
	{
		if(cnvSourceLinearUnit == "foot")
		{
			unit = "foot";
		}
		else if(cnvSourceLinearUnit == "meter")
		{
			unit = "meter";
		}
		else if(cnvSourceLinearUnit == "ussurveyfoot")
		{
			unit = "ussurveyfoot";
		}
	}
	else if(cnvSettingLinearUnit == "foot" || cnvSettingLinearUnit == "meter" || cnvSettingLinearUnit == "ussurveyfoot")
	{
		unit = SettingLinearUnit;
	}
	
	return unit; 
							
}
function CC2EE(CC)
{
	var EE = Number(CC) - Number(delta); 
	return EE;
}
function setDelta(startD, endD)
{
	delta = Number(startD)/2 + Number(endD)/2;
	return delta;
}
function getSumpDepth(node)
{
    
	var elevSump = node.item(0).getAttribute('elevSump');
	
	var objNodeList;
	objNodeList = node.item(0).childNodes;
	var minElev = Number(node.item(0).getAttribute('elevRim'));
	var i =0;
	for(i=0; i<objNodeList.length; i++)
	{
		if(objNodeList.item(i).nodeName =="Invert")
		{
			if(Number(objNodeList.item(i).getAttribute('elev')) < minElev)
			{
				minElev = Number(objNodeList.item(i).getAttribute('elev'));
			}
		}
	}	
	
	var SumpDepth = Math.abs(Number(minElev) - Number(elevSump));	
	
	return SumpDepth;
}
function getPipeType(node)
{
  var i = 0;
  var objNodeList;
  objNodeList = node.item(0).childNodes;
  var type;
    
  for(i = 0; i < objNodeList.length; i++)
  {
    var nodename = objNodeList.item(i).nodeName;
        
    if(nodename == "CircPipe")
    {
      type = "Circular";
      break;
    }
    else if (nodename == "ElliPipe")
    {
      type = "Elliptical"
      break;
    }
    else if (nodename == "RectPipe")
    {
      type = "Rectangular"
      break;
    }	
    else if (nodename == "Channel")
    {
      type = "Channel"
      break;
    }
    else if (nodename == "EggPipe")
    {
      type = "Egg"
      break;
    }
  }
    
  return type;

}
function getStructType(node)
{
    var i = 0;
    var objNodeList;
    objNodeList = node.item(0).childNodes;
    var type;
    
    for(i = 0; i < objNodeList.length; i++)
    {
        var nodename = objNodeList.item(i).nodeName;
        
		if(nodename == "CircStruct")
		{
			type = "Circular";
			break;
		}
		else if (nodename == "RectStruct")
		{
			type = "Rectangular";
			break;
		}
		else if (nodename == "InletStruct")
		{
			type = "Inlet";
			break;
		}	
		else if (nodename == "OutletStruct")
		{
			type = "Outlet";
			break;
		}
		else if(nodename == "Connection")
		{
			type = "Connection";
			break;
		}		
    }
    
    return type;
}
function getSlope(start, elev1,end, elev2)
{
    
    var xDiff = Number(getEasting(start))-Number(getEasting(end));
    var yDiff = Number(getNorthing(start))-Number(getNorthing(end));
    var zDiff = Number(elev1)-Number(elev2);
    
    var xyDiff = Math.sqrt(Math.pow(xDiff,2)+Math.pow(yDiff, 2));
    
    var slope = 100*(zDiff/xyDiff);
    
    return slope;	
	
}
function get3DLength(start, elev1,end, elev2)
{
    
    var deltaEasting = Number(getEasting(start))-Number(getEasting(end));
    var deltaNorthing = Number(getNorthing(start))-Number(getNorthing(end));
    var deltaElevation = Number(elev1)-Number(elev2);
    var length = Math.sqrt(Math.pow(deltaEasting,2)+Math.pow(deltaNorthing, 2)+Math.pow(deltaElevation, 2));    
    
    return length;	
	
}
function get2DLength(start, end)
{
    
    var deltaEasting = Number(getEasting(start))-Number(getEasting(end));
    var deltaNorthing = Number(getNorthing(start))-Number(getNorthing(end));
    
    var length = Math.sqrt(Math.pow(deltaEasting,2)+Math.pow(deltaNorthing, 2));    
    
    return length;	
	
}
function getPipesName(node)
{
   var pipename = new String();
  
   pipename = node.item(0).getAttribute('refPipe');
   return pipename;
}

function getEasting(textblock)
{
	var arr = textblock.split(" ");
	return arr[1];
}


function getNorthing(textblock)
{
	var arr = textblock.split(" ");
	return arr[0];
}

function formatUnit(inUnit)
{
	var unit = inUnit;
	var foot = /foot/i;
	
	if(foot.exec(inUnit))
	{
		unit = "foot";
	}
		
	return unit;
}
]]>
</msxsl:script>

</xsl:stylesheet>