<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet version="1.0"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">

	<xsl:param name="Legal_Descriptions.Phrasing_File.File_location">GeneralLegalPhrasings.xml</xsl:param>
	<xsl:param name="Legal_Descriptions.Phrasing_File.Metes_and_Bounds">separate line</xsl:param>
	<xsl:param name="Legal_Descriptions.Report_Units.units">foot</xsl:param>
	<xsl:param name="Legal_Descriptions.Coordinate.precision">0.000</xsl:param>
	<xsl:param name="Legal_Descriptions.Coordinate.rounding">normal</xsl:param>
	<xsl:param name="Legal_Descriptions.Distance.precision">0.0000</xsl:param>
	<xsl:param name="Legal_Descriptions.Distance.rounding">normal</xsl:param>
	<xsl:param name="Legal_Descriptions.Direction.precision">0</xsl:param>
	<xsl:param name="Legal_Descriptions.Direction.rounding">normal</xsl:param>
	<xsl:param name="Legal_Descriptions.Direction.type">N deg min sec E</xsl:param>
	<xsl:param name="Legal_Descriptions.Direction.format">Use °, ’, ” </xsl:param>
	<xsl:param name="Legal_Descriptions.Angle.precision">0.00000</xsl:param>
	<xsl:param name="Legal_Descriptions.Angle.rounding">normal</xsl:param>
	<xsl:param name="Legal_Descriptions.Angle.format">Use °, ’, ” </xsl:param>
	<xsl:param name="Legal_Descriptions.Station.Display">#+###</xsl:param>
	<xsl:param name="Legal_Descriptions.Station.precision">0.00</xsl:param>
	<xsl:param name="Legal_Descriptions.Station.rounding">normal</xsl:param>
	<xsl:param name="Legal_Descriptions.Area.unit">squareFoot</xsl:param>
	<xsl:param name="Legal_Descriptions.Area.precision">0.00</xsl:param>
	<xsl:param name="Legal_Descriptions.Area.rounding">normal</xsl:param>
	<xsl:param name="Legal_Descriptions.Header_Information.state">Arizona</xsl:param>
	<xsl:param name="Legal_Descriptions.Header_Information.Surveyor"></xsl:param>
	<xsl:param name="Legal_Descriptions.Header_Information.Survey_date">Feb 15, 2004</xsl:param>
	<xsl:param name="Legal_Descriptions.Header_Information.Heading_Field1"></xsl:param>
	<xsl:param name="Legal_Descriptions.Header_Information.Heading_Field2"></xsl:param>
	<xsl:param name="Legal_Descriptions.Header_Information.Heading_Field3"></xsl:param>
	<xsl:param name="Legal_Descriptions.Header_Information.Heading_Field4"></xsl:param>
	<xsl:param name="Legal_Descriptions.Header_Information.Heading_Field5"></xsl:param>
	<xsl:param name="Legal_Descriptions.Section_Information.Township">1</xsl:param>
	<xsl:param name="Legal_Descriptions.Section_Information.Baseline_Direction">N</xsl:param>
	<xsl:param name="Legal_Descriptions.Section_Information.Range">3</xsl:param>
	<xsl:param name="Legal_Descriptions.Section_Information.Principal_Meridian"></xsl:param>
	<xsl:param name="Legal_Descriptions.Section_Information.Meridian_Direction">E</xsl:param>
	<xsl:param name="Legal_Descriptions.Section_Information.Section">1</xsl:param>
	<xsl:param name="Legal_Descriptions.Section_Information.Section_Part"></xsl:param>
	<xsl:param name="Legal_Descriptions.Section_Information.Baseline"></xsl:param>
	<xsl:param name="Legal_Descriptions.Deed_Page.Survey"></xsl:param>
	<xsl:param name="Legal_Descriptions.Deed_Page.Abstract"></xsl:param>
	<xsl:param name="Legal_Descriptions.Deed_Page.Owner"></xsl:param>
	<xsl:param name="Legal_Descriptions.Deed_Page.Date"></xsl:param>
	<xsl:param name="Legal_Descriptions.Deed_Page.Volume">1+</xsl:param>
	<xsl:param name="Legal_Descriptions.Deed_Page.Page">1+</xsl:param>
	<xsl:param name="Legal_Descriptions.Texas_DOT.County">Anderson</xsl:param>
	<xsl:param name="Legal_Descriptions.Texas_DOT.Account_Number"></xsl:param>
	<xsl:param name="Legal_Descriptions.Texas_DOT.CSJ_Number"></xsl:param>
	<xsl:param name="Legal_Descriptions.Texas_DOT.Hwy_Number"></xsl:param>
	<xsl:param name="Legal_Descriptions.Texas_DOT.Grantor"></xsl:param>
	<xsl:param name="Legal_Descriptions.Florida_DOT.FIN_Number"></xsl:param>
	<xsl:param name="Legal_Descriptions.Florida_DOT.State_Road"></xsl:param>
	<xsl:param name="Legal_Descriptions.Florida_DOT.County">Alachua</xsl:param>


<xsl:template name="SetupLegal" >
<xsl:param name="sourceUnit"></xsl:param>

</xsl:template>

<msxsl:script language="JScript" implements-prefix="landUtils"> 
<![CDATA[
// -----------------------------------------------------------------
// XML document of Legal Description settings (the phrasing file)
var xmlDoc;
var xmlStyle;
var xmlGeom;
var geomTypeStr ;

var sourceUnit;
var convertUnit;

var coordPrec;
var coordRound;

var staPattern;
var staPrec;
var staRound;

var distPrec;
var distRound;

var dirPrec;
var dirRound;
var dirType;
var dirFormat;

var angPrec;
var angRound;
var angFormat;

var areaUnit;
var areaPrec;
var areaRound;

var ControlPointArray;


// -----------------------------------------------------------------

function SetLegalPhrasingFile(filename)
{
	xmlDoc = new ActiveXObject("Msxml2.DOMDocument.6.0");
	xmlDoc.async = false;
	var bLoaded = xmlDoc.load(filename);
	if (!bLoaded ) 
	{
   		return "Error loading phrasing file";
	}
	return "";
}

function SetDescriptionType(typename)
{
	geomTypeStr = typename;
	var geomLocStr = "./Geometry[@type='" + typename + "']";
	xmlGeom = xmlStyle.selectSingleNode(geomLocStr);
	return geomLocStr;
}


// Unit settings functions
//////////////////////////////////////////////////////

function SetUnits(unitname)
{      
	var styleLocStr;
	
	convertUnit = unitname;
	var convStr = unitname.toLowerCase();

	//when report unit is default, set to sourceUnit.
	//So SetSourceUnits(units) must be called before SetUnits(unitname)
	if(convStr =="default")
	{
		convertUnit = sourceUnit;
		convStr = sourceUnit.toLowerCase();
	}

	if(convStr =="foot")
	{
		styleLocStr = "//LegalDescription/Style[@name='Imperial']";
	}
	else if(convStr =="ussurveyfoot")
	{
		styleLocStr = "//LegalDescription/Style[@name='Imperial']";
	}
	else
	{
		styleLocStr = "//LegalDescription/Style[@name='Metric']";
	}
	xmlStyle = xmlDoc.documentElement.selectSingleNode(styleLocStr);
	return unitname ;
}

function SetSourceUnits(units)
{
	sourceUnit = units;
	return units;
}

function SetCoordPrec(prec)
{
	coordPrec = prec;
	return prec;
}

function SetCoordRound(round)
{
	coordRound = round;
	return round;
}

function SetDirectionPrecision(prec)
{
	dirPrec = prec;
	return prec;
}

function SetDirectionRound(round)
{
	dirRound = round;
	return round;
}

function SetDirectionType(type)
{
	dirType = type;
	if(type == "N deg min sec E")
	{
		SetNorthLabel("N");
		SetSouthLabel("S");
		SetEastLabel("E");
		SetWestLabel("W");
		dirType = "Bearing";
	}
	else if(type == "North deg min sec East")
	{
		SetNorthLabel("North");
		SetSouthLabel("South");
		SetEastLabel("East");
		SetWestLabel("West");
		dirType = "Bearing";
	}
	return type;
}

function SetDirectionFormat(format)
{
	dirFormat = format;
	if(format == "Use °, ’, \”")
	{
		SetDirDegreeLabel("°");
		SetDirMinuteLabel("'");
		SetDirSecondLabel("\"");
	}
	else
	{
		SetDirDegreeLabel("degrees");
		SetDirMinuteLabel("minutes");
		SetDirSecondLabel("seconds");
	}
	return format;
}

function SetDistancePrec(prec)
{
	distPrec = prec;
	return prec;
}

function SetDistanceRound(round)
{
	distRound = round;
	return round;
}

function SetAnglePrec(prec)
{
	angPrec = prec;
	return prec;
}

function SetAngleRound(round)
{
	angRound = round;
	return round;
}

function SetAngleFormat(format)
{
	angFormat = format;
	if(format == "Use °, ’, \”")
	{
		SetAngleDegreeLabel("°");
		SetAngleMinuteLabel("'");
		SetAngleSecondLabel("\"");
	}
	else
	{
		SetAngleDegreeLabel("degrees");
		SetAngleMinuteLabel("minutes");
		SetAngleSecondLabel("seconds");
	}
	return format;
}

function SetStationPattern(pattern)
{
	staPattern = pattern;
	return pattern;
}

function SetStationPrecision(prec)
{
	staPrec = prec;
	return prec;
}

function SetStationRounding(round)
{
	staRound = round;
	return round;
}

function SetAreaUnit(unit)
{
	areaUnit = unit;
	return unit;
}

function SetAreaPrec(prec)
{
	areaPrec = prec;
	return prec;
}

function SetAreaRound(round)
{
	areaRound = round;
	return round;
}
////////////////////////////////////////////////////////

function GetLegalFor(index)
{
	var legalDesc;
	var geoele = GeomArray[index];
	var nextele;
	var mode;
	
	legalDesc = ""

	if(index == 1)
	{
		legalDesc += ProcessLegalForStart(index);
    if(index == GeomArray.length - 1) // only one geometry element, both beginning and end
      legalDesc += ProcessLegalForEnd(index);
    else
		  legalDesc += ProcessLegal(index);
	}
	else if(index == GeomArray.length - 1)
	{
		legalDesc += ProcessLegalForEnd(index)
	}
	else
	{
		legalDesc += ProcessLegal(index);
	}
	
	return legalDesc;
}

function ProcessLegalForStart(index)
{
	var geoele = GeomArray[index];
	var retStr = "Uknown ";
	
	if(geoele.type == "Line")
	{
		retStr = GetLegalForStartLine(index);
	}
	else if(geoele.type == "Curve")
	{
		retStr = GetLegalForStartCurve(index);
	}
	else if(geoele.type == "Spiral")
	{
		retStr = GetLegalForStartSpiral(index)
	}
	else
	{
		retStr = GetLegalForStartPoint(index);
	}
	
	return retStr;		
}

function ProcessLegalForEnd(index)
{
	var geoele = GeomArray[index];
	var retStr = "Uknown ";
	
	if(geoele.type == "Line")
	{
		retStr = GetLegalForLineEnd(index);
	}
	else if(geoele.type == "Curve")
	{
		retStr =GetLegalForCurveEnd(index);
	}
	else if(geoele.type == "Spiral")
	{
		retStr = GetLegalForSpiralEnd(index);
	}
	return retStr;		
}

function ProcessLegal(index)
{
	var geoele = GeomArray[index];
	var retStr = "Uknown ";
	
	if(geoele.type == "Line")
	{
		retStr = GetLegalForLine(index);
	}
	else if(geoele.type == "Curve")
	{
		retStr = GetLegalForCurve(index);
	}
	else if(geoele.type == "Spiral")
	{
		retStr = GetLegalForSpiral(index);
	}
	else
	{
		retStr = GetLegalForPoint(index);
	}		
	
	return retStr;		
}

function GetLegalForStartPoint(index)
{
	var retStr = "StartPoint ";
	
	return retStr;
}

function GetLegalForStartLine(index)
{
	var retStr = "StartingLine ";
	var xmlEle = xmlGeom.selectSingleNode("Bounds/Bound[@type='POB']");
	if(xmlEle)
	{
		//var subStr = xmlEle.text;
		var subStr = GetInnerXML(xmlEle);
				
		retStr = SubstituteStartingElement(index, subStr) ;
	}
	return retStr;
}

function GetLegalForStartCurve(index)
{
	var retStr = "StartingCurve ";
	var xmlEle = xmlGeom.selectSingleNode("Bounds/Bound[@type='POB']");
	if(xmlEle)
	{
		//var subStr = xmlEle.text;
		var subStr = GetInnerXML(xmlEle);
				
		retStr = SubstituteStartingElement(index, subStr) ;
	}
	return retStr;
}

function GetLegalForStartSpiral(index)
{
	var retStr = "StartingSpiral ";
	var xmlEle = xmlGeom.selectSingleNode("Bounds/Bound[@type='POB']");
	if(xmlEle)
	{
		//var subStr = xmlEle.text;
		var subStr = GetInnerXML(xmlEle);
				
		retStr = SubstituteStartingElement(index, subStr) ;
	}
	return retStr;
}

function GetLegalForLine(index)
{
	var i = parseInt(index);
	var geoele = GeomArray[index];
	var nextEle = GeomArray[i + 1];
	var retStr = geoele.type + index + " ";
	
	if(nextEle)
	{
		if(nextEle.type == "Line")
		{
			var xmlEle = xmlGeom.selectSingleNode("Metes/Mete[@type='Line' and @connect='Line']");
			if(xmlEle)
			{
				//var subStr = xmlEle.text;
				var subStr = GetInnerXML(xmlEle);
				
				retStr = SubstituteLinearKeywords(index, subStr);
			}
			else
			{
				retStr = "LineToLine ";
			}
		}
		else if(nextEle.type == "Curve")
		{
			var xmlEle;
			var bTangent = IsNextTangent(index);
			
			if(bTangent == true)
			{
				xmlEle = xmlGeom.selectSingleNode("Metes/Mete[@type='Line' and @connect='TangentCurve']");
			}
			else
			{
				xmlEle = xmlGeom.selectSingleNode("Metes/Mete[@type='Line' and @connect='NonTangentCurve']");
			}
			
			if(xmlEle)
			{
				//var subStr = xmlEle.text;
				var subStr = GetInnerXML(xmlEle);
				
				retStr = SubstituteLinearKeywords(index, subStr);
			}
			else
			{
				retStr = "LineToCurve ";
			}
		}
		else
		{
			retStr = "LineToOther ";
		}
	}
	
	return retStr;
}

function GetLegalForCurve(index)
{      
	var i = parseInt(index);
	var geoele = GeomArray[i];
	var nextEle = GeomArray[i + 1];
	var retStr = geoele.type + index + " ";
	
	if(nextEle)
	{
		var bTangent = IsNextTangent(i);
		if(nextEle.type == "Line")
		{
			var xmlEle;
			
			if(bTangent == true)
			{
				xmlEle = xmlGeom.selectSingleNode("Metes/Mete[@type='Curve' and @connect='TangentLine']");
			}
			else
			{
				xmlEle = xmlGeom.selectSingleNode("Metes/Mete[@type='Curve' and @connect='NonTangentLine']");
			}
			//var subStr = xmlEle.text;
			var subStr = GetInnerXML(xmlEle);
				
			retStr = SubstituteCurveKeywords(index, subStr);
		}
		else if(nextEle.type == "Curve")
		{
			var xmlEle;
			
			if(bTangent == true)
			{
				var check = CompareNextCurveDirection(i);
				if(check == "Compound")
				{
					xmlEle = xmlGeom.selectSingleNode("Metes/Mete[@type='Curve' and @connect='TangentCompoundCurve']");
				}
				else if(check == "Reverse")
				{
					xmlEle = xmlGeom.selectSingleNode("Metes/Mete[@type='Curve' and @connect='TangentReverseCurve']");
				}
				//var subStr = xmlEle.text;
				var subStr = GetInnerXML(xmlEle);
				
				retStr = SubstituteCurveKeywords(index, subStr);
			}
			else
			{
				var check = CompareNextCurveDirection(i);
				if(check == "Compound")
				{
					xmlEle = xmlGeom.selectSingleNode("Metes/Mete[@type='Curve' and @connect='NonTangentCompoundCurve']");
				}
				else if(check == "Reverse")
				{
					xmlEle = xmlGeom.selectSingleNode("Metes/Mete[@type='Curve' and @connect='NonTangentReverseCurve']");
				}
				//var subStr = xmlEle.text;
				var subStr = GetInnerXML(xmlEle);
				
				retStr = SubstituteCurveKeywords(index, subStr);
			}
		}
		else
		{
			retStr = "CurveToOther ";
		}
	}
	
	return retStr;
}

function GetLegalForSpiral(index)
{
	var retStr = "Spiral ";
	
	return retStr;
}

function IsNextTangent(index)
{
	var i = parseInt(index);
	var geoele = GeomArray[i];
	
	var bIsTangent = false;
	
	if(i < GeomArray.length)
	{
		var nextEle = GeomArray[i + 1];
		if(nextEle)
		{
			if(geoele.type == "Line")
			{
				var lineDir = GetLineDirection(i);
				if(nextEle.type == "Curve")
				{
					var curDir = GetCurveStartDirection(i + 1);
					var adjDir = curDir + 90;
					var err = Math.abs(adjDir - lineDir);
					if(err > 180.0) err -= 180.0;
					if(err < .01 ||
						Math.abs(err - 180.0) < .01)
					{
						bIsTangent = true;
					}
				}
				else if(nextEle.type == "Spiral")
				{
					bIsTangent = true;
				}
			}
			else if(geoele.type == "Curve")
			{
				var curDir = GetCurveEndDirection(i);
				var adjDir = curDir + 90;
				if(nextEle.type == "Line")
				{
					var linDir = GetLineDirection(i + 1);
					var err = Math.abs(adjDir - linDir);
					if(err > 180.0) err -= 180.0;
					if(err < .01 ||
						Math.abs(err - 180.0) < .01)
					{
						bIsTangent = true;
					}
				}
				else if(nextEle.type == "Curve")
				{
					var cCurDir = GetCurveDirection(i + 1);
					if(adjDir == cCurDir + 90)
					{
						bIsTangent = false;
					}
				}
				else if(nextEle.type == "Spiral")
				{
					bIsTangent = false;
				}
			}
			else if(geoele.type == "Spiral")
			{
				bIsTangent = true;
			}
		}
	}
	
	return bIsTangent;
}

function GetLegalForLineEnd(index)
{
	var retStr = "LineEnd ";
	var xmlEleEnd;
	if(geomTypeStr == "Parcel")
	{
		xmlEleEnd = xmlGeom.selectSingleNode("Metes/Mete[@type='Line' and @connect='PointOfBegining']");
	}
	else if(geomTypeStr == "Alignment")
	{
		xmlEleEnd = xmlGeom.selectSingleNode("Metes/Mete[@type='Line' and @connect='PointOfEnding']");
	}
	else if(geomTypeStr == "LocationTraverse")
	{
		xmlEleEnd = xmlGeom.selectSingleNode("Metes/Mete[@type='Line' and @connect='PointOfEnding']");
	}
	if(xmlEleEnd)
	{
		var subStr = GetInnerXML(xmlEleEnd);
		
		retStr = SubstituteLinearKeywords(index, subStr);
	}
	
	return retStr;
}

function GetLegalForCurveEnd(index)
{
	var retStr = "CurveEnd ";
	var xmlEleEnd;
	if(geomTypeStr == "Parcel")
	{
		xmlEleEnd = xmlGeom.selectSingleNode("Metes/Mete[@type='Curve' and @connect='PointOfBegining']");
	}
	else if(geomTypeStr == "Alignment")
	{
		xmlEleEnd = xmlGeom.selectSingleNode("Metes/Mete[@type='Curve' and @connect='PointOfEnding']");
	}
	else if(geomTypeStr == "LocationTraverse")
	{
		xmlEleEnd = xmlGeom.selectSingleNode("Metes/Mete[@type='Curve' and @connect='PointOfEnding']");
	}
	if(xmlEleEnd)
	{
		var subStr = GetInnerXML(xmlEleEnd);
		
		retStr = SubstituteCurveKeywords(index, subStr);
	}
	
	return retStr;
}

function GetLegalForSpiralEnd(index)
{
	var retStr = "SpiralEnd ";
	var xmlEleEnd;
	if(geomTypeStr == "Parcel")
	{
		xmlEleEnd = xmlGeom.selectSingleNode("Metes/Mete[@type='Spiral' and @connect='PointOfBegining']");
	}
	else if(geomTypeStr == "Alignment")
	{
		xmlEleEnd = xmlGeom.selectSingleNode("Metes/Mete[@type='Spiral' and @connect='PointOfEnding']");
	}
	else if(geomTypeStr == "LocationTraverse")
	{
		xmlEleEnd = xmlGeom.selectSingleNode("Metes/Mete[@type='Spiral' and @connect='PointOfEnding']");
	}
	if(xmlEleEnd)
	{
		var subStr = GetInnerXML(xmlEleEnd);
		
		retStr = SubstituteLinearKeywords(index, subStr);
	}
	
	return retStr;
}

function SubstituteStartingElement(index, text)
{
	var retStr = "";
	var northExp = new RegExp("{startPointNorthing}", "g");
	var eastExp = new RegExp("{startPointEasting}", "g");
	
	var dNorth = GetGeomStartNorthing(index);
	var dEast = GetGeomStartEasting(index);
	
	var northStr = FormatNumber("" + dNorth, sourceUnit, convertUnit, coordPrec, coordRound)
	var eastStr = FormatNumber("" + dEast, sourceUnit, convertUnit, coordPrec, coordRound)

	var repStrOne = text.replace(northExp, northStr );
	var repStrTwo =  repStrOne.replace(eastExp , eastStr);

	return repStrTwo;
}

function SubstituteLinearKeywords(index, text)
{
	var retStr = "";
	var distMeterExp = new RegExp("{lineDistanceMeter}", "g");
	var distFootExp = new RegExp("{lineDistanceFoot}", "g");
	var distUSFootExp = new RegExp("{lineDistanceUSFoot}", "g");
	
	var dirExp = new RegExp("{lineDirection}", "g");
	
	var startNExp = new RegExp("{lineStartNorthing}", "g");
	var startEExp = new RegExp("{lineStartEasting}", "g");

	var endNExp = new RegExp("{lineEndNorthing}", "g");
	var endEExp = new RegExp("{lineEndEasting}", "g");

	var startStaExp = new RegExp("{lineStartStation}", "g");
	var endStaExp = new RegExp("{lineEndStation}", "g");
	
	var linLength = GetLineLength(index);
	
	var linDir = GetLineDirection(index);
	
	var lineStartN = GetLineStartNorthing(index);
	var lineStartE = GetLineStartEasting(index);

	var lineEndN = GetLineEndNorthing(index);
	var lineEndE = GetLineEndEasting(index);
	
	var lineStartSta = GetGeomStartingStation(index);
	var lineEndSta = GetGeomEndStation(index);
	
	var startNStr = FormatNumber("" + lineStartN , sourceUnit, convertUnit, coordPrec, coordRound);
	var startEStr = FormatNumber("" + lineStartE , sourceUnit, convertUnit, coordPrec, coordRound);

	var endNStr = FormatNumber("" + lineEndN , sourceUnit, convertUnit, coordPrec, coordRound);
	var endEStr = FormatNumber("" + lineEndE , sourceUnit, convertUnit, coordPrec, coordRound);
	
	var dirStr =  FormatDirection("" + linDir, "Conventional", dirType, dirPrec, dirRound);
	
	var distMStr = FormatNumber("" + linLength, sourceUnit, "meter", distPrec, distRound);
	var distFStr = FormatNumber("" + linLength, sourceUnit, "foot", distPrec, distRound);
	var distUSFStr = FormatNumber("" + linLength, sourceUnit, "USSurveyfoot", distPrec, distRound);
	
	var startStaCnv = ConvertNumber(lineStartSta, sourceUnit, convertUnit);
	var endStaCnv = ConvertNumber(lineEndSta, sourceUnit, convertUnit);
	var startStaStr = FormatStation(startStaCnv, staPattern, staPrec, staRound);
	var endStaStr = FormatStation(endStaCnv, staPattern, staPrec, staRound);
	
	var repStrOne = text.replace(distMeterExp, distMStr );
	var repStrD2 =repStrOne.replace(distFootExp, distFStr );
	var repStrD3 = repStrD2.replace(distUSFootExp, distUSFStr );
	
	var repStrTwo =  repStrD3.replace(dirExp , dirStr);
	
	var repStrThree = repStrTwo.replace(startNExp, startNStr);
	var repStrFour = repStrThree.replace(startEExp, startEStr);
	
	var repStrFive = repStrFour.replace(endNExp, endNStr);
	var repStrSix = repStrFive.replace(endEExp, endEStr);

	var repStrSeven = repStrSix.replace(startStaExp, startStaStr);
	var repStrEight = repStrSeven.replace(endStaExp, endStaStr);
	
	return repStrEight;
}

function SubstituteCurveKeywords(index, text)
{
	var retStr = "";
	
	var lengthFExp = new RegExp("{curveLengthFoot}", "g");
	var lengthMExp = new RegExp("{curveLengthMeter}", "g");
	var lengthUSFExp = new RegExp("{curveLengthUSFoot}", "g");
	
	var radiusFExp = new RegExp("{curveRadiusFoot}", "g");
	var radiusMExp = new RegExp("{curveRadiusMeter}", "g");
	var radiusUSFExp = new RegExp("{curveRadiusUSFoot}", "g");
	
	var chordLengthFExp = new RegExp("{curveChordLengthFoot}", "g");
	var chordLengthMExp = new RegExp("{curveChordLengthMeter}", "g");
	var chordLengthUSFExp = new RegExp("{curveChordLengthUSFoot}", "g");

	var tangtExp = new RegExp("{curveTangent}", "g");
	var angleExp = new RegExp("{curveAngle}", "g");
	var chordDirExp = new RegExp("{curveChordDirection}", "g");
	var startDirExp = new RegExp("{curveStartDirection}", "g");
	var endDirExp = new RegExp("{curveEndDirection}", "g");
	var rotationExp = new RegExp("{curveRotation}", "g");
	var externalExp = new RegExp("{curveExternal}", "g");
	var middleExp = new RegExp("{curveMiddle}", "g");
	var docExp = new RegExp("{curveDOC}", "g");

	var startNExp = new RegExp("{curveStartNorthing}", "g");
	var startEExp = new RegExp("{curveStartEasting}", "g");

	var radNExp = new RegExp("{curveRadiusNorthing}", "g");
	var radEExp = new RegExp("{curveRadiusEasting}", "g");

	var endNExp = new RegExp("{curveEndNorthing}", "g");
	var endEExp = new RegExp("{curveEndEasting}", "g");
	
	var concDirExp = new RegExp("{concavityBearing}", "g");
	var runDirExp = new RegExp("{curveRunBearing}", "g");
	
	var backDirExp = new RegExp("{curveBackDirection}", "g");
	var aheadDirExp = new RegExp("{curveAheadDirection}", "g");
	
	var startN = GetCurveStartNorthing(index);
	var startE = GetCurveStartEasting(index);
	
	var radN = GetCurveCenterNorthing(index);
	var radE = GetCurveCenterEasting(index);

	var endN = GetCurveEndNorthing(index);
	var endE = GetCurveEndEasting(index);
	
	var startStaExp = new RegExp("{curveStartStation}", "g");
	var endStaExp = new RegExp("{curveEndStation}", "g");

	var length = GetCurveLength(index);
	var tangent = GetCurveTangent(index);
	var radius = GetCurveRadius(index);
	var chordLen = GetChordLength(index);
	var external = GetCurveExternal(index);
	var middle = GetCurveMiddle(index);
	var doc = GetCurveDOC(index);
	
	var angle = GetCurveAngle(index);
	var chordDir = GetChordDirection(index);
	var startDir = GetCurveStartDirection(index);
	var endDir = GetCurveEndDirection(index);
	
	var aheadDir = GetCurveAheadDirection(index);
	var backDir = GetCurveBackDirection(index);
	
	var rotation = "left";
	var rot = GetCurveRotation(index);
	if(rot == "cw")
	{
		rotation = "right";
	}
	
	var concaveDirStr = GetDirectionOfConcavity(index);
	var runDirStr = GetDirectionOfCurveRun(index);
	
	var curveStartSta = GetGeomStartingStation(index);
	var curveEndSta = GetGeomEndStation(index);

	var lengthFStr = FormatNumber("" + length, sourceUnit, "foot", distPrec, distRound);
	var lengthMStr = FormatNumber("" + length, sourceUnit, "meter", distPrec, distRound);
	var lengthUSFStr = FormatNumber("" + length, sourceUnit, "ussurveyfoot", distPrec, distRound);

	var radiusFStr = FormatNumber("" + radius, sourceUnit, "foot", distPrec, distRound);
	var radiusMStr = FormatNumber("" + radius, sourceUnit, "meter", distPrec, distRound);
	var radiusUSFStr = FormatNumber("" + radius, sourceUnit, "ussurveyfoot", distPrec, distRound);
	
	var chordLenFStr = FormatNumber("" + chordLen, sourceUnit, "foot", distPrec, distRound);
	var chordLenMStr = FormatNumber("" + chordLen, sourceUnit, "meter", distPrec, distRound);
	var chordLenUSFStr = FormatNumber("" + chordLen, sourceUnit, "ussurveyfoot", distPrec, distRound);
	
	var tangentStr = FormatNumber("" + tangent, sourceUnit, convertUnit, distPrec, distRound);
	var externalStr = FormatNumber("" + external, sourceUnit, convertUnit, distPrec, distRound);
	var middleStr = FormatNumber("" + middle, sourceUnit, convertUnit, distPrec, distRound);
	var docStr = FormatNumber("" + doc, sourceUnit, convertUnit, distPrec, distRound);

	var angleStr = FormatAngle("" + angle, "Degrees", "Degrees", "DMS", angPrec, angRound)
	var chordDirStr = FormatDirection("" + chordDir, "Conventional", dirType, dirPrec, dirRound)
	
	var startDirStr = FormatDirection("" + startDir,"Conventional", dirType, dirPrec, dirRound);
	var endDirStr = FormatDirection("" + endDir, "Conventional", dirType, dirPrec, dirRound);
	var aheadDirStr = FormatDirection("" + aheadDir, "Conventional", dirType, dirPrec, dirRound);
	var backDirStr = FormatDirection("" + backDir, "Conventional", dirType, dirPrec, dirRound);
		
	var startNStr = FormatNumber("" + startN , sourceUnit, convertUnit, coordPrec, coordRound);
	var startEStr = FormatNumber("" + startE , sourceUnit, convertUnit, coordPrec, coordRound);

	var radNStr = FormatNumber("" + radN , sourceUnit, convertUnit, coordPrec, coordRound);
	var radEStr = FormatNumber("" + radE , sourceUnit, convertUnit, coordPrec, coordRound);

	var endNStr = FormatNumber("" + endN , sourceUnit, convertUnit, coordPrec, coordRound);
	var endEStr = FormatNumber("" + endE , sourceUnit, convertUnit, coordPrec, coordRound);
	
	var startStaCnv = ConvertNumber(curveStartSta, sourceUnit, convertUnit);
	var endStaCnv = ConvertNumber(curveEndSta, sourceUnit, convertUnit);
	var startStaStr = FormatStation(startStaCnv, staPattern, staPrec, staRound);
	var endStaStr = FormatStation(endStaCnv, staPattern, staPrec, staRound);
	
	var coordStrOne = text.replace(startNExp, startNStr);
	var coordStrTwo = coordStrOne.replace(startEExp, startEStr);
	var coordStrThree = coordStrTwo.replace(radNExp, radNStr);
	var coordStrFour = coordStrThree.replace(radEExp, radEStr);
	var coordStrFive = coordStrFour.replace(endNExp, endNStr);
	var coordStrSix = coordStrFive.replace(endEExp, endEStr);

	var repStrLen = coordStrSix.replace(lengthFExp, lengthFStr);
	var repStrLen2 = repStrLen.replace(lengthMExp, lengthMStr);
	var repStrLen3 = repStrLen2.replace(lengthUSFExp, lengthUSFStr);
	
	var repStrTang =  repStrLen3.replace( tangtExp, tangentStr);
	var repStrAngle =  repStrTang.replace( angleExp, angleStr);
	
	var repStrRadius =  repStrAngle.replace( radiusFExp,  radiusFStr);
	var repStrRadius2 =  repStrRadius.replace( radiusMExp,  radiusMStr);
	var repStrRadius3 =  repStrRadius2.replace( radiusUSFExp,  radiusUSFStr);
	
	var repStrChordLen =  repStrRadius3.replace(chordLengthFExp,  chordLenFStr);
	var repStrChordLen2 =  repStrChordLen.replace(chordLengthMExp,  chordLenMStr);
	var repStrChordLen3 =  repStrChordLen2.replace(chordLengthUSFExp,  chordLenUSFStr);
	
	var repStrChordDir =  repStrChordLen3.replace(chordDirExp,  chordDirStr);
	var repStrStartDir =  repStrChordDir.replace(startDirExp,  startDirStr);
	var repStrEndDir =  repStrStartDir.replace(endDirExp,  endDirStr);
	var repStrRotation =  repStrEndDir.replace(rotationExp,  rotation);
	var repStrExternal =  repStrRotation.replace(externalExp,  externalStr);
	var repStrMiddle =  repStrExternal.replace(middleExp,  middleStr);
	var repStrDOC =  repStrMiddle.replace(docExp,  docStr);
	
	var repStrConcDir = repStrDOC.replace(concDirExp, concaveDirStr);
	var repStrRunDir = repStrConcDir.replace(runDirExp, runDirStr);
	
	var repStrAheadDir = repStrRunDir.replace(aheadDirExp, aheadDirStr);
	var repStrBackDir = repStrAheadDir.replace(backDirExp, backDirStr);
		
	var repStrStaStart = repStrBackDir.replace(startStaExp, startStaStr);
	var repStrStaEnd = repStrStaStart.replace(endStaExp, endStaStr);
	
	return repStrStaEnd;
}

function SubstituteSpiralKeywords(index, text)
{
	var retStr = "";
	var spLengthFExp = new RegExp("{spiralLengthFoot}", "g");
	var spLengthMExp = new RegExp("{spiralLengthMeter}", "g");
	var spLengthUSFExp = new RegExp("{spiralLengthUSFoot}", "g");

	var spDirExp = new RegExp("{spiralDirection}", "g");
	
	var spStartRadFExp = new RegExp("{spiralStartRadiusFoot}", "g");
	var spStartRadMExp = new RegExp("{spiralStartRadiusMeter}", "g");
	var spStartRadUSFExp = new RegExp("{spiralStartRadiusUSFoot}", "g");
	
	var spEndRadFExp = new RegExp("{spiralEndRadiusFoot}", "g");
	var spEndRadMExp = new RegExp("{spiralEndRadiusMeter}", "g");
	var spEndRadUSFExp = new RegExp("{spiralEndRadiusUSFoot}", "g");
	
	var spRotExp = new RegExp("{spiralRotation}", "g");
	var spStartNExp = new RegExp("{spiralStartNorthing}", "g");
	var spStartEExp = new RegExp("{spiralStartEasting}", "g");
	var spPINExp = new RegExp("{spiralPINorthing}", "g");
	var spPIEExp = new RegExp("{spiralPIEasting}", "g");
	var spEndNExp = new RegExp("{spiralEndNorthing}", "g");
	var spEndEExp = new RegExp("{spiralEndEasting}", "g");
	
	var spChordLengthFExp = new RegExp("{spiralChordLengthFoot}", "g");
	var spChordLengthMExp = new RegExp("{spiralChordLengthMeter}", "g");
	var spChordLengthUSFExp = new RegExp("{spiralChordLengthUSFoot}", "g");
	
	var spAValueExp = new RegExp("{spiralAValue}", "g");
	var spTotalXExp = new RegExp("{spiralTotalX}", "g");
	var spTotalYExp = new RegExp("{spiralTotalY}", "g");
	var spXExp = new RegExp("{spiralX}", "g");
	var spYExp = new RegExp("{spiralY}", "g");
	var spLongTangentExp = new RegExp("{spiralLongTangent}", "g");
	var spShortTangentExp = new RegExp("{spiralShortTangent}", "g");
	var spThetaExp = new RegExp("{spiralTheta}", "g");
	var spPExp = new RegExp("{spiralP}", "g");
	var spKExp = new RegExp("{spiralK}", "g");
	var startStaExp = new RegExp("{spiralStartStation}", "g");
	var endStaExp = new RegExp("{spiralEndStation}", "g");
	
	var spLength = GetSpiralLength(index);
	var spDir = GetSpiralDirection(index);
	var spStartRad = GetSpiralStartRadius(index);
	var spEndRad = GetSpiralEndRadius(index);
	var spRot = GetSpiralRotation(index);
	var spStartNorthing = GetSpiralStartNorthing(index);
	var spStartEasting = GetSpiralStartEasting(index);
	var spPINorthing = GetSpiralPINorthing(index);
	var spPIEasting = GetSpiralPIEasting(index);
	var spEndNorthing = GetSpiralEndNorthing(index);
	var spEndEasting = GetSpiralEndEasting(index);
	var spChordLength = GetSpiralChordLength(index);
	var spAValue = GetSpiralA(index);
	var spTotalX = GetSpiralTotalX(index);
	var spTotalY = GetSpiralTotalY(index);
	var spX = GetSpiralX(index);
	var spY = GetSpiralY(index);
	var spLongTangent = GetSpiralLongTangent(index);
	var spShortTangent = GetSpiralShortTangent(index);
	var spTheta = GetSpiralTheta_Degrees(index);
	var spPValue = GetSpiralP(index);
	var spKValue = GetSpiralK(index);
	var spiralStartSta = GetGeomStartingStation(index);
	var spiralEndSta = GetGeomEndStation(index);
	
	var rotation = "left";
	if(spRot == "cw")
	{
		rotation = "right";
	}
	
	var spLengthFStr = FormatNumber("" + spLength, sourceUnit, "foot", distPrec, distRound);
	var spLengthMStr = FormatNumber("" + spLength, sourceUnit, "meter", distPrec, distRound);
	var spLengthUSFStr = FormatNumber("" + spLength, sourceUnit, "ussurveyfoot", distPrec, distRound);

	var spStartRadFStr = FormatNumber("" + spStartRad, sourceUnit, "foot", distPrec, distRound);
	var spStartRadMStr = FormatNumber("" + spStartRad, sourceUnit, "meter", distPrec, distRound);
	var spStartRadUSFStr = FormatNumber("" + spStartRad, sourceUnit, "ussurveyfoot", distPrec, distRound);

	var spEndRadFStr = FormatNumber("" + spEndRad, sourceUnit, "foot", distPrec, distRound);
	var spEndRadMStr = FormatNumber("" + spEndRad, sourceUnit, "meter", distPrec, distRound);
	var spEndRadUSFStr = FormatNumber("" + spEndRad, sourceUnit, "ussurveyfoot", distPrec, distRound);
	
   	var spStartNorthingStr = FormatNumber("" + spStartNorthing, sourceUnit, convertUnit, distPrec, distRound);
	var spStartEastingStr = FormatNumber("" + spStartEasting, sourceUnit, convertUnit, distPrec, distRound);
	var spPINorthingStr = FormatNumber("" + spPINorthing, sourceUnit, convertUnit, distPrec, distRound);
	var spPIEastingStr = FormatNumber("" + spPIEasting, sourceUnit, convertUnit, distPrec, distRound);
	var spEndNorthingStr = FormatNumber("" + spEndNorthing, sourceUnit, convertUnit, distPrec, distRound);
	var spEndEastingStr = FormatNumber("" + spEndEasting, sourceUnit, convertUnit, distPrec, distRound);
	
	var spChordLengthFStr = FormatNumber("" + spChordLength, sourceUnit, "foot", distPrec, distRound);
	var spChordLengthMStr = FormatNumber("" + spChordLength, sourceUnit, "meter", distPrec, distRound);
	var spChordLengthUSFStr = FormatNumber("" + spChordLength, sourceUnit, "ussurveyfoot", distPrec, distRound);

	var spAValueStr = FormatNumber("" + spAValue, sourceUnit, convertUnit, distPrec, distRound);
	var spTotalXStr = FormatNumber("" + spTotalX, sourceUnit, convertUnit, distPrec, distRound);
	var spTotalYStr = FormatNumber("" + spTotalY, sourceUnit, convertUnit, distPrec, distRound);
	var spXStr = FormatNumber("" + spX, sourceUnit, convertUnit, distPrec, distRound);
	var spYStr = FormatNumber("" + spY, sourceUnit, convertUnit, distPrec, distRound);
	var spLongTangentStr = FormatNumber("" + spLongTangent, sourceUnit, convertUnit, distPrec, distRound);
	var spShortTangentStr = FormatNumber("" + spShortTangent, sourceUnit, convertUnit, distPrec, distRound);
	var spThetaStr = FormatNumber("" + spTheta, sourceUnit, convertUnit, distPrec, distRound);
	var spPValueStr = FormatNumber("" + spPValue, sourceUnit, convertUnit, distPrec, distRound);
	var spKValueStr = FormatNumber("" + spKValue, sourceUnit, convertUnit, convertPrec, convertRound);

	var spDirStr = FormatAngle("" + spDir, "Degrees", "Degrees", "DMS", angPrec, angRound);
	
	var startStaCnv = ConvertNumber(spiralStartSta, sourceUnit, convertUnit);
	var endStaCnv = ConvertNumber(spiralEndSta, sourceUnit, convertUnit);
	var startStaStr = FormatStation(startStaConv, staPattern, staPrec, staRound);
	var endStaStr = FormatStation(endStaConv, staPattern, staPrec, staRound);
	
	var rspLen = text.replace(spLengthFExp, spLengthFStr);
 	var rspLen2 = rspLen.replace(spLengthMExp, spLengthMStr);
	var rspLen3 = rspLen2.replace(spLengthUSFExp, spLengthUSFStr);

	var rspstr = rspLen3.replace(spStartRadFExp, spStartRadFStr);
	var rspstr2 = rspstr.replace(spStartRadMExp, spStartRadMStr);
	var rspstr3 = rspstr2.replace(spStartRadUSFExp, spStartRadUSFStr);

	var rspenr = rspstr3.replace(spEndRadFExp, spEndRadFStr);
	var rspenr2 = rspenr.replace(spEndRadMExp, spEndRadMStr);
	var rspenr3 = rspenr2.replace(spEndRadUSFExp, spEndRadUSFStr);

	var rspstn =  rspenr3.replace(spStartNExp, spStartNorthingStr);
	var rspste = rspstn.replace(spStartEExp, spStartEastingStr);
	var rsppin = rspste.replace(spPINExp, spPINorthingStr);
	var rsppie = rsppin.replace(spPIEExp, spPIEastingStr);
	var rspenn = rsppie.replace(spEndNExp, spEndNorthingStr);
	var rspene = rspenn.replace(spEndEExp, spPIEastingStr);
	
	var rspchl = rspene.replace(spChordLengthFExp, spChordLengthFStr);
	var rspchl2 = rspchl.replace(spChordLengthMExp, spChordLengthMStr);
	var rspchl3 = rspchl2.replace(spChordLengthUSFExp, spChordLengthUSFStr);

	var rspava = rspchl3.replace(spAValueExp, spAValueStr);
	var rsptox = rspava.replace(spTotalXExp, spTotalXStr);
	var rsptoy = rsptox.replace(spTotalYExp, spTotalYStr);
	var rspxva = rsptoy.replace(spXExp, spXStr);
	var rspyva = rspxva.replace(spYExp, spYStr);
	var rsplta = rspyva.replace(spLongTangentExp, spLongTangentStr);
	var rspsta = rsplta.replace(spShortTangentExp, spShortTangentStr);
	var rspthe =  rspsta.replace(spThetaExp, spThetaStr);
	var rsppva = rspthe.replace(spPExp, spPValueStr);
	var rspkva = rsppva.replace(spKExp, spKValueStr);

	var rspdir = rspkva.replace(spDirExp, spDirStr);
	
	var rspRot = rspdir.replace(spRotExp, rotation);
	
	var repStrStaStart = rspRot.replace(startStaExp, startStaStr);
	var repStrStaEnd = repStrStaStart.replace(endStaExp, endStaStr);
	
	return repStrStaEnd;
}

function GetDate()
{
	var d, s = ""; 
   	d = new Date(); 
  	 s += (d.getMonth() + 1) + "/"; 
   	s += d.getDate() + "/"; 
   	s += d.getYear(); 
  	 return(s); 

}

function GetInnerXML(node)
{
    var retStr
    var oNodeList;
    var Item;
    var nType;
	
    retStr = "";
  
    if (node)
    {
        oNodeList = node.childNodes;
        for (var i=0; i<oNodeList.length; i++) 
        {
            Item = oNodeList.item(i);
            nType = Item.nodeType
            if(nType == 3)
            {
                retStr += Item.text;
            }
            else
            {
                retStr += Item.xml;
            }
        }
    }
	
    return retStr + " ";	
}

function GetAreaUnitStr(srcStr)
{
	var lcStr = srcStr.toLowerCase();
	
	if(lcStr == "squarefoot")
	{
		return "square feet";
	}
	else if(lcStr == "squaremeter")
	{
		return "square meters";
	}
	else if(lcStr == "acre")
	{
		return "acres";
	}
	else if(lcStr == "hectare")
	{
		return "hectares"
	}
	else
	{
		return srcStr;
	}
}
// ++++++++++++++++++++++++++++++
// Control Point Arrary functions
// ++++++++++++++++++++++++++++++

function AddControlPoint(name, northing, easting)
{
	if(ControlPointArray.length < 1)
	{
	 	ControlPointArray = new Array(1);
	}
	var ctrlPt = new ControlPoint(name, northing, easting);
	
	ControlPointArray.push(ctrlPt);
	
	return ControlPointArray.length;
}

function SetControlPointDescription(index, desc)
{
	if(ControlPointArray.length > 0)
	{
		var ctrlPt = ControlPointArray[index];
		ctrlPt.desc = desc;
	}
	return desc;
}

function GetControlPointName(index)
{
	var retStr = "Not Found";
	
	if(ControlPointArray.length > 0)
	{
		var ctrlPt = ControlPointArray[index];
		if(ctrlPt)
		{
			retStr = ctrlPt.name;
		}
	}
	
	return retStr;
}

function GetControlPointNorthing(name)
{
	var retNo = 0;
	for(var i = 0; i < ControlPointArray.length; i++)
	{
		var ctrlPt = ControlPointArray(i);
		if(ctrlPt.name == name)
		{
			retNo = ctrlPt.northing;
		}
	}	
	return retNo;
}

function GetControlPointEasting(name)
{
	var retNo = 0;
	for(var i = 0; i < ControlPointArray.length; i++)
	{
		var ctrlPt = ControlPointArray(i);
		if(ctrlPt.name == name)
		{
			retNo = ctrlPt.easting;
		}
	}	
	return retNo;
}

function GetControlPointDesc(name)
{
	var retStr = "Not Found";
	for(var i = 0; i < ControlPointArray.length; i++)
	{
		var ctrlPt = ControlPointArray(i);
		if(ctrlPt.name == name)
		{
			retStr = ctrlPt.desc;
		}
	}	
	return retNo;
}


// ++++++++++++++++++++++++++++++++
// Object definitions
// ++++++++++++++++++++++++++++++++

function ControlPoint(name, northing, easting)
{
	this.name = name;
	this.desc = "";
	this.northing = northing;
	this.easting = easting;
}
]]>
</msxsl:script>
	
</xsl:stylesheet>
