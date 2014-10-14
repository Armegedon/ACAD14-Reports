<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet 	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
                		xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                 		xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
                		xmlns:lxml="urn:lx_utils"
                		version="1.0">
	<xsl:param name="useStationFormat"/>
	<xsl:param name="stationFormatMinWidth"/>
	<xsl:param name="stationFormatDecPrec"/>
	<xsl:param name="stationFormatUseParens"/>
	<xsl:param name="stationFormatUseLeadZeros"/>
	<xsl:param name="stationFormatDropDecimal"/>
	<xsl:param name="stationFormatDecimalChar"/>
	<xsl:param name="stationFormatStationChar"/>
	<xsl:param name="stationFormatBase"/>
	<msxsl:script language="JScript" implements-prefix="landUtils"><![CDATA[
	
   var useStationFormat 	= new Boolean();
   var minWidth         		= new Number(3);
   var nPrec           		= new Number(3);
   var useParens        	= new Boolean();
   var useLeadZeros    	= new Boolean();
   var dropDecimal      	= new Boolean();
   var nBase           		 = new Number(1000);
   var stationChar		="+";
   var decimalChar		=".";

function setUseStationFormat(use)
{
	useStationFormat = new Boolean(use);
	return "ok";
}
function setStationMinWidth(staMinWidth)
{
	minWidth = new Number(staMinWidth);
	return "ok";
}
function setStationDecPrecision(nDecPrec)
{
	nPrec = new Number(nDecPrec);
	return "ok";
}
function setUseParens(useStaParens)
{
	useParens = new Boolean(useStaParens);
	return "ok";
}
function setUseLeadingZeros(useZeros)
{
	useLeadZeros = new Boolean(useZeros);
	return "ok";
}
function setDropDecimal(dropDec)
{
	dropDecimal = new Boolean(dropDec);
	return "ok";
}
function setStationBase(baseNumber)
{
	nBase = new Number(baseNumber);
	return "ok";
}
 function setStationCharacter(character)
 {
 	stationChar=character;
 	return "ok";
}
 function setDecimalCharacter(character)
 {
 	decimalChar=character;
	return "ok";
 }
function FormatDecimalStation(station)
{
	var str = NodeToText(station);
	var dStation = new Number(parseFloat(str));
   	var strFormatted;   
   	var isNegative = false;

   if ( dStation < 0 )
        isNegative = true;
  
   var dPositiveStation = Math.abs(dStation);

   if ( !useStationFormat )
   {
      var num = new Number(dPositiveStation);

      if ( (dropDecimal) && ( nPrec > 0 ) )
      {
          if ( Math.floor(dPositiveStation) == dPositiveStation )  
             nPrec = 0;
      }   

      strFormatted = num.toFixed(nPrec);
   }
   else
   {
     var zeros = Math.log(nBase)/Math.log(10);

     zeros = Math.round(zeros);

//   Need to round first, otherwise 1599.995 (base 100, nprec=2 ) will become 15+100.00
     var tempNum = new Number(dPositiveStation);
     var strRounded = tempNum .toFixed(nPrec);
     dPositiveStation = new Number(strRounded)

     var leftVal = Math.floor(dPositiveStation/nBase);
     var leftStr = leftVal.toString();

     var rightVal = dPositiveStation-(leftVal*nBase);

     if ( (dropDecimal != 0) && ( nPrec > 0 ) )
     {
          if ( Math.floor(rightVal) == rightVal )  
             nPrec = 0;
     }   

     var rightStr = new String(rightVal.toFixed(nPrec));

     if ( nPrec == 0 )
        var decimalPos = rightStr.length;
     else
       decimalPos = rightStr.length - nPrec - 1;

     var nMoreZeros = zeros - decimalPos;

     var i =1;
     for (i=1; i <=nMoreZeros; i++)
     {
        rightStr = "0" + rightStr;
     }

     strFormatted =  leftStr + stationChar + rightStr;

   }

   // Replace the decimal character with specified decimal character

   // strFormated must contain nPrec chars after the decimal point,
   // therefore we know that the decimal point is at 
   // strFormatted.length-nPrec-1 (unless nPrec was 0, then
   // there won't be a decimal point)

   if ( nPrec > 0 )
   {
       decimalPos = strFormatted.length - nPrec - 1;
       strFormatted = strFormatted.substr(0, decimalPos) + decimalChar + strFormatted.substr(decimalPos+1);
   }

   // Add leading zeros

   if ( useLeadZeros )
   {
     var nLeadZeros = minWidth - strFormatted.length;
     
     for (i=1; i <=nLeadZeros; i++)
     {
        strFormatted = "0" + strFormatted;
     }
   
   }

   if ( isNegative )
   {
        if ( useParens != 0)
	   strFormatted =  "(" + strFormatted + ")";
        else
	   strFormatted =  "-" + strFormatted;
   }

   return strFormatted;
	
}

function getFormattedStation(argDStation)
{
   var dStation         = new Number(argDStation);

   var strFormatted;   
   var isNegative = false;

   if ( dStation < 0 )
        isNegative = true;
  
   var dPositiveStation = Math.abs(dStation);

   if ( !useStationFormat )
   {
      var num = new Number(dPositiveStation);

      if ( (dropDecimal) && ( nPrec > 0 ) )
      {
          if ( Math.floor(dPositiveStation) == dPositiveStation )  
             nPrec = 0;
      }   

      strFormatted = num.toFixed(nPrec);
   }
   else
   {
     var zeros = Math.log(nBase)/Math.log(10);

     zeros = Math.round(zeros);

//   Need to round first, otherwise 1599.995 (base 100, nprec=2 ) will become 15+100.00
     var tempNum = new Number(dPositiveStation);
     var strRounded = tempNum .toFixed(nPrec);
     dPositiveStation = new Number(strRounded)

     var leftVal = Math.floor(dPositiveStation/nBase);
     var leftStr = leftVal.toString();

     var rightVal = dPositiveStation-(leftVal*nBase);

     if ( (dropDecimal != 0) && ( nPrec > 0 ) )
     {
          if ( Math.floor(rightVal) == rightVal )  
             nPrec = 0;
     }   

     var rightStr = new String(rightVal.toFixed(nPrec));

     if ( nPrec == 0 )
        var decimalPos = rightStr.length;
     else
       decimalPos = rightStr.length - nPrec - 1;

     var nMoreZeros = zeros - decimalPos;

     var i =1;
     for (i=1; i <=nMoreZeros; i++)
     {
        rightStr = "0" + rightStr;
     }

     strFormatted =  leftStr + stationChar + rightStr;

   }

   // Replace the decimal character with specified decimal character

   // strFormated must contain nPrec chars after the decimal point,
   // therefore we know that the decimal point is at 
   // strFormatted.length-nPrec-1 (unless nPrec was 0, then
   // there won't be a decimal point)

   if ( nPrec > 0 )
   {
       decimalPos = strFormatted.length - nPrec - 1;
       strFormatted = strFormatted.substr(0, decimalPos) + decimalChar + strFormatted.substr(decimalPos+1);
   }

   // Add leading zeros

   if ( useLeadZeros )
   {
     var nLeadZeros = minWidth - strFormatted.length;
     
     for (i=1; i <=nLeadZeros; i++)
     {
        strFormatted = "0" + strFormatted;
     }
   
   }

   if ( isNegative )
   {
        if ( useParens != 0)
	   strFormatted =  "(" + strFormatted + ")";
        else
	   strFormatted =  "-" + strFormatted;
   }

   return strFormatted;
}

function getStationWithEquation(stationEquations, internalStation)
{
// The station equations MUST be in ascending order for this routine to work properly
  internalStation = new Number(internalStation);

  var stationWithEquation = internalStation;
  var staEquInternal;
  var staEquOrder;

  var i = 0;
  for(i=0; i < stationEquations.length; i++)
  {   
     staEquInternal = stationEquations.item(i).attributes.getNamedItem("staInternal").nodeTypedValue;
     var staEquAhead    = stationEquations.item(i).attributes.getNamedItem("staAhead").nodeTypedValue;

     if ( staEquInternal <= internalStation)
     {
 	staEquOrder = stationEquations.item(i).selectSingleNode("descendant::lx:Feature[@code='Order']/lx:Property[@label='OrderValue']/@value");

        if ( staEquOrder == null || staEquOrder.text != "decreasing")
    	    stationWithEquation = internalStation + ( staEquAhead - staEquInternal );           
        else 
            stationWithEquation = staEquAhead - ( internalStation - staEquInternal);
       
      }
  }

  return stationWithEquation.toString();
 
}
]]></msxsl:script>
<xsl:template name="SetStationFormatParameters">
		<xsl:if test="$useStationFormat">
			<xsl:variable name="vardef" select="landUtils:setUseStationFormat($useStationFormat)"/>
		</xsl:if>
		<xsl:if test="$stationFormatMinWidth">
			<xsl:variable name="vardef" select="landUtils:setStationMinWidth($stationFormatMinWidth)"/>
		</xsl:if>
		<xsl:if test="$stationFormatDecPrec">
			<xsl:variable name="vardef" select="landUtils:setStationDecPrecision($stationFormatDecPrec)"/>
		</xsl:if>
		<xsl:if test="$stationFormatUseParens">
			<xsl:variable name="vardef" select="landUtils:setUseParens($stationFormatUseParens)"/>
		</xsl:if>
		<xsl:if test="$stationFormatUseLeadZeros">
			<xsl:variable name="vardef" select="landUtils:setUseLeadingZeros($stationFormatUseLeadZeros)"/>
		</xsl:if>
		<xsl:if test="$stationFormatDropDecimal">
			<xsl:variable name="vardef" select="landUtils:setDropDecimal($stationFormatDropDecimal)"/>
		</xsl:if>
		<xsl:if test="$stationFormatDecimalChar">
			<xsl:variable name="vardef" select="landUtils:setDecimalCharacter($stationFormatDecimalChar)"/>
		</xsl:if>
		<xsl:if test="$stationFormatStationChar">
			<xsl:variable name="vardef" select="landUtils:setStationCharacter($stationFormatStationChar)"/>
		</xsl:if>
		<xsl:if test="$stationFormatBase">
			<xsl:variable name="vardef" select="landUtils:setStationBase($stationFormatBase)"/>
		</xsl:if>
</xsl:template>
</xsl:stylesheet>
