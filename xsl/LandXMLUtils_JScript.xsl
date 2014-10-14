<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet 	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
                		xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                 		xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
                		xmlns:lxml="urn:lx_utils"
                		version="1.0">
<msxsl:script language="JScript" implements-prefix="landUtils"> 
<![CDATA[

function getPerimeter(alignment)
{
	var ndes = this.selectNodes("//CoordGeom");
	return 0.000;
}
function getPointNorthing(point)
{ 
  var q = new String();
  q = point.nextNode().text;
  var strs = q.split(" ");

  if (strs)
//	return formatCoordString(strs[0]);
	return strs[0];
  else
    return "0.00";
}

function getPointEasting(point)
{ 
  var q = new String();
  q = point.nextNode().text;
  var strs = q.split(" ");

  if (strs)
//	return formatCoordString(strs[1]);
	return strs[1];
  else
    return "0.00";
}

function getPointElevation(point)
{ 
  var q = new String();
  q = point.nextNode().text;
  var strs = q.split(" ");

  if (strs)
//	return formatCoordString(strs[2]);
	return strs[2];
  else
    return "0.00";
}


function getStation(sometext)
{ 
  var q = new String();
  q = sometext;
  var strs = q.split(" ");

  if (strs)
    return strs[0];
  else
    return "0.00";
}

function getElevation(sometext)
{ 
  var q = new String();
  q = sometext;
  var strs = q.split(" ");
	
  if (strs && (strs.length > 0))
    return strs[1];
  else
    return "0.00";
}

function getGradeOut(thisPVI)
{
  // arg is a node-set
  return getGradeOutInternal(thisPVI.item(0));
}

function getGradeIn(thisPVI)
{
  // arg is a node-set
  return getGradeInInternal(thisPVI.item(0));

}

function getGradeOutInternal(thisPVI)
{
// arg is one node, rather than a node-set (which is what's passed in when
// function is called from XSL)

// If there is a following PVI, calculate the grade out

  var nextPVI, nextStation, nextElevation, thisStation, theElevation, thisElevation;
  var run;


  nextPVI = thisPVI.selectSingleNode("(./following-sibling::lx:PVI | ./following-sibling::lx:ParaCurve)[1]");

  if ( null == nextPVI )
    return null;    
 // return "";


  nextStation   = getStation(nextPVI.nodeTypedValue);
  nextElevation = getElevation(nextPVI.nodeTypedValue);

  thisStation   = getStation(thisPVI.nodeTypedValue);
  thisElevation = getElevation(thisPVI.nodeTypedValue);

  run = nextStation - thisStation;

  if ( run+0 == 0.0 )
     return 0.0;
  else
     return (100 * (nextElevation - thisElevation) / run);

}

function getGradeInInternal(thisPVI)
{
// arg is one node, rather than a node-set (which is what's passed in when
// function is called from XSL)

// If there is a previous PVI, calculate the grade in

  var prevPVI, prevStation, prevElevation, thisStation, theElevation, thisElevation;
  var run;

  prevPVI = thisPVI.selectSingleNode("(./preceding-sibling::lx:PVI | ./preceding-sibling::lx:ParaCurve)[position()=last()]");

  if ( null == prevPVI )
    return null;
//    return "";

  prevStation   = getStation(prevPVI.nodeTypedValue);
  prevElevation = getElevation(prevPVI.nodeTypedValue);

  thisStation   = getStation(thisPVI.nodeTypedValue);
  thisElevation = getElevation(thisPVI.nodeTypedValue);

  run = thisStation - prevStation;

  if ( run+0 == 0.0 )
    return 0.0;
  else
    return (100 * (thisElevation - prevElevation) / run);
}

function getA(thisPVI)
{
  var gradeIn, gradeOut;

  gradeIn = getGradeIn(thisPVI);
 
  if ( gradeIn == null ) return "";
  if ( gradeIn == "" ) gradeIn = 0.0;

  gradeOut = getGradeOut(thisPVI);
	
  if ( gradeOut == null ) return "";
  if ( gradeOut == "" ) gradeOut = 0.0;

  return Math.abs(gradeIn - gradeOut);

}

function getK(paraCurveNode)
{
  var a, curveLen;

  a = getA(paraCurveNode);

  if ( a+0 < 0.000000001 ) return "Infinite";

  curveLen = paraCurveNode.item(0).getAttribute('length');

  return curveLen/a;

}

function getPvcStation(paraCurveNode)
{ 
  var station;
  var curveLen;

  station = new Number(getStation( paraCurveNode.item(0).nodeTypedValue ));
  curveLen = paraCurveNode.item(0).getAttribute('length');

  return ( station - (curveLen/2) );
}

function getPvtStation(paraCurveNode)
{ 
  var station;
  var curveLen;

  station = new Number(getStation( paraCurveNode.item(0).nodeTypedValue ));
  curveLen = paraCurveNode.item(0).getAttribute('length');

  return ( station + (curveLen/2) );
}

function getCurveType(paraCurveNode)
{
//  if (IsEGProfile() ) return AeccPVIWrap::kVerticalCurveNone;

  var gradeIn, gradeOut;

  gradeIn = getGradeIn(paraCurveNode);

  if ( gradeIn == null ) return "none";
  if ( gradeIn == "" ) gradeIn = 0.0;
// if ( gradeIn == "" ) return "none";

  gradeOut = getGradeOut(paraCurveNode);
	
  if ( gradeOut == null ) return "none";  
  if ( gradeOut == "" ) gradeOut = 0.0;
//  if ( gradeOut == "" ) return "none";


  if ( gradeIn+0 > gradeOut+0 )

      return "crest";
  
  else if (gradeIn+0 < gradeOut+0 )

      return "sag";

  return "none";

}

function getSmallestStation(pviList)
{
  var station;
  var smallestStation = 0.0;

  var firstIter = true;

  for (var pvi = pviList.nextNode(); pvi; pvi = pviList.nextNode())
  {
    station = getStation(pvi.nodeTypedValue);
    if ( firstIter )
    {
      firstIter = false;
      smallestStation = station;
    }

    if ( station < smallestStation )
      smallestStation = station;
  }

  return smallestStation;
}

function getLargestStation(pviList)
{
  var station;
  var largestStation=0.0;

  var firstIter = true;

  for (var pvi = pviList.nextNode(); pvi; pvi = pviList.nextNode())
  {
    station = getStation(pvi.nodeTypedValue);
    if ( firstIter )
    {
      firstIter = false;
      largestStation = station;
    }

    if ( station > largestStation )
      largestStation = station;
  }

  return largestStation;
}

function interpolate(x1, y1, x2, y2, x3, y3)
{
  var slope, elevation;

  slope = (y1 - y2) / (x1 - x2);

  if ( y3 == 0.0 )
      elevation = y1 - ((x1 - x3) * slope);
  else
      elevation = x1 - ((y1 - y3) * slope);

  return elevation; 
}

function getElevationAtStation(station, pviList)
{
// Calculates the elevation at station, using
// pvi information contained in pviList

// shouldn't this start at pviList.item(0)?
//  for (var pvi = pviList.nextNode(); pvi; pvi = pviList.nextNode())
  for (var pvi = pviList.item(0); pvi; pvi = pviList.nextNode())
  {
      var existingProfile = pvi.selectSingleNode("./ancestor::lx:ProfAlign[@state='existing']");

      var listStation = getStation(pvi.nodeTypedValue);

      if ( station+0 == listStation+0 )
      {
          if ( existingProfile != null )
          {

	  	// Specified PVI currently exists in list
          	return getElevation(pvi.nodeTypedValue);
   	  }
          else
	  {
             var curveLen;
             curveLen = pvi.getAttribute('length');

             if ( curveLen == null )
             {
               curveLen = 0;
             }

             if ( curveLen == 0.0 )
             {
          	return getElevation(pvi.nodeTypedValue);
             }
             else
             {
	      	return calc_vcurve_elev(pvi, station);
             }
	  }
      }

  //    nextListPVI = pvi.selectSingleNode("(./following-sibling::lx:PVI | ./following-sibling::lx:ParaCurve)[1]");
      var nextListPVI = pvi.selectSingleNode("./following-sibling::lx:*");

      if ( null == nextListPVI )
         return "";

      var nextListStation = getStation(nextListPVI.nodeTypedValue);

      if ( ( station+0 > listStation+0) && (station+0 < nextListStation+0) )
      {
          // Specified PVI found between two existing PVIs

          // it's an existing ground profile so don't have to consider curves
          if ( existingProfile != null )
          {
	     var startStation   = getStation(pvi.nodeTypedValue);
	     var startElevation = getElevation(pvi.nodeTypedValue);

	     var endStation   = nextListStation;
	     var endElevation = getElevation(nextListPVI.nodeTypedValue);

   	     return interpolate(endStation, endElevation, startStation, startElevation, station, 0.0);
          }
          else
          {
             var curveLen, nextCurveLen;
 	     var pviStation, nextPviStation, pviCurveEnd, nextPviCurveStart;
  
             curveLen = pvi.getAttribute('length');

             if ( curveLen == null )
             {
               curveLen = 0;
             }

             pviStation = new Number(getStation(pvi.nodeTypedValue));
	     pviCurveEnd = pviStation + ( curveLen/2 );
   
             nextCurveLen = nextListPVI.getAttribute('length');

             if ( nextCurveLen == null )
             {
               nextCurveLen = 0;
             }

             nextPviStation = new Number(getStation(nextListPVI.nodeTypedValue));
	     nextPviCurveStart = nextPviStation - ( nextCurveLen/2 );

	     if (station < nextPviCurveStart)   // before the next pvi vertical curve 
	     {
		if ( station >= pviCurveEnd || ( station < pviCurveEnd && curveLen == 0 ))   //shouldn't this be just > ???
		{
		   // after the current pvi vertical curve (so not on either curve)

	     	   startStation   = pviStation;
	           startElevation = getElevation(pvi.nodeTypedValue);

	           endStation   = nextPviStation;
	           endElevation = getElevation(nextListPVI.nodeTypedValue);

   	           return interpolate(endStation, endElevation, startStation, startElevation, station, 0.0);

		}
		else		
                {
		   // on the current pvi vertical curve (does this branch ever get hit?)
		   return calc_vcurve_elev(pvi, station);
                }
	     }
	     else    
             {
		// on the next pvi's vertical curve (if there is a curve)
                if ( nextCurveLen > 0 )
                {
 		   return calc_vcurve_elev(nextListPVI, station);
                 }
                else
                {

	     	   startStation   = pviStation;
	           startElevation = getElevation(pvi.nodeTypedValue);

	           endStation   = nextPviStation;
	           endElevation = getElevation(nextListPVI.nodeTypedValue);

   	           return interpolate(endStation, endElevation, startStation, startElevation, station, 0.0);

                }
             }

          }
      }

  }

  return "";
}

function trackedGetElevationAtStation(station, pviList, obj)
{
// Calculates the elevation at station, using
// pvi information contained in pviList

  // check that station is before the first PVI in list
  // don't bother searching through list if it is

  if ( pviList == null ) 
  {         
     obj.returnPVI = null;
     return "";
  }

  var firstPVI = pviList.item(0);

  if ( station < getStation( firstPVI.nodeTypedValue ) )
  {
     obj.returnPVI = firstPVI;
     return "";
  }

  for (var pvi = pviList.item(0); pvi; pvi = pviList.nextNode())
  {
      obj.returnPVI = pvi;

      var existingProfile = pvi.selectSingleNode("./ancestor::lx:ProfAlign[@state='existing']");

      var listStation = getStation(pvi.nodeTypedValue);
 
      if ( station+0 == listStation+0 )
      {
          if ( existingProfile != null )
          {
	  	// Specified PVI currently exists in list
          	return getElevation(pvi.nodeTypedValue);
   	  }
          else
	  {
             var curveLen;
             curveLen = pvi.getAttribute('length');

             if ( curveLen == null )
             {
               curveLen = 0;
             }

             if ( curveLen == 0.0 )
             { 
          	return getElevation(pvi.nodeTypedValue);
             }
             else
             {
	      	return calc_vcurve_elev(pvi, station);
             }
	  }
      }

//      nextListPVI = pvi.selectSingleNode("(./following-sibling::lx:PVI | ./following-sibling::lx:ParaCurve)[1]");
//      nextListPVI = pvi.selectSingleNode("(./following-sibling::lx:PVI | ./following-sibling::lx:ParaCurve)");
      var nextListPVI = pvi.selectSingleNode("./following-sibling::lx:*");

      if ( null == nextListPVI )
      {         
         obj.returnPVI = null;
         return "";
      }

      var nextListStation = getStation(nextListPVI.nodeTypedValue);

      if ( ( station+0 > listStation+0) && (station+0 < nextListStation+0) )
      {
          // Specified PVI found between two existing PVIs

          // it's an existing ground profile so don't have to consider curves
          if ( existingProfile != null )
          {
	     var startStation   = getStation(pvi.nodeTypedValue);
	     var startElevation = getElevation(pvi.nodeTypedValue);

	     var endStation   = nextListStation;
	     var endElevation = getElevation(nextListPVI.nodeTypedValue);
   	     return interpolate(endStation, endElevation, startStation, startElevation, station, 0.0);
          }
          else
          {
             var curveLen, nextCurveLen;
 	     var pviStation, nextPviStation, pviCurveEnd, nextPviCurveStart;
  
             curveLen = pvi.getAttribute('length');

             if ( curveLen == null )
             {
               curveLen = 0;
             }

             pviStation = new Number(getStation(pvi.nodeTypedValue));
	     pviCurveEnd = pviStation + ( curveLen/2 );
   
             nextCurveLen = nextListPVI.getAttribute('length');

             if ( nextCurveLen == null )
             {
               nextCurveLen = 0;
             }

             nextPviStation = new Number(getStation(nextListPVI.nodeTypedValue));
	     nextPviCurveStart = nextPviStation - ( nextCurveLen/2 );

	     if (station < nextPviCurveStart)   // before the next pvi vertical curve 
	     {
		if ( station >= pviCurveEnd || ( station < pviCurveEnd && curveLen == 0 ))   //shouldn't this be just > ???
		{
		   // after the current pvi vertical curve (so not on either curve)

	     	   startStation   = pviStation;
	           startElevation = getElevation(pvi.nodeTypedValue);

	           endStation   = nextPviStation;
	           endElevation = getElevation(nextListPVI.nodeTypedValue);
                   
   	           return interpolate(endStation, endElevation, startStation, startElevation, station, 0.0);

		}
		else		
                {
		   // on the current pvi vertical curve (does this branch ever get hit?)
		   return calc_vcurve_elev(pvi, station);
                }
	     }
	     else    
             {
		// on the next pvi's vertical curve (if there is a curve)
                if ( nextCurveLen > 0 )
                {
 		   return calc_vcurve_elev(nextListPVI, station);
                }
                else
                {

	     	   startStation   = pviStation;
	           startElevation = getElevation(pvi.nodeTypedValue);

	           endStation   = nextPviStation;
	           endElevation = getElevation(nextListPVI.nodeTypedValue);

   	           return interpolate(endStation, endElevation, startStation, startElevation, station, 0.0);

                }
             }

          }
      }

  }

  // should not ever get here
  obj.returnPVI = null;
  return "";
}

function getSpeedAt(station, speedList)
{
  // find largest station value that is less than station argument

  var theSpeed = 0;
  var stationPropertyNode, speedPropertyNode;
  var tmpStation;
  var prevStation = 0;
  var bFirst = true;

  if ( speedList == null ) return theSpeed;

  for (var speedNode = speedList.nextNode(); speedNode; speedNode = speedList.nextNode())
  {
      stationPropertyNode = speedNode.selectSingleNode("./lx:Property[@label='station']");
      tmpStation = stationPropertyNode.getAttribute('value');

      if ( station >= tmpStation && ( bFirst || tmpStation > prevStation ) )
      {
            speedPropertyNode = speedNode.selectSingleNode("./lx:Property[@label='speed']");
            theSpeed = speedPropertyNode.getAttribute('value');

            prevStation = tmpStation;
            bFirst = true;
      }
  }

  return theSpeed;
}

function reportIncrementInfo(profiles,     nIncrement, 
                             startStation, endStation)
{
// Builds a nodeset? that can be passed to an XSL template
// to display the increment report info.
//
// e.g.,
//
// <incrReport>
//   <profiles>
//      <profile name="center" desc="rock" />
//      <profile name="right" desc="rock" />
//      <profile name="left" desc="rock" />
//   </profiles>
//   <stationIncr station="10">
//       <elevation>100.1</elevation>
//       <elevation>102.8</elevation>
//       <elevation>131.0</elevation>
//   </stationIncr>
//   <stationIncr station="20">
//       <elevation>100.3</elevation>
//       <elevation>101.0</elevation>
//       <elevation>131.1</elevation>
//   </stationIncr>
// </incrReport>


  var xmlNewDoc = new ActiveXObject("Msxml2.DOMDocument");

  var incrementInfo = xmlNewDoc.createElement("incrReport");
  xmlNewDoc.appendChild(incrementInfo);

  var nProfiles = profiles.length;

  if ( ( nIncrement == 0 )  || ( endStation < startStation ) || nProfiles < 1 )
     return incrementInfo;

  var profilesNode = xmlNewDoc.createElement("profiles");
  incrementInfo.appendChild(profilesNode);

  var n = 0;

  for (n = 0; n < nProfiles; n++)
  {
    var profileNode = xmlNewDoc.createElement("profile");
    profilesNode.appendChild(profileNode);

    var profileAttrs = profiles.item(n).attributes;
    var outAttrs =  profileNode.attributes;

    var outNameAttr = xmlNewDoc.createAttribute("name");
    outNameAttr.text =  profileAttrs.getNamedItem("name").text;
    outAttrs.setNamedItem(outNameAttr);

    var outDescAttr = xmlNewDoc.createAttribute("desc");
    outDescAttr.text = profileAttrs.getNamedItem("desc").text;
    outAttrs.setNamedItem(outDescAttr);
  }

  var nStations;
  var stationDiff = endStation - startStation;

  if ( (stationDiff % nIncrement) == 0 )
    nStations =  (stationDiff / nIncrement) + 1;
  else
    nStations =  (stationDiff / nIncrement);

   var nSize = new Number(nStations);
   var stationIncrArray = new Array(nSize);

  var j = 0;
  for (j = 0; j < nStations; j++)
  { 
    var stationValue = startStation + (j * nIncrement);

    stationIncrArray[j] = xmlNewDoc.createElement("stationIncr");
    incrementInfo.appendChild( stationIncrArray[j]);

    var stationAttr = xmlNewDoc.createAttribute("station");
    stationAttr.text = stationValue;

    var stationAttrs =  stationIncrArray[j].attributes;

    stationAttrs.setNamedItem(stationAttr);
  }

  var pviList;
  var obj = new Object();

  var i = 0;
  for (i = 0; i < nProfiles; i++)
  {
    // initialize to the complete list of PVIs
    pviList = profiles.item(i).childNodes;

    for (j = 0; j < nStations; j++)
    {    
      stationValue = startStation + (j * nIncrement);

      var elevation = xmlNewDoc.createElement("elevation");

      var elevationValue = trackedGetElevationAtStation(stationValue, pviList , obj);

      elevation.text = elevationValue;
      stationIncrArray[j].appendChild(elevation);

      // next time trackedGetElevationAtStation() is called, won't have to 
      // search the PVI list that's already been searched

      if ( obj.returnPVI != null )
      {
         pviList = obj.returnPVI.selectNodes("( ./self::lx:* | ./following-sibling::lx:PVI | ./following-sibling::lx:ParaCurve)");
      }
      else
      {
	pviList = null;  // gone past the end og PVI list (no more stations will match)
      }

    }
  }

  return incrementInfo;
}


function calc_vcurve_elev(pviNode, station)
{
   var curveLen;
   var pviStation, prevPvi;

   var elev, elev_bvc, gradeIn, gradeOut;
   var x_dist, a_val;

   pviStation = getStation(pviNode.nodeTypedValue);

   // the * 100 is used because the grades are in %
   curveLen = pviNode.getAttribute("length");

   // grades in, out and the rate of change

   gradeIn = getGradeInInternal(pviNode);
   gradeOut = getGradeOutInternal(pviNode);

   // the change of distance from bvc to asked for station
   x_dist = (station - (pviStation - curveLen/2)) / 100.0; 

   // set up points for interpolation to get the elevation of the bvc

   prevPvi = pviNode.selectSingleNode("(./preceding-sibling::lx:PVI | ./preceding-sibling::lx:ParaCurve)[position()=last()]");

   var startStation   = getStation(prevPvi.nodeTypedValue);
   var startElevation = getElevation(prevPvi.nodeTypedValue);
   var endStation     = pviStation;
   var endElevation   = getElevation(pviNode.nodeTypedValue);

   var temp_ptX = pviStation - ( curveLen/2 );

   elev_bvc = interpolate(endStation, endElevation, startStation, startElevation, temp_ptX, 0.0);

   // do the final calculation - y = (r/2)*x*x + g1*x + elev_bvc

   var curveLen2 = curveLen / 100.0;

   a_val = (gradeOut - gradeIn) / (2.0 * curveLen2); 
   elev = (a_val * x_dist * x_dist) + (gradeIn * x_dist) + elev_bvc;

   return elev;
}

 
function getLowPointStation(paraCurveNode)
{
  var station;
  var curveLen = paraCurveNode.item(0).getAttribute('length');

  if ( null == curveLen  || curveLen <= 0.0 ) return "";

  var gradeIn, gradeOut;

  gradeIn = getGradeInInternal(paraCurveNode.item(0)) / 100.0 ;
  gradeOut = getGradeOutInternal(paraCurveNode.item(0)) / 100.0 ;
	
  if ( (gradeIn+0) < 0.0 && (gradeOut+0) > 0.0 )  
  {
     if ( Math.abs(gradeIn) == Math.abs(gradeOut) )
     {
       // if absolute value of grade in equals abs of grade out,
       // high/low point is in middle of curve

       station  = getStation(paraCurveNode.item(0).nodeTypedValue);

       return station;
     }
     else
     {
       // low point will be on side of curve with flatter gradient
       // Distance from end station to high/low point

       var A;
       A = getA(paraCurveNode) / 100.0;

       var distToLowPoint;  // from curve end point
       distToLowPoint = - ( ( (gradeIn+0) * (curveLen+0) ) / A );

       var pvcStation = getStation(paraCurveNode.item(0).nodeTypedValue) - (curveLen / 2.0);

       station = pvcStation + distToLowPoint;

       return station;
     }
   }
   else
      return "";

}

function getHighPointStation(paraCurveNode)
{
  var station;
  var curveLen = paraCurveNode.item(0).getAttribute('length');

  if ( null == curveLen  || curveLen <= 0.0 ) return "";

  var gradeIn, gradeOut;

  gradeIn = getGradeInInternal(paraCurveNode.item(0)) / 100.0 ;
  gradeOut = getGradeOutInternal(paraCurveNode.item(0)) / 100.0 ;
 	
  if ( (gradeIn+0) > 0.0 && (gradeOut+0) < 0.0 )  
  {
     if ( Math.abs(gradeIn) == Math.abs(gradeOut) )
     {
       // if absolute value of grade in equals abs of grade out,
       // high point is in middle of curve

       station  = getStation(paraCurveNode.item(0).nodeTypedValue);
       return station;
     }
     else
     {
       // high point will be on side of curve with flatter gradient
       // Distance from end station to high/low point

       var A;
       A = getA(paraCurveNode) / 100.0;

       var distToHighPoint;  // from curve end point

       distToHighPoint =  ( (gradeIn+0) * (curveLen+0) ) / A ;

       var pvcStation = getStation(paraCurveNode.item(0).nodeTypedValue) - (curveLen / 2.0);

       station = pvcStation + distToHighPoint;

       return station;
     }
   }
   else
      return "";

}
]]>
</msxsl:script>

</xsl:stylesheet>