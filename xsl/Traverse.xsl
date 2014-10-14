<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet 	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
                		xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                 		xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
                		xmlns:lxml="urn:lx_utils"
                		version="1.0">
<!-- =========== JavaScript Include==== -->
<xsl:include href="General_Formating_JScript.xsl"></xsl:include>
<xsl:include href="H_Curve_Calculations.xsl"></xsl:include>
<xsl:include href="H_Line_Calculations.xsl"></xsl:include>
<xsl:include href="Cogo_Point.xsl"></xsl:include>
<!-- ================================= -->

<msxsl:script language="JScript" implements-prefix="landUtils"><![CDATA[
var occupiedNorthing = 0;
var occupiedEasting = 0;

var backsightNorthing = 0;
var backsightEasting = 0;

var radialPointNorthing = 0;
var radialPointEasting = 0;


]]></msxsl:script>
</xsl:stylesheet>
