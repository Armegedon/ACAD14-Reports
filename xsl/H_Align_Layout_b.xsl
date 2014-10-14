<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet version="1.0"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com"
	xmlns:lxml="urn:lx_utils">
	
<xsl:include href="NodeConversion_JScript.xsl"/>
<xsl:include href="CoordGeometry.xsl"/>

<!-- =========== Formatting parameters==== -->
<xsl:param name="Alignment.Coordinate.unit"/>
<xsl:param name="Alignment.Coordinate.precision"/>
<xsl:param name="Alignment.Coordinate.rounding"/>
<xsl:param name="Alignment.Elevation.unit"/>
<xsl:param name="Alignment.Elevation.precision"/>
<xsl:param name="Alignment.Elevation.rounding"/>
<xsl:param name="Alignment.Linear.unit"/>
<xsl:param name="Alignment.Linear.precision"/>
<xsl:param name="Alignment.Linear.rounding"/>
<xsl:param name="Alignment.Angular.unit"/>
<xsl:param name="Alignment.Angular.precision"/>
<xsl:param name="Alignment.Angular.rounding"/>
<xsl:param name="Alignment.Station.Display"/>
<xsl:param name="Alignment.Station.unit"/>
<xsl:param name="Alignment.Station.precision"/>
<xsl:param name="Alignment.Station.rounding"/>
<!-- ================================= -->

</xsl:stylesheet>
