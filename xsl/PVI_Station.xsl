<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">

<!--Description:PVI Station report.
This form provides PVI station, elevation, and other design information for profile alignments in your LandXML data.

This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
<!--CreatedBy:Autodesk Inc. -->
<!--DateCreated:06/15/2002 -->
<!--LastModifiedBy:Autodesk Inc. -->
<!--DateModified:04/04/2005 $ -->
<!--OutputExtension:html -->

<xsl:output method="html" encoding="UTF-8"/>

<!-- ==== JavaScript Includes ==== -->
<xsl:include href="header.xsl"/>
<xsl:include href="V_Alignment_Layout_b.xsl"/>
<xsl:include href="Number_Formatting.xsl"/>
<!-- ============================= -->

<xsl:param name="SourceLinearUnit" select="//lx:Units/*/@linearUnit"/>

<xsl:template match="/">

<html>
<head>
	<title>PVI Stations report</title>
</head>
<body>
	<div style="width:7in">

	<xsl:call-template name="AutoHeader">
		<xsl:with-param name="ReportTitle">PVI Stations Report</xsl:with-param>
		<xsl:with-param name="ReportDesc"><xsl:value-of select="//lx:Project/@name"/></xsl:with-param>
	</xsl:call-template>
	
	<xsl:apply-templates select="//lx:Alignment"/>
	
	</div>
</body>
</html>
</xsl:template>

<xsl:template match="lx:Alignment">
	<p/>
	<xsl:if test="./lx:Profile/lx:ProfAlign">

	<table width="75%">
		<tr>
			<td colspan="2"><u>Horizontal Alignment Information</u></td>
		</tr>
		<tr>
			<td>Name:</td>
			<td><xsl:value-of select="@name"/></td>
		</tr>
		<tr>
			<td>Station Range:</td>
			<xsl:variable name="start" select="@staStart"/>
			<xsl:variable name="AlignmentLength" select="@length"/>
<!--			<xsl:variable name="end" select="landUtils:GetVEndStation()"/>-->
			<xsl:variable name="end" select="landUtils:GetEndAlignmentStation(string($start), string($AlignmentLength))"/>
			<td><xsl:value-of select="landUtils:FormatStation(string($start), string($Profile.Station.Display), string($Profile.Station.precision), string($Profile.Station.rounding))"/> to <xsl:value-of select="landUtils:FormatStation(string($end), string($Profile.Station.Display), string($Profile.Station.precision), string($Profile.Station.rounding))"/></td>
		</tr>
	
	<!-- Load the station Equations -->
	<xsl:if test="./lx:StaEquation">	
		<tr bgcolor="#eeeeee">
			<th>Back Station</th>
			<th>Ahead Station</th>
		</tr>

		<xsl:for-each select="./lx:StaEquation">
		<tr>
			<td><xsl:value-of select="@staBack"/></td>
			<td><xsl:value-of select="@staAhead"/></td>
		</tr>
		</xsl:for-each>
	</xsl:if>
	
	</table>

    <!-- Loop over each profile -->
    <xsl:for-each select="lx:Profile">
			<xsl:for-each select="lx:ProfAlign">
				<xsl:apply-templates select="." mode="set"/>

				<!-- for every profile, check all the StaEquations and update stations-->
				<xsl:for-each select="../../lx:StaEquation">
					<xsl:variable name="loadStaEquation" select="landUtils:AddVStationEquation(@staInternal, @staBack, @staAhead)"/>
				</xsl:for-each>

				<!-- Apply the station equations to the alignment geometry -->
				<xsl:variable name="appstaeq" select="landUtils:ApplyVStationEquations()"/>
				
				<xsl:apply-templates select="."/>
			</xsl:for-each>
		</xsl:for-each>
	
	</xsl:if>
</xsl:template>

<xsl:template match="lx:ProfAlign">
	<h3>Vertical Alignment: <xsl:value-of select="@name"/></h3>
	<table border="1" width="100%" cellpadding="7" cellspacing="1" bgcolor="#eeeeeff">
		<tr align="right">
			<th>PVI</th>
			<th>Station</th>
			<xsl:choose>
				<xsl:when test="$Profile.Elevation.unit='default'">
					<xsl:choose>
						<xsl:when test="$SourceLinearUnit='foot'">
							<th>Elevation (ft)</th>
						</xsl:when>
						<xsl:when test="$SourceLinearUnit='meter'">
							<th>Elevation (m)</th>
						</xsl:when>
						<xsl:otherwise>
							<th>Elevation</th>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="$Profile.Elevation.unit='foot'">
					<th>Elevation (ft)</th>
				</xsl:when>
				<xsl:when test="$Profile.Elevation.unit='meter'">
					<th>Elevation (m)</th>
				</xsl:when>
				<xsl:otherwise>
					<th>Elevation</th>
				</xsl:otherwise>
			</xsl:choose>
			<th>Grade Out (%)</th>
			<xsl:choose>
				<xsl:when test="$Profile.Linear.unit='default'">
					<xsl:choose>
						<xsl:when test="$SourceLinearUnit='foot'">
							<th>Curve Length (ft)</th>
						</xsl:when>
						<xsl:when test="$SourceLinearUnit='meter'">
							<th>Curve Length (m)</th>
						</xsl:when>
						<xsl:otherwise>
							<th>Curve Length</th>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="$Profile.Linear.unit='foot'">
					<th>Curve Length (ft)</th>
				</xsl:when>
				<xsl:when test="$Profile.Linear.unit='meter'">
					<th>Curve Length (m)</th>
				</xsl:when>
				<xsl:otherwise>
					<th>Curve Length</th>
				</xsl:otherwise>
			</xsl:choose>
		</tr>
		<xsl:for-each select="./node()">
		<xsl:variable name="number" select="position()"/>
		<xsl:variable name="sta" select="landUtils:GetVStation($number)"/>
		<xsl:variable name="elev" select="landUtils:GetVElevation($number)"/>
		<xsl:variable name="grade" select="landUtils:GetGradeOut($number)"/>
		<xsl:variable name="length" select="landUtils:GetVCurveLength($number)"/>

		<xsl:choose >
			<xsl:when test="position() != last()">
				<tr align="right">
					<td><xsl:value-of select="$number"/></td>
					<td><xsl:value-of select="landUtils:FormatStation(string($sta), string($Profile.Station.Display), string($Profile.Station.precision), string($Profile.Station.rounding))"/></td>
					<td><xsl:value-of select="landUtils:FormatNumber(string($elev), string($SourceLinearUnit), string($Profile.Elevation.unit), string($Profile.Elevation.precision), string($Profile.Elevation.rounding))"/></td>
					<td><xsl:value-of select="landUtils:FormatNumber(string($grade), string($SourceLinearUnit), string($Profile.Linear.unit), string($Profile.Linear.precision), string($Profile.Linear.rounding))"/> %</td>
					<td><xsl:value-of select="landUtils:FormatNumber(string($length), string($SourceLinearUnit), string($Profile.Linear.unit), string($Profile.Linear.precision), string($Profile.Linear.rounding))"/></td>
				</tr>
			</xsl:when>
			<xsl:otherwise>
				<tr align="right">
					<td><xsl:value-of select="$number"/></td>
					<td><xsl:value-of select="landUtils:FormatStation(string($sta), string($Profile.Station.Display), string($Profile.Station.precision), string($Profile.Station.rounding))"/></td>
					<td><xsl:value-of select="landUtils:FormatNumber(string($elev), string($SourceLinearUnit), string($Profile.Elevation.unit), string($Profile.Elevation.precision), string($Profile.Elevation.rounding))"/></td>
					<td>&#160;</td>
					<td>&#160;</td>
				</tr>
			</xsl:otherwise>
		</xsl:choose>
		</xsl:for-each>
	</table>
	<hr size="2" color="red"/>
</xsl:template>

</xsl:stylesheet>
