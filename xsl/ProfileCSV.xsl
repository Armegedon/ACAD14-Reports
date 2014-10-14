<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">

<!--Description:Profile CSV report.
This form provides PVI station, elevation, and other design information for profile alignments in your LandXML data in a comma delimited format, suitable for spreadsheet import.

This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
<!--CreatedBy:Autodesk Inc. -->
<!--DateCreated:09/06/2002 -->
<!--LastModifiedBy:Autodesk Inc. -->
<!--DateModified:09/30/2002 -->
<!--OutputExtension:csv -->
<xsl:output method="text" media-type="iso-8859-1" encoding="us-ascii"/>

<!-- ==== JavaScript Includes ==== -->
<xsl:include href="header.xsl"/>
<xsl:include href="V_Alignment_Layout_b.xsl"/>
<xsl:include href="Number_Formatting.xsl"/>
<!-- ============================= -->

<xsl:param name="SourceLinearUnit" select="//lx:Units/*/@linearUnit"/>

<xsl:template match="/">
	<!-- make the header with unit labels -->
	<xsl:text>"PVI",</xsl:text>
	<xsl:choose>
		<xsl:when test="$Profile.Linear.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceLinearUnit='foot'">
					<xsl:text>"station (ft)",</xsl:text>
				</xsl:when>
				<xsl:when test="$SourceLinearUnit='meter'">
					<xsl:text>"station (m)",</xsl:text>
				</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Profile.Linear.unit='foot'">
			<xsl:text>"station (ft)",</xsl:text>
		</xsl:when>
		<xsl:when test="$Profile.Linear.unit='meter'">
			<xsl:text>"station (m)",</xsl:text>
		</xsl:when>
	</xsl:choose>
	<xsl:choose>
		<xsl:when test="$Profile.Elevation.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceLinearUnit='foot'">
					<xsl:text>"elevation (ft)",</xsl:text>
				</xsl:when>
				<xsl:when test="$SourceLinearUnit='meter'">
					<xsl:text>"elevation (m)",</xsl:text>
				</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Profile.Elevation.unit='foot'">
			<xsl:text>"elevation (ft)",</xsl:text>
		</xsl:when>
		<xsl:when test="$Profile.Elevation.unit='meter'">
			<xsl:text>"elevation (m)",</xsl:text>
		</xsl:when>
	</xsl:choose>
	<xsl:text>"grade out (%)",</xsl:text>
	<xsl:choose>
		<xsl:when test="$Profile.Linear.unit='default'">
			<xsl:choose>
				<xsl:when test="$SourceLinearUnit='foot'">
					<xsl:text>"curve length (ft)"</xsl:text>
				</xsl:when>
				<xsl:when test="$SourceLinearUnit='meter'">
					<xsl:text>"curve length (m)"</xsl:text>
				</xsl:when>
			</xsl:choose>
		</xsl:when>
		<xsl:when test="$Profile.Linear.unit='foot'">
			<xsl:text>"curve length (ft)"</xsl:text>
		</xsl:when>
		<xsl:when test="$Profile.Linear.unit='meter'">
			<xsl:text>"curve length (m)"</xsl:text>
		</xsl:when>
	</xsl:choose>
	<xsl:text>&#xa;</xsl:text>
	<xsl:apply-templates select="//lx:Alignment"/>
</xsl:template>

<xsl:template match="lx:Alignment">
  <xsl:for-each select="./lx:Profile/lx:ProfAlign">
    <!-- Put vertical alignment in memory -->
    <xsl:apply-templates select="." mode="set"/>
    <xsl:for-each select="../../lx:StaEquation">
      <xsl:variable name="loadStaEquation" select="landUtils:AddVStationEquation(@staInternal, @staBack, @staAhead)"/>
    </xsl:for-each>

    <!-- Apply the station equations to the alignment geometry -->
    <xsl:variable name="appstaeq" select="landUtils:ApplyVStationEquations()"/>
    <xsl:apply-templates select="."/>
  </xsl:for-each>
</xsl:template>

<xsl:template match="lx:ProfAlign">
	
	<xsl:text/>"Vertical Alignment: <xsl:value-of select="../@name"/> - <xsl:value-of select="@name"/>"&#xa;<xsl:text/>
		<xsl:for-each select="./node()">
			<xsl:variable name="number" select="position()"/>
			<xsl:variable name="sta" select="landUtils:GetVStation($number)"/>
			<xsl:variable name="elev" select="landUtils:GetVElevation($number)"/>
			<xsl:variable name="grade" select="landUtils:GetGradeOut($number)"/>
			<xsl:variable name="length" select="landUtils:GetVCurveLength($number)"/>
			<xsl:text/>"<xsl:value-of select="$number"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($sta), string($SourceLinearUnit), string($Profile.Linear.unit), string($Profile.Linear.precision), string($Profile.Linear.rounding))"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($elev), string($SourceLinearUnit), string($Profile.Elevation.unit), string($Profile.Elevation.precision), string($Profile.Elevation.rounding))"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($grade), 'default', string($Profile.Linear.unit), string($Profile.Linear.precision), string($Profile.Linear.rounding))"/>",<xsl:text/>
			<xsl:text/>"<xsl:value-of select="landUtils:FormatNumber(string($length), string($SourceLinearUnit), string($Profile.Linear.unit), string($Profile.Linear.precision), string($Profile.Linear.rounding))"/>"<xsl:text/>
			<xsl:text>&#xa;</xsl:text>
		</xsl:for-each>
</xsl:template>

</xsl:stylesheet>