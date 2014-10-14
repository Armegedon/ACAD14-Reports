<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">

	<!--Description:Alignment Curve
	
This form provides mathematical design data for the curves, spirals and tangents contained in the selected horizontal alignment(s).  

This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.
Spiral data in this report is accurate for clothoid spirals only.-->
<!--CreatedBy:Autodesk Inc. -->
<!--DateCreated:06/15/2002 -->
<!--LastModifiedBy:Autodesk Inc. -->
<!--DateModified:09/16/2002 -->
<!--OutputExtension:html -->

<xsl:output method="html" encoding="UTF-8"/>

<!-- =========== JavaScript Include==== -->
<xsl:include href="header.xsl"/>
<xsl:include href="H_Align_Layout_b.xsl"/>
<xsl:include href="Number_Formatting.xsl"/>


<!-- ================================= -->
<xsl:param name="SourceLinearUnit" select="//lx:Units/*/@linearUnit"/>
<!-- ================================= -->
<xsl:template match="/">
<!-- ================================= -->

<html>
<head>
	<title>Horizontal Alignment Station and Curve Report</title>
</head>
<body>
<div style="width:7in">

<!-- This places an automatic header in your output -->
<xsl:call-template name="AutoHeader">
<xsl:with-param name="ReportTitle">Alignment Curve Report</xsl:with-param>
<xsl:with-param name="ReportDesc"><xsl:value-of select="//lx:Project/@name"/></xsl:with-param>
</xsl:call-template>

<xsl:apply-templates select="//lx:Alignment"/>
<!--	<xsl:apply-templates select="."/>-->
</div>
</body>
</html>
</xsl:template>

<xsl:template match="lx:Alignment">
	<h3>Alignment: <xsl:value-of select="@name"/></h3>
	<h3>Description: <xsl:value-of select="@desc"/></h3>
	<hr color = "black"/>

	<!-- Load and set stationing for the alignment geometry -->
	<!-- Loads the geometry -->
	<xsl:apply-templates select="./lx:CoordGeom" mode="set"></xsl:apply-templates>
	
	<!-- Sets the internal stationing of the geometry -->
 	<xsl:variable name="AppStations" select="landUtils:ApplyStationing(@staStart)"/>
	
	<!-- Load the station Equations -->
	<xsl:if test="./lx:StaEquation">	
		<h3>Station Equations</h3>
		<table border="2" width="100%">
			<tbody>
				<tr bgcolor="#eeeeee">
					<th>Internal Station</th>
					<th>Back Station</th>
					<th>Ahead Station</th>
				</tr>

			<xsl:for-each select="./lx:StaEquation">
				<tr>
					<td><xsl:value-of select="@staInternal"/></td>
					<td><xsl:value-of select="@staBack"/></td>
					<td><xsl:value-of select="@staAhead"/></td>
				</tr>
				<xsl:variable name="loadStaEquation" select="landUtils:AddStationEquation(string(@staInternal), string(@staBack), string(@staAhead))"/>	
			</xsl:for-each>
		</tbody>
	</table>
	<p></p>
	</xsl:if>
	
	<!-- Apply the station equations to the alignment geometry -->
 	<xsl:variable name="appstaeq" select="landUtils:ApplyStationEquations()"/>
	
	<!-- Begin iteration of the geometry -->
	<xsl:apply-templates select="./lx:CoordGeom"></xsl:apply-templates> 
</xsl:template>

<xsl:template match="lx:CoordGeom">
	<table width="100%" >
		<xsl:for-each select="./node()">
			<xsl:variable name="pos" select="position()"/>

		<xsl:choose>
			<xsl:when test="name()='Line'">

				<tr>
					<td colspan="4" align="center"><hr/></td>
				</tr>
				<tr>
					<td colspan="3" align="center"><u>Tangent Data</u></td>
				</tr>
				<tr>
					<td>Length:</td>
					<xsl:variable name="len" select="landUtils:GetLineLength($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($len), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
					<td>Course:</td>
					<xsl:variable name="angle" select="landUtils:GetLineDirection($pos)"/>
					<xsl:variable name="Bearing">Bearing</xsl:variable>
					<xsl:variable name="Conventional">Conventional</xsl:variable>
					<td><xsl:value-of select="landUtils:FormatDirection(string($angle), string($Conventional), string($Bearing), string($Alignment.Angular.precision), string($Alignment.Angular.rounding), string($Alignment.Angular.unit))"></xsl:value-of></td>
				</tr>
			</xsl:when>

			<xsl:when test="name()='Curve'">
				<tr>
					<td colspan="4" align="center"><hr/></td>
				</tr>
				<tr>
					<td colspan="3" align="center"><u>Circular Curve Data</u></td>
				</tr>
				<tr>
					<td>Delta:</td>
					<xsl:variable name="CurveAngle" select="landUtils:GetCurveAngle($pos)"/>
					<td><xsl:value-of select="landUtils:FormatDMS(string($CurveAngle),  string($Alignment.Angular.precision), string($Alignment.Angular.rounding))"/></td>
					<td>Type:</td>
					<xsl:variable name="rotation" select="landUtils:GetCurveRotation($pos)"/>
					<xsl:choose>
						<xsl:when test="string($rotation) = 'ccw'">
							<td>LEFT</td>
						</xsl:when>
						<xsl:otherwise>
							<td>RIGHT</td>
						</xsl:otherwise>
					</xsl:choose>
				</tr>
				<tr>
					<td>Radius:</td>
					<xsl:variable name="rad" select="landUtils:GetCurveRadius($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($rad), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>

					<!--Don't report DOC if units are metric-->
					<xsl:if test="$Alignment.Linear.unit = 'default'"> <!--If Linear units are set to default...-->
						<xsl:if test="$SourceLinearUnit = 'foot'">      <!--...and the XML file is imperial, then display the DOC-->
							<td>DOC:</td>
							<xsl:variable name="DOC" select="landUtils:GetCurveDOC($pos)"/>
							<td><xsl:value-of select="landUtils:FormatDMS(string($DOC), string($Alignment.Angular.precision), string($Alignment.Angular.rounding))"/></td>
						</xsl:if>
					</xsl:if>
					<xsl:if test="$Alignment.Linear.unit = 'foot'">   <!--If Linear units are set to foot, always display the DOC-->
						<td>DOC:</td>
						<xsl:variable name="DOC" select="landUtils:GetCurveDOC($pos)"/>
						<td><xsl:value-of select="landUtils:FormatDMS(string($DOC), string($Alignment.Angular.precision), string($Alignment.Angular.rounding))"/></td>
					</xsl:if>
				</tr>
				<tr>
					<td>Length:</td>
					<xsl:variable name="CurveLen" select="landUtils:GetCurveLength($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($CurveLen), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
					<td>Tangent:</td>
					<xsl:variable name="tan" select="landUtils:GetCurveTangent($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($tan), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
				</tr>
				<tr>
					<td>Mid-Ord:</td>
					<xsl:variable name="ord" select="landUtils:GetCurveMiddle($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($ord), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
					<td>External:</td>
					<xsl:variable name="ext" select="landUtils:GetCurveExternal($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($ext), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
				</tr>
				<tr>
					<td>Chord:</td>
					<xsl:variable name="chord" select="landUtils:GetChordLength($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($chord), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
					<td>Course:</td>
					<xsl:variable name="chdir" select="landUtils:GetChordDirection($pos)"/>
					<xsl:variable name="Bearing">Bearing</xsl:variable>
					<xsl:variable name="Conventional">Conventional</xsl:variable>
					<td><xsl:value-of select="landUtils:FormatDirection(string($chdir), string($Conventional), string($Bearing), string($Alignment.Angular.precision), string($Alignment.Angular.rounding), string($Alignment.Angular.unit))"></xsl:value-of></td>
			
				</tr>
			</xsl:when>


			<xsl:when test="name()='Spiral'">

				<tr>
					<td colspan="4" align="center"><hr/></td>
				</tr>
				<xsl:variable name="SpiralType" select="landUtils:GetSpiralType($pos)"/>
				<xsl:choose>
					<xsl:when test="$SpiralType!='clothoid'">
						<tr>
							<td colspan="3" align="center"><u>Spiral Curve: <xsl:value-of select="landUtils:GetSpiralType($pos)"/></u></td>
						</tr>
						<tr>
							<td colspan="3">Values not reported for non-clothoid spirals</td>
						</tr>
					</xsl:when>
				<xsl:otherwise>
				<tr>
					<td colspan="3" align="center"><u>Spiral Curve Data: <xsl:value-of select="landUtils:GetSpiralType($pos)"/></u></td>
				</tr>
				
				<tr>
					<td>Length:</td>
					<xsl:variable name="SpiralLen" select="landUtils:GetSpiralLength($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($SpiralLen), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
					<td>L Tan:</td>
					<xsl:variable name="lTan" select="landUtils:GetSpiralLongTangent($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($lTan), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
				</tr>
				
				<tr>
					<td>Radius:</td>
					<xsl:variable name="SpiralRad" select="landUtils:GetSpiralRadius($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($SpiralRad), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/>	</td>
					<td>S Tan:</td>
					<xsl:variable name="sTan" select="landUtils:GetSpiralShortTangent($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($sTan), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
				</tr>
				
				<tr>
					<td>Theta:</td>
					<xsl:variable name="theta" select="landUtils:GetSpiralTheta_Degrees($pos)"/>
					<td><xsl:value-of select="landUtils:FormatDMS(string($theta), string($Alignment.Angular.precision), string($Alignment.Angular.rounding))"/></td>
					<td>P:</td>
					<xsl:variable name="P" select="landUtils:GetSpiralP($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($P), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
				</tr>
	
				<tr>
					<td>X:</td>
					<xsl:variable name="X" select="landUtils:GetSpiralX($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($X), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
					<td>K:</td>
					<xsl:variable name="K" select="landUtils:GetSpiralK($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($K), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
				</tr>
				
				<tr>
					<td>Y:</td>
					<xsl:variable name="Y" select="landUtils:GetSpiralY($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($Y), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
					<td>A:</td>
					<xsl:variable name="A" select="landUtils:GetSpiralA($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($A), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
				</tr>
				
				<tr>
					<td>Chord:</td>
					<xsl:variable name="chord" select="landUtils:GetSpiralChordLength($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($chord), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
					<td>Course:</td>
					<xsl:variable name="chdir" select="landUtils:GetSpiralDirection($pos)"/>
					<xsl:variable name="Bearing">Bearing</xsl:variable>
					<xsl:variable name="Conventional">Conventional</xsl:variable>
					<xsl:choose>
						<xsl:when test="string($chdir) = 'NaN'">
							<td></td>
						</xsl:when>
						<xsl:otherwise>
							<td><xsl:value-of select="landUtils:FormatDirection(string($chdir), string($Conventional), string($Bearing), string($Alignment.Angular.precision), string($Alignment.Angular.rounding), string($Alignment.Angular.unit))"></xsl:value-of></td>
						</xsl:otherwise>
					</xsl:choose>
				</tr>
				
				</xsl:otherwise>
				</xsl:choose> <!--End test for spiral type-->
			</xsl:when>	
			<xsl:otherwise><tr></tr></xsl:otherwise>
		</xsl:choose> <!--End test for element type-->
	
	</xsl:for-each>
	</table>
	<hr color = "black"/>
</xsl:template>


</xsl:stylesheet>
