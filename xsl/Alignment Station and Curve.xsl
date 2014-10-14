<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">

<!--Description:Station and Curve

This form provides mathematical and station design data for the curves, spiral and tangents contained in the selected horizontal alignment(s).

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

<xsl:template match="/">

<html>
<head>
	<title>Horizontal Alignment Station and Curve Report</title>
</head>
<body>
<div style="width:7in">

<xsl:call-template name="AutoHeader">
<xsl:with-param name="ReportTitle">Alignment Station and Curve Report</xsl:with-param>
<xsl:with-param name="ReportDesc"><xsl:value-of select="//lx:Project/@name"/></xsl:with-param>
</xsl:call-template>

<xsl:apply-templates select="//lx:Alignment"/>
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
<table width="100%">
	<xsl:for-each select="./node()">
		<xsl:variable name="pos" select="position()"/>
		<xsl:variable name="prev" select="position()-1"/>
		<xsl:variable name="next" select="position()+1"/>
		<xsl:variable name="startStation" select="landUtils:GetGeomStartingStation($pos)"/>
		<xsl:variable name="endStation" select="landUtils:GetGeomEndStation($pos)"/>
		<xsl:variable name="Bearing">Bearing</xsl:variable>
		<xsl:variable name="Conventional">Conventional</xsl:variable>

		<xsl:choose>
			<xsl:when test="name()='Line'">
				
				<tr>
					<td colspan="4" align="center"><hr color = "#ddddff" size="4"/><u>Tangent Data</u></td>
				</tr>
				<tr  bgcolor="eeeeee">
						<th>Description</th>
						<th>PT Station</th>
						<th>Northing</th>
						<th>Easting</th>
				</tr>	
				<tr>
					<td>Start:</td>					
					<td><xsl:value-of select="landUtils:FormatStation(string($startStation), string($Alignment.Station.Display), string($Alignment.Station.precision), string($Alignment.Station.rounding))"/></td>
					<xsl:variable name="startN" select="landUtils:GetLineStartNorthing($pos)"/>
					<xsl:variable name="startE" select="landUtils:GetLineStartEasting($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($startN), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string($Alignment.Coordinate.rounding))"/></td>
					<td><xsl:value-of select="landUtils:FormatNumber(string($startE), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string($Alignment.Coordinate.rounding))"/></td>				
				</tr>				
				<tr>
					<td>End:</td>					
					<td><xsl:value-of select="landUtils:FormatStation(string($endStation), string($Alignment.Station.Display), string($Alignment.Station.precision), string($Alignment.Station.rounding))"/></td>
					<xsl:variable name="endN" select="landUtils:GetLineEndNorthing($pos)"/>
					<xsl:variable name="endE" select="landUtils:GetLineEndEasting($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($endN), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string($Alignment.Coordinate.rounding))"/></td>
					<td><xsl:value-of select="landUtils:FormatNumber(string($endE), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string($Alignment.Coordinate.rounding))"/></td>				
				</tr>								
				<tr>
					<td colspan="4" align="center"><u>Tangent Data</u></td>
				</tr>			
				<tr  bgcolor="eeeeee">
					<th>Parameter</th>
					<th>Value</th>
					<th>Parameter</th>
					<th>Value</th>
				</tr>
				<tr>
					<td>Length:</td>					
					<xsl:variable name="len" select="landUtils:GetLineLength($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($len), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
					<td>Course:</td>					
					<xsl:variable name="angle" select="landUtils:GetLineDirection($pos)"/>
					<td><xsl:value-of select="landUtils:FormatDirection(string($angle), string($Conventional), string($Bearing), string($Alignment.Angular.precision), string($Alignment.Angular.rounding), string($Alignment.Angular.unit))"></xsl:value-of></td>
				</tr>				
			</xsl:when>

		<xsl:when test="name()='Curve'">		
				<tr>
					<td colspan="4" align="center"><hr color = "#ddddff" size="4"/><u>Curve Point Data</u></td>
				</tr>
				<tr  bgcolor="eeeeee">
					<th>Description</th>
					<th>Station </th>
					<th>Northing</th>
					<th>Easting</th>
				</tr>
				
				<xsl:variable name="PrevElement" select="landUtils:GetGeomElementType($prev)"/>
				<xsl:variable name="NextElement" select="landUtils:GetGeomElementType($next)"/>
									
				<tr>
					<xsl:choose>
						<xsl:when test="string($PrevElement) = 'Spiral'">
							<td>SC:</td>
						</xsl:when>
						<xsl:when test="string($PrevElement) = 'Curve'">
							<xsl:variable name="PrevCurveDirection" select="landUtils:ComparePreviousCurveDirection($pos)"/>
							<xsl:choose>
								<xsl:when test="string($PrevCurveDirection) = 'Compound'">
									<td>PCC:</td>
								</xsl:when>
								<xsl:otherwise>
									<td>PRC:</td>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:otherwise>
							<td>PC:</td>
						</xsl:otherwise>
					</xsl:choose>
					
					<td><xsl:value-of select="landUtils:FormatStation(string($startStation), string($Alignment.Station.Display), string($Alignment.Station.precision), string($Alignment.Station.rounding))"/></td>
					<xsl:variable name="sNorth" select="landUtils:GetCurveStartNorthing($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($sNorth), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string($Alignment.Coordinate.rounding))"/>	</td>
					<xsl:variable name="sEast" select="landUtils:GetCurveStartEasting($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($sEast), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string($Alignment.Coordinate.rounding))"/>	</td>
				</tr>
				<tr>
					<td>RP:</td>
					<td></td>
					<xsl:variable name="cNorth" select="landUtils:GetCurveCenterNorthing($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($cNorth), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string($Alignment.Coordinate.rounding))"/>	</td>
					<xsl:variable name="cEast" select="landUtils:GetCurveCenterEasting($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($cEast), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string($Alignment.Coordinate.rounding))"/>	</td>
				</tr>				
					
				<tr>
					<xsl:choose>
						<xsl:when test="string($NextElement) = 'Spiral'">
							<td>CS:</td>
						</xsl:when>
						<xsl:when test="string($NextElement) = 'Curve'">
							<xsl:variable name="NextCurveDirection" select="landUtils:CompareNextCurveDirection($pos)"/>
							<xsl:choose>
								<xsl:when test="string($NextCurveDirection) = 'Compound'">
									<td>PCC:</td>
								</xsl:when>
								<xsl:otherwise>
									<td>PRC:</td>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:otherwise>
							<td>PT:</td>
						</xsl:otherwise>
					</xsl:choose>
				
					<td><xsl:value-of select="landUtils:FormatStation(string($endStation), string($Alignment.Station.Display), string($Alignment.Station.precision), string($Alignment.Station.rounding))"/></td>
					<xsl:variable name="eNorth" select="landUtils:GetCurveEndNorthing($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($eNorth), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string($Alignment.Coordinate.rounding))"/>	</td>
					<xsl:variable name="eEast" select="landUtils:GetCurveEndEasting($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($eEast), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string($Alignment.Coordinate.rounding))"/>	</td>
				</tr>
				
				<tr>
					<td colspan="4" align="center"><u>Circular Curve Data</u></td>
				</tr>			
				<tr  bgcolor="eeeeee">
					<th>Parameter</th>
					<th>Value</th>
					<th>Parameter</th>
					<th>Value</th>
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
					<td><xsl:value-of select="landUtils:FormatDirection(string($chdir), string($Conventional), string($Bearing), string($Alignment.Angular.precision), string($Alignment.Angular.rounding), string($Alignment.Angular.unit))"></xsl:value-of></td>
				</tr>
			</xsl:when>
												
						
			<xsl:when test="name()='Spiral'">		
				<xsl:variable name="SpiralType" select="landUtils:GetSpiralType($pos)"/>	
						
				<tr>
					<td colspan="4" align="center"><hr color = "#ddddff" size="4"/><u>Spiral Point Data</u></td>
				</tr>
				<tr  bgcolor="eeeeee">
					<th>Description</th>
					<th>Station </th>
					<th>Northing</th>
					<th>Easting</th>
				</tr>
				
				<xsl:variable name="PrevElement" select="landUtils:GetGeomElementType($prev)"/>
				<xsl:variable name="NextElement" select="landUtils:GetGeomElementType($next)"/>
								
				<tr>
					<xsl:choose>
						<xsl:when test="string($PrevElement) = 'Curve'">
							<td>CS:</td>
						</xsl:when>
						<xsl:when test="string($PrevElement) = 'Spiral'">
							<td>SS:</td>
						</xsl:when>
						<xsl:otherwise>
							<td>TS:</td>
						</xsl:otherwise>
					</xsl:choose>							
				
					<td><xsl:value-of select="landUtils:FormatStation(string($startStation), string($Alignment.Station.Display), string($Alignment.Station.precision), string($Alignment.Station.rounding))"/></td>
					<xsl:variable name="ssn" select="landUtils:GetSpiralStartNorthing($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($ssn), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string	($Alignment.Coordinate.rounding))"/></td>
					<xsl:variable name="sse" select="landUtils:GetSpiralStartEasting($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($sse), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string	($Alignment.Coordinate.rounding))"/></td>
				</tr>
				<tr>
					<td>SPI:</td>
					<td></td>
					<xsl:variable name="spn" select="landUtils:GetSpiralPINorthing($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($spn), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string	($Alignment.Coordinate.rounding))"/></td>
					<xsl:variable name="spe" select="landUtils:GetSpiralPIEasting($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($spe), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string	($Alignment.Coordinate.rounding))"/></td>
				</tr>											
				<tr>
					<xsl:choose>
						<xsl:when test="string($NextElement) = 'Curve'">
							<td>SC:</td>
						</xsl:when>
						<xsl:when test="string($NextElement) = 'Spiral'">
							<td>SS:</td>
						</xsl:when>
						<xsl:otherwise>
							<td>ST:</td>
						</xsl:otherwise>
					</xsl:choose>
				
					<td><xsl:value-of select="landUtils:FormatStation(string($endStation), string($Alignment.Station.Display), string($Alignment.Station.precision), string($Alignment.Station.rounding))"/></td>
					<xsl:variable name="sen" select="landUtils:GetSpiralEndNorthing($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($sen), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string	($Alignment.Coordinate.rounding))"/></td>
					<xsl:variable name="see" select="landUtils:GetSpiralEndEasting($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($see), string($SourceLinearUnit), string($Alignment.Coordinate.unit), string($Alignment.Coordinate.precision), string	($Alignment.Coordinate.rounding))"/></td>
				</tr>
				
				<xsl:choose> 
					<xsl:when test="$SpiralType!='clothoid'">
						<tr>
							<td colspan="4" align="center"><u>Spiral Curve: <xsl:value-of select="landUtils:GetSpiralType($pos)"/></u></td>
						</tr>
						<tr>
							<td colspan="3">Values not reported for non-clothoid spirals</td>
						</tr>
					</xsl:when>
				<xsl:otherwise>
				
				<tr>
					<td colspan="4" align="center"><u>Spiral Curve Data:  <xsl:value-of select="landUtils:GetSpiralType($pos)"/></u></td>
				</tr>
				<tr  bgcolor="eeeeee">
					<th>Parameter</th>
					<th>Value</th>
					<th>Parameter</th>
					<th>Value</th>
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
					<td><xsl:value-of select="landUtils:FormatNumber(string($SpiralRad), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
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
					<xsl:variable name="X" select="landUtils:GetSpiralTotalX($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($X), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
					<td>K:</td>
					<xsl:variable name="K" select="landUtils:GetSpiralK($pos)"/>
					<td><xsl:value-of select="landUtils:FormatNumber(string($K), string($SourceLinearUnit), string($Alignment.Linear.unit), string($Alignment.Linear.precision), string($Alignment.Linear.rounding))"/></td>
				</tr>
				
				<tr>
					<td>Y:</td>
					<xsl:variable name="Y" select="landUtils:GetSpiralTotalY($pos)"/>
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
					<xsl:choose>
						<xsl:when test="string($chdir) = 'NaN'">
							<td>0</td>
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