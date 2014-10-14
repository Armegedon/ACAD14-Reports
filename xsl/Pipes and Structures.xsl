<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0" xmlns:lx="http://www.landxml.org/schema/LandXML-1.2" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">
	<!--Description:Tabulation of all pipes and structures in the selected networks.
This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
	<!--CreatedBy:Autodesk Inc. -->
	<!--DateCreated:12/13/2004 -->
	<!--LastModifiedBy:Autodesk Inc. -->
	<!--DateModified:03/07/2005 -->
	<!--OutputExtension:html -->
	<xsl:output method="html" encoding="UTF-8" />
	<!-- ==== JavaScript Includes ==== -->
	<xsl:include href="header.xsl" />
	<xsl:include href="Pipework_Layout.xsl" />
	<xsl:param name="SourceLinearUnit" select="//lx:Units/*/@linearUnit" />
	<xsl:param name="DisplaySourceLinearUnit" select="landUtils:formatUnit(string($SourceLinearUnit))" />
	<xsl:param name="PipeLengthUnit" select="landUtils:getPipeLengthUnit(string($SourceLinearUnit),string($Pipe_Reports.Linear.unit))" />
	<xsl:param name="DisplayPipeLengthUnit" select="landUtils:formatUnit(string($PipeLengthUnit))" />
	<xsl:param name="DisplayReportsElevationUnit" select="landUtils:formatUnit(string($Pipe_Reports.Elevation.unit))" />
	<xsl:param name="DisplayReportsCoordinateUnit" select="landUtils:formatUnit(string($Pipe_Reports.Coordinate.unit))" />
	<xsl:param name="DiameterUnit">
		<xsl:choose>
			<xsl:when test="//lx:Units/*/@diameterUnit">
				<!--if there is diameterUnit, use it as pipe diameter dimension unit-->
				<xsl:value-of select="//lx:Units/*/@diameterUnit" />
			</xsl:when>
			<xsl:otherwise>
				<!--otherwise, use linearUnit-->
				<xsl:value-of select="//lx:Units/*/@linearUnit" />
			</xsl:otherwise>
		</xsl:choose>
	</xsl:param>
	<xsl:param name="PipeSizeUnit">
		<xsl:choose>
			<xsl:when test="$DisplayPipeLengthUnit = 'meter'">
				<xsl:value-of select="string('milimeter')"/>
			</xsl:when>					
			<xsl:when test="$DisplayPipeLengthUnit = 'foot'">
				<xsl:value-of select="string('inch')"/>
			</xsl:when>
		</xsl:choose>	
	</xsl:param>
	<!-- ============================= -->
	<xsl:template match="/">
		<html>
			<head>
				<title>Pipes and Structures report for <xsl:value-of select="//lx:Project/@name" /></title>
				<style type="text/css">
					div{
					width:10.5in;
					font-family:Verdana;
					}
					th{
					font-size:12pt;
					text-align:center;
					text-decoration:underline;
					}
					td{
					padding:0.02in 0.15in;
					text-align:right;
					}
					tr{
					font-size:10pt;
					}
				</style>
			</head>
			<body>
				<div>
					<xsl:call-template name="AutoHeader">
						<xsl:with-param name="ReportTitle">Pipes and Structures Report</xsl:with-param>
						<xsl:with-param name="ReportDesc">
							<xsl:value-of select="//lx:Project/@name" />
						</xsl:with-param>
					</xsl:call-template>
					<br />
					<xsl:for-each select="//lx:PipeNetworks/lx:PipeNetwork">					
					<b>Pipe Network: <xsl:value-of select="@name" /></b>
					<br />Pipes
					<table bordercolor="black" border="1" cellspacing="0" width="100%">
							<tr>
								<th>Name</th>
								<th>Shape</th>
								<xsl:choose>
									<xsl:when test="$PipeSizeUnit='milimeter'">
										<th>Size (mm)</th>
									</xsl:when>
									<xsl:when test="$PipeSizeUnit='inch'">
										<th>Size (in)</th>
									</xsl:when>
								</xsl:choose>
								
								<th>Material</th>
								<th>US Node</th>
								<th>DS Node</th>
								<xsl:choose>
									<xsl:when test="$Pipe_Reports.Elevation.unit='default'">
										<xsl:choose>
											<xsl:when test="$DisplaySourceLinearUnit='foot'">
												<th>US Invert (ft)</th>
												<th>DS Invert (ft)</th>
											</xsl:when>
											<xsl:when test="$DisplaySourceLinearUnit='meter'">
												<th>US Invert (m)</th>
												<th>DS Invert (m)</th>
											</xsl:when>
										</xsl:choose>
									</xsl:when>
									<xsl:when test="$DisplayReportsElevationUnit='foot'">
										<th>US Invert (ft)</th>
										<th>DS Invert (ft)</th>
									</xsl:when>
									<xsl:when test="$DisplayReportsElevationUnit='meter'">
										<th>US Invert (m)</th>
										<th>DS Invert (m)</th>
									</xsl:when>
								</xsl:choose>
								<xsl:choose>
									<xsl:when test="$Pipe_Reports.Pipe_Length.type='2-D'">
										<xsl:choose>
											<xsl:when test="$DisplayPipeLengthUnit='foot'">
												<th>2D Length (ft)
										<br>center-to-center</br>											
										<br>edge-to-edge</br>											
										</th>
											</xsl:when>
											<xsl:when test="$DisplayPipeLengthUnit='meter'">
												<th>2D Length (m)
										<br>center-to-center</br>											
										<br>edge-to-edge</br>											
										</th>
											</xsl:when>
										</xsl:choose>
									</xsl:when>
									<xsl:when test="$Pipe_Reports.Pipe_Length.type='3-D'">
										<xsl:choose>
											<xsl:when test="$DisplayPipeLengthUnit='foot'">
												<th>3D Length (ft)
										<br>center-to-center</br>											
										<br>edge-to-edge</br>											
										</th>
											</xsl:when>
											<xsl:when test="$DisplayPipeLengthUnit='meter'">
												<th>3D Length (m)
										<br>center-to-center</br>											
										<br>edge-to-edge</br>											
										</th>
											</xsl:when>
										</xsl:choose>
									</xsl:when>
								</xsl:choose>
								<th>% Slope</th>
							</tr>
							<xsl:for-each select="./lx:Pipes/lx:Pipe">
								<tr>
									<xsl:variable name="PipeType" select="landUtils:getPipeType(.)" />
									<xsl:variable name="name" select="@name" />
									<xsl:variable name="start" select="@refStart" />
									<xsl:variable name="end" select="@refEnd" />
									<!--Define the following two variables to convert pipe length from center-to-center to edge-to-edge-->
									<xsl:variable name="startStructDimension">
										<xsl:choose>
											<xsl:when test="./parent::*/parent::*/lx:Structs/lx:Struct[@name=$start]/lx:CircStruct/@diameter">
												<xsl:value-of select="landUtils:FormatNumber( string(./parent::*/parent::*/lx:Structs/lx:Struct[@name=$start]/lx:CircStruct/@diameter), string($DiameterUnit), string($SourceLinearUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
											</xsl:when>
											<xsl:when test="./parent::*/parent::*/lx:Structs/lx:Struct[@name=$start]/lx:RectStruct/@width">
												<xsl:value-of select="landUtils:FormatNumber( string(./parent::*/parent::*/lx:Structs/lx:Struct[@name=$start]/lx:RectStruct/@width), string($SourceLinearUnit), string($SourceLinearUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
											</xsl:when>
											<xsl:otherwise><xsl:value-of select="string('0')"/></xsl:otherwise>
										</xsl:choose>
									</xsl:variable>
									<xsl:variable name="endStructDimension">
										<xsl:choose>
											<xsl:when test="./parent::*/parent::*/lx:Structs/lx:Struct[@name=$end]/lx:CircStruct/@diameter">
												<xsl:value-of select="landUtils:FormatNumber( string(./parent::*/parent::*/lx:Structs/lx:Struct[@name=$end]/lx:CircStruct/@diameter), string($DiameterUnit), string($SourceLinearUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
											</xsl:when>
											<xsl:when test="./parent::*/parent::*/lx:Structs/lx:Struct[@name=$end]/lx:RectStruct/@width">
												<xsl:value-of select="landUtils:FormatNumber( string(./parent::*/parent::*/lx:Structs/lx:Struct[@name=$end]/lx:RectStruct/@width), string($SourceLinearUnit), string($SourceLinearUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
											</xsl:when>
											<xsl:otherwise><xsl:value-of select="string('0')"/></xsl:otherwise>
										</xsl:choose>
									</xsl:variable>
									<xsl:variable name="DeltaFromCC2EE" select="landUtils:setDelta(string($startStructDimension), string($endStructDimension))" />
									<xsl:variable name="US_Invert" select="./parent::*/parent::*/lx:Structs/lx:Struct[@name=$start]/lx:Invert[@refPipe=$name]/@elev" />
									<xsl:variable name="DS_Invert" select="./parent::*/parent::*/lx:Structs/lx:Struct[@name=$end]/lx:Invert[@refPipe=$name]/@elev" />
									<xsl:variable name="center1">
										<xsl:choose>
											<xsl:when test="./parent::*/parent::*/lx:Structs/lx:Struct[@name=$start]/lx:Center/@pntRef">
												<!--if there are pntRef, then use it-->
												<xsl:variable name="pntRef" select="./parent::*/parent::*/lx:Structs/lx:Struct[@name=$start]/lx:Center/@pntRef" />
												<xsl:copy-of select="//lx:CgPoints/lx:CgPoint[@name=$pntRef]" />
											</xsl:when>
											<xsl:otherwise>
												<!--othersize, use the value of node ./lx:Center-->
												<xsl:copy-of select="./parent::*/parent::*/lx:Structs/lx:Struct[@name=$start]/lx:Center" />
											</xsl:otherwise>
										</xsl:choose>
									</xsl:variable>
									<xsl:variable name="center2">
										<xsl:choose>
											<xsl:when test="./parent::*/parent::*/lx:Structs/lx:Struct[@name=$end]/lx:Center/@pntRef">
												<xsl:variable name="pntRef" select="./parent::*/parent::*/lx:Structs/lx:Struct[@name=$end]/lx:Center/@pntRef" />
												<xsl:copy-of select="//lx:CgPoints/lx:CgPoint[@name=$pntRef]" />
											</xsl:when>
											<xsl:otherwise>
												<xsl:copy-of select="./parent::*/parent::*/lx:Structs/lx:Struct[@name=$end]/lx:Center" />
											</xsl:otherwise>
										</xsl:choose>
									</xsl:variable>
									<td>
										<xsl:value-of select="$name" />
									</td>
									<td>
										<xsl:value-of select="$PipeType" />
									</td>
									<td>
										<xsl:choose>
											<!--Circular Pipe-->
											<xsl:when test="$PipeType='Circular'">
												D:<xsl:value-of select="landUtils:FormatNumber( string(./lx:CircPipe/@diameter),string($DiameterUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
											</xsl:when>
											<!--Elliptical Pipe-->
											<xsl:when test="$PipeType='Elliptical'">
												H:<xsl:value-of select="landUtils:FormatNumber( string(./lx:ElliPipe/@height), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
												<br />
												S:<xsl:value-of select="landUtils:FormatNumber( string(./lx:ElliPipe/@span), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
											</xsl:when>
                      <!--Egg Pipe-->
                      <xsl:when test="$PipeType='Egg'">
                        H:<xsl:value-of select="landUtils:FormatNumber( string(./lx:EggPipe/@height), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
                        <br />
                        S:<xsl:value-of select="landUtils:FormatNumber( string(./lx:EggPipe/@span), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
                      </xsl:when>
											<!--Rectangular Pipe-->
											<xsl:when test="$PipeType='Rectangular'">
												H:<xsl:value-of select="landUtils:FormatNumber( string(./lx:RectPipe/@height), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
												<br />
												W:<xsl:value-of select="landUtils:FormatNumber( string(./lx:RectPipe/@width), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
											</xsl:when>
											<!--Channel Pipe-->
											<xsl:when test="$PipeType='Channel'">
												H:<xsl:value-of select="landUtils:FormatNumber( string(./lx:Channel/@height), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
												<br />
												WT:<xsl:value-of select="landUtils:FormatNumber( string(./lx:Channel/@widthTop), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
												<br />
												WB:<xsl:value-of select="landUtils:FormatNumber( string(./lx:Channel/@widthBottom), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
											</xsl:when>
											<xsl:otherwise>...</xsl:otherwise>
										</xsl:choose>
									</td>
									
									
									<td>
										<xsl:choose>
											<xsl:when test="./*/@material"><xsl:value-of select="./*/@material" /></xsl:when>
											<xsl:otherwise>...</xsl:otherwise>
										</xsl:choose>
									</td>
									<td>
										<xsl:value-of select="$start" />
									</td>
									<td>
										<xsl:value-of select="$end" />
									</td>
									<td>
										<xsl:value-of select="landUtils:FormatNumber(string($US_Invert), string($SourceLinearUnit), string($Pipe_Reports.Elevation.unit), string($Pipe_Reports.Elevation.precision), string($Pipe_Reports.Elevation.rounding))" />
									</td>
									<td>
										<xsl:value-of select="landUtils:FormatNumber(string($DS_Invert), string($SourceLinearUnit), string($Pipe_Reports.Elevation.unit), string($Pipe_Reports.Elevation.precision), string($Pipe_Reports.Elevation.rounding))" />
									</td>
									<xsl:choose>
										<xsl:when test="$Pipe_Reports.Pipe_Length.type='2-D'">
											<td>
												<!--SourceLinearUnit format-->
												<xsl:variable name="TwoDLength" select="landUtils:get2DLength(string($center1), string($center2))" />
												<xsl:value-of select="landUtils:FormatNumber( string($TwoDLength), string($SourceLinearUnit), string($PipeLengthUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
												<br />
												<xsl:value-of select="landUtils:FormatNumber( string(landUtils:CC2EE(string($TwoDLength))),string($SourceLinearUnit),string($PipeLengthUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
												<br />
											</td>
										</xsl:when>
										<xsl:when test="$Pipe_Reports.Pipe_Length.type='3-D'">
											<td>
												<xsl:variable name="ThreeDLength" select="landUtils:get3DLength(string($center1), string($US_Invert), string($center2), string($DS_Invert))" />
												<xsl:value-of select="landUtils:FormatNumber(string($ThreeDLength), string($SourceLinearUnit), string($PipeLengthUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
												<br />
												<xsl:value-of select="landUtils:FormatNumber(string(landUtils:CC2EE(string($ThreeDLength))), string($SourceLinearUnit), string($PipeLengthUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
											</td>
										</xsl:when>
									</xsl:choose>
									<td>
										<xsl:value-of select="landUtils:FormatNumber(string(landUtils:getSlope(string($center1), string($US_Invert), string($center2), string($DS_Invert))), string($Pipe_Reports.Angular.unit), string($Pipe_Reports.Angular.unit), string($Pipe_Reports.Angular.precision), string($Pipe_Reports.Angular.rounding))" />
									</td>
								</tr>
							</xsl:for-each>
						</table>
					<br />					
					Structures
					<table bordercolor="black" border="1" cellspacing="0" width="100%">
							<tr>
								<th>Structure</th>
								<th>Type</th>
								<xsl:choose>
									<xsl:when test="$DisplayPipeLengthUnit='foot'">
										<th>Size (ft)</th>
									</xsl:when>
									<xsl:when test="$DisplayPipeLengthUnit='meter'">
										<th>Size (m)</th>
									</xsl:when>
								</xsl:choose>
								<th>Material</th>
								<xsl:choose>
									<xsl:when test="$Pipe_Reports.Coordinate.unit='default'">
										<xsl:choose>
											<xsl:when test="$DisplaySourceLinearUnit='foot'">
												<th>Northing (ft)</th>
												<th>Easting (ft)</th>
											</xsl:when>
											<xsl:when test="$DisplaySourceLinearUnit='meter'">
												<th>Northing (m)</th>
												<th>Easting (m)</th>
											</xsl:when>
										</xsl:choose>
									</xsl:when>
									<xsl:when test="$DisplayReportsCoordinateUnit='foot'">
										<th>Northing (ft)</th>
										<th>Easting (ft)</th>
									</xsl:when>
									<xsl:when test="$DisplayReportsCoordinateUnit='meter'">
										<th>Northing (m)</th>
										<th>Easting (m)</th>
									</xsl:when>
								</xsl:choose>
								<xsl:choose>
									<xsl:when test="$Pipe_Reports.Elevation.unit='default'">
										<xsl:choose>
											<xsl:when test="$DisplaySourceLinearUnit='foot'">
												<th>Rim Elev (ft)</th>
												<th>Sump Elev (ft)</th>
											</xsl:when>
											<xsl:when test="$DisplaySourceLinearUnit='meter'">
												<th>Rim Elev (m)</th>
												<th>Sump Elev (m)</th>
											</xsl:when>
										</xsl:choose>
									</xsl:when>
									<xsl:when test="$DisplayReportsElevationUnit='foot'">
										<th>Rim Elev (ft)</th>
										<th>Sump Elev (ft)</th>
									</xsl:when>
									<xsl:when test="$DisplayReportsElevationUnit='meter'">
										<th>Rim Elev (m)</th>
										<th>Sump Elev (m)</th>
									</xsl:when>
								</xsl:choose>
								<xsl:choose>
									<xsl:when test="$DisplayPipeLengthUnit='foot'">
										<th>Sump Depth (ft)</th>
									</xsl:when>
									<xsl:when test="$DisplayPipeLengthUnit='meter'">
										<th>Sump Depth (m)</th>
									</xsl:when>
								</xsl:choose>
								<th>Pipes</th>
							</tr>
							<xsl:for-each select="./lx:Structs/lx:Struct">
								<tr>
									<xsl:variable name="name" select="@name" />
									<xsl:variable name="CoordinateValue">
										<xsl:choose>
											<xsl:when test="./lx:Center/@pntRef">
												<xsl:variable name="pntRef" select="./lx:Center/@pntRef" />
												<xsl:copy-of select="//lx:CgPoints/lx:CgPoint[@name=$pntRef]" />
											</xsl:when>
											<xsl:otherwise>
												<xsl:copy-of select="./lx:Center" />
											</xsl:otherwise>
										</xsl:choose>
									</xsl:variable>
									<xsl:variable name="Northing" select="landUtils:getNorthing(string($CoordinateValue))" />
									<xsl:variable name="Easting" select="landUtils:getEasting(string($CoordinateValue))" />
									<xsl:variable name="elevRim" select="@elevRim" />
									<xsl:variable name="elevSump" select="@elevSump" />
									<xsl:variable name="SumpDepth" select="landUtils:getSumpDepth(.)" />
									<xsl:variable name="StructType" select="landUtils:getStructType(.)" />
									<td>
										<xsl:value-of select="$name" />
									</td>
									<td>
										<xsl:value-of select="$StructType" />
									</td>
									<td>
										<xsl:choose>
											<!--Circular Struct-->
											<xsl:when test="$StructType='Circular'">
												D:<xsl:value-of select="landUtils:FormatNumber( string(./lx:CircStruct/@diameter), string($DiameterUnit), string($PipeLengthUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
											</xsl:when>
											<!--Rectangular Struct-->
											<xsl:when test="$StructType='Rectangular'">
												L:<xsl:value-of select="landUtils:FormatNumber( string(./lx:RectStruct/@length), string($SourceLinearUnit), string($PipeLengthUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
												<br></br>
												W:<xsl:value-of select="landUtils:FormatNumber( string(./lx:RectStruct/@width), string($SourceLinearUnit), string($PipeLengthUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
											</xsl:when>
											<!--Inlet Struct-->
											<xsl:when test="$StructType='Inlet'">...</xsl:when>
											<!--Outlet Struct-->
											<xsl:when test="$StructType='Outlet'">...</xsl:when>
											<!--Connection Struct-->
											<xsl:when test="$StructType='Connection'">...</xsl:when>
											<xsl:otherwise>...</xsl:otherwise>
										</xsl:choose>
									</td>									
									<td>
										<xsl:choose>
											<xsl:when test="./*/@material"><xsl:value-of select="./*/@material" /></xsl:when>
											<xsl:otherwise>...</xsl:otherwise>
										</xsl:choose>
									</td>
									<td>
										<xsl:value-of select="landUtils:FormatNumber(string($Northing), string($SourceLinearUnit), string($Pipe_Reports.Coordinate.unit), string($Pipe_Reports.Coordinate.precision), string($Pipe_Reports.Coordinate.rounding))" />
									</td>
									<td>
										<xsl:value-of select="landUtils:FormatNumber(string($Easting), string($SourceLinearUnit), string($Pipe_Reports.Coordinate.unit), string($Pipe_Reports.Coordinate.precision), string($Pipe_Reports.Coordinate.rounding))" />
									</td>
									<td>
										<xsl:value-of select="landUtils:FormatNumber(string($elevRim), string($SourceLinearUnit), string($Pipe_Reports.Elevation.unit), string($Pipe_Reports.Elevation.precision), string($Pipe_Reports.Elevation.rounding))" />
									</td>
									<td>
										<xsl:value-of select="landUtils:FormatNumber(string($elevSump), string($SourceLinearUnit), string($Pipe_Reports.Elevation.unit), string($Pipe_Reports.Elevation.precision), string($Pipe_Reports.Elevation.rounding))" />
									</td>
									<td>
										<xsl:value-of select="landUtils:FormatNumber(string($SumpDepth), string($SourceLinearUnit), string($Pipe_Reports.Linear.unit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
									</td>
									<td>
										<xsl:for-each select="./lx:Invert">
											<xsl:value-of select="landUtils:getPipesName(.)" />
											<br></br>
										</xsl:for-each>
									</td>
								</tr>
							</xsl:for-each>
						</table>
					<br />
					</xsl:for-each>
				</div>
			</body>
		</html>
	</xsl:template>
</xsl:stylesheet>