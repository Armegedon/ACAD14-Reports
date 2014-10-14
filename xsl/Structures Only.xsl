<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0" xmlns:lx="http://www.landxml.org/schema/LandXML-1.2" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">
	<!--Description:Tabulation of all structures in the selected networks.
This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
	<!--CreatedBy:Autodesk Inc. -->
	<!--DateCreated:12/13/2004 -->
	<!--LastModifiedBy:Autodesk Inc. -->
	<!--DateModified:02/20/2005 -->
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
	<!-- ============================= -->
	<xsl:template match="/">
		<html>
			<head>
				<title>Pipes report for <xsl:value-of select="//lx:Project/@name" /></title>
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
						<xsl:with-param name="ReportTitle">Structures Report</xsl:with-param>
						<xsl:with-param name="ReportDesc">
							<xsl:value-of select="//lx:Project/@name" />
						</xsl:with-param>
					</xsl:call-template>
					<br />
					<xsl:for-each select="//lx:PipeNetworks/lx:PipeNetwork">
						<b>Pipe Network: <xsl:value-of select="@name" /></b>
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