<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0" xmlns:lx="http://www.landxml.org/schema/LandXML-1.2" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">
	<!--Description:Structures data in comma delimited format, suitable for spreadsheet import.
This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
	<!--CreatedBy:Autodesk Inc. -->
	<!--DateCreated:12/13/2004 -->
	<!--LastModifiedBy:Autodesk Inc. -->
	<!--DateModified:02/28/2005 -->
	<!--OutputExtension:csv -->
	<xsl:output method="text" media-type="iso-8859-1" encoding="us-ascii" />
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
				<xsl:value-of select="//lx:Units/*/@linearUnit"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:param>		
	<!-- ============================= -->
	<xsl:template match="/">
		<xsl:for-each select="//lx:PipeNetworks/lx:PipeNetwork">
			<xsl:text>"Pipe Network: </xsl:text><xsl:value-of select="@name" /><xsl:text>"&#xa;</xsl:text>
			<xsl:text>"Structure",</xsl:text>
			<xsl:text>"Type",</xsl:text>
			<xsl:choose>
				<xsl:when test="$DisplayPipeLengthUnit='foot'">
					<xsl:text>"Size (ft)",</xsl:text>
				</xsl:when>
				<xsl:when test="$DisplayPipeLengthUnit='meter'">
					<xsl:text>"Size (m)",</xsl:text>
				</xsl:when>
			</xsl:choose>
			<xsl:text>"Material",</xsl:text>
			<xsl:choose>
				<xsl:when test="$Pipe_Reports.Coordinate.unit='default'">
					<xsl:choose>
						<xsl:when test="$DisplaySourceLinearUnit='foot'">
							<xsl:text>"Northing (ft)","Easting (ft)",</xsl:text>
						</xsl:when>
						<xsl:when test="$DisplaySourceLinearUnit='meter'">
							<xsl:text>"Northing (m)","Easting (m)",</xsl:text>
						</xsl:when>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="$DisplayReportsCoordinateUnit='foot'">
					<xsl:text>"Northing (ft)","Easting (ft)",</xsl:text>
				</xsl:when>
				<xsl:when test="$DisplayReportsCoordinateUnit='meter'">
					<xsl:text>"Northing (m)","Easting (m)",</xsl:text>
				</xsl:when>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="$Pipe_Reports.Elevation.unit='default'">
					<xsl:choose>
						<xsl:when test="$DisplaySourceLinearUnit='foot'">
							<xsl:text>"Rim Ele (ft)","Sump Elev (ft)",</xsl:text>
						</xsl:when>
						<xsl:when test="$DisplaySourceLinearUnit='meter'">
							<xsl:text>"Rim Ele (m)","Sump Elev (m)",</xsl:text>
						</xsl:when>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="$DisplayReportsElevationUnit='foot'">
					<xsl:text>"Rim Ele (ft)","Sump Elev (ft)",</xsl:text>
				</xsl:when>
				<xsl:when test="$DisplayReportsElevationUnit='meter'">
					<xsl:text>"Rim Ele (m)","Sump Elev (m)",</xsl:text>
				</xsl:when>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="$DisplayPipeLengthUnit='foot'">
					<xsl:text>"Sump Depth (ft)",</xsl:text>
				</xsl:when>
				<xsl:when test="$DisplayPipeLengthUnit='meter'">
					<xsl:text>"Sump Depth (m)",</xsl:text>
				</xsl:when>
			</xsl:choose>
			<xsl:text>"Pipes"&#xa;</xsl:text>
			<xsl:for-each select="./lx:Structs/lx:Struct">
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
	
		<xsl:text />"<xsl:value-of select="$name" />",<xsl:text />
		<xsl:text />"<xsl:value-of select="$StructType" />",<xsl:text />		
		<xsl:choose>
			<!--Circular Struct-->
			<xsl:when test="$StructType='Circular'">
				<xsl:text />"D:<xsl:value-of select="landUtils:FormatNumber( string(./lx:CircStruct/@diameter), string($DiameterUnit), string($PipeLengthUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />",<xsl:text />
			</xsl:when>
			<!--Rectangular Struct-->
			<xsl:when test="$StructType='Rectangular'">
				<xsl:text />"L:<xsl:value-of select="landUtils:FormatNumber( string(./lx:RectStruct/@length), string($SourceLinearUnit), string($PipeLengthUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />						
				<xsl:text>&#xa;W:</xsl:text>
					<xsl:value-of select="landUtils:FormatNumber( string(./lx:RectStruct/@width), string($SourceLinearUnit), string($PipeLengthUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />",<xsl:text />
			</xsl:when>
			<!--Inlet Struct-->
			<xsl:when test="$StructType='Inlet'">
				<xsl:text />"...",<xsl:text />
			</xsl:when>
			<!--Outlet Struct-->
			<xsl:when test="$StructType='Outlet'">
				<xsl:text />"...",<xsl:text />
			</xsl:when>
			<!--Connection Struct-->
			<xsl:when test="$StructType='Connection'">
				<xsl:text />"...",<xsl:text />
			</xsl:when>
		</xsl:choose>
		<xsl:choose>
			<xsl:when test="./*/@material">
				<xsl:text />"<xsl:value-of select="./*/@material" />",<xsl:text />
			</xsl:when>
			<xsl:otherwise>
				<xsl:text />"...",<xsl:text />
			</xsl:otherwise>
		</xsl:choose>
		<xsl:text />"<xsl:value-of select="landUtils:FormatNumber(string($Northing), string($SourceLinearUnit), string($Pipe_Reports.Coordinate.unit), string($Pipe_Reports.Coordinate.precision), string($Pipe_Reports.Coordinate.rounding))" />",<xsl:text />
		<xsl:text />"<xsl:value-of select="landUtils:FormatNumber(string($Easting), string($SourceLinearUnit), string($Pipe_Reports.Coordinate.unit), string($Pipe_Reports.Coordinate.precision), string($Pipe_Reports.Coordinate.rounding))" />",<xsl:text />
		<xsl:text />"<xsl:value-of select="landUtils:FormatNumber(string($elevRim), string($SourceLinearUnit), string($Pipe_Reports.Elevation.unit), string($Pipe_Reports.Elevation.precision), string($Pipe_Reports.Elevation.rounding))" />",<xsl:text />
		<xsl:text />"<xsl:value-of select="landUtils:FormatNumber(string($elevSump), string($SourceLinearUnit), string($Pipe_Reports.Elevation.unit), string($Pipe_Reports.Elevation.precision), string($Pipe_Reports.Elevation.rounding))" />",<xsl:text />
		<xsl:text />"<xsl:value-of select="landUtils:FormatNumber(string($SumpDepth), string($SourceLinearUnit), string($Pipe_Reports.Linear.unit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />",<xsl:text />			
		<xsl:text>"</xsl:text>
		
		<xsl:variable name="Counter" select="landUtils:ResetInternalCounterAndGetCurNodeCounter(./lx:Invert)" />
		<xsl:for-each select="./lx:Invert">
					<xsl:variable name="CurCounter" select="landUtils:DecreCounter()" />
					<xsl:text />
					<xsl:value-of select="landUtils:getPipesName(.)" />
					<xsl:text />
					<xsl:choose>
						<xsl:when test="$CurCounter > 0">
							<xsl:text>&#xa;</xsl:text>
						</xsl:when>
					</xsl:choose>
				</xsl:for-each>
		<xsl:text>"&#xa;</xsl:text>		
</xsl:for-each>
<xsl:text>&#xa;</xsl:text>
</xsl:for-each>
	</xsl:template>
</xsl:stylesheet>