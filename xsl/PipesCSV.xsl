<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0" xmlns:lx="http://www.landxml.org/schema/LandXML-1.2" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">
	<!--Description:Pipe data in comma delimited format, suitable for spreadsheet import.
This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
	<!--CreatedBy:Autodesk Inc. -->
	<!--DateCreated:12/13/2004 -->
	<!--LastModifiedBy:Autodesk Inc. -->
	<!--DateModified:03/07/2005 -->
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
		<xsl:for-each select="//lx:PipeNetworks/lx:PipeNetwork">
			<xsl:text>"Pipe Network: </xsl:text>
			<xsl:value-of select="@name" />
			<xsl:text>"&#xa;</xsl:text>
			<xsl:text>"Name",</xsl:text>
			<xsl:text>"Shape",</xsl:text>
			<xsl:choose>
				<xsl:when test="$PipeSizeUnit='milimeter'">
					<xsl:text>"Size (mm)",</xsl:text>
				</xsl:when>
				<xsl:when test="$PipeSizeUnit='inch'">
					<xsl:text>"Size (in)",</xsl:text>
				</xsl:when>
			</xsl:choose>			
			<xsl:text>"Material",</xsl:text>
			<xsl:text>"US Node",</xsl:text>
			<xsl:text>"DS Node",</xsl:text>
			<xsl:choose>
				<xsl:when test="$Pipe_Reports.Elevation.unit='default'">
					<xsl:choose>
						<xsl:when test="$DisplaySourceLinearUnit='foot'">
							<xsl:text>"US Invert (ft)","DS Invert (ft)",</xsl:text>
						</xsl:when>
						<xsl:when test="$DisplaySourceLinearUnit='meter'">
							<xsl:text>"US Invert (m)","DS Invert (m)",</xsl:text>
						</xsl:when>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="$DisplayReportsElevationUnit='foot'">
					<xsl:text>"US Invert (ft)","DS Invert (ft)",</xsl:text>
				</xsl:when>
				<xsl:when test="$DisplayReportsElevationUnit='meter'">
					<xsl:text>"US Invert (m)","DS Invert (m)",</xsl:text>
				</xsl:when>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="$Pipe_Reports.Pipe_Length.type='2-D'">
					<xsl:choose>
						<xsl:when test="$DisplayPipeLengthUnit='foot'">
							<xsl:text>"2D Length (ft)
					center-to-center
					edge-to-edge",</xsl:text>
						</xsl:when>
						<xsl:when test="$DisplayPipeLengthUnit='meter'">
							<xsl:text>"2D Length (m)
					center-to-center
					edge-to-edge",</xsl:text>
						</xsl:when>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="$Pipe_Reports.Pipe_Length.type='3-D'">
					<xsl:choose>
						<xsl:when test="$DisplayPipeLengthUnit='foot'">
							<xsl:text>"3D Length (ft)
					center-to-center
					edge-to-edge",</xsl:text>
						</xsl:when>
						<xsl:when test="$DisplayPipeLengthUnit='meter'">
							<xsl:text>"3D Length (m)
					center-to-center
					edge-to-edge",</xsl:text>
						</xsl:when>
					</xsl:choose>
				</xsl:when>
			</xsl:choose>
			<xsl:text>"% Slope"&#xa;</xsl:text>
		
		<xsl:for-each select="./lx:Pipes/lx:Pipe">
	
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
							<!--get pntRef, and get coordinate-->
							<xsl:variable name="pntRef" select="./parent::*/parent::*/lx:Structs/lx:Struct[@name=$start]/lx:Center/@pntRef" />
							<xsl:copy-of select="//lx:CgPoints/lx:CgPoint[@name=$pntRef]" />
						</xsl:when>
						<xsl:otherwise>
							<!--get the string(./lx:Center)-->
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
		
		<xsl:text />"<xsl:value-of select="$name" />",<xsl:text />
		<xsl:text />"<xsl:value-of select="$PipeType" />",<xsl:text />
      <xsl:choose>
        <xsl:when test="$PipeType='Circular'">
          <xsl:text />"D:<xsl:value-of select="landUtils:FormatNumber( string(./lx:CircPipe/@diameter),string($DiameterUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />",<xsl:text />
        </xsl:when>
        <xsl:when test="$PipeType='Elliptical'">
          <xsl:text />"H:<xsl:value-of select="landUtils:FormatNumber( string(./lx:ElliPipe/@height), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
          <xsl:text>&#xa;S:</xsl:text>
          <xsl:value-of select="landUtils:FormatNumber( string(./lx:ElliPipe/@span), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />",<xsl:text />
        
        </xsl:when>
        <xsl:when test="$PipeType='Egg'">
          <xsl:text />"H:<xsl:value-of select="landUtils:FormatNumber( string(./lx:EggPipe/@height), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
          <xsl:text>&#xa;S:</xsl:text>
          <xsl:value-of select="landUtils:FormatNumber( string(./lx:EggPipe/@span), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />",<xsl:text />
        </xsl:when>
        <xsl:when test="$PipeType='Rectangular'">
          <xsl:text />"H:<xsl:value-of select="landUtils:FormatNumber( string(./lx:RectPipe/@height), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
          <xsl:text>&#xa;W:</xsl:text>
          <xsl:value-of select="landUtils:FormatNumber( string(./lx:RectPipe/@width), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />",<xsl:text />
				
        </xsl:when>
        <xsl:when test="$PipeType='Channel'">
          <xsl:text />"H:<xsl:value-of select="landUtils:FormatNumber( string(./lx:Channel/@height), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
          <xsl:text>&#xa;WT:</xsl:text>
          <xsl:value-of select="landUtils:FormatNumber( string(./lx:Channel/@widthTop), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />
          <xsl:text>&#xa;WB:</xsl:text>
          <xsl:value-of select="landUtils:FormatNumber( string(./lx:Channel/@widthBottom), string($SourceLinearUnit), string($PipeSizeUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />",<xsl:text />
				
        </xsl:when>
        <xsl:otherwise>
          <xsl:text />"...",<xsl:text />
        </xsl:otherwise>
      </xsl:choose>


      <xsl:choose>
			<xsl:when test="./*/@material">
				<xsl:text />"<xsl:value-of select="./*/@material" />",<xsl:text />
			</xsl:when>
			<xsl:otherwise>
				<xsl:text />"...",<xsl:text />
			</xsl:otherwise>
		</xsl:choose>
		
		
		<xsl:text />"<xsl:value-of select="$start" />",<xsl:text />
		<xsl:text />"<xsl:value-of select="$end" />",<xsl:text />
		<xsl:text />"<xsl:value-of select="landUtils:FormatNumber(string($US_Invert), string($SourceLinearUnit), string($Pipe_Reports.Elevation.unit), string($Pipe_Reports.Elevation.precision), string($Pipe_Reports.Elevation.rounding))" />",<xsl:text />
		<xsl:text />"<xsl:value-of select="landUtils:FormatNumber(string($DS_Invert), string($SourceLinearUnit), string($Pipe_Reports.Elevation.unit), string($Pipe_Reports.Elevation.precision), string($Pipe_Reports.Elevation.rounding))" />",<xsl:text />								
		
		<xsl:choose>
					<xsl:when test="$Pipe_Reports.Pipe_Length.type='2-D'">
				<!--SourceLinearUnit format-->
				<xsl:variable name="TwoDLength" select="landUtils:get2DLength(string($center1), string($center2))" />
				<xsl:text />"<xsl:value-of select="landUtils:FormatNumber( string($TwoDLength), string($SourceLinearUnit), string($PipeLengthUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" /><xsl:text />
				<xsl:text>&#xa;</xsl:text>
				<xsl:text /><xsl:value-of select="landUtils:FormatNumber( string(landUtils:CC2EE(string($TwoDLength))),string($SourceLinearUnit),string($PipeLengthUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />",<xsl:text />
			</xsl:when>
					<xsl:when test="$Pipe_Reports.Pipe_Length.type='3-D'">					
					<xsl:variable name="ThreeDLength" select="landUtils:get3DLength(string($center1), string($US_Invert), string($center2), string($DS_Invert))" />
					<xsl:text>"</xsl:text>
					<xsl:text /><xsl:value-of select="landUtils:FormatNumber(string($ThreeDLength), string($SourceLinearUnit), string($PipeLengthUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" /><xsl:text />
					<xsl:text>&#xa;</xsl:text>
					<xsl:text /><xsl:value-of select="landUtils:FormatNumber(string(landUtils:CC2EE(string($ThreeDLength))), string($SourceLinearUnit), string($PipeLengthUnit), string($Pipe_Reports.Linear.precision), string($Pipe_Reports.Linear.rounding))" />",<xsl:text />					
			</xsl:when>
				</xsl:choose>
		
		<xsl:text />"<xsl:value-of select="landUtils:FormatNumber(string(landUtils:getSlope(string($center1), string($US_Invert), string($center2), string($DS_Invert))), string($Pipe_Reports.Angular.unit), string($Pipe_Reports.Angular.unit), string($Pipe_Reports.Angular.precision), string($Pipe_Reports.Angular.rounding))" />"&#xa;<xsl:text />
	
</xsl:for-each>
<xsl:text>&#xa;</xsl:text>
		</xsl:for-each>
	</xsl:template>
</xsl:stylesheet>
