<?xml version="1.0" encoding="UTF-8"?>
<?altova_samplexml D:\Fenway Dwgs\764324\pg_pntref.xml?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:lx="http://www.landxml.org/schema/LandXML-1.2" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit" xmlns:lxml="urn:lx_utils">
	<!--Description:COGO point list report.  Listing of COGO points in table format.
This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
	<!--CreatedBy:Autodesk Inc. -->
	<!--DateCreated:07/26/2002 -->
	<!--LastModifiedBy:Autodesk Inc. -->
	<!--DateModified:10/15/2004 -->
	<!--OutputExtension:html -->
	<xsl:output method="html" encoding="UTF-8"/>
	<!-- ==== JavaScript Includes ==== -->
	<xsl:include href="LandXMLUtils_JScript.xsl"/>
	<xsl:include href="General_Formating_JScript.xsl"/>
	<xsl:include href="Number_Formatting.xsl"/>
	<xsl:include href="Cogo_Point.xsl"/>
	<xsl:include href="header.xsl"/>
	<!-- ============================= -->
	<xsl:param name="SourceLinearUnit" select="//lx:Units/*/@linearUnit"/>
	<xsl:template match="/">
		<xsl:variable name="TotalPoints" select="count(//lx:CgPoint)"/>
		<xsl:call-template name="SetGeneralFormatParameters"/>
		<html>
			<head>
				<title>Point report for <xsl:value-of select="//lx:Project/@name"/>
				</title>
				<style type="text/css">
				div{
				width:7in;
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
						<xsl:with-param name="ReportTitle">Point Report</xsl:with-param>
						<xsl:with-param name="ReportDesc">
							<xsl:value-of select="//lx:Project/@name"/>
						</xsl:with-param>
					</xsl:call-template>
					<h2>Point Report</h2>
					<h3>Project: <xsl:value-of select="//lx:Project/@name"/>
					</h3>
					<p>Total COGO Points:<xsl:value-of select="$TotalPoints"/>
					</p>
					<table bordercolor="black" border="1" cellspacing="0">
						<tr>
							<th>Number</th>
							<xsl:choose>
								<xsl:when test="$Point.Coordinate.unit='default'">
									<xsl:choose>
										<xsl:when test="$SourceLinearUnit='foot'">
											<th>Northing (ft)</th>
											<th>Easting (ft)</th>
										</xsl:when>
										<xsl:when test="$SourceLinearUnit='USSurveyFoot'">
											<th>Northing (USSurveyFoot)</th>
											<th>Easting (USSurveyFoot)</th>
										</xsl:when>
										<xsl:when test="$SourceLinearUnit='inch'">
											<th>Northing (in)</th>
											<th>Easting (in)</th>
										</xsl:when>
										<xsl:when test="$SourceLinearUnit='mile'">
											<th>Northing (mi)</th>
											<th>Easting (mi)</th>
										</xsl:when>
										<xsl:when test="$SourceLinearUnit='millimeter'">
											<th>Northing (mm)</th>
											<th>Easting (mm)</th>
										</xsl:when>
										<xsl:when test="$SourceLinearUnit='centimeter'">
											<th>Northing (cm)</th>
											<th>Easting (cm)</th>
										</xsl:when>
										<xsl:when test="$SourceLinearUnit='meter'">
											<th>Northing (m)</th>
											<th>Easting (m)</th>
										</xsl:when>
										<xsl:when test="$SourceLinearUnit='kilometer'">
											<th>Northing (km)</th>
											<th>Easting (km)</th>
										</xsl:when>
									</xsl:choose>
								</xsl:when>
								<xsl:when test="$Point.Coordinate.unit='foot'">
									<th>Northing (ft)</th>
									<th>Easting (ft)</th>
								</xsl:when>
								<xsl:when test="$Point.Coordinate.unit='meter'">
									<th>Northing (m)</th>
									<th>Easting (m)</th>
								</xsl:when>
							</xsl:choose>
							<xsl:choose>
								<xsl:when test="$Point.Elevation.unit='default'">
									<xsl:choose>
										<xsl:when test="$SourceLinearUnit='foot'">
											<th>Elevation (ft)</th>
										</xsl:when>
										<xsl:when test="$SourceLinearUnit='USSurveyFoot'">
											<th>Elevation (USSurveyFoot)</th>
										</xsl:when>
										<xsl:when test="$SourceLinearUnit='inch'">
											<th>Elevation (in)</th>>
										</xsl:when>
										<xsl:when test="$SourceLinearUnit='mile'">
											<th>Elevation (mi)</th>
										</xsl:when>
										<xsl:when test="$SourceLinearUnit='millimeter'">
											<th>Elevation (mm)</th>
										</xsl:when>
										<xsl:when test="$SourceLinearUnit='centimeter'">
											<th>Elevation (cm)</th>
										</xsl:when>
										<xsl:when test="$SourceLinearUnit='meter'">
											<th>Elevation (m)</th>
										</xsl:when>
										<xsl:when test="$SourceLinearUnit='kilometer'">
											<th>Elevation (km)</th>
										</xsl:when>
									</xsl:choose>
								</xsl:when>
								<xsl:when test="$Point.Elevation.unit='foot'">
									<th>Elevation (ft)</th>
								</xsl:when>
								<xsl:when test="$Point.Elevation.unit='meter'">
									<th>Elevation (m)</th>
								</xsl:when>
							</xsl:choose>
		              <th width="75%">Description</th>
						</tr>
						<xsl:for-each select="//lx:CgPoint[not(@pntRef)]">
							<xsl:apply-templates select="." mode="set"/>
							<xsl:variable name="Northing" select="landUtils:GetCogoPointNorthing()"/>
							<xsl:variable name="Easting" select="landUtils:GetCogoPointEasting()"/>
							<xsl:variable name="Elevation" select="landUtils:GetCogoPointElevation()"/>
							<tr>
								<td>
									<xsl:value-of select="@name"/>
								</td>
								<td>
									<xsl:value-of select="landUtils:FormatNumber(string($Northing), string($SourceLinearUnit), string($Point.Coordinate.unit), string($Point.Coordinate.precision), string($Point.Coordinate.rounding))"/>
								</td>
								<td>
									<xsl:value-of select="landUtils:FormatNumber(string($Easting), string($SourceLinearUnit), string($Point.Coordinate.unit), string($Point.Coordinate.precision), string($Point.Coordinate.rounding))"/>
								</td>
								<td>
									<xsl:value-of select="landUtils:FormatNumber(string($Elevation), string($SourceLinearUnit), string($Point.Elevation.unit), string($Point.Elevation.precision), string($Point.Elevation.rounding))"/>
								</td>
								<xsl:choose>
									<xsl:when test="@desc">
										<!-- use "desc" attribute if available -->
										<td style="text-align:left">
											<xsl:value-of select="@desc"/>
										</td>
									</xsl:when>
									<xsl:when test="@code">
										<!-- else try the "code" attribute -->
										<td style="text-align:left">
											<xsl:value-of select="@code"/>
										</td>
									</xsl:when>
									<xsl:otherwise>
										<!-- otherwise, print "NA" -->
										<td style="text-align:center">NA</td>
									</xsl:otherwise>
								</xsl:choose>
							</tr>
						</xsl:for-each>
					</table>
				</div>
			</body>
		</html>
	</xsl:template>
</xsl:stylesheet>
