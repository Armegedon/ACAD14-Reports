<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:lx="http://www.landxml.org/schema/LandXML-1.2" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit" xmlns:lxml="urn:lx_utils">
  <xsl:output method="html" encoding="UTF-8"/>
  <!--Description:Surface Report
    
This form provides a list of properties for the selected surface(s) which include description, 2D area, 3D area, Max. and Min. Elevations, and total number of points and triangles.
 
This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
  <!--CreatedBy:Autodesk Inc. -->
  <!--DateCreated:06/15/2002 -->
  <!--LastModifiedBy:Autodesk Inc. -->
  <!--DateModified:01/16/2006 -->
  <!--OutputExtension:html -->
  <!-- =========== JavaScript Include==== -->
  <xsl:include href="LandXMLUtils_JScript.xsl"/>
  <xsl:include href="Surface_Layout.xsl"/>
  <xsl:include href="General_Formating_JScript.xsl"/>
  <xsl:include href="Conversion_JScript.xsl"/>
  <xsl:include href="header.xsl"/>
  <xsl:include href="Number_Formatting.xsl"/>
  <!-- ================================= -->

  <xsl:param name="SourceLinearUnit" select="//lx:Units/*/@linearUnit"/>
  <xsl:param name="SourceAreaUnit" select="//lx:Units/*/@areaUnit"/>
  <xsl:param name="SourceVolumeUnit" select="//lx:Units/*/@volumeUnit"/>

  <xsl:template match="/">
    <xsl:call-template name="SetGeneralFormatParameters"/>
    <html>
      <body>
        <div style="width:7in">
          <xsl:call-template name="AutoHeader">
            <xsl:with-param name="ReportTitle">Surface Report</xsl:with-param>
            <xsl:with-param name="ReportDesc">
              <xsl:value-of select="//lx:Project/@name"/>
            </xsl:with-param>
          </xsl:call-template>
          <table width="100%" border="2">
            <tr>
              <td>
                <b>Linear Units: </b><xsl:value-of select="//lx:Units/*/@linearUnit"/></td>
              <td>
                <b>Area Units: </b><xsl:value-of select="//lx:Units/*/@areaUnit"/></td>
              <td>
                <b>Volume Units: </b><xsl:value-of select="//lx:Units/*/@volumeUnit"/></td>
            </tr>
          </table>
          <xsl:apply-templates select="//lx:Surface"/>
          <xsl:apply-templates select="//lx:SurfVolume"/>
          <hr color="black"/>
        </div>
      </body>
    </html>
  </xsl:template>
  <xsl:template match="//lx:Surface">
    <hr color="black"/>
    <b>Surface: <xsl:value-of select="@name"/></b><br/>
    Description: <xsl:value-of select="@desc"/>
    <table width="100%" bgcolor="lightgreen">
      <tr>
        <td>Area 2D: <xsl:value-of select="landUtils:FormatNumber(string(./lx:Definition/@area2DSurf), string($SourceAreaUnit), string($Surface.2D_Area.unit), string($Surface.2D_Area.precision), string($Surface.2D_Area.rounding))"/>
        </td>
        <td>Area 3D: <xsl:value-of select="landUtils:FormatNumber(string(./lx:Definition/@area3DSurf), string($SourceAreaUnit), string($Surface.3D_Area.unit), string($Surface.3D_Area.precision), string($Surface.3D_Area.rounding))"/>
        </td>
      </tr>
      <tr>
        <td>Elevation Max:  <xsl:value-of select="landUtils:FormatNumber(string(./lx:Definition/@elevMax), string($SourceLinearUnit), string($Surface.Elevations.unit), string($Surface.Elevations.precision), string($Surface.Elevations.rounding))"/>
        </td>
        <td>Elevation Min: <xsl:value-of select="landUtils:FormatNumber(string(./lx:Definition/@elevMin), string($SourceLinearUnit), string($Surface.Elevations.unit), string($Surface.Elevations.precision), string($Surface.Elevations.rounding))"/>
        </td>
      </tr>
      <tr>
        <td>Number of Points: <xsl:value-of select="count(.//lx:P)"/>
        </td>
        <td>Number of Triangles: <xsl:value-of select="count(.//lx:F[not(@i)])"/>
        </td>
      </tr>
    </table>
  </xsl:template>
  <xsl:template match="//lx:SurfVolume">
    <hr color="black"/>
    <b>Volume Surface: <xsl:value-of select="@name"/></b><br/>
    Description: <xsl:value-of select="@desc"/>
    <table width="100%" bgcolor="cyan">
      <tr>
        <td>Volume Cut: <xsl:value-of select="landUtils:FormatNumber(string(@volCut), string($SourceVolumeUnit), string($Surface.Volume.unit), string($Surface.Volume.precision), string($Surface.Volume.rounding))"/>
        </td>
        <td>Volume Fill: <xsl:value-of select="landUtils:FormatNumber(string(@volFill), string($SourceVolumeUnit), string($Surface.Volume.unit), string($Surface.Volume.precision), string($Surface.Volume.rounding))"/>
        </td>
        <td>Volume Total: <xsl:value-of select="landUtils:FormatNumber(string(@volTotal), string($SourceVolumeUnit), string($Surface.Volume.unit), string($Surface.Volume.precision), string($Surface.Volume.rounding))"/>
        </td>
      </tr>
      <tr>
        <td>Compare Surface: <xsl:value-of select="@surfCompare"/><br/>
          Base Surface: <xsl:value-of select="@surfBase"/>
        </td>
      </tr>
    </table>
  </xsl:template>
</xsl:stylesheet>
