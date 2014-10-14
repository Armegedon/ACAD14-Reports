<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:lx="http://www.landxml.org/schema/LandXML-1.2" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit" xmlns:lxml="urn:lx_utils">
	<!--Description:General Legal Description for Parcels report.
	
This is a sample report form for generating metes and bounds legal descriptions for parcels.  This report form (stylesheet) is not intended to be a replacement for a trained, experienced professional.  It is intended as a example to make the professional more productive.
	
This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
	<!--CreatedBy:Autodesk Inc. -->
	<!--DateCreated:12/15/2003 -->
	<!--LastModifiedBy:Autodesk Inc. -->
	<!--DateModified:1/13/2004 -->
	<!--OutputExtension:html -->

	<!-- Parameters for unit conversion -->
	<xsl:param name="SourceLinearUnit" select="//lx:Units/*/@linearUnit"/>
	<xsl:param name="SourceAreaUnit" select="//lx:Units/*/@areaUnit"/>
	<!-- =========== JavaScript Includes ==== -->
	<xsl:include href="LandXMLUtils_JScript.xsl"/>
	<xsl:include href="General_Formating_JScript.xsl"/>
	<xsl:include href="Plan_Comp_JScript.xsl"/>
	<xsl:include href="Conversion_JScript.xsl"/>
	<xsl:include href="Number_Formatting.xsl"/>
	<xsl:include href="Parcel_Layout.xsl"/>
	<xsl:include href="CoordGeometry.xsl"/>
	<xsl:include href="LegalDescription_Layout.xsl"/>
	
	<xsl:template match="/">
		<xsl:variable name="FileSet" select="landUtils:SetLegalPhrasingFile(string($Legal_Descriptions.Phrasing_File.File_location))"/>
		<xsl:variable name="setSrcUnit" select="landUtils:SetSourceUnits(string($SourceLinearUnit))"/>
		<xsl:variable name="setUnit" select="landUtils:SetUnits(string($Legal_Descriptions.Report_Units.units))"/>
		<xsl:variable name="coordP" select="landUtils:SetCoordPrec(string($Legal_Descriptions.Coordinate.precision))"/>
		<xsl:variable name="coordR" select="landUtils:SetCoordRound(string($Legal_Descriptions.Coordinate.rounding))"/>
		<xsl:variable name="distP" select="landUtils:SetDistancePrec(string($Legal_Descriptions.Distance.precision))"/>
		<xsl:variable name="distR" select="landUtils:SetDistanceRound(string($Legal_Descriptions.Coordinate.rounding))"/>
		<xsl:variable name="dirP" select="landUtils:SetDirectionPrecision(string($Legal_Descriptions.Direction.precision))"/>
		<xsl:variable name="dirR" select="landUtils:SetDirectionRound(string($Legal_Descriptions.Direction.rounding))"/>
		<xsl:variable name="dirTyp" select="landUtils:SetDirectionType(string($Legal_Descriptions.Direction.type))"/>
		<xsl:variable name="dirFormat" select="landUtils:SetDirectionFormat(string($Legal_Descriptions.Direction.format))"/>
		<xsl:variable name="angleP" select="landUtils:SetAnglePrec(string($Legal_Descriptions.Angle.precision))"/>
		<xsl:variable name="angleR" select="landUtils:SetAngleRound(string($Legal_Descriptions.Angle.rounding))"/>
		<xsl:variable name="angleFormat" select="landUtils:SetAngleFormat(string($Legal_Descriptions.Angle.format))"/>
		<xsl:variable name="areaUnit" select="$Legal_Descriptions.Area.unit"/>
		<xsl:variable name="areaP" select="$Legal_Descriptions.Area.precision"/>
		<xsl:variable name="areaR" select="$Legal_Descriptions.Area.rounding"/>
		<xsl:variable name="areaU" select="landUtils:SetAreaUnit(string($Legal_Descriptions.Area.unit))"/>
		<xsl:variable name="staDisp" select="landUtils:SetStationPattern(string($Legal_Descriptions.Station.Display))"/>
		<xsl:variable name="staPrec" select="landUtils:SetStationPrecision(string($Legal_Descriptions.Station.precision))"/>
		<xsl:variable name="staRound" select="landUtils:SetStationRounding(string($Legal_Descriptions.Station.rounding))"/>
		<html>
			<head>
				<title>General Legal Description Generation</title>
				<script language="VBScript">
				Dim xLoc
				Dim yLoc
				
				Dim bIsShrunk
		
			sub appendOutput
 				dim parcelSelected
 				Dim chk
 				Dim bDone
  				bDone = False
				Dim i
				i = 0
 				
				chk = Document.all("selectAll").Value
				if(chk = "ON") then
 					parcelSelected = Document.forms("appendForm").parcelCombo.Value
 					appendLegalDiv parcelSelected 
				else
					dim sz
					sz = Document.forms("appendForm").parcelCombo.options.length '20
					do until i = sz
						parcelSelected = Document.forms("appendForm").parcelCombo(i).text
						if(parcelSelected = "") then
							bDone = True
						else
	 						appendLegalDiv parcelSelected 
	 						i = i + 1
						end if
					loop 
				end if
				hideInput
			end sub
			
			function GetHeader(parcelName)
				Dim headStr
				headStr = "<p><b>LEGAL DESCRIPTION OF: " + parcelName + "</b></p>"
				GetHeader = headStr + "<p/>"
			end function
			
			
			function GetFooter
				Dim footStr
				footStr = ""
				GetFooter = footStr
			end function

			sub appendLegalDiv(parcelName)
				Dim dataDoc
				Dim parcelNode
				Dim legalDesc
				Dim pName
				Dim pobDesc
				Dim cNodes
				Dim cnode
				
				legalDesc = getHeader(parcelName)
				set pobDesc = Document.all("POBDescription")
				
				aliUse = Document.forms("appendForm").AlignmentCombo.selectedIndex
				
				Dim chk
				chk = Document.all("selectAll").Value
				if(chk = "OFF") then
					aliUse = 0
				else
				
				end if

				
				set dataDoc= Document.all("Data").XMLDocument
				for each parcelNode in dataDoc.selectNodes("//Parcel[@name='" + parcelName + "']")
					if(aliUse > 0) then
						aliName = Document.forms("appendForm").AlignmentCombo(aliUse).text
						for each aliNode in dataDoc.selectNodes("//Alignment[@name='" + aliName + "']")
							legalDesc = legalDesc + aliNode.text + "<br/>"
						next
					end if
					if(chk = "ON") then
						legalDesc = legalDesc +  pobDesc.innerText + "<br/>"
					end if
					
					set cNodes = parcelNode.childNodes
					for each cnode in cNodes
						if(cnode.nodeType = 3)then
							legalDesc = legalDesc + cnode.text
						else
							legalDesc = legalDesc + cnode.xml
						end if
					next

					legalDesc = legalDesc + "<p/>" + parcelNode.getAttributeNode("area").text
					legalDesc = legalDesc  + GetFooter + "<hr/>"
				next
				
				appendDiv legalDesc
			end sub
			
			<![CDATA[
			Function appendDiv(text)
				dim styleAttr
				dim styleVal
				dim appendedDiv

				set appendedDiv =document.createElement("<DIV style='page-break-after: always'>")
				'appendedDiv.setAttribute "style", "Pat Coleman"

				'set styleAttr = document.createAttribute("style")
				'styleAttr.Value = "Test"
				'appendedDiv.setAttributeNode styleAttr

				appendedDiv.innerHTML = text
				document.all("output").appendChild appendedDiv

				appendDiv = "done"
			end Function
			]]>
			
			sub SelectAllOnChange
				Dim chk
				chk = Document.all("selectAll").Value
				if(chk = "ON") then
					Document.all("selectAll").Value = "OFF"
					Document.all("parcelCombo").disabled = "disabled"
					Document.all("AlignmentCombo").disabled = "disabled"
					Document.all("pointDescCombo").disabled = "disabled"
					Document.all("insertButton").disabled = "disabled"
					
				else
					Document.all("selectAll").Value = "ON"				
					Document.all("parcelCombo").disabled = ""
					Document.all("AlignmentCombo").disabled = ""
					Document.all("pointDescCombo").disabled = ""
					Document.all("insertButton").disabled = ""
				end if
			end sub
			
			sub outputScrolled
				Document.all("input").style.posTop = document.body.scrollTop
				if bIsShrunk then
					Document.all("input").style.overflow = "hidden"
				else
					Document.all("input").style.posHeight = document.body.clientHeight
					Document.all("input").style.overflow = "auto"
				end if
			end sub
			
			sub clearReport
				Dim iResult
				iResult = MsgBox ("Are you sure you would like to clear the current report?", 4, "Confirm Clear Report")
				if( iResult = 6) then
					document.all("output").innerHTML = ""
				end if
			end sub
			
			sub printToWindow
				dim doc
				set wind = window.open ("",null, "status=yes,toolbar=no,menubar=yes,location=no,resizable=yes,scrollbars=yes")
				wind.document.write  document.all("output").outerHTML
			end sub
			
			sub onClickPOB
				Dim offsetInfo
				
				xLoc = window.event.x
				yLoc = window.event.y
			end sub
			
			sub doPOBInsert
				Dim textBefore
				Dim iAdd
				Dim addText
				
				textBefore = Document.all("POBDescription").innerText 
				iAdd = Document.all("pointDescCombo").selectedIndex
								
				if(iAdd > 0) then
					addText= Document.forms("appendForm").pointDescCombo(iAdd).Value
					textBefore = textBefore + addText
					Document.all("POBDescription").innerText = textBefore
				end if
			end sub
			
			sub hideInput
				if bIsShrunk = false then
					Document.all("input").style.height = "10"
					Document.all("titleDiv").innerText = "_"
					Document.all("input").style.overflow = "hidden"
					Document.all("input").scrollTop = 0
					bIsShrunk = true
				else
					Document.all("input").style.height = ""
					Document.all("titleDiv").innerText = "General Legal Description (Click to collapse or expand)"
					Document.all("input").style.overflow = "auto"
					bIsShrunk = false
				end if
				outputScrolled
			end sub
			
   		</script>
   		<style type="text/css" media="print">
   			.noprint 
   			{ 
   				display: none
   			}
		</style>
		<style type="text/css">
   		#parcelSelectDiv
   		{
   			margin-left: 5;
   		}
   		#alignDiv
    		{
   			margin-left: 5;
   		}
  		
   		#selectPointDiv
    		{
   			margin-left: 5;
   		}
  		
   		#descPOBDiv
   		{
   			margin-left: 5;
   		}
   		
   		#buttonDiv
   		{
   			margin-left: 5;
   		}
   		
   		#input
   		{
			position:absolute;
			overflow: hidden;
			width:230px; 
 			border:black thin solid; 
 			background-color:lightgrey
  		}
	
		#output
		{
			position:absolute;
		}

		#parcelCombo
		{
			width:200px;
		}
		#AlignmentCombo
		{
			width:200px;
		}
		#pointDescCombo
		{
			width:200px;
		}
		#insertButton
		{
			width:200px;
		}
		#POBDescription
		{
			width:200px;
		}
		#appendButton
		{
			width:200px;
		}
		#clearButton
		{
			width:200px;
		}
		#saveButton
		{
			width:200px;
		}
		</style>
			</head>
			<body onscroll="outputScrolled" onload="outputScrolled" onresize="outputScrolled">
				<div id="output" style="width: 7in"></div>
				<div id="input" class="noprint">
					<xsl:call-template name="InsertForm"/>
				</div>

				<!-- Making the data island for alignments -->
				<xsl:element name="xml">
					<xsl:attribute name="id">Data</xsl:attribute>
					<xsl:element name="root">
						<xsl:element name="Alignments">
							<!-- iterate each alignment -->
							<xsl:for-each select="//lx:Alignment">
								<xsl:apply-templates select="."/>
							</xsl:for-each>
						</xsl:element>
						<!-- ============================= -->
						<!-- making the data for Parcels -->
						<xsl:element name="Parcels">
							<!-- interate each Parcel -->
							<xsl:for-each select="//lx:Parcel">
								<xsl:apply-templates select="."/>
							</xsl:for-each>
						</xsl:element>
						<!-- ============================= -->
					</xsl:element>
				</xsl:element>
			</body>
		</html>
	</xsl:template>
	<xsl:template name="InsertForm">
		<form id="appendForm">
			<b><div id="titleDiv" onclick="hideInput" style="background-color: silver">
				General Legal Description <br/>(Click to collapse or expand)
			</div></b>
			<p/>
			<div id="parcelSelectDiv">
				<input type="checkbox" name="allParcels" value="ON" tabindex="1" class="selectAll" id="selectAll" onclick="SelectAllOnChange" align="left"/>Select  all Parcels<br/>
      Parcel:<br/>
				<select size="1" name="ParcelCombo" id="parcelCombo">
					<xsl:for-each select="//lx:Parcel">
						<xsl:variable name="name" select="@name"/>
						<xsl:element name="option">
							<xsl:attribute name="value"><xsl:value-of select="$name"/></xsl:attribute>
							<xsl:value-of select="$name"/>
						</xsl:element>
					</xsl:for-each>
				</select>
			</div>
			<p/>
			<div id="alignDiv">
      POB location alignment:<br/>
				<select size="1" name="AlignmentCombo" id="AlignmentCombo">
					<option selected="selected" value="none">none</option>
					<xsl:for-each select="//lx:Alignment">
						<xsl:variable name="name">
							<xsl:value-of select="@name"/>
						</xsl:variable>
						<xsl:element name="option">
							<xsl:attribute name="value"><xsl:value-of select="$name"/></xsl:attribute>
							<xsl:value-of select="$name"/>
						</xsl:element>
					</xsl:for-each>
				</select>
			</div>
			<br/>
			<div id="selectPointDiv">Select point description for POB<br/>
				<select size="1" name="pointDescription" id="pointDescCombo">
					<option selected="selected" value="none">none</option>
					<xsl:for-each select="//lx:CgPoint[not(@pntRef)]">
						<xsl:if test="@desc">
							<xsl:element name="option">
								<xsl:attribute name="value"><xsl:value-of select="@desc"/></xsl:attribute>
								<xsl:value-of select="@name"/> - <xsl:value-of select="@desc"/>
							</xsl:element>
						</xsl:if>
					</xsl:for-each>
				</select><br/>
				<button name="insertPOBDesc" id="insertButton" onclick="doPOBInsert">Insert point description</button>
			</div>
			<p/>
			<div id="descPOBDiv">
      Description of POB<br/>
				<textarea rows="10" name="POBDescription" id="POBDescription" onclick="onClickPOB"/>
			</div>
			<p/>
			<div id="buttonDiv">
				<button name="appendReport" onclick="appendOutput" id="appendButton">Append to Report</button>
				<p/>
				<button name="clearDiv" id="clearButton" onclick="clearReport">Clear Report</button>
				<p/>
				<button name="saveDiv" id="saveButton" onclick="printToWindow">Save/Print Report</button>
			</div>
		</form>
	</xsl:template>
	
	<xsl:template match="lx:Alignment">
		<xsl:variable name="setGeomType" select="landUtils:SetDescriptionType('LocationTraverse')"/>
		<xsl:element name="Alignment">
			<xsl:attribute name="name"><xsl:value-of select="@name"/></xsl:attribute>
			<xsl:apply-templates select="./lx:CoordGeom"/>
		</xsl:element>
	</xsl:template>
	
	<xsl:template match="lx:Parcel">
		<xsl:variable name="setGeomType" select="landUtils:SetDescriptionType('Parcel')"/>
		<xsl:variable name="areaAcres" select="landUtils:FormatNumber(string(@area), string($SourceAreaUnit), string($Legal_Descriptions.Area.unit), string($Legal_Descriptions.Area.precision), string($Legal_Descriptions.Area.rounding))"/>
		<xsl:element name="Parcel">
			<xsl:attribute name="name"><xsl:value-of select="@name"/></xsl:attribute>
			<xsl:attribute name="area"><xsl:text>Containing </xsl:text><xsl:value-of select="$areaAcres"/>
				<xsl:choose>
					<xsl:when test="$Legal_Descriptions.Area.unit = 'default'">
						<xsl:text> </xsl:text><xsl:value-of select="landUtils:GetAreaUnitStr(string($SourceAreaUnit)) "/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:text> </xsl:text><xsl:value-of select="landUtils:GetAreaUnitStr(string($Legal_Descriptions.Area.unit)) "/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>
			<xsl:apply-templates select="./lx:CoordGeom"/>
		</xsl:element>
	</xsl:template>
	
	<xsl:template match="lx:CoordGeom">
		<xsl:apply-templates mode="set" select="."/>
		<xsl:for-each select="./*">
			<xsl:variable name="pos" select="position()"/>
			<xsl:value-of select="landUtils:GetLegalFor(string($pos))"/>
			<xsl:if test="$Legal_Descriptions.Phrasing_File.Metes_and_Bounds='separate line'"><p/></xsl:if>
		</xsl:for-each>
	</xsl:template>
</xsl:stylesheet>
