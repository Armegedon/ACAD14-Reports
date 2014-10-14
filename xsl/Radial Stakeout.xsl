<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit">
<!--Description:Radial Stakeout report.

This is a sample report that provides radial stakeout information for COGO points or alignment stations. The stakeout data that is generated is relative to a specified occupied COGO point, a specified backsight COGO point or direction, a specified maximum stakeout distance, and specified COGO point range or alignemnt station range.

This form is valid for LandXML 0.88, 1.0, 1.1 and 1.2 data.-->
<!--CreatedBy:Autodesk Inc. -->
<!--DateCreated:01/05/2004 -->
<!--LastModifiedBy:Autodesk Inc. -->
<!--DateModified:10/15/2004 -->
<!--OutputExtension:html -->
<xsl:output method="html" encoding="UTF-8"/>

<xsl:param name="Radial_Stakeout.Coordinate.unit">default</xsl:param>
<xsl:param name="Radial_Stakeout.Coordinate.precision">0.00</xsl:param>
<xsl:param name="Radial_Stakeout.Coordinate.rounding">normal</xsl:param>

<xsl:param name="Radial_Stakeout.Elevation.unit">default</xsl:param>
<xsl:param name="Radial_Stakeout.Elevation.precision">0.00</xsl:param>
<xsl:param name="Radial_Stakeout.Elevation.rounding">normal</xsl:param>

<xsl:param name="Radial_Stakeout.Angle.type">Angle Right</xsl:param>
<xsl:param name="Radial_Stakeout.Angle.unit">Degrees DMS</xsl:param>
<xsl:param name="Radial_Stakeout.Angle.precision">0.00</xsl:param>
<xsl:param name="Radial_Stakeout.Angle.rounding">normal</xsl:param>

<xsl:param name="Radial_Stakeout.Direction.type">Azimuth North</xsl:param>
<xsl:param name="Radial_Stakeout.Direction.unit">Degrees DMS</xsl:param>

<xsl:param name="Radial_Stakeout.Distance.unit">default</xsl:param>
<xsl:param name="Radial_Stakeout.Distance.precision">0.00</xsl:param>
<xsl:param name="Radial_Stakeout.Distance.rounding">normal</xsl:param>

<xsl:param name="Radial_Stakeout.Station.Display">####</xsl:param>
<xsl:param name="Radial_Stakeout.Station.precision">0.00</xsl:param>
<xsl:param name="Radial_Stakeout.Station.rounding">normal</xsl:param>
<xsl:param name="Radial_Stakeout.Station.Increment">50</xsl:param>

<xsl:param name="Radial_Stakeout.Alignment_Geometry.Tolerance">0.2</xsl:param>
<xsl:param name="Radial_Stakeout.Alignment_Geometry.Point_of_Curvature">PC</xsl:param>
<xsl:param name="Radial_Stakeout.Alignment_Geometry.Point_of_Tangency">PT</xsl:param>
<xsl:param name="Radial_Stakeout.Alignment_Geometry.Tangent_Spiral">TS</xsl:param>
<xsl:param name="Radial_Stakeout.Alignment_Geometry.Spiral_Curve">SC</xsl:param>
<xsl:param name="Radial_Stakeout.Alignment_Geometry.Curve_Spiral">CS</xsl:param>
<xsl:param name="Radial_Stakeout.Alignment_Geometry.Spiral_Tangent">ST</xsl:param>
<xsl:param name="Radial_Stakeout.Alignment_Geometry.Spiral_Spiral">SS</xsl:param>
<xsl:param name="Radial_Stakeout.Alignment_Geometry.Point_of_Compound_Curvature">PCC</xsl:param>
<xsl:param name="Radial_Stakeout.Alignment_Geometry.Point_of_Reverse_Curvature">PRC</xsl:param>

<xsl:include href="Cogo_Point.xsl"/>
<xsl:include href="header.xsl"/>
<xsl:include href="CoordGeometry.xsl"/>

<xsl:template match="/">
<xsl:param name="SourceLinearUnit" select="//lx:Units/*/@linearUnit"/>
<html>
	<head id="headus">
		<title>Radial Stakeout report for <xsl:value-of select="//lx:Project/@name"/></title>
		
	<script type="text/vbscript">
	Dim unit
	Dim CoordUnit, CoordPrecision, CoordRounding
	Dim ElevUnit, ElevPrecision, ElevRounding
	Dim AngType, AngUnit, AngPrecision, AngRounding
	Dim DirectionType, DirectionUnit
	Dim DistUnit, DistPrecision, DistRounding
	Dim StaDisplay, StaPrecision, StaRounding, StaIncrement
	Dim tolerance,PC,PT,TS,SC,CS,ST,SS,PCC,PRC
	Dim strStaInternal(),strStaBack(),strStaAhead()

	Const PI = 3.1415926535897932384
	unit="<xsl:value-of select="$SourceLinearUnit"/>"

	CoordUnit = "<xsl:value-of select="$Radial_Stakeout.Coordinate.unit"/>"
	CoordPrecision = "<xsl:value-of select="$Radial_Stakeout.Coordinate.precision"/>"
	CoordRounding = "<xsl:value-of select="$Radial_Stakeout.Coordinate.rounding"/>"

	ElevUnit = "<xsl:value-of select="$Radial_Stakeout.Elevation.unit"/>"
	ElevPrecision = "<xsl:value-of select="$Radial_Stakeout.Elevation.precision"/>"
	ElevRounding = "<xsl:value-of select="$Radial_Stakeout.Elevation.rounding"/>"

	AngType = "<xsl:value-of select="$Radial_Stakeout.Angle.type"/>"
	AngUnit = "<xsl:value-of select="$Radial_Stakeout.Angle.unit"/>"
	AngPrecision = "<xsl:value-of select="$Radial_Stakeout.Angle.precision"/>"
	AngRounding = "<xsl:value-of select="$Radial_Stakeout.Angle.rounding"/>"

	DirectionType = "<xsl:value-of select="$Radial_Stakeout.Direction.type"/>"
	DirectionUnit = "<xsl:value-of select="$Radial_Stakeout.Direction.unit"/>"

	DistUnit = "<xsl:value-of select="$Radial_Stakeout.Distance.unit"/>"
	DistPrecision = "<xsl:value-of select="$Radial_Stakeout.Distance.precision"/>"
	DistRounding = "<xsl:value-of select="$Radial_Stakeout.Distance.rounding"/>"

	StaDisplay = "<xsl:value-of select="$Radial_Stakeout.Station.Display"/>"
	StaPrecision = "<xsl:value-of select="$Radial_Stakeout.Station.precision"/>"
	StaRounding = "<xsl:value-of select="$Radial_Stakeout.Station.rounding"/>"
	StaIncrement ="<xsl:value-of select="$Radial_Stakeout.Station.Increment"/>"

	tolerance = "<xsl:value-of select="$Radial_Stakeout.Alignment_Geometry.Tolerance"/>"
	PC =  "<xsl:value-of select="$Radial_Stakeout.Alignment_Geometry.Point_of_Curvature"/>"
	PT =  "<xsl:value-of select="$Radial_Stakeout.Alignment_Geometry.Point_of_Tangency"/>"
	TS =  "<xsl:value-of select="$Radial_Stakeout.Alignment_Geometry.Tangent_Spiral"/>"
	SC = "<xsl:value-of select="$Radial_Stakeout.Alignment_Geometry.Spiral_Curve"/>"
	CS =  "<xsl:value-of select="$Radial_Stakeout.Alignment_Geometry.Curve_Spiral"/>"
	ST =  "<xsl:value-of select="$Radial_Stakeout.Alignment_Geometry.Spiral_Tangent"/>"
	SS = "<xsl:value-of select="$Radial_Stakeout.Alignment_Geometry.Spiral_Spiral"/>"
	PCC = "<xsl:value-of select="$Radial_Stakeout.Alignment_Geometry.Point_of_Compound_Curvature"/>"
	PRC = "<xsl:value-of select="$Radial_Stakeout.Alignment_Geometry.Point_of_Reverse_Curvature"/>"

	<![CDATA[
'##########################################
	Dim bIsShrunk
	
	Dim oldScrollVal,blnNoMaxDistance
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
			sub outputScrolled
				Document.all("allForms").style.posTop = document.body.scrollTop
				if bIsShrunk then
					Document.all("allForms").style.overflow = "hidden"
				else
					Document.all("allForms").style.posHeight = document.body.clientHeight
					Document.all("allForms").style.overflow = "auto"
				end if
			end sub
'############################################
			sub hideInput
				'MsgBox "Hideinput"
				if bIsShrunk = false then
					Document.all("allForms").style.height = "10"
					Document.all("titleDiv").innerText = "_"
					Document.all("allForms").style.overflow = "hidden"
					bIsShrunk = true
				else
					Document.all("allForms").style.height = ""
					Document.all("titleDiv").innerText = "Radial Stakeout (Click to collapse or expand)"
					Document.all("allForms").style.overflow = "auto"
					bIsShrunk = false
				end if
				outputScrolled
			end sub
'############################################
	Function Scrolling()
		Dim dblHeight, scrollDirection,strCaption

		If (document.body.scrollTop - oldScrollVal)> 0 Then
			scrollDirection = "D"
			dblHeight = document.body.scrollTop - Document.forms("StakeFormCommonDown").offsetHeight'80
		Elseif (document.body.scrollTop - oldScrollVal)< 0 Then
			scrollDirection = "U"
			dblHeight = document.body.scrollTop
		Else
			oldScrollVal = document.body.scrollTop
			Exit function
		End if
		if dblHeight < 0 Then dblHeight = 0
		
		Document.forms("Dummy").style.height = dblHeight


		oldScrollVal = document.body.scrollTop
		'Document.forms("StakeFormCommonDown").style.posTop = Document.forms("StakeFormCommonDown").style.posTop + strCaption + 0
                'strCaption = Document.forms("StakeFormCommonDown").style.posTop
		'Document.forms("StakeFormCommonDown").Append.setAttribute "value",(Document.forms("StakeFormCommonDown").offsetHeight) & ""
	End Function
'########################################
	Function FormatAngle_verVVB(strAngle,strOccupied_StakeoutAngle)
		'if strAngle<0 Then 
		'	signAngle = -1
		'else
		'	signAngle = 1
		'end if
		'strAngle = CStr(signAngle * strAngle)

		'MsgBox strOccupied_StakeoutAngle
		Dim DMSprecision
		Select Case AngType 
			Case "Azimuth"
				strAngle = 360 + 90 - strOccupied_StakeoutAngle
			Case "Angle Right"
				strAngle =(1) * strAngle 'Fs_az - Bs_az
			Case "Deflection"
				strAngle = strAngle + 180
			Case "South Azimuth"
				If (strOccupied_StakeoutAngle + 0) < 0 Then strOccupied_StakeoutAngle = 360 + strOccupied_StakeoutAngle
				strAngle = 270 - strOccupied_StakeoutAngle
			Case Else
				FormatAngle_verVVB = "Unknown type"
				Exit Function
		End Select

		If strAngle < 0 Then
			strAngle = strAngle + 360
		Elseif strAngle >= 360 Then
			strAngle = strAngle - 360
		End If

		Select Case UCase(AngUnit)
			Case "DEGREES DMS"
				'MsgBox strAngle
				strAngle =FormatAngle1(CStr(strAngle), "Degrees", "Degrees", "DMS","0.00000000", CStr(AngRounding))
				If Len(Cstr(AngPrecision)) = 1 Then
					'MsgBox InStr(1, strAngle, "-", 1) - 1
					strAngle  = Mid(strAngle,1,InStr(1, strAngle, "-", 1) - 1 )
				Else
					If (Len(Cstr(AngPrecision)) Mod 2 <> 0) And (Len(Cstr(AngPrecision))<6)  Then
						AngPrecision = AngPrecision & "0"
					End if
					DMSprecision = Len(Cstr(AngPrecision)) - 2 'degrees out
					IF DMSprecision = 2 Then
						strAngle  = Mid(strAngle,1,InStr(1, strAngle, "-", 1) + 2 )
					Elseif DMSprecision = 4 Then
						strAngle  = Mid(strAngle,1,InStr(1, strAngle, "-", 1) + 5 )
					Else
						strAngle  = Mid(strAngle,1,InStr(1, strAngle, "-", 1) + DMSprecision + 2 )
					End If
				End if

			Case "GRADS"
				strAngle = 360 * strAngle/400
				strAngle =FormatAngle1(CStr(strAngle), "Degrees", "Degrees", "-",CStr(AngPrecision), CStr(AngRounding))
			Case "DECIMAL DEGREES"
				'MsgBox strTemp	
				strAngle =FormatAngle1(CStr(strAngle), "Degrees", "Degrees", "-",CStr(AngPrecision), CStr(AngRounding))
			Case "RADIANS"
				strAngle =FormatAngle1(CStr(strAngle), "Degrees", "Radians", "-",CStr(AngPrecision), CStr(AngRounding))
			Case "MILS"
				strAngle = 360 * strAngleValue/6400
				strAngle =FormatAngle1(CStr(strAngle), "Degrees", "Degrees", "-",CStr(AngPrecision), CStr(AngRounding))
		End Select

	'If (signAngle = -1) Then
	'	strAngle  = "-" & strAngle
	'End If
	FormatAngle_verVVB = strAngle

	End Function
'########################################
	Function FillTableReportPoints(strPoint, strAngle, strDistance, strNorth, strEast, strZ,strOccupied_StakeoutAngle)
	On Error Resume Next
	Dim strTemp
	Dim signAngle

	strNorth= FormatNumber1(CStr(strNorth) , CStr(unit) , CStr(CoordUnit) ,  CStr(CoordPrecision) ,  CStr(CoordRounding))
	strDistance = FormatNumber1(CStr(strDistance) , CStr(unit) , CStr(DistUnit) ,  CStr(DistPrecision) ,  CStr(DistRounding))
	strEast= FormatNumber1(CStr(strEast) , CStr(unit) , CStr(CoordUnit) ,  CStr(CoordPrecision) ,  CStr(CoordRounding))
	strZ= FormatNumber1(CStr(strZ) , CStr(unit) , CStr(ElevUnit) ,  CStr(ElevPrecision) ,  CStr(ElevRounding))
	strAngle = FormatAngle_verVVB (strAngle,strOccupied_StakeoutAngle)

	
	strNorth = "" & strNorth 
	strEast = "" & strEast 
	strZ = "" & strZ 
	strDistance = "" & strDistance
	strAngle = "" & strAngle 
	
	strTemp = "<TR><TD SCOPE=""row"">" + strPoint + "</TD><TD>" + strAngle + "</TD><TD>" + strDistance + "</TD><TD>" + strNorth + "</TD><TD>" + strEast + "</TD><TD>" + strZ + "</TD></TR>"
	FillTableReportPoints = strTemp
	Err.Clear
	End Function
'#######################################
	Function GetLocalStation(strStation,strStationSuffix)
		Dim n
		For n = 0 to MaxDim(strStaInternal)
			If (strStation + 0)< (0+strStaInternal(n)) Then
				Exit For
			End If
		Next
		strStationSuffix = n + 1
		strStationSuffix = "(" & strStationSuffix & ")"		
		If n = 0 Then
			GetLocalStation = strStation
			Exit Function
		Else
			strStation = strStation - strStaInternal(n-1) + strStaAhead(n-1)
			GetLocalStation = strStation
		End If

	End Function
'######################################
	Function FillTableReportAlignment(strStation, strAngle, strDistance, strNorth, strEast,objCoordGeom,strOccupied_StakeoutAngle)
	On Error Resume Next
	Dim strTemp
	Dim signAngle,strStationPrefix,strStationSuffix
	strStationSuffix = ""
	strNorth= FormatNumber1(CStr(strNorth) , CStr(unit) , CStr(CoordUnit) ,  CStr(CoordPrecision) ,  CStr(CoordRounding))
	strDistance = FormatNumber1(CStr(strDistance) , CStr(unit) , CStr(DistUnit) ,  CStr(DistPrecision) ,  CStr(DistRounding))
	strEast= FormatNumber1(CStr(strEast) , CStr(unit) , CStr(CoordUnit) ,  CStr(CoordPrecision) ,  CStr(CoordRounding))
	strStationPrefix = GetStationPrefix(strStation,objCoordGeom)
	if Document.forms("StaEq").useStationEquation.checked Then
		strStation = GetLocalStation(strStation,strStationSuffix)
	End if
	strStation= FormatStation1(CStr(strStation) , CStr(StaDisplay) , CStr(StaPrecision) ,  CStr(StaRounding))
	strStation = strStationPrefix &  strStation & strStationSuffix

	strAngle = FormatAngle_verVVB (strAngle,strOccupied_StakeoutAngle)

	strNorth = "" & strNorth 
	strEast = "" & strEast 
	strStation = "" & strStation
	strDistance = "" & strDistance
	strAngle = "" & strAngle 
	
	strTemp = "<TR><TD SCOPE=""row"">" + strStation + "</TD><TD>" + strAngle + "</TD><TD>" + strDistance + "</TD><TD>" + strNorth + "</TD><TD>" + strEast + "</TD></TR>"
	FillTableReportAlignment = strTemp
	Err.Clear
	End Function
'######################################
	sub ChangeAppeareance
		'Msgbox CStr(StaIncrement)
		if Document.forms("StakeFormPoints").style.display = "none" then
			Document.forms("StakeFormPoints").style.display = ""		
			Document.forms("StakeFormAlignmentA").style.display = "none"	
			Document.forms("StaEq").style.display = "none"
			Document.forms("StaEq").useStationEquation.checked = False

			Document.forms("StakeFormAlignmentB").style.display = "none"	
		else
			Document.forms("StakeFormPoints").style.display = "none"		
			Document.forms("StakeFormAlignmentA").style.display = ""
			Document.forms("StakeFormAlignmentB").style.display = ""
			FillAlignmentDropBox	
		
		end if
	end sub	
'#####################################
	Function FillAlignmentDropBox
		Dim xmlDoc,objAlignment,objAlignmentList,strAlignmentName
		Dim n, newOption
		On Error Resume Next
		set xmlDoc = document.all("IslandCgPoints").XMLDocument
		Document.forms("StakeFormAlignmentA").AlignmentName.length = 0

		Set objAlignmentList = xmlDoc.selectNodes("//Alignments/Alignment")

		If objAlignmentList Is Nothing Then
			Document.forms("StakeFormAlignmentA").AlignmentName.length = 1			
			Document.forms("StakeFormAlignmentA").AlignmentName.options(0).text = "-------------none-------------"
		Else
			For n = 0 to objAlignmentList.length-1
				Set objAlignment = objAlignmentList.item(n)
				strAlignmentName = objAlignment.attributes.getNamedItem("name").nodeValue
				Document.forms("StakeFormAlignmentA").AlignmentName.length = n + 1 
				Document.forms("StakeFormAlignmentA").AlignmentName.options(n).text = strAlignmentName 
			Next
			ChangeAlignment
		End If
		If Err.Number<>0 Then
			MsgBox "No Alignments Found"
			'for future investigate pointer to algn. node
			Err.clear
			Document.forms("StakeFormAlignmentA").AlignmentName.length = 0			
			Document.forms("StakeFormAlignmentA").AlignmentName.length = 1			
			Document.forms("StakeFormAlignmentA").AlignmentName.options(0).text = "-------------none-------------"
		End If

	End Function
'####################################################
	Function FoundInPointList(strRandomPoint,strPointListFilter, strMinRange,strMaxRange,strPrefixes)
		Dim n, strPrefix,strSuffix
		FoundInPointList = False
		If (Instr(1,strPointListFilter,"," & strRandomPoint & ",",1 ) <> 0) Then
			FoundInPointList = True
			Exit Function
		End If
		GetPrefixAndSuffixFromName strRandomPoint, strPrefix, strSuffix
		For n = 0 To MaxDim(strPrefixes)
			If strPrefixes(n) = strPrefix Then
				If (strMaxRange(n) & "") = "INF" Then
					'MsgBox strSuffix & "---" & strMinRange(n) & "--" & ((strSuffix + 0) > strMinRange(n))
					If (strSuffix + 0 >= strMinRange(n) + 0) Then
						FoundInPointList = True
						Exit function
					End if
				Elseif IsNumeric(strMaxRange(n)) Then
					If (strSuffix + 0 >= strMinRange(n) +0 ) And (strSuffix + 0 <= strMaxRange(n) + 0) Then
						FoundInPointList = True
						Exit function
					End If
				End If
			End If
		Next
	End Function
'#############################################
	Function showProgress()

			'Document.forms("StakeFormCommonDown").Append.setAttribute "value","No Processing..." 
			'MsgBox "Got"
	End Function
'#############################################
	Function changeCaption()
			Document.forms("StakeFormCommonDown").Append.setAttribute "value","Processing..." 
			'MainTable.refresh
			'RepaintMe
			'Call Document.execCommand("Refresh")
			oldScrollVal = document.body.scrollTop 
			'window.onscroll = GetRef("Scrolling")
	End Function
'#############################################
	function AppendReportMain(strCase)
		On Error Resume Next
		Dim dblMaxDistance,oPopup,oPopupBody
		Dim xmlDoc
		Dim strOccupiedPoint, strRandomPoint,strBacksightPoint,strBacksightDirection
		Dim strXQL,staIndx,timeoutID,blnMessageShown

		Dim strSelectedAlignment,objAlignment,strStations(),objCoordGeom
		Dim objOccupiedNode,objRandomNode, objBacksightNode
		Dim objPointList
		Dim strPointListFilter
		Dim dblNOccupied, dblEOccupied, dblZOccupied
		Dim dblNBacksight,dblEBacksight, dblZBacksight
		Dim dblNRandom, dblERandom, dblZRandom
		Dim strReferenceDirection,strOccupied_StakeoutAngle
		Dim n, strOutputHTML, strMinRange(),strMaxRange(),strPrefixes()
		Dim strNorthDisplay, strEastDisplay, strElevationDisplay,strAngleDisplay,strDistanceDisplay 
		Dim strPoint, strAngle, strDistance, strNorth, strEast, strZ, strStation

		blnMessageShown = False		
		dblMaxDistance = GetMaxDist(Document.forms("StakeFormCommonDown").maximumDistance.Value)
		strOccupiedPoint = Document.forms("StakeFormCommonUp").occupiedPoint.Value
		strBacksightPoint = Document.forms("StakeFormCommonUp").backsightPointName.Value
		strPointListFilter = Document.forms("StakeFormPoints").stakePointList.Value
		ExpandPointListFilter strPointListFilter,strMinRange,strMaxRange,strPrefixes
		strBacksightDirection = "Unsolved"
		set xmlDoc = document.all("IslandCgPoints").XMLDocument
         strXQL = "//Points/Point[@id='" + strOccupiedPoint + "']"
         '  msgbox (xmlDoc.documentElement.childNodes.length)
		set objOccupiedNode = xmlDoc.SelectSingleNode(strXQL)
         strXQL = "//Points/Point[@id='" + strBacksightPoint + "']"
		set objBacksightNode = xmlDoc.SelectSingleNode(strXQL)
		
		if objOccupiedNode Is Nothing Then
			MsgBox "Failed to retrieve occupied point"
		else
			dblNOccupied = objOccupiedNode.attributes.getNamedItem("Northing").nodeValue
			dblEOccupied = objOccupiedNode.attributes.getNamedItem("Easting").nodeValue
			dblZOccupied = objOccupiedNode.attributes.getNamedItem("Elevation").nodeValue

			if Document.forms("StakeFormCommonUp").backsightPointOption.checked Then
				if objBacksightNode Is Nothing Then
					MsgBox "Failed to retrieve backsight point"
				Else
					dblNBacksight = objBacksightNode.attributes.getNamedItem("Northing").nodeValue
					dblEBacksight = objBacksightNode.attributes.getNamedItem("Easting").nodeValue
					dblZBacksight = objBacksightNode.attributes.getNamedItem("Elevation").nodeValue
					strBacksightDirection = GetDirection(dblNOccupied,dblEOccupied,dblNBacksight,dblEBacksight)
					'MsgBox "Backsight-->" & strBacksightDirection
				End If
			else
				strBacksightDirection = TransformBacksightDirection(Document.forms("StakeFormCommonUp").backsightDirectionAngle.Value)
				'msgbox strBacksightDirection
			end if
			
			strOutputHTML = Output.innerHTML + FillHeader(dblNOccupied,dblEOccupied,dblZOccupied,dblNBacksight,dblEBacksight,dblZBacksight,strCase)
			strNorthDisplay = "Northing" & GetUnitPrefix(CoordUnit)
			strEastDisplay = "Easting" & GetUnitPrefix(CoordUnit)
			strElevationDisplay = "Elevation" & GetUnitPrefix(ElevUnit)
			strAngleDisplay = AngType
			strDistanceDisplay = "Distance" & GetUnitPrefix(DistUnit)
			strStation = "Station"
			'Document.forms("StakeFormCommonDown").Append.fireEvent "ondrop"
			'MsgBox strOutputHTML, xmlDoc.documentElement.childNodes.length-1
			If strCase = "Points" Then

				strOutputHTML = strOutputHTML + "<TABLE border=""1"" > <COLGROUP><COLGROUP SPAN=3><THEAD>"
				strOutputHTML = strOutputHTML + "<TR><TH SCOPE=col>Point #</TH>"
				strOutputHTML = strOutputHTML + "<TH SCOPE=col>" & strAngleDisplay & "</TH>"
				strOutputHTML = strOutputHTML + "<TH SCOPE=col>" & strDistanceDisplay & "</TH>"
				strOutputHTML = strOutputHTML + "<TH SCOPE=col>" & strNorthDisplay & "</TH>"
				strOutputHTML = strOutputHTML + "<TH SCOPE=col>" & strEastDisplay & "</TH>"
				strOutputHTML = strOutputHTML + "<TH SCOPE=col>" & strElevationDisplay & "</TH></TR></THEAD><TBODY>"  
				Set objPointList = xmlDoc.selectNodes("//Points/Point")
				For n = 0 to objPointList.length-1
					Set objRandomNode = objPointList.item(n)  
					strRandomPoint = objRandomNode.attributes.getNamedItem("id").nodeValue
					dblNRandom = objRandomNode.attributes.getNamedItem("Northing").nodeValue
					dblERandom = objRandomNode.attributes.getNamedItem("Easting").nodeValue
					dblZRandom = objRandomNode.attributes.getNamedItem("Elevation").nodeValue
					If strRandomPoint <> strOccupiedPoint And (FoundInPointList(strRandomPoint,strPointListFilter, strMinRange,strMaxRange,strPrefixes) Or Document.forms("StakeFormPoints").stakePointList.Value = "") Then
						If InRadius (dblNOccupied,dblEOccupied,dblNRandom,dblERandom,dblMaxDistance) Then
							strDistance = GetDistance(dblNOccupied,dblEOccupied,dblNRandom,dblERandom)
							strAngle = GetDirection(dblNOccupied,dblEOccupied,dblNRandom,dblERandom)
							strOccupied_StakeoutAngle = strAngle
							If IsNumeric(strBacksightDirection) Then
								strAngle = strBacksightDirection - strAngle
							Else
								strAngle = "Unsolved"
							End If
							strOutputHTML =strOutputHTML + FillTableReportPoints(strRandomPoint, strAngle, strDistance, dblNRandom, dblERandom ,dblZRandom,strOccupied_StakeoutAngle)
	
						End If
					End If
				Next
			Elseif strCase = "Alignments" Then

				strOutputHTML = strOutputHTML + "<TABLE border=""1"" > <COLGROUP><COLGROUP SPAN=3><THEAD>"
				strOutputHTML = strOutputHTML + "<TR><TH SCOPE=col>" & strStation & "</TH>"
				strOutputHTML = strOutputHTML + "<TH SCOPE=col>" & strAngleDisplay & "</TH>"
				strOutputHTML = strOutputHTML + "<TH SCOPE=col>" & strDistanceDisplay & "</TH>"
				strOutputHTML = strOutputHTML + "<TH SCOPE=col>" & strNorthDisplay & "</TH>"
				strOutputHTML = strOutputHTML + "<TH SCOPE=col>" & strEastDisplay & "</TH></TR></THEAD><TBODY>" 
				strSelectedAlignment = Document.forms("StakeFormAlignmentA").AlignmentName.options(Document.forms("StakeFormAlignmentA").AlignmentName.selectedIndex).text
				strXQL = "//Alignments/Alignment[@name='" + strSelectedAlignment + "']"
				set objAlignment = xmlDoc.SelectSingleNode(strXQL)
				If Not objAlignment Is Nothing Then
					PopulateStations objAlignment,strStations
					Set objCoordGeom = objAlignment.selectSingleNode("CoordGeom")
					For n = 0 to MaxDim(strStations)
						'MsgBox CStr(strStations(n))
						dblNRandom = GetCoord("N",strStations(n),objAlignment,staIndx)
						dblERandom = GetCoord("E",strStations(n),objAlignment,staIndx)
						'MsgBox Cstr(dblERandom)
						If dblNRandom <> "Not Supported" And dblERandom <> "Not Supported" Then
							If dblNRandom <> "Error" And dblERandom <> "Error" Then
								If InRadius (dblNOccupied,dblEOccupied,dblNRandom,dblERandom,dblMaxDistance) Then
									strDistance = GetDistance(dblNOccupied,dblEOccupied,dblNRandom,dblERandom)
									strAngle = GetDirection(dblNOccupied,dblEOccupied,dblNRandom,dblERandom)
									strOccupied_StakeoutAngle = strAngle
									If IsNumeric(strBacksightDirection) Then
										strAngle = strBacksightDirection - strAngle
									Else
										strAngle = "Unsolved"
									End If
									strOutputHTML =strOutputHTML + FillTableReportAlignment(strStations(n), strAngle, strDistance, dblNRandom, dblERandom,objCoordGeom,strOccupied_StakeoutAngle)
			
								End If
									
							Else
								MsgBox "Error for station: " & strStations(n)
								Exit For
							End If
						Else
							If Not blnMessageShown Then
								blnMessageShown = True
								MsgBox "Spiral type not supported encountered." & Chr(13) &  "All spiral not supported will be skipped."
							End If
						End If
					Next 
				End If
			End If
			strOutputHTML = strOutputHTML + "</TBODY></TABLE></BODY></HTML>"
			
			output.innerHTML = strOutputHTML
			
		end if
		Document.forms("StakeFormCommonDown").Append.Value = "Append to report"
		
		If Err.Number<> 0 Then
			MsgBox "Error creating report" 
		End If
	end function	

'#######################################
	Function SaveReportA
		On Error Resume next
		Dim str
		str = CStr(reportHTML.outerHTML)
		Launch str & ""


		If Err.number <> 0 then
			MsgBox "Error saving report!"	& Err.number		
		End if
	End Function
'######################################
	Function SaveReportB
		On Error Resume next
   		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		Dim fso, f
		Dim strLocation
		strLocation = InputBox("Please specify full path!","Path")
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile(strLocation, ForWriting, True)
   		f.Write reportHTML.outerHTML
		f.Close

		If Err.number <> 0 then
			MsgBox "Error saving report!" & char(13) &  "Please make sure the folder specified exists already."	
		Else
			Open strLocation		
		End if
	End Function
'######################################
	Function ChangeAlignment
		Dim strAlignment, objAlignment,strXQL,objStaEqList,n,objStaEq,strBeginLocal,strEndLocal,strBeginRaw,strEndRaw
		Dim staInternal,staBack,staAhead,strStaEq,strStartStation,strEndStation,objCoordGeom
		strAlignment = Document.forms("StakeFormAlignmentA").AlignmentName.options(Document.forms("StakeFormAlignmentA").AlignmentName.selectedIndex).text
		strXQL = "//Alignments/Alignment[@name='" + strAlignment + "']"
		Erase strStaInternal
		Erase strStaBack
		Erase strStaAhead

		set objAlignment = document.all("IslandCgPoints").XMLDocument.SelectSingleNode(strXQL)
		Document.forms("StaEq").style.display = "none"	
		'Document.forms("StaEq").StationEquationListTest.length = 0
		Document.forms("StakeFormAlignmentB").StationEquationList.value = ""
		Document.forms("StaEq").useStationEquation.checked = False
		If Not objAlignment Is Nothing Then
			Set objCoordGeom = objAlignment.selectSingleNode("CoordGeom")
			'MsgBox objCoordGeom.childNodes.Length - 1
			strStartStation = objCoordGeom.childNodes(0).attributes.getNamedItem("StartStation").nodeValue
			strEndStation = objCoordGeom.childNodes(objCoordGeom.childNodes.Length - 1).attributes.getNamedItem("EndStation").nodeValue
			Set objStaEqList = objAlignment.SelectNodes("StaEquation")
			Document.forms("StakeFormAlignmentB").StationEquationList.value = "#" & Space(6) & "Begin" & Space(40-Len("Begin")) & "End" & chr(13)

			If objStaEqList.length > 0 Then
				Document.forms("StaEq").style.display = ""	
								
				staBeginRaw = strStartStation
				
				For n = 0 to objStaEqList.length-1
					Set objStaEq = objStaEqList.item(n)
					staInternal = objStaEq.attributes.getNamedItem("staInternal").nodeValue
					staBack = objStaEq.attributes.getNamedItem("staBack").nodeValue
					staAhead = objStaEq.attributes.getNamedItem("staAhead").nodeValue

					Redim Preserve strStaInternal(MaxDim(strStaInternal)+1)
					strStaInternal(MaxDim(strStaInternal)) = staInternal
					Redim Preserve strStaAhead(MaxDim(strStaAhead)+1)
					strStaAhead(MaxDim(strStaAhead)) = staAhead
					Redim Preserve strStaBack(MaxDim(strStaBack)+1)
					strStaBack(MaxDim(strStaBack)) = staBack +0

					If n = 0 Then
						strBeginLocal = strStartStation
						strEndLocal = strStaInternal(n)
						strBeginRaw = strStartStation
						strEndRaw = strStaInternal(n)
					Else
						strBeginLocal = strStaAhead(n-1)
						strEndLocal = (strStaAhead(n-1)+0) + (strStaInternal(n)+0) - (strStaInternal(n-1)+0)
						strBeginRaw = strStaInternal(n-1)
						strEndRaw = strStaInternal(n)

					End if

					strBeginLocal= FormatStation1(CStr(strBeginLocal) , CStr(StaDisplay) , CStr(StaPrecision) ,  CStr(StaRounding))
					strEndLocal = FormatStation1(CStr(strEndLocal ) , CStr(StaDisplay) , CStr(StaPrecision) ,  CStr(StaRounding))
					strBeginRaw = FormatStation1(CStr(strBeginRaw ) , CStr(StaDisplay) , CStr(StaPrecision) ,  CStr(StaRounding))
					strEndRaw = FormatStation1(CStr(strEndRaw ) , CStr(StaDisplay) , CStr(StaPrecision) ,  CStr(StaRounding))

					If n = 0 Then
						strStaEq = "(" & (n+1) & ")" & Space(5-Len(CStr(n+1))) & strBeginLocal & Space(40-Len(CStr(strBeginLocal))) & strEndLocal
					Else
						strStaEq = "(" & (n+1) & ")" & Space(5-Len(CStr(n+1))) & strBeginLocal & " [" & strBeginRaw & "]" & Space(40-Len(CStr(strBeginLocal & " [" & strBeginRaw & "]"))) &  strEndLocal & " [" & strEndRaw & "]  "

					End If

					'Document.forms("StaEq").StationEquationListTest.length = n + 1 
					'Document.forms("StaEq").StationEquationListTest.options(n).text = strStaEq 
					'msgbox strStaEq
					Document.forms("StakeFormAlignmentB").StationEquationList.value = Document.forms("StakeFormAlignmentB").StationEquationList.value & strStaEq & chr(13)

				Next

				'add last region if is the case
				strBeginLocal = strStaAhead(n-1)
				strEndLocal = (strStaAhead(n-1) +0) + (strEndStation +0 )- (strStaInternal(n-1)-0)
				strBeginRaw = strStaInternal(n-1)
				strEndRaw = strEndStation 
				strBeginLocal= FormatStation1(CStr(strBeginLocal) , CStr(StaDisplay) , CStr(StaPrecision) ,  CStr(StaRounding))
				strEndLocal = FormatStation1(CStr(strEndLocal ) , CStr(StaDisplay) , CStr(StaPrecision) ,  CStr(StaRounding))
				strBeginRaw = FormatStation1(CStr(strBeginRaw ) , CStr(StaDisplay) , CStr(StaPrecision) ,  CStr(StaRounding))
				strEndRaw = FormatStation1(CStr(strEndRaw ) , CStr(StaDisplay) , CStr(StaPrecision) ,  CStr(StaRounding))

				strStaEq = "(" & (n+1) & ")" & Space(5-Len(CStr(n+1))) & strBeginLocal & " [" & strBeginRaw & "]" & Space(40-Len(CStr(strBeginLocal & " [" & strBeginRaw & "]"))) &  strEndLocal & " [" & strEndRaw & "]  "
				Document.forms("StakeFormAlignmentB").StationEquationList.value = Document.forms("StakeFormAlignmentB").StationEquationList.value & strStaEq & chr(13)

			Else 'no station equation
				n = 0
				strBeginLocal= strStartStation
				strEndLocal = strEndStation
				strBeginLocal= FormatStation1(CStr(strBeginLocal) , CStr(StaDisplay) , CStr(StaPrecision) ,  CStr(StaRounding))
				strEndLocal = FormatStation1(CStr(strEndLocal ) , CStr(StaDisplay) , CStr(StaPrecision) ,  CStr(StaRounding))

				strStaEq = "(" & (n+1) & ")" & Space(5-Len(CStr(n+1))) & strBeginLocal & Space(40-Len(CStr(strBeginLocal))) & strEndLocal
				Document.forms("StakeFormAlignmentB").StationEquationList.value = Document.forms("StakeFormAlignmentB").StationEquationList.value & strStaEq & chr(13)

			End If
		End If

	End Function
'#####################################
	Function InStationRange(strRandomStation,strMinRangeStations,strMaxRangeStations)
		
		If (Document.forms("StakeFormAlignmentB").stakeStationList.Value ="") Then 'No Range Stations
			InStationRange = True
		Else
			If (MaxDim(strMinRangeStations)<0) Then
				InStationRange = False
			Else

			For n = 0 To Ubound(strMinRangeStations)
				If (strMinRangeStations(n)-strRandomStation<=0) And (strMaxRangeStations(n) - strRandomStation >= 0) Then
					InStationRange = True
					Exit For
				End If

			Next
			End If
		End If
		
	End Function
'######################################
	Function FillRangeStations (strMinRangeStations,strMaxRangeStations,strOddStations,strStartStation,strEndStation)
		On Error Resume Next
		Dim strCompressRange
		Dim strArrA,strArrB
		Dim n,m
		Erase strMinRangeStations
		Erase strMaxRangeStations
		Erase strOddStations
		
		strCompressRange = "," & Document.forms("StakeFormAlignmentB").stakeStationList.Value & ","
		strCompressRange = Replace(strCompressRange,"+","",1,-1,1 )

		strArrA = Split(strCompressRange,",",-1,1)
		For n = 0 to UBound(strArrA)
			If Left(strArrA(n),1) = "-" Then
				strArrA(n) = "<minus>" & Mid(strArrA(n),2)
			End If
			strArrA(n) = Replace(strArrA(n),"--","-<minus>",1,-1,1 )
			If InStr(1, strArrA(n), "-", 1) <> 0 Then
				strArrB = Split(strArrA(n),"-",-1,1)
				strArrB(0) = Replace(strArrB(0),"<minus>","-",1,-1,1 )
				strArrB(1) = Replace(strArrB(1),"<minus>","-",1,-1,1 )
				strArrB(0) = GetGlobalStation(strArrB(0),strStartStation,strEndStation)
				strArrB(1) = GetGlobalStation(strArrB(1),strStartStation,strEndStation)
				If IsNumeric(strarrB(0)) And IsNumeric(strarrB(1)) Then
					If strArrB(0) - strArrB(1) <=0 Then
						Redim Preserve strMinRangeStations(MaxDim(strMinRangeStations)+1)
						strMinRangeStations(MaxDim(strMinRangeStations))= strArrB(0)	

						Redim Preserve strMaxRangeStations(MaxDim(strMaxRangeStations)+1)
						strMaxRangeStations(MaxDim(strMaxRangeStations))= strArrB(1)	
					
					End If
				End If
			Else
				If strArrA(n)<>"" Then
					strArrA(n) = GetGlobalStation(strArrA(n),strStartStation,strEndStation)
					If IsNumeric(strArrA(n)) Then
						'550-950,1010,1010-1090,1020,1030,1050-1070,1011
						Redim Preserve strOddStations(MaxDim(strOddStations)+1)
						strOddStations(MaxDim(strOddStations))	= strArrA(n) + 0
	
					End If		
				End If
			End If
		Next

		'For n = 0 to MaxDim(strOddStations)
			'MsgBox strOddStations(n)
		'Next

		If Err.Number <> 0 Then
			MsgBox "Error reading Range Stations"
		End If
	End Function
'#######################################
	Function GetGlobalStationFirstRelease(strStation,strStartStation,strEndStation)
		Dim strRegion,strSta,strTemp

		'23+444(9)-55+66(7),66+33(4)
		strTemp = strStation
		GetGlobalStationFirstRelease = "NonValid"
		If (InStr(1, strStation, ")", 1) <> 0) And InStr(1, strStation, "(", 1) <> 0 Then
			strRegion = Mid(strStation, InStr(1, strStation, "(", 1) + 1 )
			strRegion = Replace(strRegion,")","",1,-1,1 )
			strSta = Mid(strStation,1, InStr(1, strStation, "(", 1) - 1 )
			If Not IsNumeric(strRegion) then
				MsgBox "Error reading station: " & strStation
			Else
				If ((strRegion + 0) > (MaxDim(strStaAhead) + 2 )) Or ((strRegion + 0) <=0) Then
					MsgBox "Station region: (" & strRegion & ") was not found."
				Elseif (strRegion+0) = 1 Then
					GetGlobalStationFirstRelease = strSta
				Else
					'get the global (no checking if is valid and not finished by another equation
					strStation = strStaInternal(strRegion-2) + (strSta - strStaAhead(strRegion-2))
					'MsgBox strStaInternal(strRegion-2) & "==" & strSta & "==" & strStaAhead(strRegion-2)
					'MsgBox strStation
					If ((strRegion + 0) = (MaxDim(strStaAhead) + 2 )) Then
						'let it go, is the last sta eq, don't check
						GetGlobalStationFirstRelease = strStation
					Else
						'check if next staEquation is validating this local station 
						If (strSta + 0)<= strStaInternal(strRegion-1) - strStaInternal(strRegion-2) +  strStaAhead(strRegion-2) Then
							GetGlobalStationFirstRelease = strStation
						Else
							MsgBox strTemp & " can not be found!"
						End If
					End If
				End if
			End If
		Elseif IsNumeric(strStation) Then
			GetGlobalStationFirstRelease = strStation 'Global
		Else
			MsgBox "Error reading station: " & strStation
		End If
		'MsgBox strTemp & " ###### " & strStation
	End Function
'#######################################
	Function GetGlobalStation(strStation,strStartStation,strEndStation)
		Dim strRegion,strSta,strTemp

		'23+444(9)-55+66(7),66+33(4)
		
		GetGlobalStation = "NonValid"
		'if don't use sta eq:
		if Not Document.forms("StaEq").useStationEquation.checked Then
			If IsNumeric(strStation) Then
				GetGlobalStation = strStation 'Global
			Else
				MsgBox "Error reading station: " & strStation & chr(13) & "Please check the checkbox if you want to use Station equations"
			End if
			Exit Function
		End if
		strTemp = strStation

		'if use station equation :
		If IsNumeric(strStation) Then
			'strStation  =  strStation & "(1)" 'assume region 1
			If (strStation + 0)<= (strEndStation + 0) Then
				If (strStation + 0) >= (strStartStation + 0) Then
					GetGlobalStation = strStation
				Else
					MsgBox strTemp & " can not be found!"
				End If
			Else
				MsgBox strTemp & " can not be found!"
			End If
		End If

		If (InStr(1, strStation, ")", 1) <> 0) And InStr(1, strStation, "(", 1) <> 0 Then
			strRegion = Mid(strStation, InStr(1, strStation, "(", 1) + 1 )
			strRegion = Replace(strRegion,")","",1,-1,1 )
			strSta = Mid(strStation,1, InStr(1, strStation, "(", 1) - 1 )
			If Not IsNumeric(strRegion) then
				MsgBox "Error reading station: " & strStation
			Else
				If ((strRegion + 0) > (MaxDim(strStaAhead) + 2 )) Or ((strRegion + 0) <=0) Then
					MsgBox "Station region: (" & strRegion & ") was not found."
				Elseif (strRegion+0) = 1 Then
						'MsgBox strSta & "===" & strStaInternal(strRegion-1)
						If (strSta + 0)<= (strStaInternal(strRegion-1)+0) Then
							If (strSta + 0) >= (strStartStation + 0) Then
								GetGlobalStation = strSta
							Else
								MsgBox strTemp & " can not be found!"
							End If
						Else
							MsgBox strTemp & " can not be found!"
						End If
				Else
					'get the global (no checking if is valid and not finished by another equation
					strStation = strStaInternal(strRegion-2) + (strSta - strStaAhead(strRegion-2))
					'MsgBox strStaInternal(strRegion-2) & "==" & strSta & "==" & strStaAhead(strRegion-2)
					'MsgBox strStation
					If ((strRegion + 0) = (MaxDim(strStaAhead) + 2 )) Then
						'let it go, is the last sta eq if not greater than the end station
							If (strStation + 0) <= (strEndStation + 0) Then
								GetGlobalStation = strStation
							Else
								MsgBox strTemp & " can not be found!"
							End If
					Else
						'check if next staEquation is validating this local station 
						If (strSta + 0)<= strStaInternal(strRegion-1) - strStaInternal(strRegion-2) +  strStaAhead(strRegion-2) Then
							GetGlobalStation = strStation
						Else
							MsgBox strTemp & " can not be found!"
						End If
					End If
				End if
			End If
		Else
			MsgBox "Error reading station: " & strStation
		End If
		'MsgBox strTemp & " ###### " & strStation
	End Function
'######################################
	Function MaxDim(strArray)
		on Error Resume Next
		MaxDim = Ubound(strArray)

		If Err.Number <>0 Then
			Err.clear
			MaxDim = -1
		End If
	End Function
'######################################
    Function EvenStation(dblStation, increm)
        Dim dblTemp, dblFraction

	If dblStation < 0 Then
            dblTemp = (-1) * Fix(dblStation)
            dblFraction = dblStation + dblTemp
            dblTemp = dblTemp - Fix(dblTemp / increm) * increm
            dblStation = dblStation + dblTemp - dblFraction
        Else
            dblTemp = Fix(dblStation)
            dblFraction = dblStation - dblTemp
            dblTemp = dblTemp - Fix(dblTemp / increm) * increm
            dblStation = dblStation - dblTemp - dblFraction
        End If
        EvenStation = dblStation
    End Function
'#############################################
	Function ValidStation(strStationTemp,strRegion,strStartStation,strEndStation)
		
		Dim blnValid,strStation, strSta
		blnValid = False

		'get the global (no checking if is valid and not finished by another equation
		strSta = strStationTemp

		if (strRegion+0) = 1 Then
			If (strSta + 0)< (strStaInternal(strRegion-1)+0) Then
				If (strSta + 0) >= (strStartStation + 0) Then
					blnValid = True
				End If
			End If
		Else
			strStation = strStaInternal(strRegion-2) + (strSta - strStaAhead(strRegion-2))

			If ((strRegion + 0) = (MaxDim(strStaAhead) + 2 )) Then
				'let it go, is the last sta eq if not greater than the end station
				If (strStation + 0) <= (strEndStation + 0) Then
					blnValid = True
				End If
			Else
				'check if next staEquation is validating this local station 
				If (strSta + 0)< strStaInternal(strRegion-1) - strStaInternal(strRegion-2) +  strStaAhead(strRegion-2) Then
					blnValid = True
				End If
			End If
		End If

		ValidStation = blnValid
	End Function
'######################################
	Function EvenStationLocally (dblStation,StaIncrement,strStationSuffix,strStartStation,strEndStation) 
		On Error Resume Next
		Dim strLocalStation,strStationTemp,strNewStationSuffix
		if Not Document.forms("StaEq").useStationEquation.checked Then
			EvenStationLocally = dblStation
			Exit function
		end if

		strStationTemp = dblStation
		strLocalStation = GetLocalStation(strStationTemp,strNewStationSuffix)
		If (Trim(strNewStationSuffix & "") = Trim(strStationSuffix & "")) Then
			EvenStationLocally = dblStation
		Else
			strNewStationSuffix = strStationSuffix
			strNewStationSuffix = Replace(strNewStationSuffix,"(","",1,-1,1 )
			strNewStationSuffix = Replace(strNewStationSuffix,")","",1,-1,1 )
			'check how many stations I jumped over, force it to not be more than one
			strNewStationSuffix  = strNewStationSuffix + 1
			strStationSuffix = "(" & (strNewStationSuffix +0 ) & ")"

			strStationTemp = strStaAhead(strNewStationSuffix-2)
			EvenStation strStationTemp,StaIncrement
			'MsgBox strNewStationSuffix
			If (strStationTemp + 0)< (strStaAhead(strNewStationSuffix-2) +0) Then
				strStationTemp = strStationTemp + StaIncrement
			End If
			If ValidStation(strStationTemp,strNewStationSuffix,strStartStation,strEndStation) then			
				strStationTemp = strStationTemp & strStationSuffix
				'mSGbOX strStationTemp & "====" & strStationSuffix			

				strStationTemp = GetGlobalStation (strStationTemp,strStartStation,strEndStation )
				EvenStationLocally = strStationTemp
				'mSGbOX strStationTemp
			Else
				'increase region, but check first if is not the last one
				if ((strNewStationSuffix + 0) = (MaxDim(strStaAhead) + 2 )) Then
					'output the last global station, will be filtered out by calling procedure
					EvenStationLocally = strEndStation
				Else
					strNewStationSuffix = strNewStationSuffix + 1
					strStationTemp = strStaAhead(strNewStationSuffix-2)
					strStationTemp = strStationTemp & "(" & strNewStationSuffix & ")"
					strStationTemp = GetGlobalStation (strStationTemp,strStartStation,strEndStation )
					EvenStationLocally = EvenStationLocally (strStationTemp,StaIncrement,strStationSuffix,strStartStation,strEndStation) 
				End If
			End If
		End If
		If Err.Number <> 0 Then
			MsgBox "Error Evening stations locally"
		End if
	End Function
'#####################################
    Function GM_LocateOnSpiral(l, xsp, ysp, azm, spiralLength, spiralXStart, spiralYStart, spiralDirectionBack, spiralRadius1, spiralRadius2, spiralDirection)
    
    Dim u(8), p(8), i
    Dim x, y, xw, yw, az90, th,D1,D2
    Const PTTOL = 0.000001

    If Not IsNumeric(spiralRadius1) Then
    	spiralRadius1 = -1
    End If
    If Not IsNumeric(spiralRadius2) Then
    	spiralRadius2 = -1
    End If
    If spiralRadius1 + 0 > 0 Then
	If (unit = "meter") Then
		D1 = 30.480/spiralRadius1
	Else
		D1 = 100/spiralRadius1
	End If
    Else
	D1 = 0
    End If

    If spiralRadius2 + 0 > 0 Then
	If (unit = "meter") Then
		D2 = 30.480/spiralRadius2
	Else
		D2 = 100/spiralRadius2
	End If
    Else
	D2 = 0
    End If


    spiralLength = spiralLength + 0 
    spiralXStart = spiralXStart + 0
    spiralYStart = spiralYStart+ 0  
    spiralDirectionBack = spiralDirectionBack + 0
    D1 = D1 +0 
    D2 = D2 +0 
    spiralDirection = spiralDirection + 0     
    If ((l < -PTTOL) Or ((l - spiralLength) > PTTOL)) Then
        GM_LocateOnSpiral = False
        Exit Function
    End If
    
    If (Abs(l) < PTTOL) Then
        l = 0
    Else
        If (Abs(l - spiralLength) < PTTOL) Then
            l = spiralLength
        End If
    End If
    
    'modifications for handling zero-length spiral -- 11/08/98
    If (Abs(spiralLength) < PTTOL) Then
        If (Abs(l) < PTTOL) Then
            xsp = spiralXStart
            ysp = spiralYStart
            azm = spiralDirectionBack
            GM_LocateOnSpiral = True
            Exit Function
        End If
        GM_LocateOnSpiral = False
        Exit Function
     End If
    
    If (unit = "meter") Then
        u(1) = l * l * (D2 - D1) / (2 * 30.48 * spiralLength)
        p(1) = D1 * l / 30.48
    Else
        u(1) = l * l * (D2 - D1) / (200 * spiralLength)
        p(1) = D1 * l / 100
    End If
    
    For i = 2 To 7
        u(i) = u(i - 1) * u(1)
        p(i) = p(i - 1) * p(1)
    Next
    
    y = spiralDirection * l * (u(1) / 3 + p(1) / 2 -          (1 / 6) * (u(3) / 7 + u(2) * p(1) / 2 + 3 * u(1) * p(2) / 5 + p(3) / 4) +          (1 / 120) * (u(5) / 11 + u(4) * p(1) / 2 + 10 * u(3) * p(2) / 9 + 5 / 4 * u(2) * p(3) +                       5 / 7 * u(1) * p(4) + p(5) / 6) -          (1 / 5040) * (u(7) / 15 + u(6) * p(1) / 2 + 21 * u(5) * p(2) / 13 + 35 * u(4) * p(3) / 12 +                       35 * u(3) * p(4) / 11 + 21 * u(2) * p(5) / 10 + 7 * u(1) * p(6) / 9 + p(7) / 8))
    x = l * (1 - (1 / 2) * (u(2) / 5 + u(1) * p(1) / 2 + p(2) / 3) +        (1 / 24) * (u(4) / 9 + u(3) * p(1) / 2 + 6 * u(2) * p(2) / 7 + 2 * u(1) * p(3) / 3 +                      p(4) / 5) -          (1 / 720) * (u(6) / 13 + u(5) * p(1) / 2 + 15 * u(4) * p(2) / 11 + 2 * u(3) * p(3) +                      5 * u(2) * p(4) / 3 + 3 * u(1) * p(5) / 4 + p(6) / 7))
    
    If (unit = "meter") Then
        th = spiralDirection * l / 30.48 * (D1 + 0.5 * l / spiralLength * (D2 - D1))
    Else
        th = spiralDirection * l / 100 * (D1 + 0.5 * l / spiralLength * (D2 - D1))
    End If
    xw = x '+ s.dir * s.offset * sin(th);
    yw = y '- s.dir * s.offset * cos(th);
    
    az90 = PI / 2 - spiralDirectionBack
    
    xsp = xw * Cos(az90) + yw * Sin(az90) + spiralXStart
    ysp = xw * Sin(az90) - yw * Cos(az90) + spiralYStart
    
    azm = spiralDirectionBack + th
    If (azm < 0) Then azm = azm + 2 * PI
    If (azm > 2 * PI) Then azm = azm - 2 * PI
    GM_LocateOnSpiral = True
    
    End Function
'#############################################
	Function PopulateStations (objAlignment,strStations)
		On Error Resume Next
		Dim strStartStation
		Dim strEndStation
		Dim objCoordGeom
		Dim strMinRangeStations(),strMaxRangeStations()
		Dim strOddStations(),strTempStations()
		Dim dblStation,n,m,strStationSuffix
		If StaIncrement <= 0 Then
			Erase strStations
			Exit Function
		End If
		Set objCoordGeom = objAlignment.selectSingleNode("CoordGeom")
		'MsgBox objCoordGeom.childNodes.Length - 1
		strStartStation = objCoordGeom.childNodes(0).attributes.getNamedItem("StartStation").nodeValue
		strEndStation = objCoordGeom.childNodes(objCoordGeom.childNodes.Length - 1).attributes.getNamedItem("EndStation").nodeValue
		Erase strStations
		strStationSuffix = "(1)"
		FillRangeStations strMinRangeStations,strMaxRangeStations,strOddStations,strStartStation,strEndStation

		If InStationRange(strStartStation,strMinRangeStations,strMaxRangeStations) Then		
			Redim Preserve strTempStations(0)
			strTempStations(0) = strStartStation 
		End If
		dblStation  = strStartStation + 0
		
		EvenStation dblStation,StaIncrement

		If dblStation  > (strStartStation + 0) Then
			dblStation = dblStation - StaIncrement
		End If

		Do While dblStation < (strEndStation - StaIncrement)
			dblStation = dblStation + StaIncrement
			dblStation = EvenStationLocally (dblStation,StaIncrement,strStationSuffix,strStartStation,strEndStation) 
			If (dblStation + 0) < (strEndStation) Then 		
				If InStationRange(dblStation,strMinRangeStations,strMaxRangeStations) Then
					Redim Preserve strTempStations(MaxDim(strTempStations) + 1)
					strTempStations (MaxDim(strTempStations)) = dblStation			
				End If
			End If
		Loop

		If InStationRange(strEndStation,strMinRangeStations,strMaxRangeStations) Then
			If strTempStations(Ubound(strTempStations))< strEndStation Then
				Redim Preserve strTempStations(MaxDim(strTempStations) + 1)
				strTempStations (MaxDim(strTempStations)) = strEndStation
			End If
		End If
		SortArray strOddStations
		'transfer the array, filling odd stations in the same time
		For n = 0 to MaxDim(strTempStations)
			For m = 0 to MaxDim(strOddStations)
				If strOddStations(m) <> "-" Then
					If strOddStations(m) - strTempStations(n) < 0 Then
						Redim Preserve strStations(MaxDim(strStations) + 1)
						strStations (MaxDim(strStations)) = strOddStations(m)
						strOddStations(m) = "-"
					End If			
				End If
			Next

			Redim Preserve strStations(MaxDim(strStations) + 1)
			strStations (MaxDim(strStations)) = strTempStations(n)

		Next
		'Fill the odd stations left (greater than any range specified)
		For m = 0 to MaxDim(strOddStations)
			If strOddStations(m) <> "-" Then
				Redim Preserve strStations(MaxDim(strStations) + 1)
				strStations (MaxDim(strStations)) = strOddStations(m)
				strOddStations(m) = "-"
			End If
		Next
	End Function
'#######################################
	Function SortArray(avarIn)
	   
 	  Dim intLowBounds
	  Dim intHighBounds
	  Dim intX
	  Dim intY
	  Dim varTmp
	  If MaxDim(avarIn)<0 Then
		Exit Function
	  End If
	  intLowBounds = LBound(avarIn)
	  intHighBounds = UBound(avarIn)

	  For intX = intLowBounds To (intHighBounds-1)
	    For intY = (intX + 1) To intHighBounds
	      If (avarIn(intX) > avarIn(intY)) Then
	        varTmp = avarIn(intX)
	        avarIn(intX) = avarIn(intY)
	        avarIn(intY) = varTmp
	      End If
	    Next 
	  Next 

	End function
'########################################
	Function GetCoord(strCoordinateType,strStation,objAlignment,staIndx)
		
		Select Case strCoordinateType
		Case "N"
			GetCoord = GetNorthingAtSta (strStation,objAlignment,staIndx )
		Case "E"
			GetCoord = GetEastingAtSta (strStation,objAlignment,staIndx )
		End Select
		
	End Function
'##########################################
	Function GetStationIndex(strSta,objAlignment)
		Dim strStartStation
		Dim strEndStation
		Dim objCoordGeom,n
		GetStationIndex = -1
		Set objCoordGeom = objAlignment.selectSingleNode("CoordGeom")
		'MsgBox objCoordGeom.childNodes.Length - 1
		For n = 0 To objCoordGeom.childNodes.Length - 1
			strStartStation = objCoordGeom.childNodes(n).attributes.getNamedItem("StartStation").nodeValue
			strEndStation = objCoordGeom.childNodes(n).attributes.getNamedItem("EndStation").nodeValue

			If (strStartStation - strSta <= 0) And (strEndStation - strSta>=0 ) Then
				GetStationIndex = n
				Exit For
			End If
		Next
		
	End Function
'#############################################
	Function GetStationPrefix(strSta,objCoordGeom)
		Dim n
		Dim strStartStation
		Dim strEndStation
		Dim blnFound
		blnFound = False
		
		For n = 0 To objCoordGeom.childNodes.Length - 1
			strStartStation = objCoordGeom.childNodes(n).attributes.getNamedItem("StartStation").nodeValue
			strEndStation = objCoordGeom.childNodes(n).attributes.getNamedItem("EndStation").nodeValue
			If (abs(strStartStation - strSta) <= (tolerance+0)) Then
				'MsgBox tolerance & ",,,,," & abs(strStartStation - strSta)
				blnFound = True

				If n = 0 Then
					Select Case UCase(objCoordGeom.childNodes(n).nodeName)
						Case "SPIRAL"
							GetStationPrefix = TS & ":"	
						Case "CURVE"
							GetStationPrefix = PC & ":"	
					End Select
				Else
					Select Case UCase(objCoordGeom.childNodes(n).nodeName)
						Case "LINE"
							Select Case UCase(objCoordGeom.childNodes(n-1).nodeName)
								Case "CURVE"
									GetStationPrefix = PT & ":"
								Case "SPIRAL"
									GetStationPrefix = ST & ":"
							End Select
						Case "SPIRAL"
							Select Case UCase(objCoordGeom.childNodes(n-1).nodeName)
								Case "LINE"
									GetStationPrefix = TS & ":"
								Case "CURVE"
									GetStationPrefix = CS & ":"
								Case "SPIRAL"
									GetStationPrefix = SS & ":"
							End Select

						Case "CURVE"
							Select Case UCase(objCoordGeom.childNodes(n-1).nodeName)
								Case "LINE"
									GetStationPrefix = PC & ":"
								Case "CURVE"
									If objCoordGeom.childNodes(n-1).attributes.getNamedItem("rotation").nodeValue = objCoordGeom.childNodes(n).attributes.getNamedItem("rotation").nodeValue Then
										If objCoordGeom.childNodes(n-1).attributes.getNamedItem("radius").nodeValue <> objCoordGeom.childNodes(n).attributes.getNamedItem("radius").nodeValue Then
										'MsgBox objCoordGeom.childNodes(n-1).attributes.getNamedItem("radius").nodeValue
										'MsgBox objCoordGeom.childNodes(n).attributes.getNamedItem("radius").nodeValue

											GetStationPrefix = PCC & ":"
										End If
									Else
										GetStationPrefix = PRC & ":"
									End If
								Case "SPIRAL"
									GetStationPrefix = SC & ":"
							End Select
					End Select
				End if
				'If strSta = 38991 then
				'msgbox UCase(objCoordGeom.childNodes(n-1).nodeName) & "---" & UCase(objCoordGeom.childNodes(n).nodeName)
				'End if
				Exit For
			End If
		Next
		'check the end
		If Not blnFound Then
			If (abs(strEndStation - strSta) <= (tolerance+0)) Then
				blnFound = True
				Select Case UCase(objCoordGeom.childNodes(objCoordGeom.childNodes.Length - 1).nodeName)
					Case "SPIRAL"
						GetStationPrefix = ST & ":"	
					Case "CURVE"
						GetStationPrefix = PT & ":"	
				End Select
			End If
		End If	
		'MsgBox 	GetStationPrefix
	End Function
'##########################################
function GetNorthingAtSta(strStation,objAlignment,staIndx)
	On Error Resume Next
	Dim north,eleIndex,dLen,spN,delta,length,theta
	Dim spiralXStart,spiralYStart,spiralDirection,spiralDirectionBack,PIN,PIE,dblAngleTemp
	Dim spiralLength,spiralRadius1,spiralRadius2,y,x,azm,spiType

	Dim rad,cpN,sLen,sRad,Temp1, Temp2, Temp3
	Dim fSta, theEle
	fSta = strStation + 0
	staIndx = GetStationIndex(fSta,objAlignment)
	
	if(staIndx >= 0) Then
		eleIndex = staIndx
	else
		GetNorthingAtSta = "Error"
		Exit Function
	End If	
	Set theEle = objAlignment.selectSingleNode("CoordGeom").childNodes(eleIndex)
	dLen = fSta - theEle.attributes.getNamedItem("StartStation").nodeValue

		if(UCase(theEle.nodeName) = "LINE") Then
			spN = theEle.selectSingleNode("Coordinates").attributes.getNamedItem("StartNorth").nodeValue
			delta = theEle.attributes.getNamedItem("angle").nodeValue
			north = spN + (dLen * sin(((delta) * PI) / 180 )) 

		elseif (UCase(theEle.nodeName) = "CURVE") Then
			delta = Abs(theEle.attributes.getNamedItem("delta").nodeValue)
			length = Abs(theEle.attributes.getNamedItem("length").nodeValue)
			If theEle.attributes.getNamedItem("rotation").nodeValue = "ccw" then
				theta = delta * (dLen / length) + theEle.attributes.getNamedItem("startDirection").nodeValue
			ElseIf theEle.attributes.getNamedItem("rotation").nodeValue = "cw" then
				theta =  theEle.attributes.getNamedItem("startDirection").nodeValue - delta * (dLen / length) 
			End if
			rad = theEle.attributes.getNamedItem("radius").nodeValue
			cpN = theEle.attributes.getNamedItem("centerN").nodeValue

	'if fSta = 825 Then
	'	MsgBox "Expr :::"   & ((-1) * (dLen * cos((delta * PI) / 180 )) )
	'	MsgBox "1000 - spN: " & spN
	'	MsgBox "1000 - dLen: " & dLen
	'	MsgBox "1000 - delta: " & delta
	'	MsgBox "theta: " & theta

	'End if
			north = cpN + (rad * sin((theta *PI) / 180) )
		elseif(UCase(theEle.nodeName) = "SPIRAL") Then
			spiType = theEle.attributes.getNamedItem("spiType").nodeValue
			If spiType & "" <> "clothoid" then 
				GetNorthingAtSta = "Not Supported"
				Exit Function
			End If
			spiralLength = theEle.attributes.getNamedItem("length").nodeValue
			spiralRadius1 = theEle.attributes.getNamedItem("startRadius").nodeValue
			spiralRadius2 = theEle.attributes.getNamedItem("endRadius").nodeValue
			spiralYStart = theEle.selectSingleNode("Coordinates").attributes.getNamedItem("StartNorth").nodeValue
			spiralXStart = theEle.selectSingleNode("Coordinates").attributes.getNamedItem("StartEast").nodeValue
			If theEle.attributes.getNamedItem("rotation").nodeValue = "ccw" then
				spiralDirection = -1
			ElseIf theEle.attributes.getNamedItem("rotation").nodeValue = "cw" then
				spiralDirection = 1
			End if

			PIE = theEle.selectSingleNode("Coordinates").attributes.getNamedItem("PIEast").nodeValue
			PIN = theEle.selectSingleNode("Coordinates").attributes.getNamedItem("PINorth").nodeValue
			dblAngleTemp = GetDirection(spiralYStart,spiralXStart,PIN,PIE)
			'transform from degrees measured ccw from East, to azimuth in radians
			If dblAngleTemp < 0 Then dblAngleTemp = 360 + dblAngleTemp '(it shouldn't happend, but...)
			dblAngleTemp = 90 + (360 - dblAngleTemp)
			If dblAngleTemp => 360 Then
				dblAngleTemp = dblAngleTemp - 360
			End If
			spiralDirectionBack = dblAngleTemp * PI/180
			If GM_LocateOnSpiral(dLen, x, y, azm, spiralLength, spiralXStart, spiralYStart, spiralDirectionBack, spiralRadius1, spiralRadius2, spiralDirection) Then
				north = y
			Else
				north = "Error"
			End If
                        'If Err.Number <> 0 Then
			'MsgBox dLen & "==" &  x & "==" &  y & "==" &  azm & "==" &  spiralLength & "==" &  spiralXStart & "==" &  spiralYStart & "==" &  spiralDirectionBack & "==" &  spiralRadius1 & "==" &  spiralRadius2 & "==" &  spiralDirection
			'MsgBox Err.Description
			'End If
		End If
		GetNorthingAtSta = north
	If Err.Number <> 0 Then
		'mSGBOX sStartRad & "===" & sEndRad & "===" & Temp1 & "===" & Temp2 & "===" & Temp3
		If Err.Number = -2147467260 then
			Msgbox "User aborted process!"
		End if
		GetNorthingAtSta = "Error"
	End If
End Function
'##########################################
function GetEastingAtSta(strStation,objAlignment,staIndx)
	On Error Resume Next
	Dim east, eleIndex,theEle, dLen,fSta,spE,delta 
	Dim spiralXStart,spiralYStart,spiralDirection,spiralDirectionBack,PIN,PIE,dblAngleTemp
	Dim spiralLength,spiralRadius1,spiralRadius2,y,x,azm,spiType

	Dim length,theta,rad,cpE,sStartRad,sEndRad,sRad,Temp1,Temp2,Temp3
	fSta = strStation + 0
	if(staIndx >= 0) Then
		eleIndex = staIndx
	else
		GetEastingAtSta = "Error"
		Exit Function
	End If
	Set theEle = objAlignment.selectSingleNode("CoordGeom").childNodes(eleIndex)
	dLen = fSta - theEle.attributes.getNamedItem("StartStation").nodeValue
		if(UCase(theEle.nodeName) = "LINE") Then
			spE = theEle.selectSingleNode("Coordinates").attributes.getNamedItem("StartEast").nodeValue
			delta = theEle.attributes.getNamedItem("angle").nodeValue
			east = spE + (dLen * cos((delta * PI) / 180 )) 
		elseif (UCase(theEle.nodeName) = "CURVE") Then
			delta = Abs(theEle.attributes.getNamedItem("delta").nodeValue)
			length = Abs(theEle.attributes.getNamedItem("length").nodeValue)
			If theEle.attributes.getNamedItem("rotation").nodeValue = "ccw" then
				theta = delta * (dLen / length) + theEle.attributes.getNamedItem("startDirection").nodeValue
			ElseIf theEle.attributes.getNamedItem("rotation").nodeValue = "cw" then
				theta =  theEle.attributes.getNamedItem("startDirection").nodeValue - delta * (dLen / length) 
			End if
			rad = theEle.attributes.getNamedItem("radius").nodeValue
			cpE = theEle.attributes.getNamedItem("centerE").nodeValue
			east = cpE + (rad * cos((theta * PI) / 180) )
		elseif(UCase(theEle.nodeName) = "SPIRAL") Then
			spiType = theEle.attributes.getNamedItem("spiType").nodeValue
			If spiType & "" <> "clothoid" then 
				GetEastingAtSta = "Not Supported"
				Exit Function
			End If
			spiralLength = theEle.attributes.getNamedItem("length").nodeValue
			spiralRadius1 = theEle.attributes.getNamedItem("startRadius").nodeValue
			spiralRadius2 = theEle.attributes.getNamedItem("endRadius").nodeValue
			spiralYStart = theEle.selectSingleNode("Coordinates").attributes.getNamedItem("StartNorth").nodeValue
			spiralXStart = theEle.selectSingleNode("Coordinates").attributes.getNamedItem("StartEast").nodeValue
			If theEle.attributes.getNamedItem("rotation").nodeValue = "ccw" then
				spiralDirection = -1
			ElseIf theEle.attributes.getNamedItem("rotation").nodeValue = "cw" then
				spiralDirection = 1
			End if
			PIE = theEle.selectSingleNode("Coordinates").attributes.getNamedItem("PIEast").nodeValue
			PIN = theEle.selectSingleNode("Coordinates").attributes.getNamedItem("PINorth").nodeValue
			dblAngleTemp = GetDirection(spiralYStart,spiralXStart,PIN,PIE)
			'transform from degrees measured ccw from East, to azimuth in radians
			If dblAngleTemp < 0 Then dblAngleTemp = 360 + dblAngleTemp '(it shouldn't happend, but...)
			dblAngleTemp = 90 + (360 - dblAngleTemp)
			If dblAngleTemp => 360 Then
				dblAngleTemp = dblAngleTemp - 360
			End If
			spiralDirectionBack = dblAngleTemp * PI/180
			If GM_LocateOnSpiral(dLen, x, y, azm, spiralLength, spiralXStart, spiralYStart, spiralDirectionBack, spiralRadius1, spiralRadius2, spiralDirection) Then
				east = x
			Else
				east = "Error"
			End If
		end if
		GetEastingAtSta = east
	If Err.Number <> 0 Then
		GetEastingAtSta = "Error"
	End If
End Function
'######################################
	function ChangeOptionBacksightPoint
		if Document.forms("StakeFormCommonUp").backsightPointOption.checked = false then
			Document.forms("StakeFormCommonUp").backsightPointOption.checked = true
			Document.forms("StakeFormCommonUp").backsightDirectionOption.checked =false
		end if 
	End Function	
'#####################################
	function ChangeOptionBacksightDirection
		if Document.forms("StakeFormCommonUp").backsightDirectionOption.checked = false then
			Document.forms("StakeFormCommonUp").backsightDirectionOption.checked = true
			Document.forms("StakeFormCommonUp").backsightPointOption.checked = false
		end if 
	End Function	
'##########################################
    Function FromDMS_to_DecimalDegrees(strAngleValue)
	
        Dim strResult
        Dim strDegrees, strMinutes, strSeconds
	If Instr(strAngleValue,".")<> 0 Then
		strAngleValue = strAngleValue & "0000"
	else
		strAngleValue = strAngleValue & ".0000"
	End If
        strDegrees = Fix(strAngleValue)
        strMinutes = Mid(strAngleValue, InStr(1, strAngleValue, ".", 1) + 1, 2)
        strSeconds = Mid(strAngleValue, InStr(1, strAngleValue, ".", 1) + 3, 2)
        strMinutes = strMinutes / 60
        strSeconds = strSeconds / 3600

	'MsgBox strDegrees & "==" & strMinutes  & "==" & strSeconds
        strResult = strDegrees + strMinutes + strSeconds
	'MsgBox strResult
        FromDMS_to_DecimalDegrees = strResult
    End Function
'##########################################
	Function GetDegreesFromAnyAngle(strAngleValue,strTemp)
		On Error Resume Next
		If Not IsNumeric(strAngleValue) Then
			GetDegreesFromAnyAngle = False
			Exit Function
		End if
		Select Case UCase(DirectionUnit)
			Case "DEGREES DMS"
				strTemp = (FromDMS_to_DecimalDegrees(strAngleValue))
			Case "GRADS"
				strTemp = 360 * strAngleValue/400
			Case "DECIMAL DEGREES"
				'MsgBox strTemp	
				strTemp = strAngleValue
			Case "RADIANS"
				strTemp = 180 * strAngleValue/3.14159265358979323
			Case "MILS"
				strTemp = 360 * strAngleValue/6400
		End Select
		If Err.Number <> 0 then
			GetDegreesFromAnyAngle = False
			'MsgBox Err.description
		else
			GetDegreesFromAnyAngle = True		
		end if

	End Function
'############################################
	function TransformBacksightDirection(strAngleValue)
		'based on the settings, will trasform this angle in decimal degrees, 
		'get quadrant	     90
		'		2    |    1
		'		     |
		'		     |
		'	     180_____|______0
		'		     |
		'		     |
		'		3    |    4
		'		     |
		'		    270

		On Error Resume Next
		Dim strTemp , strStart, strEnd
		TransformBacksightDirection = "Error"
 
		Select Case UCase(DirectionType)
			Case "AZIMUTH NORTH"
				If GetDegreesFromAnyAngle(strAngleValue,strTemp) Then
					If strTemp < 0 Then strTemp = 360 + strTemp
					strTemp = 90 + (360 - strTemp)
					'MsgBox strTemp
					If strTemp => 360 Then
						strTemp = strTemp - 360
					End If
				Else
					MsgBox "Please verify back direction angle!"
					Exit Function
				End If
			Case "AZIMUTH SOUTH"
				If GetDegreesFromAnyAngle(strAngleValue,strTemp) Then
					If strTemp < 0 Then strTemp = 360 + strTemp
					strTemp = 270 - strTemp
					'MsgBox strTemp
					If strTemp < 0  Then
						strTemp = strTemp + 360
					End If
				Else
					MsgBox "Please verify back direction angle!"
					Exit Function
				End If
			Case "BEARING"
				strStart = Left(strAngleValue,1)
				strEnd  = Right(strAngleValue,1)
				strAngleValue = Mid(strAngleValue,2,Len(strAngleValue)-2)
				strAngleValue = Replace(strAngleValue," ","",1,-1,1)
				
				If GetDegreesFromAnyAngle(strAngleValue,strTemp) Then
					
					If (strTemp+ 0) <= 90 And (0 + strTemp) >= 0 Then
						If UCase(strStart) = "N" And UCase(strEnd) = "E" Then
							strTemp = 90 - strTemp
						ElseIf UCase(strStart) = "S" And UCase(strEnd) = "E" Then
							strTemp = 270 + strTemp
						ElseIf UCase(strStart) = "S" And UCase(strEnd) = "W" Then
							strTemp = 270 - strTemp
						ElseIf UCase(strStart) = "N" And UCase(strEnd) = "W" Then
							strTemp = 90 + strTemp
						Else
							MsgBox "Please verify back direction angle!"
							Exit Function
						End If
					Else
						MsgBox "Please verify back direction angle!"
						Exit Function
					End If
				Else
					MsgBox "Please verify back direction angle!"
					Exit Function
				End If
				
			Case Else
				MsgBox DirectionType & " ->Not Implemented"
				Exit Function
		End Select
		If Isnumeric(strTemp) Then
			TransformBacksightDirection = strTemp
		End If
	end function
'########################################
	Function GetDirection(dblNOccupied,dblEOccupied,dblNRandom,dblERandom)
		Dim dblTemp
		'get quadrant	     90
		'		2    |    1
		'		     |
		'		     |
		'	     180_____|______0
		'		     |
		'		     |
		'		3    |    4
		'		     |
		'		    270
		
		If (cdbl(dblERandom) = cdbl(dblEOccupied)) And (cdbl(dblNRandom) >= cdbl(dblNOccupied)) Then
			dblTemp = 90
		ElseIf (cdbl(dblERandom) = cdbl(dblEOccupied)) And (cdbl(dblNRandom) < cdbl(dblNOccupied)) Then
			dblTemp = 270
		elseIf (cdbl(dblNRandom) = cdbl(dblNOccupied)) And (cdbl(dblERandom) >= cdbl(dblEOccupied)) Then
			dblTemp = 0
		elseIf (cdbl(dblNRandom) = cdbl(dblNOccupied)) And (cdbl(dblERandom) < cdbl(dblEOccupied)) Then
			dblTemp = 180
		elseIf (cdbl(dblNRandom) > cdbl(dblNOccupied)) And (cdbl(dblERandom) > cdbl(dblEOccupied)) Then 'Q1
			dblTemp = atn((cdbl(dblNRandom) - cdbl(dblNOccupied))/(cdbl(dblERandom) - cdbl(dblEOccupied))) * 180/3.14159265358979323

		elseIf (cdbl(dblNRandom) < cdbl(dblNOccupied)) And (cdbl(dblERandom) > cdbl(dblEOccupied)) Then 'Q4
			dblTemp = atn((cdbl(dblNRandom) - cdbl(dblNOccupied))/(cdbl(dblERandom) - cdbl(dblEOccupied))) * 180/3.14159265358979323
			dblTemp = 360 + dblTemp
		elseIf (cdbl(dblNRandom) < cdbl(dblNOccupied)) And (cdbl(dblERandom) < cdbl(dblEOccupied)) Then 'Q3
			dblTemp = 180 + atn((cdbl(dblNRandom) - cdbl(dblNOccupied))/(cdbl(dblERandom) - cdbl(dblEOccupied))) * 180/3.14159265358979323
		elseIf (cdbl(dblNRandom) > cdbl(dblNOccupied)) And (cdbl(dblERandom) < cdbl(dblEOccupied)) Then 'Q2
			dblTemp = 180 + atn((cdbl(dblNRandom) - cdbl(dblNOccupied))/(cdbl(dblERandom) - cdbl(dblEOccupied))) * 180/3.14159265358979323
		end if

		GetDirection = dblTemp
	End Function
'#################################################
	Function GetMaxDist(dblMaxDist)
		Dim dblTemp
		'Msgbox "" & dblMaxDist
		if Not IsNumeric(dblMaxDist) Then
			dblMaxDist = 0
		End if
		dblMaxDist = 0 + dblMaxDist
		If dblMaxDist = 0 Then
			blnNoMaxDistance = True
			GetMaxDist = 0 
			Exit Function
		Else
			blnNoMaxDistance = False
		End If
		Select Case LCase(DistUnit)
			Case "default"
				Select Case unit
					Case "meter"
						dblTemp = dblMaxDist + 0
					Case "foot"
						dblTemp = dblMaxDist + 0
				End Select
			Case "meter"
				Select Case unit
					Case "meter"
						dblTemp = dblMaxDist + 0
					Case "foot"
						dblTemp = dblMaxDist/0.3048
				End Select
			Case "foot"
				Select Case unit
					Case "meter"
						dblTemp = dblMaxDist * 0.3048
					Case "foot"
						dblTemp = dblMaxDist + 0
				End Select
			Case Else
		End Select
		'MsgBox LCase(DistUnit) & "--" & unit

		GetMaxDist = dblTemp + 0		
		'MsgBox blnNoMaxDistance & "--" & dblTemp
	End Function
'#######################################
	function GetUnitPrefix(strParameterReference)
		Dim strTemp
		Select Case strParameterReference
			Case "default"
				Select Case unit
					Case "meter"
						strTemp = " (m)"
					Case "foot"
						strTemp = " (ft)"
				End Select
			Case "meter"
				strTemp = " (m)"
			Case "foot"
				strTemp = " (ft)"
		End Select
		GetUnitPrefix = strTemp		
	end function
'#####################################
	function FillHeader(dblNOccupied,dblEOccupied,dblZOccupied,dblNBacksight,dblEBacksight,dblZBacksight,strCase)
		Dim strTemp
		Dim strNOccupied,strEOccupied,strZOccupied,strNBacksight,strEBacksight,strZBacksight
		'"<TABLE border=""1"" > <COLGROUP><COLGROUP SPAN=3><THEAD><TR><TH SCOPE=col>Point #</TH><TH SCOPE=col>Angle</TH><TH SCOPE=col>Distance</TH><TH SCOPE=col>Northing</TH><TH SCOPE=col>Easting</TH><TH SCOPE=col>Elevation</TH></TR></THEAD><TBODY>

		Dim strNorthDisplay, strEastDisplay, strElevationDisplay
		strNorthDisplay = "Northing" & GetUnitPrefix(CoordUnit)
		strEastDisplay = "Easting" & GetUnitPrefix(CoordUnit)
		strElevationDisplay = "Elevation" & GetUnitPrefix(ElevUnit)

		strNOccupied= FormatNumber1(CStr(dblNOccupied) , CStr(unit) , CStr(CoordUnit) ,  CStr(CoordPrecision) ,  CStr(CoordRounding))
		strEOccupied= FormatNumber1(CStr(dblEOccupied) , CStr(unit) , CStr(CoordUnit) ,  CStr(CoordPrecision) ,  CStr(CoordRounding))
		strZOccupied= FormatNumber1(CStr(dblZOccupied) , CStr(unit) , CStr(ElevUnit) ,  CStr(ElevPrecision) ,  CStr(ElevRounding))

		strNBacksight= FormatNumber1(CStr(dblNBacksight) , CStr(unit) , CStr(CoordUnit) ,  CStr(CoordPrecision) ,  CStr(CoordRounding))
		strEBacksight= FormatNumber1(CStr(dblEBacksight) , CStr(unit) , CStr(CoordUnit) ,  CStr(CoordPrecision) ,  CStr(CoordRounding))
		strZBacksight= FormatNumber1(CStr(dblZBacksight) , CStr(unit) , CStr(ElevUnit) ,  CStr(ElevPrecision) ,  CStr(ElevRounding))
		if strNBacksight = "NaN" Then strNBacksight="-"
		if strEBacksight = "NaN" Then strEBacksight="-"
		if strZBacksight = "NaN" Then strZBacksight="-"

	     strTemp = "<html><p></p><p></p><p></p><p></p><p></p><body><table border=""1"" >"

	     strTemp = strTemp + "<tr><td><b>Occupied Point:</b></td><td>" + Document.forms("StakeFormCommonUp").occupiedPoint.Value + "</td></tr>"
	     strTemp = strTemp + "<tr><td> </td><td><b>" + strNorthDisplay + "</b></td><td><b>" + strEastDisplay + "</b></td><td><b>" + strElevationDisplay + "</b></td></tr>"
	     strTemp = strTemp + "<tr><td> </td><td>" + strNOccupied + "</td><td>" + strEOccupied + "</td><td>" + strZOccupied + "</td></tr>"
	     strTemp = strTemp + "<tr><td><b>Backsight Point:</b></td><td>" + Document.forms("StakeFormCommonUp").backsightPointName.Value + " </td></tr>"
	     strTemp = strTemp + "<tr><td> </td><td><b>" + strNorthDisplay + "</b></td><td><b>" + strEastDisplay + "</b></td><td><b>" + strElevationDisplay + "</b></td></tr>"
	     strTemp = strTemp + "<tr><td> </td><td>" + strNBacksight + "</td><td>" + strEBacksight + "</td><td>" + strZBacksight + "</td></tr>"

	     strTemp = strTemp + "<tr><td><b>Backsight Direction:</b></td><td>" + Document.forms("StakeFormCommonUp").backsightDirectionAngle.Value + "</td></tr>"

	     If strCase = "Alignments" Then
		strAlignment = Document.forms("StakeFormAlignmentA").AlignmentName.options(Document.forms("StakeFormAlignmentA").AlignmentName.selectedIndex).text
	     	strTemp = strTemp + "<tr><td><b>Alignment:</b></td><td>" + strAlignment + "</td></tr>"
	     End If
		strTemp = strTemp + "</table>"
		FillHeader = strTemp
	end function
'######################################		
	function InRadius(dblNOccupied,dblEOccupied,dblNRandom,dblERandom,MaxDistance)
		If blnNoMaxDistance Then
			InRadius = True
			Exit Function
		End If 
		If ((dblNOccupied - dblNRandom)^2 + (dblEOccupied - dblERandom)^2) > (MaxDistance^2) Then
			InRadius = False
		Else
			Inradius = True
		End If
	end function
'########################################
	Function GetDistance(dblNOccupied,dblEOccupied,dblNRandom,dblERandom)
		Dim dblDistance
		dblDistance = sqr((dblNOccupied - dblNRandom)^2 + (dblEOccupied - dblERandom)^2)
'		
		GetDistance = (dblDistance)
	End Function
'#########################################
'zhangc-modify at 12/22/2004-to add confirm message box
	function ClearReport
		Dim iResult
		iResult = MsgBox ("Are you sure you would like to clear the current report?", 4, "Confirm Clear Report")
		if( iResult = 6) then
			output.innerHTML = ""
		end if
	End Function	
'###########################################
    Function GetPrefixAndSuffixFromName(strFullName, strPrefix, strSuffix)
        Dim i
	strPrefix = ""
	strSuffix = ""
	If IsNumeric(strFullname) Then
		strSuffix = strFullName + 0 
		Exit Function
	End If
        For i = Len(strFullName) To 1 Step -1
            If Not IsNumeric(Mid(strFullName, i, 1)) Then
                Exit For
            End If
        Next
        strPrefix = Mid(strFullName, 1, i)

        strSuffix = (Mid(strFullName, Len(strPrefix) + 1))

    End Function
'#########################################
	Function ExpandPointListFilter (strPointListFilter,strMinRange,strMaxRange,strPrefixes)
		Dim strArr
		Dim strStart, strPrefixStart
		Dim strEnd , strPrefixEnd
		Dim dblSuffixStart,dblSuffixEnd
		Dim n,i
		strPointListFilter = "," & strPointListFilter
		strArr = Split(strPointListFilter,",",-1,1)
		strPointListFilter = ""
		If UBound(strArr)<> -1 Then
			For n = 0 To Ubound(strArr)
				If Instr(1,strArr(n),"-",1) <> 0 Then
					strStart = Left(strArr(n),Instr(1,strArr(n),"-",1)-1)
					strEnd = Mid(strArr(n),Instr(1,strArr(n),"-",1)+1)
					GetPrefixAndSuffixFromName strStart, strPrefixStart, dblSuffixStart
					GetPrefixAndSuffixFromName strEnd, strPrefixEnd, dblSuffixEnd
					If strPrefixStart = strPrefixEnd Then
						if IsNumeric(dblSuffixStart) And IsNumeric(dblSuffixEnd) Then
							Redim Preserve strPrefixes(MaxDim(strPrefixes)+1)
							strPrefixes(MaxDim(strPrefixes)) = strPrefixStart
							Redim Preserve strMinRange(MaxDim(strMinRange)+1)
							strMinRange(MaxDim(strMinRange)) = dblSuffixStart
							Redim Preserve strMaxRange(MaxDim(strMaxRange)+1)
							strMaxRange(MaxDim(strMaxRange)) = dblSuffixEnd
						End If
					End If
				Elseif Instr(1,strArr(n),">=",1)<> 0 Then
					strStart = Mid(strArr(n),Instr(1,strArr(n),">=",1)+2)
					GetPrefixAndSuffixFromName strStart, strPrefixStart, dblSuffixStart
						if IsNumeric(dblSuffixStart) Then
							Redim Preserve strPrefixes(MaxDim(strPrefixes)+1)
							strPrefixes(MaxDim(strPrefixes)) = strPrefixStart
							Redim Preserve strMinRange(MaxDim(strMinRange)+1)
							strMinRange(MaxDim(strMinRange)) = dblSuffixStart
							Redim Preserve strMaxRange(MaxDim(strMaxRange)+1)
							strMaxRange(MaxDim(strMaxRange)) = "INF"
						End If
				Elseif Instr(1,strArr(n),"<=",1)<> 0 Then
					strEnd = Mid(strArr(n),Instr(1,strArr(n),"<=",1)+2)
					GetPrefixAndSuffixFromName strEnd, strPrefixEnd, dblSuffixEnd
						if IsNumeric(dblSuffixEnd) Then
							Redim Preserve strPrefixes(MaxDim(strPrefixes)+1)
							strPrefixes(MaxDim(strPrefixes)) = strPrefixEnd
							Redim Preserve strMinRange(MaxDim(strMinRange)+1)
							strMinRange(MaxDim(strMinRange)) = 0
							Redim Preserve strMaxRange(MaxDim(strMaxRange)+1)
							strMaxRange(MaxDim(strMaxRange)) = dblSuffixEnd
						End If
				Elseif Instr(1,strArr(n),">",1)<> 0 Then
					strStart = Mid(strArr(n),Instr(1,strArr(n),">",1)+1)
					GetPrefixAndSuffixFromName strStart, strPrefixStart, dblSuffixStart
						if IsNumeric(dblSuffixStart) Then
							Redim Preserve strPrefixes(MaxDim(strPrefixes)+1)
							strPrefixes(MaxDim(strPrefixes)) = strPrefixStart
							Redim Preserve strMinRange(MaxDim(strMinRange)+1)
							strMinRange(MaxDim(strMinRange)) = dblSuffixStart + 1
							Redim Preserve strMaxRange(MaxDim(strMaxRange)+1)
							strMaxRange(MaxDim(strMaxRange)) = "INF"
						End If
				Elseif Instr(1,strArr(n),"<",1)<> 0 Then
					strEnd = Mid(strArr(n),Instr(1,strArr(n),"<",1)+1)
					GetPrefixAndSuffixFromName strEnd, strPrefixEnd, dblSuffixEnd
						if IsNumeric(dblSuffixEnd) Then
							Redim Preserve strPrefixes(MaxDim(strPrefixes)+1)
							strPrefixes(MaxDim(strPrefixes)) = strPrefixEnd
							Redim Preserve strMinRange(MaxDim(strMinRange)+1)
							strMinRange(MaxDim(strMinRange)) = 0
							Redim Preserve strMaxRange(MaxDim(strMaxRange)+1)
							strMaxRange(MaxDim(strMaxRange)) = dblSuffixEnd - 1
						End If
				Else
					strPointListFilter = strPointListFilter & strArr(n) & ","
				end If
			Next
		Else
			strPointListFilter = "," & strPointListFilter & ","
		End If
		'For n = 0 to MaxDim(strPrefixes)
		'	MsgBox strPrefixes(n) & "-" & strMinRange(n) & "-" & strMaxRange(n)
		'Next
	End Function
'#########################################
	function AppendReport()
		On Error Resume Next
		if Document.forms("StakeFormCommonUp").reportType.selectedIndex = 0 Then
			AppendReportMain "Points" 
		else
			AppendReportMain "Alignments" 
		End if
		hideinput
	end function 
'############################################	
    function isDigit(character)
        asciiValue = CInt(Asc(character))
        if (48 <= asciiValue and asciiValue <= 57) then
            isDigit = true
        else
            isDigit = false
        end if
    end function
'############################################	
' Reset decimal symbal to dot ".", 
' The decimal symbal may not "." according to regional options in Windows. And JScript function parseFloat(...) can't deal with decimal symbal other then ".".
' This function reset the decimal symbal to "."
    function resetDecimalSymbal(stringValue)
        returnVal = mid(stringValue, 1, 1)
        for i = 2 to len(stringValue)
            curChar = mid(stringValue, i, 1)
            if isDigit(curChar) then
                returnVal = returnVal & curChar
            else
                returnVal = returnVal & "."
            end if
        next
        resetDecimalSymbal = returnVal
    end function
'############################################	
	]]>
	</script>

	<script type="text/JScript">
<![CDATA[

function Launch(strContent)
{
	var win = window.open(); // a window object
	var doc = win.document;
	doc.open("text/html", "replace");
	doc.write(strContent);
	doc.close();
}
/////////////////////////////////////////////////////////////////////////
// Number formatting

function getInsets() {
   // Store the old document position
   var oldScreenLeft = window.top.screenLeft;
   var oldScreenTop = window.top.screenTop;

   // if no previous inset calculated assume one
   if (window.top._insets == null)
      window.top._insets = {left: 5, top: 80};
   
   // move to a known position
   window.top.moveTo(oldScreenLeft - window.top._insets.left,
                 oldScreenTop - window.top._insets.top);
   
   // Measure the new document position
   var newScreenLeft = window.top.screenLeft;
   var newScreenTop = window.top.screenTop;
   
   // ... and store the insets result
   var res = {
      left:   newScreenLeft - oldScreenLeft + window.top._insets.left,
      top:   newScreenTop - oldScreenTop + window.top._insets.top
   };
   
   // move back the window to its original place
   window.top.moveTo(oldScreenLeft - res.left, oldScreenTop - res.top);
   
   // and backup the insets for next time
   window.top._insets = res;
   
   return res;
}

function getLeft() {
   return window.top.screenLeft - getInsets().left;
}

function getTop() {
   return window.top.screenTop - getInsets().top;
}

function RepaintMe()
{
	repaint();
}
function FormatNumber1(numberStr, fromStr, toStr, precisionStr, roundingStr)
{
	try
	{
	var numStr = numberStr;
	var fStr = fromStr;
	var tStr = toStr;
	var pStr = precisionStr;
	var rStr = roundingStr;
	
	var origNumber= new Number(parseFloat(resetDecimalSymbal(numStr)));
	
	var convNumber = 0;
	var precDepth;
	var fixedDec;
	var rNumber = numberStr;
	
	// Do conversion if necessary
	if(tStr == "default")
	{
		convNumber = origNumber;
	}
	else
	{
		convNumber = ConvertNumber(origNumber, fStr, tStr);
	}
	
	rNumber = FormatPrecision(convNumber, pStr, rStr);
	
	return rNumber.toString();
	//return "gjgkgk";

	}
	catch (e)
	{
		return "Format---->" + e.description;
	}

	//return pStr;
}

function ConvertNumber(num, fromStr, toStr)
{
	var convNumber = new Number(parseFloat(resetDecimalSymbal(num)));
	var numer = new Number(parseFloat(resetDecimalSymbal(num)));
	
	if(fromStr == "foot")
	{
		if(toStr == "meter")
		{
			convNumber = FeetToMeters(numer)
		}
	}
	else if(fromStr == "meter")
	{
		if(toStr == "foot")
		{
			convNumber = MToFeet(numer);
		}
	}
	else if(fromStr == "squareFoot")
	{
		if(toStr == "squareMeter")
		{
			convNumber = SqFeetToSqMeters(numer);
		}
		if(toStr == "acre")
		{
			convNumber = SqFtToAcres(numer);
		}
		if(toStr == "hectare")
		{
			convNumber = SqFeetToSqMeters(numer) / 10000.;
		}
	}
	else if(fromStr == "squareMeter")
	{
		if(toStr == "squareFoot")
		{
			convNumber = SqMetersToSqFeet(numer);
		}
		if(toStr == "acre")
		{
			convNumber = SqMetersToAcres(numer);
		}
		if(toStr == "hectare")
		{
			convNumber = numer / 10000.;
		}
	}
	else
	{
		convNumber = numer;
	}
	
	return convNumber;
}

function ParsePrecision(precStr)
{
	var decPrec;
	var per = ".";
	var s = precStr.indexOf(".");
	if(s < 0)
	{
		decPrec= 0;
	}
	else
	{
		var d = precStr.length - s - 1;
		decPrec= d;
	}	
	return decPrec;
}
/////////////////////////////////////////////////////////////////////////
// Direction formatting
function FormatDirection(directionStr, fromStr, toStr, precisionStr, roundingStr)
{
	var rDir = directionStr;
	var dirNum = new Number(parseFloat(resetDecimalSymbal(directionStr)));
	var convNumber = 0;
	
	if(fromStr == "Conventional")
	{
		if(toStr == "Bearing")
		{
			if(dirNum <= 90 && dirNum >= 0)
			{
				var adjDir = 90 - dirNum;
				var dirStr = FormatDMS(adjDir , precisionStr, roundingStr);
				return "N " + dirStr + " E";
			}
			else if(dirNum <= 180 && dirNum > 90)
			{
				var adjDir =  dirNum - 90;
				var dirStr = FormatDMS(adjDir , precisionStr, roundingStr);
				return "N " + dirStr + " W";
			}
			else if(dirNum <= 270 && dirNum > 180)
			{
				var adjDir = 270 - dirNum;
				var dirStr = FormatDMS(adjDir , precisionStr, roundingStr);
				return "S " + dirStr + " W";
			}
			else if(dirNum < 360 && dirNum > 270)
			{
				var adjDir = dirNum - 270;
				var dirStr = FormatDMS(adjDir , precisionStr, roundingStr);
				return "S " + dirStr + " E";
			}
		}
	}
	else if(fromStr == "North Azimuth")
	{
		if(toStr == "Bearing")
		{
			if(dirNum <= 90 && dirNum >= 0)
			{
				var adjDir = 90 - dirNum;
				var dirStr = FormatDMS(adjDir , precisionStr, roundingStr);
				return "N " + dirStr + " E";
			}
			else if(dirNum <= 180 && dirNum > 90)
			{
				var adjDir = 180 - dirNum;
				var dirStr = FormatDMS(adjDir , precisionStr, roundingStr);
				return "S " + dirStr + " E";
			}
			else if(dirNum <= 270 && dirNum > 180)
			{
				var adjDir = dirNum - 180;
				var dirStr = FormatDMS(adjDir , precisionStr, roundingStr);
				return "S " + dirStr + " W";
			}
			else if(dirNum < 360 && dirNum > 270)
			{
				var adjDir = 360 - dirNum;
				var dirStr = FormatDMS(adjDir , precisionStr, roundingStr);
				return "N " + dirStr + " W";
			}
		}
	}
	else if(fromStr == "Bearing")
	{
		if(toStr == "North Azimuth")
		{
		}
	}
	
	return rDir;
}

/////////////////////////////////////////////////////////////////////////
// Angle formatting

function FormatAngle1(angleStr, fromStr, toStr, formatStr, precisionStr, roundingStr)
{
	var retStr = angleStr;
	
	var angNumber = new Number(parseFloat(resetDecimalSymbal(angleStr)));
	var convNumber = angNumber;
	var rNumber;
	
	// Convert Number
	if(fromStr == "Degrees")
	{
		if(toStr == "Radians")
		{
			convNumber = (angNumber * Math.PI) / 180;
		}
	}
	else if(fromStr == "Radians")
	{
		if(toStr == "Degrees")
		{
			convNumber = (angNumber * 180) / Math.PI;
		}
	}
	
	if(formatStr == "DMS")
	{
		rNumber = FormatDMS(convNumber, precisionStr, roundingStr);
	}
	else
	{
		rNumber = FormatPrecision(convNumber, precisionStr, roundingStr);
	}
	
	return rNumber.toString();
}

function FormatDMS(angle, precisionStr, roundingStr)
{
	//var angle = FormatPrecision(angleGot,  precisionStr, roundingStr);
	var degrees = Math.floor(angle);

	var dMin = 60. * (angle - degrees);
	var minutes = Math.floor(dMin);

	var dSec = 60. * (dMin - minutes);
	var seconds = FormatPrecision(dSec,  precisionStr, roundingStr);

	if(degrees < 10) { Sdeg = "0" + String(degrees); } else { Sdeg = String(degrees); }
	if(minutes < 10) { Smin = "0" + String(minutes); } else { Smin = String(minutes); }
	if(seconds < 10) { Ssec = "0" + String(seconds); } else { Ssec = String(seconds); }	
	
	return Sdeg + "-" + Smin + "-" + Ssec;
}

function FormatPrecision(numStr, precisionStr, roundingStr)
{
	var convNumber = new Number(parseFloat(resetDecimalSymbal(numStr)));
	var rNumber = numStr;
	
	// Set Precision value
	var precDepth = ParsePrecision(precisionStr);
	
	// Do rounding
	if(roundingStr == "normal")
	{
		rNumber = convNumber.toFixed(precDepth);
	}
	else if(roundingStr == "round up")
	{
		var depthMulti = 1;
		var i;
		for(i = 0; i < precDepth; i++)
		{
			depthMulti *= 10;
		}
		var cNumber = convNumber * depthMulti;
		var ceilNumber = Math.ceil(cNumber);
		var rcNumber = ceilNumber / depthMulti;
		rNumber = rcNumber.toFixed(precDepth);
	}
	else if(roundingStr == "round down")
	{
		var depthMulti = 1;
		var i;
		for(i = 0; i < precDepth; i++)
		{
			depthMulti *= 10;
		}
		var cNumber = convNumber * depthMulti;
		var floorNumber = Math.floor(cNumber);
		var rcNumber = floorNumber / depthMulti;
		rNumber = rcNumber.toFixed(precDepth);
	}
	else if(roundingStr == "truncate")
	{
		var depthMulti = 1;
		var i;
		for(i = 0; i < precDepth; i++)
		{
			depthMulti *= 10;
		}
		var cNumber = convNumber * depthMulti;
		var roundNumber = Math.floor(cNumber);
		var rcNumber = roundNumber / depthMulti;
		rNumber = rcNumber.toFixed(precDepth);
	}	
	else
	{
		rNumber = convNumber;
	}
	return rNumber;
} 

/////////////////////////////////////////////////////////////////////////
// Station formatting

function FormatStation1(stationStr, patterStr, precisionStr, roundingStr)
{
	var convNumber = stationStr;
	var precDepth;
	var rNumber;

	rNumber = FormatPrecision(stationStr,  precisionStr, roundingStr);
	
	// Do Station Pattern
	var numStr = rNumber.toString();
	
	var plusIndex;
	var dotIndexPattern;
	var dotIndexNumber;
	var insertIndex;
	
	plusIndex = FindPlusIndex(patterStr);
	dotIndexPattern = FindDotIndex(patterStr);
	dotIndexNumber = FindDotIndex(numStr);
	
	insertIndex = dotIndexNumber - (dotIndexPattern - plusIndex - 1)
	
	var rString = "";
	for(var i = 0; i < numStr.length; i++)
	{
		if(i == insertIndex && i !=0 && plusIndex > 0)
		{
			rString = rString + '+';
		}
		rString = rString + numStr.charAt(i);
	}
	
	return rString;
}
function FindPlusIndex(stationPattern)
{
	var decPrec;
	var s = stationPattern.indexOf("+");
	if(s < 0)
	{
		decPrec= -1; //stationPattern.length;
	}
	else
	{
		var d = s;
		decPrec= d;
	}	
	return decPrec;
}

function FindDotIndex(dotStr)
{
	var decPrec;
	var s = dotStr.indexOf(".");
	if(s < 0)
	{
		decPrec=dotStr.length;
	}
	else
	{
		var d = s;
		decPrec= d;
	}	
	return decPrec;

}
/////////////////////////////////////////////////////////////////////////
// Unit Conversion Functions
//--------------------------------------------------------------
// English to English - Linear
//--------------------------------------------------------------
function FeetToMiles(feet)
{
	var ft = new Number(parseFloat(resetDecimalSymbal(feet)));
	return ft / 5280;
}

function MilesToFeet(miles)
{
	var mls = new Number(parseFloat(resetDecimalSymbal(miles)));
	return 5280 * mls;
}
//--------------------------------------------------------------
// English to English - Area
//--------------------------------------------------------------
function SqFtToAcres(sqFt)
{
	var sq = new Number(parseFloat(resetDecimalSymbal(sqFt)));
	return sq / 43560;
}
function AcresToSqFeet(acres)
{
	var acrs = new Number(parseFloat(resetDecimalSymbal(acres)));
	return acrs * 43560;
}

function AcresToSqMiles(acres)
{
	var acrs = new Number(parseFloat(resetDecimalSymbal(acres)));
	return acrs / 640;
}
//--------------------------------------------------------------
// Metric to Metric - Length
//--------------------------------------------------------------
function CmToMeter(d)
{
	return d/100;
}
function CmToKm(d)
{
	return d/100000;
}
function CmToMm(d)
{
	return d*10;
}
function MmToCm(d)
{
	return d/10;
}
//--------------------------------------------------------------
// Metric to English - Length
//--------------------------------------------------------------
function MmToInches(d)
{
	return d/.254;
}
function CmToFeet(d)
{
	return d / 30.48;
}
function MToFeet(d)
{
	return d/.3048;
}

function CmToInch(d)
{
	return d/2.54;
}
function CmToYards(d)
{
	return d/91.44;
}
function CmToMiles(d)
{
	return d/160934.4;
}
function KmToMiles(d)
{
	return d / 1.609344;
}
//--------------------------------------------------------------
// English to Metric - Length
//--------------------------------------------------------------

function InchesToMm(d)
{
	return d * 25.4;
}

function InchesToCm(d)
{
	return d * 2.54;
}
function FeetToMm(d)
{
	return d * 304.8;
}
function FeetToCm(d)
{
	return d * 30.48;
}

function FeetToMeters(d)
{
	return d * .3048;
}
function FeetToKm(d)
{
	return d * .003048;
}

function YardsToMeters(d)
{
	return d * .9144;
}

function MilesToKm(d)
{
	return d * 1.609344;
}

//--------------------------------------------------------------
// English to Metric - Area
//--------------------------------------------------------------
function SqInchesToSqCm(d)
{
	return d * 6.4516;
}

function SqFeetToSqMeters(d)
{
	return d * 0.09290304;
}

function SqYardsToSqMeters(d)
{
	return d * 0.83612736;
}

function SqMilesToSqKm(d)
{
	return d * 2.589988110336;
}

function AcresToHectares(d)
{
	return d * 0.40468564224;
}

//--------------------------------------------------------------
// Metric to English - Area
//--------------------------------------------------------------
function SqCmToSqInches(d)
{
	return d / 6.4516;
}
function SqMetersToSqFeet(d)
{
	return d / 0.09290304;
}
function SqMetersToSqYards(d)
{
	return d / 0.83612736;
}
function SqKmToSqMiles(d)
{
	return d / 2.589988110336;
}
function HectaresToAcres(d)
{
	return d / 0.40468564224;
}
function SqMetersToAcres(d)
{
	return d / 4046.8564224;
}
function Try(d,y)
{
	return d + "ooo";
}

//->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Get Northing and Easting <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

]]>

	</script>
   	<style type="text/css" media="print">
   		.noprint{display: none}
	</style>
	<style>
		td{
		padding:0.02in 0.15in;
		text-align:right;
		}
	</style>
	</head>

	<body onscroll="outputScrolled" onload="outputScrolled" onresize="outputScrolled">
				<div id="reportHTML" style="width: 7in;position:absolute">
					<xsl:call-template name="AutoHeader">
					<xsl:with-param name="ReportTitle">Radial Stakeout Report</xsl:with-param>
					<xsl:with-param name="ReportDesc"><xsl:value-of select="//lx:Project/@name"/></xsl:with-param>
					</xsl:call-template>
					<div id="output" style="width: 7in"></div>
				</div>

				<div id = "allForms" style="width: 225px;margin-left: 5;border:black thin solid; overflow: hidden;background-color:lightgrey; position:absolute"  class="noprint">
					<xsl:call-template name="makeform"/>
				</div>

 	<xml id="IslandCgPoints">
	<xsl:element name="Root">
 	
		<xsl:element name="Points">
			<xsl:for-each select="//lx:CgPoint[not(@pntRef)]">
			
				<xsl:apply-templates select="." mode="set"/>
				<xsl:variable name="Northing" select="landUtils:GetCogoPointNorthing()"/>
				<xsl:variable name="Easting" select="landUtils:GetCogoPointEasting()"/>
				<xsl:variable name="Elevation" select="landUtils:GetCogoPointElevation()"/>			

				<xsl:element name="Point">
					<xsl:attribute name="id"><xsl:value-of select="@name"/></xsl:attribute>
					<xsl:attribute name="Northing"><xsl:value-of select="$Northing"/></xsl:attribute>
					<xsl:attribute name="Easting"><xsl:value-of select="$Easting"/></xsl:attribute>
					<xsl:attribute name="Elevation"><xsl:value-of select="$Elevation"/></xsl:attribute>
	
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
		<xsl:element name="Alignments">
			<xsl:for-each select="//lx:Alignment">
				<!-- Loads the geometry of the alignment -->
				<xsl:element name="Alignment">
					<xsl:attribute name="name"><xsl:value-of select="@name"/></xsl:attribute>
					<xsl:attribute name="desc"><xsl:value-of select="@desc"/></xsl:attribute>
					<xsl:if test="./lx:StaEquation">
						<xsl:for-each select="./lx:StaEquation">
						<xsl:element name="StaEquation">	
							<xsl:attribute name="staInternal"><xsl:value-of select="@staInternal"/></xsl:attribute>
							<xsl:attribute name="staBack"><xsl:value-of select="@staBack"/></xsl:attribute>
							<xsl:attribute name="staAhead"><xsl:value-of select="@staAhead"/></xsl:attribute>
						</xsl:element>
						</xsl:for-each>
					</xsl:if>

					<xsl:element name="CoordGeom">

<!--@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
				<xsl:apply-templates select="./lx:CoordGeom" mode="set"></xsl:apply-templates>
			 	<xsl:variable name="AppStations" select="landUtils:ApplyStationing(@staStart)"/>
			 	<xsl:variable name="appstaeq" select="landUtils:ApplyStationEquations()"/>

				<xsl:for-each select="./lx:CoordGeom/node()">
					<xsl:variable name="pos" select="position()"/>
					<xsl:variable name="startStation" select="landUtils:GetGeomStartingStation($pos)"/>
					<xsl:variable name="endStation" select="landUtils:GetGeomEndStation($pos)"/>
					<xsl:choose>
						<xsl:when test="name()='Line'">
							
							<xsl:element name="Line">
								<xsl:attribute name="position"><xsl:value-of select="$pos"/></xsl:attribute>
								<xsl:attribute name="StartStation"><xsl:value-of select="$startStation"/></xsl:attribute>
								<xsl:attribute name="EndStation"><xsl:value-of select="$endStation"/></xsl:attribute>
								<xsl:variable name="angle" select="landUtils:GetLineDirection($pos)"/>
								<xsl:attribute name="angle"><xsl:value-of select="$angle"/></xsl:attribute>

								<xsl:element name="Coordinates">
									<xsl:variable name="startN" select="landUtils:GetLineStartNorthing($pos)"/>
									<xsl:variable name="startE" select="landUtils:GetLineStartEasting($pos)"/>
									<xsl:attribute name="StartNorth"><xsl:value-of select="$startN"/></xsl:attribute>
									<xsl:attribute name="StartEast"><xsl:value-of select="$startE"/></xsl:attribute>
									<xsl:variable name="endN" select="landUtils:GetLineEndNorthing($pos)"/>
									<xsl:variable name="endE" select="landUtils:GetLineEndEasting($pos)"/>
									<xsl:attribute name="EndNorth"><xsl:value-of select="$endN"/></xsl:attribute>
									<xsl:attribute name="EndEast"><xsl:value-of select="$endE"/></xsl:attribute>								
								</xsl:element>

							</xsl:element>
						</xsl:when>	
						<xsl:when test="name()='Curve'">
							<xsl:element name="Curve">

								<xsl:attribute name="position"><xsl:value-of select="$pos"/></xsl:attribute>
								<xsl:attribute name="StartStation"><xsl:value-of select="$startStation"/></xsl:attribute>
								<xsl:attribute name="EndStation"><xsl:value-of select="$endStation"/></xsl:attribute>
								<xsl:variable name="delta" select="landUtils:GetCurveAngle($pos)"/>
								<xsl:attribute name="delta"><xsl:value-of select="$delta"/></xsl:attribute>

								<xsl:variable name="radius" select="landUtils:GetCurveRadius($pos)"/>
								<xsl:attribute name="radius"><xsl:value-of select="$radius"/></xsl:attribute>

								<xsl:variable name="theta" select="landUtils:GetCurveAngle($pos)"/>
								<xsl:attribute name="delta"><xsl:value-of select="$delta"/></xsl:attribute>

								<xsl:variable name="length" select="landUtils:GetCurveLength($pos)"/>
								<xsl:attribute name="length"><xsl:value-of select="$length"/></xsl:attribute>

								<xsl:variable name="centerN" select="landUtils:GetCurveCenterNorthing($pos)"/>
								<xsl:attribute name="centerN"><xsl:value-of select="$centerN"/></xsl:attribute>

								<xsl:variable name="centerE" select="landUtils:GetCurveCenterEasting($pos)"/>
								<xsl:attribute name="centerE"><xsl:value-of select="$centerE"/></xsl:attribute>

								<xsl:variable name="startDirection" select="landUtils:GetCurveStartDirection($pos)"/>
								<xsl:attribute name="startDirection"><xsl:value-of select="$startDirection"/></xsl:attribute>


								<xsl:variable name="rotation" select="landUtils:GetCurveRotation($pos)"/>
								<xsl:attribute name="rotation"><xsl:value-of select="$rotation"/></xsl:attribute>

								<xsl:element name="Coordinates">
									<xsl:variable name="startN" select="landUtils:GetCurveStartNorthing($pos)"/>
									<xsl:variable name="startE" select="landUtils:GetCurveStartEasting($pos)"/>
									<xsl:attribute name="StartNorth"><xsl:value-of select="$startN"/></xsl:attribute>
									<xsl:attribute name="StartEast"><xsl:value-of select="$startE"/></xsl:attribute>
									<xsl:variable name="endN" select="landUtils:GetCurveEndNorthing($pos)"/>
									<xsl:variable name="endE" select="landUtils:GetCurveEndEasting($pos)"/>
									<xsl:attribute name="EndNorth"><xsl:value-of select="$endN"/></xsl:attribute>
									<xsl:attribute name="EndEast"><xsl:value-of select="$endE"/></xsl:attribute>								
								</xsl:element>

							</xsl:element>
						</xsl:when>
						<xsl:when test="name()='Spiral'">
							
							<xsl:element name="Spiral">

								<xsl:attribute name="position"><xsl:value-of select="$pos"/></xsl:attribute>
								<xsl:attribute name="spiType"><xsl:value-of select="@spiType"/></xsl:attribute>
								<xsl:attribute name="StartStation"><xsl:value-of select="$startStation"/></xsl:attribute>
								<xsl:attribute name="EndStation"><xsl:value-of select="$endStation"/></xsl:attribute>
								<xsl:attribute name="length"><xsl:value-of select="@length"/></xsl:attribute>
								<xsl:attribute name="startRadius"><xsl:value-of select="@radiusStart"/></xsl:attribute>
								<xsl:attribute name="endRadius"><xsl:value-of select="@radiusEnd"/></xsl:attribute>
								<xsl:variable name="theta" select="landUtils:GetSpiralTheta_Degrees($pos)"/>
								<xsl:attribute name="theta"><xsl:value-of select="$theta"/></xsl:attribute>

								<xsl:variable name="rotation" select="landUtils:GetSpiralRotation($pos)"/>
								<xsl:attribute name="rotation"><xsl:value-of select="$rotation"/></xsl:attribute>

								<xsl:element name="Coordinates">
									<xsl:variable name="startN" select="landUtils:GetSpiralStartNorthing($pos)"/>
									<xsl:variable name="startE" select="landUtils:GetSpiralStartEasting($pos)"/>
									<xsl:attribute name="StartNorth"><xsl:value-of select="$startN"/></xsl:attribute>
									<xsl:attribute name="StartEast"><xsl:value-of select="$startE"/></xsl:attribute>
									<xsl:variable name="endN" select="landUtils:GetSpiralEndNorthing($pos)"/>
									<xsl:variable name="endE" select="landUtils:GetSpiralEndEasting($pos)"/>
									<xsl:attribute name="EndNorth"><xsl:value-of select="$endN"/></xsl:attribute>
									<xsl:attribute name="EndEast"><xsl:value-of select="$endE"/></xsl:attribute>								
									<xsl:variable name="PIN" select="landUtils:GetSpiralPINorthing($pos)"/>
									<xsl:variable name="PIE" select="landUtils:GetSpiralPIEasting($pos)"/>
									<xsl:attribute name="PINorth"><xsl:value-of select="$PIN"/></xsl:attribute>
									<xsl:attribute name="PIEast"><xsl:value-of select="$PIE"/></xsl:attribute>								
								</xsl:element>

							</xsl:element>
						</xsl:when>
					</xsl:choose> 
				</xsl:for-each>

<!--@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->

					</xsl:element><!-- coordGeom element-->
				</xsl:element><!-- alignment element-->
			</xsl:for-each>
		</xsl:element>		
	
	</xsl:element>		

	</xml>
	</body>
</html>
</xsl:template>

<xsl:template name="makeform">
<!-- Upper Form ##############################-->
<form id="Dummy"></form>
<form id="StakeFormCommonUp">
		<b><div id="titleDiv" onclick="hideInput" style="background-color: silver">
			Radial Stakeout (Click to collapse or expand)
		</div></b>
		<p/>
		
		<div id="div_StakeFormCommonUp" style="margin-left: 5">
			<label>Radial stakeout type:</label><br/>
			<select id="reportType" onchange="changeAppeareance" style="WIDTH: 200px">
					<option>Points</option>
					<option>Alignment from points</option>
			</select><br/>
			<label>Occupied point:</label><br/>
			<input id="occupiedPoint" style="WIDTH: 200px" value=""></input><br/>
			<input type="radio" id="backsightPointOption" onclick="ChangeOptionBacksightPoint" checked="checked" >Backsight Point:</input><br/>
			<input id="backsightPointName" value="" style="WIDTH: 200px"></input><br/>
			<input type="radio" id="backsightDirectionOption"  onclick="ChangeOptionBacksightDirection">Backsight Direction:</input><br/>
			<input id="backsightDirectionAngle" style="WIDTH: 200px"></input><br/>
		</div>
		<p/>
</form>

<!-- Point Specific form ####################################-->
<form id="StakeFormPoints">
<div id="div_StakeFormPoints" style="margin-left: 5">
	<label id="stakePointListLabel" >Stakeout points:</label><br/>
	<input id="stakePointList" style="WIDTH: 200px"></input><br/>
</div>
	<p/>
</form>

<!-- Alignment Specific form ###############################-->
<form id="StakeFormAlignmentA" style="display = none" >
<div id="div_StakeFormAlignmentA" style="margin-left: 5">
	<label>Alignment:</label><br/>
	<select id="AlignmentName" style="WIDTH: 200px" onchange="ChangeAlignment">
					<option>-------------none-------------</option>
			</select><br/>
</div>
<p/>
</form>
<!--######################################################-->
<form id="StaEq" style="display = none" >
<div id="div_StaEq" style="margin-left: 5">
	<input id="useStationEquation"  type="checkbox">Use station equations</input><br/>

</div>
<p/>
</form>

<!--###########################################################-->
<form id="StakeFormAlignmentB" style="display = none" >
<div id="div_StakeFormAlignmentB" style="margin-left: 5">
		<label>Station Regions:</label><br/>
		<TEXTAREA id="StationEquationList" CONTENTEDITABLE = "false" style="WIDTH: 200px; OVERFLOW-X: scroll" rows="3" wrap="off"></TEXTAREA><br/>
		<label id="stakeStationListLabel">Stakeout station range:</label><br/>
		<input id="stakeStationList" style="WIDTH: 200px"></input><br/>
</div>
<p/>
</form>
<!-- Lower part form, common ##############################-->
<form id="StakeFormCommonDown" onresize = "RefreshMe" POSITION="absolute">
<div id="div_StakeFormCommonDown" style="margin-left: 5">
		<label>Maximum distance:</label><br/>
		<input id="maximumDistance" value="" style="WIDTH: 200px"></input><br/>
		<hr size="2" color="black" style="WIDTH: 200px"/><br/>		
		<input type="button" id="append" value="Append to report" onclick="AppendReport"  onmousedown="changeCaption" ondrop="showProgress" style="WIDTH: 200px"></input><br/><p/>
		<input type="button" id="clear" value="  Clear report     " onclick="ClearReport" style="WIDTH: 200px"></input><br/><p/>
		<input type="button" id="saveA" value="  Save report     " onclick="SaveReportA" style="WIDTH: 200px"></input><br/>
</div>
<p/>
</form>
</xsl:template>

</xsl:stylesheet>
