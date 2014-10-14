VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportForm_CorridorSlopeStake 
   Caption         =   "Create Reports - Slope Stake Report"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10395
   OleObjectBlob   =   "ReportForm_CorridorSlopeStake.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReportForm_CorridorSlopeStake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





''//
''// (C) Copyright 2005 by Autodesk, Inc.
''//
''// Permission to use, copy, modify, and distribute this software in
''// object code form for any purpose and without fee is hereby granted,
''// provided that the above copyright notice appears in all copies and
''// that both that copyright notice and the limited warranty and
''// restricted rights notice below appear in all supporting
''// documentation.
''//
''// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
''// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
''// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
''// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
''// UNINTERRUPTED OR ERROR FREE.
''//
''// Use, duplication, or disclosure by the U.S. Government is subject to
''// restrictions set forth in FAR 52.227-19 (Commercial Computer
''// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
''// (Rights in Technical Data and Computer Software), as applicable.
'
'
'
Option Explicit
'
Private Type ItemStationInfo
    index As Long
    oSampleLineGroup As AeccSampleLineGroup '
    startStation As Double
    endStation As Double
    curStartStation As Double
    curEndStation As Double
End Type

Private g_StationInfo As ItemStationInfo

Private g_sFileName As String
Private g_reportFileName As String

Private g_Corridors As New Dictionary

Private Const nNameIndex = 0
Private Const nAlignmentIndex = 1
Private Const nSampleLineGroup = 2
Private Const nLinkCodes = 3
Private Const nStationStartIndex = 4
Private Const nStationEndIndex = 5
Private Const nLastIndex = nStationEndIndex

Private sCaptionArr() As String


Private Sub DeleteButton_Click()
    Dim i As Long
    Dim item As ListItem
    Dim colItems As ListItems
    Dim oCorridorInfo As CCorridor
    Dim oBaselineInfo As CBaseline
    Dim oSampleLineGroupInfo As CSampleLineGroup
    Dim oLinkCodeInfo As CLinkCodes
    
    Set colItems = ListView.ListItems
    For i = 1 To colItems.count Step 1
        Set item = colItems.item(i)
        If item.Selected = True Then
            Set oCorridorInfo = g_Corridors.item(item.Text)
            Set oBaselineInfo = oCorridorInfo.mBaselines.item(item.SubItems(nAlignmentIndex))
            Set oSampleLineGroupInfo = oBaselineInfo.mSampleLineGroups.item(item.SubItems(nSampleLineGroup))
            Set oLinkCodeInfo = oBaselineInfo.mLinkCodesGroups.item(item.SubItems(nLinkCodes))
            oLinkCodeInfo.mAdded = False
            If item.index = g_StationInfo.index Then
                StartStaEdit.Enabled = False
                StartStaEdit.Value = ""
                EndStaEdit.Enabled = False
                EndStaEdit.Value = ""
            End If
            colItems.Remove i
            i = i - 1
        End If
        If i = colItems.count Then
            Exit For
        End If
    Next i
     
  
End Sub
Private Sub AddButton_Click()
    If Combo_LineGroup.ListCount = 0 Or Combo_Link.ListCount = 0 Then
        Exit Sub
    End If
    
    Dim oSampleLineGroupInfo As CSampleLineGroup
    Dim oSampleLineGroup As AeccSampleLineGroup
    Set oSampleLineGroupInfo = FindSampleLineGroupInfo(Combo_Corridor.SelText, _
                                               Combo_Align.SelText, _
                                               Combo_LineGroup.SelText)
    Set oSampleLineGroup = oSampleLineGroupInfo.mObject
    
    Dim oLinkCodesInfo As CLinkCodes
    Dim sCodeName As String
    Set oLinkCodesInfo = FindLinkCodesInfo(Combo_Corridor.SelText, _
                                    Combo_Align.SelText, _
                                    Combo_Link.SelText)
    sCodeName = oLinkCodesInfo.mCodeName
    
    If oLinkCodesInfo.mAdded = False Then
        Dim item As ListItem
        Dim dStartStation As Double, dEndStation As Double
        Call StartStationOfSampleLineGroup(oSampleLineGroup, dStartStation, dEndStation)
        
        Set item = ListView.ListItems.Add(, , Combo_Corridor.SelText)
        item.SubItems(nAlignmentIndex) = Combo_Align.SelText
        item.SubItems(nSampleLineGroup) = Combo_LineGroup.SelText
        item.SubItems(nLinkCodes) = Combo_Link.SelText
        item.SubItems(nStationStartIndex) = oSampleLineGroup.Parent.GetStationStringWithEquations(dStartStation)
        item.SubItems(nStationEndIndex) = oSampleLineGroup.Parent.GetStationStringWithEquations(dEndStation)
        oLinkCodesInfo.mAdded = True
    Else
        MsgBox "The component has already been to be reported."
    End If
    
End Sub

Private Sub Combo_LineGroup_Change()
    Dim oSampleLineGroupInfo As CSampleLineGroup
    If Combo_LineGroup.ListCount = 0 Then
        Exit Sub
    End If
    Set oSampleLineGroupInfo = FindSampleLineGroupInfo(Combo_Corridor.SelText, _
                                               Combo_Align.SelText, _
                                               Combo_LineGroup.SelText)
    If oSampleLineGroupInfo.mObject.SampleLines.count = 0 Then
        AddButton.Enabled = False
        MsgBox "The Selected Sample Line Group has no Sample line."
    Else
        AddButton.Enabled = True
    End If
        
  
End Sub

Private Sub Combo_Align_Change()
    Dim oCorridorInfo As CCorridor
    Dim oBaselineInfo As CBaseline
    Dim sSampleLineGroupName
    Dim sLinkCodeName
 
        
    If Combo_Align.ListCount = 0 Then
        Exit Sub
    End If
    
    Set oCorridorInfo = g_Corridors.item(Combo_Corridor.SelText)
    Set oBaselineInfo = oCorridorInfo.mBaselines.item(Combo_Align.SelText)
    
    Combo_LineGroup.Clear
    For Each sSampleLineGroupName In oBaselineInfo.mSampleLineGroups.Keys
        Combo_LineGroup.AddItem sSampleLineGroupName
    Next
    
    Combo_Link.Clear
    For Each sLinkCodeName In oBaselineInfo.mLinkCodesGroups.Keys
        Combo_Link.AddItem sLinkCodeName
    Next
   
    If Combo_LineGroup.ListCount > 0 Then
        Combo_LineGroup.ListIndex = 0
    Else
        AddButton.Enabled = False
        MsgBox "No Sample Line Group available in select alignment."
    End If
    
    If Combo_Link.ListCount > 0 Then
        Combo_Link.ListIndex = 0
    Else
        AddButton.Enabled = False
        MsgBox "No Link available in select alignment."
    End If
End Sub

Private Sub Combo_Corridor_Change()
    Dim oCorridorInfo As CCorridor
    Dim sAlignmentName
    
    If Combo_Corridor.ListCount = 0 Then
        Exit Sub
    End If
    
    Set oCorridorInfo = g_Corridors.item(Combo_Corridor.SelText)
    Combo_Align.Clear
    For Each sAlignmentName In oCorridorInfo.mBaselines.Keys
        Combo_Align.AddItem sAlignmentName
    Next
    
    If Combo_Align.ListCount > 0 Then
        Combo_Align.ListIndex = 0
    Else
        AddButton.Enabled = False
        MsgBox "No Alignment available."
    End If
End Sub



Private Sub HelpButton_Click()
    ReportsHH_hWnd = ReportsHtmlHelp(0, getReportsHelpFile, HH_HELP_CONTEXT, IDH_AECC_REPORTS_CREATE_SLOPE_GRADE_REPORT)
End Sub

Private Sub SelectButton_Click()
    On Error GoTo Cancel
    Dim oEntity As Object
    Dim pt(2) As Double
    Me.Hide
    ThisDrawing.Utility.GetEntity oEntity, pt, "Select a Corridor."
    Do While Not TypeOf oEntity Is AeccCorridor
        ThisDrawing.Utility.GetEntity oEntity, pt, "Please select a Corridor."
    Loop
    Combo_Corridor.Text = oEntity.Name
Cancel:
    Me.Show
End Sub

' Initialize the main dialog
Private Sub UserForm_Initialize()
    Dim tempPath As String
    Dim oCorridors As AeccCorridors
    Dim oCorridor As AeccCorridor
    Dim oCorridorInfo As CCorridor
    g_sFileName = g_oCivilRoadwayApp.Path

    tempPath = Environ("temp")
    g_reportFileName = tempPath & "\civilreport.html"


    ReportFileNameEdit.Text = g_reportFileName

    'initial Select report Components

    g_Corridors.removeAll
    On Error GoTo NoCorridor
    Set oCorridors = g_oRoadwayDocument.Corridors
    On Error GoTo 0
    For Each oCorridor In oCorridors
        If g_Corridors.Exists(oCorridor.Name) = False Then
            Set oCorridorInfo = New CCorridor
            BuildCorridorInfo oCorridor, oCorridorInfo
            g_Corridors.Add oCorridor.Name, oCorridorInfo
            Combo_Corridor.AddItem oCorridor.Name
        End If
    Next
    If g_Corridors.count > 0 Then
        Combo_Corridor.ListIndex = 0
    Else
        AddButton.Enabled = False
        GoTo NoCorridor
    End If
    

    ' progress bar settings
    ProgressBar.Min = 0
    ProgressBar.Max = 100
    ProgressBar.Visible = False
    progressBarLabel.Visible = False

    
    ' Get Captions
    GetCaptions
    
    ' list view grid setup
    ' setup columns
    
    Call ListView.ColumnHeaders.Add(1, , sCaptionArr(nNameIndex), 100, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(2, , sCaptionArr(nAlignmentIndex), 80, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(3, , sCaptionArr(nSampleLineGroup), 100, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(4, , sCaptionArr(nLinkCodes), 80, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(5, , sCaptionArr(nStationStartIndex), 80, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(6, , sCaptionArr(nStationEndIndex), 80, lvwColumnLeft)

    StartStaEdit.Enabled = False
    EndStaEdit.Enabled = False
    Exit Sub
NoCorridor:
    MsgBox "No corridors in " & ThisDrawing.Name & " exist."
    End
End Sub


' Do the real work of gnerating the report here
Private Sub ExecuteButton_Click()
    Dim t, pbarInc, cCount As Integer
    Dim stat As Boolean
    Dim curListItem As ListItem
    Dim exeStr As String
    Dim inStartStation As Double
    Dim inEndStation As Double
    Dim oCorridorInfo As CCorridor
    Dim oBaselineInfo As CBaseline
    Dim oSampleLineGroupInfo As CSampleLineGroup
    Dim oLinkCodesInfo As CLinkCodes

    If ListView.ListItems.count = 0 Then
        Exit Sub
    End If

    For t = 1 To ListView.ListItems.count
        If ListView.ListItems(t).Checked = True Then
            cCount = cCount + 1
            pbarInc = 80 / cCount
        End If
    Next

    ProgressBar.Value = 0
    
    progressBarLabel.Enabled = False
    progressBarLabel.Enabled = True

    progressBarLabel.Visible = True
    ProgressBar.Visible = True

    Call WriteReportHeader(g_reportFileName)

    'update progress bar
    ProgressBar.Value = 10

    For t = 1 To ListView.ListItems.count
        Set curListItem = ListView.ListItems(t)
        inStartStation = GetRawStation(curListItem.SubItems(nStationStartIndex))
        inEndStation = GetRawStation(curListItem.SubItems(nStationEndIndex))
        Set oCorridorInfo = g_Corridors.item(curListItem.Text)
        Set oBaselineInfo = oCorridorInfo.mBaselines.item(curListItem.SubItems(nAlignmentIndex))
        Set oSampleLineGroupInfo = oBaselineInfo.mSampleLineGroups.item(curListItem.SubItems(nSampleLineGroup))
        Set oLinkCodesInfo = oBaselineInfo.mLinkCodesGroups.item(curListItem.SubItems(nLinkCodes))
        stat = AppendReport(oCorridorInfo.mObject, oBaselineInfo.mObject, oSampleLineGroupInfo.mObject, inStartStation, inEndStation, g_reportFileName, oLinkCodesInfo.mCodeName)
        ProgressBar.Value = ProgressBar.Value + pbarInc
      
    Next

    'update progress bar
    ProgressBar.Value = 100

    ProgressBar.Visible = False
    progressBarLabel.Visible = False


    Call WriteReportFooter(g_reportFileName)
    
    OpenFileByDefaultBrowser g_reportFileName
    
End Sub

Private Sub CancelButton_Click()

    Unload Me

End Sub

Private Sub UserForm_Terminate()

    ' clean up created objects
    Set g_oCivilApp = Nothing
    Set g_oAeccDb = Nothing
    Set g_oCivilRoadwayApp = Nothing
    Set g_oRoadwayDocument = Nothing
    Set g_Corridors = Nothing
    Set g_oSlopeStakeData = Nothing

End Sub

Private Sub ListView_ItemClick(ByVal item As ListItem)
    If MultiItemsSelected = True Then
        StartStaEdit.Enabled = False
        EndStaEdit.Enabled = False
    Else
        Dim oSampleLineGroupInfo As CSampleLineGroup
        Dim dStartStation As Double, dEndStation As Double
        Set oSampleLineGroupInfo = FindSampleLineGroupInfo(item.Text, _
                                                           item.SubItems(nAlignmentIndex), _
                                                           item.SubItems(nSampleLineGroup))
        Call StartStationOfSampleLineGroup(oSampleLineGroupInfo.mObject, dStartStation, dEndStation)
        Set g_StationInfo.oSampleLineGroup = oSampleLineGroupInfo.mObject
        g_StationInfo.index = item.index
        g_StationInfo.startStation = GetRawStation(oSampleLineGroupInfo.mObject.Parent.GetStationStringWithEquations(dStartStation))
        g_StationInfo.endStation = GetRawStation(oSampleLineGroupInfo.mObject.Parent.GetStationStringWithEquations(dEndStation))
        g_StationInfo.curStartStation = GetRawStation(item.SubItems(nStationStartIndex))
        g_StationInfo.curEndStation = GetRawStation(item.SubItems(nStationEndIndex))
        StartStaEdit.Enabled = True
        EndStaEdit.Enabled = True
        StartStaEdit.Value = item.SubItems(nStationStartIndex)
        EndStaEdit.Value = item.SubItems(nStationEndIndex)
    End If
End Sub

Private Sub EndStaEdit_Change()
    Dim curListItem As ListItem
    Set curListItem = ListView.SelectedItem
    If Not curListItem Is Nothing Then
        curListItem.SubItems(nStationEndIndex) = EndStaEdit.Value
    End If
End Sub

Private Sub EndStaEdit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Dim oAlignment As AeccAlignment
    Dim rawStation As Double
    Dim curListItem As ListItem

    Set curListItem = ListView.SelectedItem
    If curListItem Is Nothing Then
       StartStaEdit.Value = ""
       Exit Sub
    End If
    Set oAlignment = g_StationInfo.oSampleLineGroup.Parent
    'verify the station is in right format
    rawStation = CDbl(EndStaEdit.Value)
    If Err Then
        Err.Clear
        rawStation = GetRawStation(EndStaEdit.Value)
        If Err Then
            MsgBox "Please input right format for station."
            rawStation = g_StationInfo.curEndStation
        End If
    End If

    'verify the startStation >= default value
    If rawStation > g_StationInfo.endStation Or rawStation < g_StationInfo.curStartStation Then
        MsgBox "The EndStation Value is out of range"
        rawStation = g_StationInfo.endStation
    End If
    g_StationInfo.curEndStation = rawStation
    EndStaEdit.Value = oAlignment.GetStationStringWithEquations(rawStation)
End Sub

Private Sub EndStaEdit_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> Asc("+") And KeyAscii <> Asc("-") Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub StartStaEdit_Change()
    Dim curListItem As ListItem
    Set curListItem = ListView.SelectedItem
    If Not curListItem Is Nothing Then
        curListItem.SubItems(nStationStartIndex) = StartStaEdit.Value
    End If
End Sub

Private Sub StartStaEdit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Dim oAlignment As AeccAlignment
    Dim rawStation As Double
    Dim curListItem As ListItem

    Set curListItem = ListView.SelectedItem
    If curListItem Is Nothing Then
       StartStaEdit.Value = ""
       Exit Sub
    End If

    Set oAlignment = g_StationInfo.oSampleLineGroup.Parent
    'verify the station is in right format
    rawStation = CDbl(StartStaEdit.Value)
    If Err Then
        Err.Clear
        rawStation = GetRawStation(StartStaEdit.Value)
        If Err Then
            MsgBox "Please input right format for station."
            rawStation = g_StationInfo.curStartStation
        End If
    End If

    'verify the startStation >= default value
    If rawStation > g_StationInfo.curEndStation Or rawStation < g_StationInfo.startStation Then
        MsgBox "The StartStation Value is out of range"
        rawStation = g_StationInfo.startStation
    End If
    g_StationInfo.curStartStation = rawStation
    StartStaEdit.Value = oAlignment.GetStationStringWithEquations(rawStation)
End Sub



Private Sub StartStaEdit_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> Asc("+") And KeyAscii <> Asc("-") Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Frame3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Frame3.ActiveControl = StartStaEdit Then
        StartStaEdit_Exit Cancel
    ElseIf Frame3.ActiveControl = EndStaEdit Then
        EndStaEdit_Exit Cancel
    End If
End Sub


Private Sub FileSelectionButton_Click()

    Dim fileDialog As New CFileDialog
    
    With fileDialog

        .DefaultExt = "html"

        .DialogTitle = "Select a HTML File"

        .Filter = "HTML Files|*.HTML"

        .FilterIndex = 1

        .MaxFileSize = 255

        .FileName = g_reportFileName

        If .Show(True) Then

            g_reportFileName = fileDialog.FileName
            ReportFileNameEdit.Text = g_reportFileName

        Else

            Exit Sub

        End If

    End With

End Sub




Private Sub GetCaptions()
    ReDim sCaptionArr(nLastIndex)
 
    sCaptionArr(nNameIndex) = "Name"
    sCaptionArr(nAlignmentIndex) = "Alignment"
    sCaptionArr(nSampleLineGroup) = "Sample Line Group"
    sCaptionArr(nLinkCodes) = "Link"
    sCaptionArr(nStationStartIndex) = "Station Start"
    sCaptionArr(nStationEndIndex) = "Station End"
End Sub



'Get the start Station and end Station of the Sample Line Group. Assume that the lines in group are sorted.
' if there is no sampleline in group, the function return false and the startstation and endstation is invalid.
Public Function StartStationOfSampleLineGroup(oSampleLineGroup As AeccSampleLineGroup, dStartStation As Double, dEndStation As Double) As Boolean
    Dim oSampleLine As AeccSampleLine
    If oSampleLineGroup.SampleLines.count < 0 Then
        StartStationOfSampleLineGroup = False
    Else
        dStartStation = oSampleLineGroup.SampleLines(0).Station
        dEndStation = dStartStation
        StartStationOfSampleLineGroup = True
        For Each oSampleLine In oSampleLineGroup.SampleLines
            If oSampleLine.Station > dEndStation Then
                dEndStation = oSampleLine.Station
            End If
            If oSampleLine.Station < dStartStation Then
                dStartStation = oSampleLine.Station
            End If
        Next
    End If
    
End Function

Private Sub BuildCorridorInfo(oCorridor As AeccCorridor, oCorridorInfo As CCorridor)
    Dim dicBaselines As Dictionary
    Dim oBaselineInfo As CBaseline
    
    Dim oBaseline As AeccBaseline
    
    
    Set dicBaselines = oCorridorInfo.mBaselines
    Set oCorridorInfo.mObject = oCorridor
    dicBaselines.removeAll
    
    For Each oBaseline In oCorridor.Baselines
        If dicBaselines.Exists(oBaseline.Alignment.Name) = False Then
            Set oBaselineInfo = New CBaseline
            BuildBaselineInfo oCorridor, oBaseline, oBaselineInfo
            dicBaselines.Add oBaseline.Alignment.Name, oBaselineInfo
        End If
    Next
End Sub
Private Sub BuildBaselineInfo(oCorridor As AeccCorridor, oBaseline As AeccBaseline, oBaselineInfo As CBaseline)
    Dim dicSampleLineGroups As Dictionary
    Dim oSampleLineGroupInfo As CSampleLineGroup
    
    Dim oSampleLineGroup As AeccSampleLineGroup
    Dim oBaselineRegion As AeccBaselineRegion
    
    Set dicSampleLineGroups = oBaselineInfo.mSampleLineGroups
    Set oBaselineInfo.mObject = oBaseline
    dicSampleLineGroups.removeAll
    
    Dim dicLinkCodesGroups As Dictionary
    Dim oLinkCodesInfo As CLinkCodes
    Set dicLinkCodesGroups = oBaselineInfo.mLinkCodesGroups
    Set oBaselineInfo.mObject = oBaseline
    dicLinkCodesGroups.removeAll
    
    If oBaseline.Alignment.SampleLineGroups.count > 0 And oBaseline.BaselineRegions.count > 0 Then
        For Each oSampleLineGroup In oBaseline.Alignment.SampleLineGroups
            If dicSampleLineGroups.Exists(oSampleLineGroup.Name) = False Then
                Set oSampleLineGroupInfo = New CSampleLineGroup
                Set oSampleLineGroupInfo.mObject = oSampleLineGroup
                dicSampleLineGroups.Add oSampleLineGroup.Name, oSampleLineGroupInfo
            End If
        Next
        
        Dim oSubAssembly As AeccSubassembly
        For Each oSubAssembly In g_oRoadwayDocument.Subassemblies
            Dim oRoadwayLinks As AeccRoadwayLinks
            Set oRoadwayLinks = oSubAssembly.Links
            Dim oRoadwayLink As AeccRoadwayLink
            For Each oRoadwayLink In oRoadwayLinks
                Dim iIndex As Integer
                For iIndex = 0 To oRoadwayLink.RoadwayCodes.count - 1
                    Dim sCodeName As String
                    sCodeName = oRoadwayLink.RoadwayCodes.item(iIndex)
                    If dicLinkCodesGroups.Exists(sCodeName) = False Then
                        Set oLinkCodesInfo = New CLinkCodes
                        oLinkCodesInfo.mCodeName = sCodeName
                        dicLinkCodesGroups.Add sCodeName, oLinkCodesInfo
                    End If
                Next iIndex
            Next
        Next
        
    End If
End Sub

Private Function FindSampleLineGroupInfo(sCorridor As String, sAlignment As String, sSampleLineGroupName As String) As CSampleLineGroup
    On Error GoTo ErrHandle
    Dim oCorridorInfo As CCorridor
    Dim oBaselineInfo As CBaseline
   
    Set oCorridorInfo = g_Corridors.item(sCorridor)
    Set oBaselineInfo = oCorridorInfo.mBaselines.item(sAlignment)
    Set FindSampleLineGroupInfo = oBaselineInfo.mSampleLineGroups.item(sSampleLineGroupName)
    
    Exit Function
ErrHandle:
    Set FindSampleLineGroupInfo = Nothing
End Function

Private Function FindLinkCodesInfo(sCorridor As String, sAlignment As String, sCodeName As String) As CLinkCodes
    On Error GoTo ErrHandle
    Dim oCorridorInfo As CCorridor
    Dim oBaselineInfo As CBaseline
   
    Set oCorridorInfo = g_Corridors.item(sCorridor)
    Set oBaselineInfo = oCorridorInfo.mBaselines.item(sAlignment)
    Set FindLinkCodesInfo = oBaselineInfo.mLinkCodesGroups.item(sCodeName)
    
    Exit Function
ErrHandle:
    Set FindLinkCodesInfo = Nothing
End Function

Private Function MultiItemsSelected() As Boolean
    Dim count As Integer
    Dim item As ListItem
    MultiItemsSelected = False
    For Each item In ListView.ListItems
        If item.Selected = True Then
            count = count + 1
        End If
        If count = 2 Then
            MultiItemsSelected = True
            Exit Function
        End If
    Next
End Function
