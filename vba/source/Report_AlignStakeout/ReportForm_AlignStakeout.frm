VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportForm_AlignStakeout 
   Caption         =   "Create Reports - Stakeout Alignment Report"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9405
   OleObjectBlob   =   "ReportForm_AlignStakeout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReportForm_AlignStakeout"
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

Private g_oAlignment As AeccAlignment
Private g_StartStation As Double
Private g_EndStation As Double
Private g_sFileName As String
Private g_reportFileName As String


Private Const nIncludeIndex = 0
Private Const nNameIndex = 1
Private Const nDescriptionIndex = 2
Private Const nStationStartIndex = 3
Private Const nStationEndIndex = 4
Private Const nLastIndex = nStationEndIndex

Private sCaptionArr() As String
Private g_alignArr()
Private pointDict As Dictionary



Private Sub HelpButton_Click()
    ReportsHH_hWnd = ReportsHtmlHelp(0, getReportsHelpFile, HH_HELP_CONTEXT, IDH_AECC_REPORTS_CREATE_STAKEOUT_ALIGNMENT_REPORT)
End Sub

' Initialize the main dialog
Private Sub UserForm_Initialize()

    Dim tempPath As String


    g_sFileName = g_oCivilApp.Path

    tempPath = Environ("temp")
    g_reportFileName = tempPath & "\civilreport.html"

' get the point dictionary from current drawing database
    If g_oAeccDb.Points.Count < 2 Then
        MsgBox "At least 2 Points exist in current drawing!"
        Err.Number = 1
        Exit Sub
    End If

    Dim iIndex As Integer
    Dim currentPoint As AeccPoint
    Set pointDict = New Dictionary
    For iIndex = 0 To g_oAeccDb.Points.Count - 1
        Set currentPoint = g_oAeccDb.Points.Item(iIndex)
        If Not currentPoint Is Nothing And Not pointDict.Exists(currentPoint.Number) Then
            pointDict.Add currentPoint.Number, currentPoint
        End If
    Next iIndex
    
    If pointDict.Count < 2 Then
        MsgBox "At least 2 Points exist in current drawing!"
        Err.Number = 1
        Exit Sub
    End If

' set report options here
    g_staInc = GetDefaultInc()
    g_offSet = 0#
' set angle type option here
    OptionTurnedPuls.Value = True
    g_angleType = TurnedPlus
    
' init controls
    StaIncEdit.Value = CStr(g_staInc)
    OffsetEdit.Value = CStr(g_offSet)
    
    ReportFileNameEdit.Text = g_reportFileName

    ' progress bar settings
    ProgressBar.min = 0
    ProgressBar.Max = 100
    ProgressBar.Visible = False
    progressBarLabel.Visible = False


    ' Get Captions
    GetCaptions

    ' list view grid setup
    ' setup columns
    Call ListView.ColumnHeaders.Add(, , sCaptionArr(nIncludeIndex), 40, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(, , sCaptionArr(nNameIndex), 100, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(, , sCaptionArr(nDescriptionIndex), 100, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(, , sCaptionArr(nStationStartIndex), 80, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(, , sCaptionArr(nStationEndIndex), 80, lvwColumnLeft)

    FillGrid

    StartStaEdit.Enabled = False
    EndStaEdit.Enabled = False

End Sub
''
' Radio button for angle type
Private Sub OptionDeflectMinus_Click()
    g_angleType = DeflectMinus
End Sub

Private Sub OptionDeflectPlus_Click()
    g_angleType = DeflectPlus
End Sub

Private Sub OptionDirection_Click()
    g_angleType = Direction
End Sub

Private Sub OptionTurnedMinus_Click()
    g_angleType = TurnedMinus
End Sub

Private Sub OptionTurnedPuls_Click()
    g_angleType = TurnedPlus
End Sub

Private Sub PointOccupiedEdit_KeyPress(ByVal KeyAscii As ReturnInteger)

    If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
        Beep
        KeyAscii = 0
    End If
End Sub
Private Sub PointOccupiedEdit_Exit(ByVal Cancel As ReturnBoolean)
    Dim oPoints As AeccPoints
    
    If PointOccupiedEdit.Value = "" Then
        Exit Sub
    End If
        
    If pointDict.Exists(CLng(PointOccupiedEdit.Value)) Then
        Set g_oOccupiedPt = pointDict.Item(CLng(PointOccupiedEdit.Value))
    Else
        MsgBox "Please input the correct Point Number"
        If g_oOccupiedPt Is Nothing Then
            PointOccupiedEdit.Value = ""
        Else
            PointOccupiedEdit.Value = g_oOccupiedPt.Number
        End If
    End If
    
End Sub

Private Sub BacksightPointEdit_KeyPress(ByVal KeyAscii As ReturnInteger)
    If KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
        Beep
        KeyAscii = 0
    End If
End Sub
Private Sub BacksightPointEdit_Exit(ByVal Cancel As ReturnBoolean)
    Dim oPoints As AeccPoints
    
    If BacksightPointEdit.Value = "" Then
        Exit Sub
    End If
    
    If pointDict.Exists(CLng(BacksightPointEdit.Value)) Then
        Set g_oBacksightPt = pointDict.Item(CLng(BacksightPointEdit.Value))
    Else
        MsgBox "Please input the correct Point Number"
        If g_oBacksightPt Is Nothing Then
            BacksightPointEdit.Value = ""
        Else
            BacksightPointEdit.Value = g_oBacksightPt.Number
        End If
    End If
    
End Sub

Private Sub FrameStakeout_Exit(ByVal Cancel As ReturnBoolean)
    If FrameStakeout.ActiveControl = PointOccupiedEdit Then
        PointOccupiedEdit_Exit Cancel
    ElseIf FrameStakeout.ActiveControl = BacksightPointEdit Then
        BacksightPointEdit_Exit Cancel
    End If
End Sub

Private Sub Frame3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Frame3.ActiveControl = StartStaEdit Then
        StartStaEdit_Exit Cancel
    ElseIf Frame3.ActiveControl = EndStaEdit Then
        EndStaEdit_Exit Cancel
    End If
End Sub

Private Sub BacksightButton_Click()
    Dim sSet As AcadSelectionSet         'Define sset as a SelectionSet object
    Dim oPoint As AeccPoint
    'Set sset to a new selection set named SS1 (the name doesn't matter here)
    
    On Error Resume Next
    If Not IsNull(ThisDrawing.SelectionSets.Item("Points")) Then
        Set sSet = ThisDrawing.SelectionSets.Item("Points")
     sSet.Delete
    End If
  
    Set sSet = ThisDrawing.SelectionSets.Add("Points")
   
    
    Me.Hide  'Hide the current dialog
    
    sSet.SelectOnScreen       'Prompt user to select objects
    ' continue select on screen until select nothing or select a single point
    Do While Not sSet.Count = 0
        If sSet.Count = 1 Then
            If TypeOf sSet.Item(0) Is AeccPoint Then
                Exit Do
            End If
        End If
        MsgBox "Please select only one point."
        sSet.Clear
        sSet.SelectOnScreen       'Prompt user to select objects
    Loop
    
    
    If sSet.Count = 1 Then
        Dim iNum As Long
        Set oPoint = sSet.Item(0)
        iNum = oPoint.Number
        Set g_oBacksightPt = pointDict.Item(iNum)
        BacksightPointEdit.Value = iNum
    End If
         
    Me.Show ' restore the dialog
End Sub

Private Sub OccupiedButton_Click()
    
    Dim sSet As AcadSelectionSet         'Define sset as a SelectionSet object
    Dim oPoint As AeccPoint
    'Set sset to a new selection set named SS1 (the name doesn't matter here)
    
    On Error Resume Next
    If Not IsNull(ThisDrawing.SelectionSets.Item("Points")) Then
        Set sSet = ThisDrawing.SelectionSets.Item("Points")
     sSet.Delete
    End If
  
    Set sSet = ThisDrawing.SelectionSets.Add("Points")
   
    
    Me.Hide  'Hide the current dialog
    
    sSet.SelectOnScreen       'Prompt user to select objects
    
    
   ' continue select on screen until select nothing or select a single point
    Do While Not sSet.Count = 0
        If sSet.Count = 1 Then
            If TypeOf sSet.Item(0) Is AeccPoint Then
                Exit Do
            End If
        End If
        MsgBox "Please select only one point."
        sSet.Clear
        sSet.SelectOnScreen       'Prompt user to select objects
    Loop
    
    
    If sSet.Count = 1 Then
        Dim iNum As Long
        Set oPoint = sSet.Item(0)
        iNum = oPoint.Number
        Set g_oOccupiedPt = pointDict.Item(iNum)
        PointOccupiedEdit.Value = iNum
    End If
         
    Me.Show ' restore the dialog
End Sub

Private Sub StaIncEdit_Change()

    On Error Resume Next
    If StaIncEdit.Value <> "" Then
        g_staInc = CDbl(StaIncEdit.Value)
        If Err Or g_staInc < 0# Then
            g_staInc = GetDefaultInc()
            StaIncEdit.Value = g_staInc
        End If
    End If

End Sub
'


Private Sub StaIncSpinner_SpinDown()

    If StaIncEdit.Value > 0 Then
        StaIncEdit.Value = StaIncEdit.Value - 1
    End If

End Sub

Private Sub StaIncSpinner_SpinUp()
    If StaIncEdit.Value < 999 Then
        StaIncEdit.Value = StaIncEdit.Value + 1
    End If
End Sub
Private Sub OffsetEdit_Change()
    On Error Resume Next

    If OffsetEdit.Value <> "" Then
        g_offSet = CDbl(OffsetEdit.Value)
        If Err Or g_offSet < 0# Then
            g_offSet = 0#
            OffsetEdit.Value = 0#
        End If
    End If
End Sub
Private Sub OffsetSpinner_SpinDown()

    If OffsetEdit.Value > 0 Then
        OffsetEdit.Value = OffsetEdit.Value - 1
    End If

End Sub

Private Sub OffsetSpinner_SpinUp()
    If OffsetEdit.Value < 999 Then
        OffsetEdit.Value = OffsetEdit.Value + 1
    End If
End Sub
' Do the real work of gnerating the report here
Private Sub ExecuteButton_Click()
    Dim t, cCount As Integer
    Dim pbarInc As Single
    Dim oAlignment As AeccAlignment
    Dim curListItem As ListItem
    Dim exeStr As String
    Dim inStartStation As Double
    Dim inEndStation As Double
    
    If VerifyBeforeExe = False Then
        Exit Sub
    End If
    
    If ListView.ListItems.Count = 0 Then
        Exit Sub
    End If

    ' check if the increment is too small
    If g_staInc < 0.01 Then
        MsgBox "The station increment value is invalid."
        Exit Sub
    End If
    
    For t = 1 To ListView.ListItems.Count
        If ListView.ListItems(t).Checked = True Then
            cCount = cCount + 1
        End If
    Next

    pbarInc = 90 / cCount
    
    progressBarLabel.Enabled = False
    progressBarLabel.Enabled = True

    progressBarLabel.Visible = True
    ProgressBar.Visible = True
    
    Call WriteReportHeader(g_reportFileName)

    'update progress bar
    ProgressBar.Value = 0

    For t = 1 To ListView.ListItems.Count
        Set curListItem = ListView.ListItems(t)
        If curListItem.Checked = True Then
            Set oAlignment = g_alignArr(curListItem.Index - 1)
            If Not oAlignment Is Nothing Then
                inStartStation = GetRawStation(curListItem.SubItems(nStationStartIndex))
                inEndStation = GetRawStation(curListItem.SubItems(nStationEndIndex))
                Call AppendReport(oAlignment, inStartStation, inEndStation, g_reportFileName)
                ProgressBar.Value = ProgressBar.Value + pbarInc
            End If
        End If
    Next

    'update progress bar
    ProgressBar.Value = 100

    ProgressBar.Visible = False
    progressBarLabel.Visible = False


    Call WriteReportFooter(g_reportFileName)

    OpenFileByDefaultBrowser g_reportFileName

End Sub
'
Private Sub CancelButton_Click()

    Unload Me

End Sub
'

'
Private Sub UserForm_Terminate()

    ' clean up created objects
    Set g_oCivilApp = Nothing
    Set g_oAeccDoc = Nothing
    Set g_oAeccDb = Nothing
    Set g_oAlignment = Nothing
    Set pointDict = Nothing
    Set g_oOccupiedPt = Nothing
    Set g_oBacksightPt = Nothing
    
End Sub
Private Sub ListView_ItemClick(ByVal Item As ListItem)
    'MsgBox ""
    StartStaEdit.Enabled = True
    EndStaEdit.Enabled = True
    StartStaEdit.Value = Item.SubItems(nStationStartIndex)
    EndStaEdit.Value = Item.SubItems(nStationEndIndex)
    g_StartStation = GetRawStation(StartStaEdit.Value)
    g_EndStation = GetRawStation(EndStaEdit.Value)
    Set g_oAlignment = g_alignArr(Item.Index - 1)
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
    Dim rawStation As Double
    Dim defaultStation As Double
    Dim curListItem As ListItem
    
    Set curListItem = ListView.SelectedItem
    If g_oAlignment Is Nothing Or curListItem Is Nothing Then
       StartStaEdit.Value = ""
       Exit Sub
    End If
    
    'verify the station is in right format
    rawStation = CDbl(EndStaEdit.Value)
    If Err Then
        Err.Clear
        rawStation = GetRawStation(EndStaEdit.Value)
        If Err Then
            MsgBox "Please input right format for station."
            rawStation = g_EndStation
        End If
    End If
    
    'verify the startStation >= default value
    If rawStation > GetRoundedStationVal(g_oAlignment.EndingStation, g_oAlignment) Or rawStation < g_StartStation Then
        MsgBox "The EndStation Value is out of range"
        rawStation = g_oAlignment.EndingStation
    End If
    g_EndStation = rawStation
    EndStaEdit.Value = g_oAlignment.GetStationStringWithEquations(rawStation)
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
    Dim rawStation As Double
    Dim defaultStation As Double
    Dim curListItem As ListItem

    Set curListItem = ListView.SelectedItem
    If g_oAlignment Is Nothing Or curListItem Is Nothing Then
       StartStaEdit.Value = ""
       Exit Sub
    End If

    'verify the station is in right format
    rawStation = CDbl(StartStaEdit.Value)
    If Err Then
        Err.Clear
        rawStation = GetRawStation(StartStaEdit.Value)
        If Err Then
            MsgBox "Please input right format for station."
            rawStation = g_StartStation
        End If
    End If

    'verify the startStation >= default value
    If rawStation > g_EndStation Or rawStation < GetRoundedStationVal(g_oAlignment.StartingStation, g_oAlignment) Then
        MsgBox "The StartStation Value is out of range"
        rawStation = g_oAlignment.StartingStation
    End If
    g_StartStation = rawStation
    StartStaEdit.Value = g_oAlignment.GetStationStringWithEquations(rawStation)
End Sub



Private Sub StartStaEdit_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> Asc("+") And KeyAscii <> Asc("-") Then
        Beep
        KeyAscii = 0
    End If
End Sub
'


'
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
'
'' Fill the data to Grid

Private Sub FillGrid()
    Dim oAlignment As AeccAlignment
    Dim oSite As AeccSite
    If g_oAeccDb.Sites.Count = 0 And g_oAeccDoc.AlignmentsSiteless.Count = 0 Then
       GoTo NoObject
    End If
    Dim li As ListItem
    Dim nCount As Long
    Dim nCurIndex As Long
    
    'get alignments from Document
    nCurIndex = 0
    nCount = g_oAeccDoc.AlignmentsSiteless.Count
    ReDim Preserve g_alignArr(nCount)
    For Each oAlignment In g_oAeccDoc.AlignmentsSiteless
        Set li = ListView.ListItems.Add
        li.SubItems(nNameIndex) = oAlignment.Name
        li.SubItems(nDescriptionIndex) = oAlignment.Description
        li.SubItems(nStationStartIndex) = oAlignment.GetStationStringWithEquations(oAlignment.StartingStation)
        li.SubItems(nStationEndIndex) = oAlignment.GetStationStringWithEquations(oAlignment.EndingStation)
        li.Checked = True
        Set g_alignArr(nCurIndex) = oAlignment
        nCurIndex = nCurIndex + 1
    Next
    
    'get alignments from Site
    For Each oSite In g_oAeccDb.Sites
        nCount = nCount + oSite.Alignments.Count
        ReDim Preserve g_alignArr(nCount)
        For Each oAlignment In oSite.Alignments
             Set li = ListView.ListItems.Add
             li.SubItems(nNameIndex) = oAlignment.Name
             li.SubItems(nDescriptionIndex) = oAlignment.Description
             li.SubItems(nStationStartIndex) = oAlignment.GetStationStringWithEquations(oAlignment.StartingStation)
             li.SubItems(nStationEndIndex) = oAlignment.GetStationStringWithEquations(oAlignment.EndingStation)
             li.Checked = True
             Set g_alignArr(nCurIndex) = oAlignment
             nCurIndex = nCurIndex + 1
        Next
    Next
    If nCount = 0 Then
        GoTo NoObject
    End If
    Exit Sub
NoObject:
        MsgBox "No alignments in " & ThisDrawing.Name & " exist."
        Err.Number = 1
End Sub

Private Sub GetCaptions()
    ReDim sCaptionArr(nLastIndex)
    sCaptionArr(nIncludeIndex) = "Include"
    sCaptionArr(nNameIndex) = "Name"
    sCaptionArr(nDescriptionIndex) = "Description"
    sCaptionArr(nStationStartIndex) = "Station Start"
    sCaptionArr(nStationEndIndex) = "Station End"
End Sub
'
Private Function GetDefaultInc() As Double
    Dim fInc As Double
    If GetUnitStr() = "Feet" Then
        fInc = 50#
    Else
        fInc = 25#
    End If
    GetDefaultInc = fInc
End Function

Private Function VerifyBeforeExe() As Boolean
    VerifyBeforeExe = False
    If g_oOccupiedPt Is Nothing Then
        MsgBox "Please select a occupied Point"
    ElseIf Not g_angleType = Direction Then
        If g_oBacksightPt Is Nothing Then
            MsgBox "Please select a backsight Point"
        ElseIf g_oOccupiedPt.Northing = g_oBacksightPt.Northing And g_oOccupiedPt.Easting = g_oBacksightPt.Easting Then
            MsgBox "The Backsight Point can not be the same as Occupied point"
        Else
            VerifyBeforeExe = True
        End If
    Else
        VerifyBeforeExe = True
    End If
End Function

Private Function GetRoundedStationVal(rawVal As Double, oAlignment As AeccAlignment) As Double
    GetRoundedStationVal = GetRawStation(oAlignment.GetStationStringWithEquations(rawVal))
End Function
