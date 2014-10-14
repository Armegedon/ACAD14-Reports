VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportForm_ProfStaInc 
   Caption         =   "Create Reports -  Incremental Stationing Report"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   OleObjectBlob   =   "ReportForm_ProfStaInc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReportForm_ProfStaInc"
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

Private g_oProfile As AeccProfile
Private g_StartStation As Double
Private g_EndStation As Double
Private g_sFileName As String
Private g_reportFileName As String


Private Const nIncludeIndex = 0
Private Const nNameIndex = 1
Private Const nDescriptionIndex = 2
Private Const nStationStartIndex = 3
Private Const nStationEndIndex = 4
Private Const nAlignNameIndex = 5
Private Const nLastIndex = nAlignNameIndex

Private sCaptionArr() As String
Private g_profileArr()


Private Sub HelpButton_Click()
    ReportsHH_hWnd = ReportsHtmlHelp(0, getReportsHelpFile, HH_HELP_CONTEXT, IDH_AECC_REPORTS_CREATE_INCREMENTAL_STATIONING_REPORT_PROFILES)
End Sub

' Initialize the main dialog
Private Sub UserForm_Initialize()

    Dim tempPath As String

    
    g_sFileName = g_oCivilApp.Path

    tempPath = Environ("temp")
    g_reportFileName = tempPath & "\civilreport.html"
    
' init controls
    ' set report options here
    g_staInc = GetDefaultInc()
    StaIncEdit.Value = CStr(g_staInc)
    
    
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
    Call ListView.ColumnHeaders.Add(, , sCaptionArr(nAlignNameIndex), 100, lvwColumnLeft)

    FillGrid
    
    StartStaEdit.Enabled = False
    EndStaEdit.Enabled = False
 
End Sub
''

' Do the real work of gnerating the report here
Private Sub ExecuteButton_Click()
    Dim t, cCount As Integer
    Dim pbarInc As Single
    Dim oProfile As AeccProfile
    Dim curListItem As ListItem
    Dim exeStr As String
    Dim inStartStation As Double
    Dim inEndStation As Double
   
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
            Set oProfile = g_profileArr(curListItem.Index - 1)
            If Not oProfile Is Nothing Then
                inStartStation = GetRawStation(curListItem.SubItems(nStationStartIndex))
                inEndStation = GetRawStation(curListItem.SubItems(nStationEndIndex))
                Call AppendReport(oProfile, inStartStation, inEndStation, g_reportFileName)
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
    Set g_oProfile = Nothing
    Set g_oCivilApp = Nothing
    Set g_oAeccDoc = Nothing
    Set g_oAeccDb = Nothing

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
    Dim oSite As AeccSite
    Dim oAlignment As AeccAlignment
    Dim oProfile As AeccProfile
    If g_oAeccDb.Sites.Count = 0 And g_oAeccDoc.AlignmentsSiteless.Count = 0 Then
       GoTo NoObject
    End If
    Dim li As ListItem
    Dim nCount As Long
    Dim nCurIndex As Long
    
    'get alignments from Document
    nCurIndex = 0
    For Each oAlignment In g_oAeccDoc.AlignmentsSiteless
        nCount = nCount + oAlignment.Profiles.Count
        ReDim Preserve g_profileArr(nCount)
        For Each oProfile In oAlignment.Profiles
            If oProfile.Type = aeccFinishedGround Then
                Set li = ListView.ListItems.Add
                li.SubItems(nNameIndex) = oProfile.Name
                li.SubItems(nDescriptionIndex) = oProfile.Description
                li.SubItems(nStationStartIndex) = oAlignment.GetStationStringWithEquations(oProfile.StartingStation)
                li.SubItems(nStationEndIndex) = oAlignment.GetStationStringWithEquations(oProfile.EndingStation)
                li.SubItems(nAlignNameIndex) = oProfile.Alignment.Name
                li.Checked = True
                Set g_profileArr(nCurIndex) = oProfile
                nCurIndex = nCurIndex + 1
            End If
        Next
    Next
    
    'get alignments from Site
    For Each oSite In g_oAeccDb.Sites
        For Each oAlignment In oSite.Alignments
            nCount = nCount + oAlignment.Profiles.Count
            ReDim Preserve g_profileArr(nCount)
            For Each oProfile In oAlignment.Profiles
                If oProfile.Type = aeccFinishedGround Then
                    Set li = ListView.ListItems.Add
                    li.SubItems(nNameIndex) = oProfile.Name
                    li.SubItems(nDescriptionIndex) = oProfile.Description
                    li.SubItems(nStationStartIndex) = oAlignment.GetStationStringWithEquations(oProfile.StartingStation)
                    li.SubItems(nStationEndIndex) = oAlignment.GetStationStringWithEquations(oProfile.EndingStation)
                    li.SubItems(nAlignNameIndex) = oProfile.Alignment.Name
                    li.Checked = True
                    Set g_profileArr(nCurIndex) = oProfile
                    nCurIndex = nCurIndex + 1
                End If
            Next
        Next
    Next
    If nCurIndex = 0 Then
        GoTo NoObject
    End If
    Exit Sub
NoObject:
        MsgBox "No FG profiles in " & ThisDrawing.Name & " exist."
        Err.Number = 1
End Sub

Private Sub GetCaptions()
    ReDim sCaptionArr(nLastIndex)
    sCaptionArr(nIncludeIndex) = "Include"
    sCaptionArr(nNameIndex) = "Name"
    sCaptionArr(nDescriptionIndex) = "Description"
    sCaptionArr(nStationStartIndex) = "Station Start"
    sCaptionArr(nStationEndIndex) = "Station End"
    sCaptionArr(nAlignNameIndex) = "Alignment"
End Sub
'

Private Sub Frame3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Frame3.ActiveControl = StartStaEdit Then
        StartStaEdit_Exit Cancel
    ElseIf Frame3.ActiveControl = EndStaEdit Then
        EndStaEdit_Exit Cancel
    End If
End Sub

Private Sub ListView_ItemClick(ByVal Item As ListItem)
    'MsgBox ""
    StartStaEdit.Enabled = True
    EndStaEdit.Enabled = True
    StartStaEdit.Value = Item.SubItems(nStationStartIndex)
    EndStaEdit.Value = Item.SubItems(nStationEndIndex)
    g_StartStation = GetRawStation(StartStaEdit.Value)
    g_EndStation = GetRawStation(EndStaEdit.Value)
    Set g_oProfile = g_profileArr(Item.Index - 1)
End Sub

Private Sub StartStaEdit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Dim rawStation As Double
    Dim defaultStation As Double
    Dim curListItem As ListItem
    
    Set curListItem = ListView.SelectedItem
    If g_oProfile Is Nothing Or curListItem Is Nothing Then
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
    If rawStation < GetRoundedStationVal(g_oProfile.StartingStation, g_oProfile.Alignment) Or rawStation > g_EndStation Then
        MsgBox "The StartStation Value is out of range"
        rawStation = g_oProfile.StartingStation
    End If
    g_StartStation = rawStation
    StartStaEdit.Value = g_oProfile.Alignment.GetStationStringWithEquations(rawStation)
End Sub


Private Sub StartStaEdit_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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

Private Sub EndStaEdit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Dim rawStation As Double
    Dim defaultStation As Double
    Dim curListItem As ListItem
    
    Set curListItem = ListView.SelectedItem
    If g_oProfile Is Nothing Or curListItem Is Nothing Then
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
    If rawStation > GetRoundedStationVal(g_oProfile.EndingStation, g_oProfile.Alignment) Or rawStation < g_StartStation Then
        MsgBox "The EndStation Value is out of range"
        rawStation = g_oProfile.EndingStation
    End If
    g_EndStation = rawStation
    EndStaEdit.Value = g_oProfile.Alignment.GetStationStringWithEquations(rawStation)
End Sub


Private Sub EndStaEdit_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> Asc("+") And KeyAscii <> Asc("-") Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub EndStaEdit_Change()
    Dim curListItem As ListItem
    Set curListItem = ListView.SelectedItem
    If Not curListItem Is Nothing Then
        curListItem.SubItems(nStationEndIndex) = EndStaEdit.Value
    End If
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


Private Function GetDefaultInc() As Double
    Dim fInc As Double
    If GetUnitStr() = "Feet" Then
        fInc = 50#
    Else
        fInc = 25#
    End If
    GetDefaultInc = fInc
End Function

Private Function GetRoundedStationVal(rawVal As Double, oAlignment As AeccAlignment) As Double
    GetRoundedStationVal = GetRawStation(oAlignment.GetStationStringWithEquations(rawVal))
End Function
