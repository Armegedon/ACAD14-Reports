VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportForm_PointsStaOffset 
   Caption         =   "Create Reports - Station Offset to Points Report"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12375
   OleObjectBlob   =   "ReportForm_PointsStaOffset.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReportForm_PointsStaOffset"
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
Private g_AlignsArr()

Private g_sFileName As String
Private g_reportFileName As String

Private g_curAlign As AeccAlignment


Private Const nIncludeIndex = 0
Private Const nPtNumberIndex = 1
Private Const nNorthingIndex = 2
Private Const nEastingIndex = 3
Private Const nElevationIndex = 4
Private Const nNameIndex = 5
Private Const nRawDesriptionIndex = 6
Private Const nFullDescriptionIndex = 7
Private Const nLastIndex = nFullDescriptionIndex

Private sCaptionArr() As String


Private Sub Combo_Align_Change()
    Set g_curAlign = g_AlignsArr(Combo_Align.ListIndex)
    Label_AlignName = g_curAlign.name
    Label_StaRange = "Start:" & g_curAlign.GetStationStringWithEquations(g_curAlign.StartingStation) & "--" _
                      & "End:" & g_curAlign.GetStationStringWithEquations(g_curAlign.EndingStation)
  
    If g_curAlign.StationEquations.Count = 0 Then
        Label_StaEquation = "None"
    Else
        Label_StaEquation = "Applied"
    End If
End Sub



Private Sub DeselAllButton_Click()
    Dim lItem As listItem
    For Each lItem In ListView.ListItems
        lItem.Checked = False
    Next
End Sub



Private Sub HelpButton_Click()
    ReportsHH_hWnd = ReportsHtmlHelp(0, getReportsHelpFile, HH_HELP_CONTEXT, IDH_AECC_REPORTS_CREATE_STATION_OFFSET_TO_POINTS_REPORT)
End Sub

Private Sub SelAlignButton_Click()
    On Error GoTo Cancel
    Dim oEntity As Object
    Dim pt(2) As Double
    Me.Hide
    ThisDrawing.Utility.GetEntity oEntity, pt, "Select a Alignment."
    Do While Not TypeOf oEntity Is AeccAlignment
        ThisDrawing.Utility.GetEntity oEntity, pt, "Please select a Alignment."
    Loop
    Combo_Align.Text = oEntity.name
Cancel:
    Me.Show
End Sub

Private Sub SelAllButton_Click()
    Dim lItem As listItem
    For Each lItem In ListView.ListItems
        lItem.Checked = True
    Next
End Sub

' Initialize the main dialog
Private Sub UserForm_Initialize()
    Dim tempPath As String
    Dim oSite As AeccSite
    Dim oAlign As AeccAlignment
    Dim nAlignCount As Long
    Dim nCurAlignIndex As Long
    
    If g_oAeccDb.Points.Count = 0 Then
        MsgBox "No points in " & ThisDrawing.name & " exist."
        End
    End If
    
    
    g_sFileName = g_oCivilApp.Path

    tempPath = Environ("temp")
    g_reportFileName = tempPath & "\civilreport.html"


' init controls
    ReportFileNameEdit.Text = g_reportFileName

    ' progress bar settings
    ProgressBar.Min = 0
    ProgressBar.Max = 100
    ProgressBar.Visible = False
    progressBarLabel.Visible = False


    
    
    ' Get Captions
    GetCaptions
    
    ' list view grid setup
    ' setup columns
    Call ListView.ColumnHeaders.Add(, , sCaptionArr(nIncludeIndex), 40, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(, , sCaptionArr(nPtNumberIndex), 80, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(, , sCaptionArr(nNorthingIndex), 60, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(, , sCaptionArr(nEastingIndex), 60, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(, , sCaptionArr(nElevationIndex), 80, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(, , sCaptionArr(nNameIndex), 40, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(, , sCaptionArr(nRawDesriptionIndex), 100, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(, , sCaptionArr(nFullDescriptionIndex), 100, lvwColumnLeft)
    
    

    FillGrid

    
    'Alignments
    nAlignCount = 0
    nCurAlignIndex = 0
    
    
    nAlignCount = g_oAeccDoc.AlignmentsSiteless.Count
    If nAlignCount > 0 Then
        ReDim Preserve g_AlignsArr(nAlignCount - 1)
    End If
    For Each oAlign In g_oAeccDoc.AlignmentsSiteless
        Set g_AlignsArr(nCurAlignIndex) = oAlign
        Combo_Align.AddItem (oAlign.name)
        nCurAlignIndex = nCurAlignIndex + 1
    Next

    
    For Each oSite In g_oAeccDb.Sites
        nAlignCount = oSite.Alignments.Count + nAlignCount
        If nAlignCount > 0 Then
            ReDim Preserve g_AlignsArr(nAlignCount - 1)
        End If
        For Each oAlign In oSite.Alignments
            Set g_AlignsArr(nCurAlignIndex) = oAlign
            Combo_Align.AddItem (oAlign.name)
            nCurAlignIndex = nCurAlignIndex + 1
        Next
    Next
    If nAlignCount > 0 Then
        Combo_Align.ListIndex = 0
    Else
        MsgBox "No alignments in " & ThisDrawing.name & " exist."
        End
    End If

End Sub


' Do the real work of gnerating the report here
Private Sub ExecuteButton_Click()
    Dim t, cCount As Integer
    Dim pbarInc As Single
    Dim oPoint As AeccPoint
    Dim curListItem As listItem
    Dim exeStr As String
    

    If ListView.ListItems.Count = 0 Then
        Exit Sub
    End If

    For t = 1 To ListView.ListItems.Count
        If ListView.ListItems(t).Checked = True Then
            cCount = cCount + 1
        End If
    Next
'
    If cCount = 0 Then
        MsgBox "Please select at least one Point."
        Exit Sub
    End If
    pbarInc = 90 / cCount
'
    progressBarLabel.Enabled = False
    progressBarLabel.Enabled = True

    progressBarLabel.Visible = True
    ProgressBar.Visible = True
'
        'update progress bar
    ProgressBar.Value = 0
    
    Call WriteReportHeader(g_curAlign, g_reportFileName)
'
    For t = 1 To ListView.ListItems.Count
        Set curListItem = ListView.ListItems(t)
        If curListItem.Checked = True Then
            Set oPoint = g_oAeccDb.Points.Find(curListItem.SubItems(nPtNumberIndex))
            If Not oPoint Is Nothing Then
                Call AppendReport(g_curAlign, oPoint, g_reportFileName)
                ProgressBar.Value = ProgressBar.Value + pbarInc
            End If
        End If
    Next
'
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
    Set g_oAeccDoc = Nothing
    Set g_oAeccDb = Nothing
    Set g_curAlign = Nothing
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

' Fill the data to Grid
Private Sub FillGrid()
    Dim oPoints As AeccPoints
    Dim oPoint As AeccPoint
    Dim li As listItem
    For Each oPoint In g_oAeccDb.Points
        Set li = ListView.ListItems.Add
        li.SubItems(nPtNumberIndex) = oPoint.Number
        li.SubItems(nNorthingIndex) = FormatPtCoordSettings(oPoint.Northing)
        li.SubItems(nEastingIndex) = FormatPtCoordSettings(oPoint.Easting)
        li.SubItems(nElevationIndex) = FormatPtElevationSettings(oPoint.elevation)
        On Error Resume Next
        Dim name As String
        name = ""
        name = oPoint.name
        On Error GoTo 0
        li.SubItems(nNameIndex) = name
        li.SubItems(nRawDesriptionIndex) = oPoint.RawDescription
        li.SubItems(nFullDescriptionIndex) = oPoint.FullDescription
        li.Checked = True
    Next


End Sub


Private Sub GetCaptions()
    ReDim sCaptionArr(nLastIndex)
    sCaptionArr(nIncludeIndex) = "Include"
    sCaptionArr(nPtNumberIndex) = "Point Number"
    sCaptionArr(nNorthingIndex) = "Northing"
    sCaptionArr(nEastingIndex) = "Easting"
    sCaptionArr(nElevationIndex) = "Point Elevation"
    sCaptionArr(nNameIndex) = "Name"
    sCaptionArr(nRawDesriptionIndex) = "Raw Description"
    sCaptionArr(nFullDescriptionIndex) = "Full Description"
End Sub
