VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportForm_HECRAS 
   Caption         =   "Create Reports - HEC-RAS Geographical Data Report"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
   OleObjectBlob   =   "ReportForm_HECRAS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReportForm_HECRAS"
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

Private g_sFileName As String
Private g_reportFileName As String


Private g_Steams As New Dictionary

Private Const nNameIndex = 0
Private Const nBeginPtIndex = 1
Private Const nEndPtIndex = 2
Private Const nLastIndex = nEndPtIndex

Private sCaptionArr() As String


Private Sub DeleteButton_Click()

    Dim i As Long
    Dim item As ListItem
    Dim colItems As ListItems
    Dim oSteam As CSTEAM
    
    If Not g_Steams.Exists(Combo_LineGroup.Text) Then
        Exit Sub
    End If
    Set oSteam = g_Steams.item(Combo_LineGroup.Text)
    
    Set colItems = ListView.ListItems
    For i = 1 To colItems.Count Step 1
        Set item = colItems.item(i)
        If item.Selected = True Then
            Call oSteam.RemoveReach(item.Text)
            colItems.Remove i
            i = i - 1
        End If
        If i = colItems.Count Then
            Exit For
        End If
    Next i
    If ListView.ListItems.Count = 0 Then
        DeleteButton.Visible = False
    End If

End Sub
Private Sub AddButton_Click()
    Dim oSteam As CSTEAM
    Dim lBeginPtNum As Long
    Dim lEndPtNum As Long
    
    If Not g_Steams.Exists(Combo_LineGroup.Text) Then
        Exit Sub
    End If
    Set oSteam = g_Steams.item(Combo_LineGroup.Text)
    lBeginPtNum = Combo_BeginPt.ListIndex + 1
    lEndPtNum = Combo_EndPt.ListIndex + 1
    
    If oSteam.AddReach(TextBox_Name.Text, lBeginPtNum, lEndPtNum) Then
        Dim item As ListItem
        Set item = ListView.ListItems.Add(, , TextBox_Name.Text)
        item.SubItems(nBeginPtIndex) = Combo_BeginPt.Text
        item.SubItems(nEndPtIndex) = Combo_EndPt.Text
        DeleteButton.Visible = True
    Else
        MsgBox "The Reach's name is existing or the Begin and End Point input error."
    End If
    
End Sub

Private Sub Combo_LineGroup_Change()
    TextBox_Name.Text = ""
    Dim oSteam As CSTEAM
    Dim lBeginPtNum As Long
    Dim lEndPtNum As Long
    Dim dicReaches As Dictionary
    Dim oAlignment As AeccAlignment
    Dim sReachName As String
    Dim oReachRegion As CReachRegion
    Dim i As Long
        
    If Not g_Steams.Exists(Combo_LineGroup.Text) Then
        Exit Sub
    End If
    Set oSteam = g_Steams.item(Combo_LineGroup.Text)
    Set dicReaches = oSteam.mReaches
    Set oAlignment = oSteam.SampleLineGroup.Parent
    
    ListView.ListItems.Clear
    For i = 0 To dicReaches.Count - 1 Step 1
        Dim item As ListItem
        sReachName = dicReaches.Keys(i)
        Set oReachRegion = dicReaches.item(sReachName)
        Set item = ListView.ListItems.Add(, , sReachName)
        item.SubItems(nBeginPtIndex) = oAlignment.GetStationStringWithEquations(oReachRegion.BeginPt)
        item.SubItems(nEndPtIndex) = oAlignment.GetStationStringWithEquations(oReachRegion.EndPt)
    Next i
    
End Sub

Private Sub Combo_Align_Change()
    Dim oAlignment As AeccAlignment
    Dim oSampleLineGroup As AeccSampleLineGroup
    Dim Station As AeccAlignmentStation
    Dim oAlignmentPIs As CAlignmentPIs
    Dim i As Long
    
    Set oAlignment = g_Alignments.item(Combo_Align.SelText).Alignment
    Combo_LineGroup.Clear
    
    For Each oSampleLineGroup In oAlignment.SampleLineGroups
        Combo_LineGroup.AddItem oSampleLineGroup.Name
    Next
    
    If Combo_LineGroup.ListCount > 0 Then
        Combo_LineGroup.ListIndex = 0
    End If
    
    Combo_BeginPt.Clear
    Combo_EndPt.Clear

    Set oAlignmentPIs = g_Alignments.item(Combo_Align.Text)
    For i = 1 To oAlignmentPIs.PICounts Step 1
        Combo_BeginPt.AddItem oAlignment.GetStationStringWithEquations(oAlignmentPIs.GetStationByNum(i))
        Combo_EndPt.AddItem oAlignment.GetStationStringWithEquations(oAlignmentPIs.GetStationByNum(i))
    Next
    ' Initial Reach begin and end points list box
    
    If Combo_BeginPt.ListCount > 0 Then
        Combo_BeginPt.ListIndex = 0
    End If
    If Combo_EndPt.ListCount > 0 Then
        Combo_EndPt.ListIndex = Combo_EndPt.ListCount - 1
    End If
End Sub



Private Sub HelpButton_Click()
    ReportsHH_hWnd = ReportsHtmlHelp(0, getReportsHelpFile, HH_HELP_CONTEXT, IDH_AECC_REPORTS_CREATE_HEC_RAS_REPORT)
End Sub

Private Sub TextBox_Name_Change()
    If TextBox_Name.Text <> "" Then
        Combo_BeginPt.Enabled = True
        Combo_EndPt.Enabled = True
        AddButton.Enabled = True
    Else
        Combo_BeginPt.Enabled = False
        Combo_EndPt.Enabled = False
        AddButton.Enabled = False
    End If
End Sub

' Initialize the main dialog
Private Sub UserForm_Initialize()
    Dim tempPath As String
    Dim oSite As AeccSite
    Dim oAlignment As AeccAlignment
    Dim oAlignmentPIs As CAlignmentPIs
    Dim oSteam As CSTEAM
    Dim oSampleLineGroup As AeccSampleLineGroup
   
    g_sFileName = g_oCivilApp.Path

    tempPath = Environ("temp")
    g_reportFileName = tempPath & "\HEC-RAS.geo"


    ReportFileNameEdit.Text = g_reportFileName

    'initial Select report Components
    g_Alignments.removeAll
    g_Steams.removeAll
    
    ' siteless alignment
    For Each oAlignment In g_oAeccDoc.AlignmentsSiteless
        If oAlignment.SampleLineGroups.Count > 0 Then
            If g_Alignments.Exists(oAlignment.Name) = False Then
                Set oAlignmentPIs = New CAlignmentPIs
                oAlignmentPIs.Alignment = oAlignment
                g_Alignments.Add oAlignment.Name, oAlignmentPIs
                Combo_Align.AddItem oAlignment.Name
                For Each oSampleLineGroup In oAlignment.SampleLineGroups
                    Set oSteam = New CSTEAM
                    oSteam.SampleLineGroup = oSampleLineGroup
                    g_Steams.Add oSampleLineGroup.Name, oSteam
                Next
            End If
        End If
    Next
    
    'alignment from sites
    For Each oSite In g_oAeccDb.Sites
        For Each oAlignment In oSite.Alignments
            If oAlignment.SampleLineGroups.Count > 0 Then
                If g_Alignments.Exists(oAlignment.Name) = False Then
                    Set oAlignmentPIs = New CAlignmentPIs
                    oAlignmentPIs.Alignment = oAlignment
                    g_Alignments.Add oAlignment.Name, oAlignmentPIs
                    Combo_Align.AddItem oAlignment.Name
                    For Each oSampleLineGroup In oAlignment.SampleLineGroups
                        Set oSteam = New CSTEAM
                        oSteam.SampleLineGroup = oSampleLineGroup
                        g_Steams.Add oSampleLineGroup.Name, oSteam
                    Next
                End If
            End If
        Next
    Next
    
    If g_Alignments.Count > 0 Then
        Combo_Align.ListIndex = 0
    Else
        MsgBox "There is no Alignment with Sample Line Group."
        End
    End If

    Combo_BeginPt.Enabled = False
    Combo_EndPt.Enabled = False
    AddButton.Enabled = False
    DeleteButton.Visible = False
   
    
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
    Call ListView.ColumnHeaders.Add(2, , sCaptionArr(nBeginPtIndex), 140, lvwColumnLeft)
    Call ListView.ColumnHeaders.Add(3, , sCaptionArr(nEndPtIndex), 140, lvwColumnLeft)

End Sub


' Do the real work of gnerating the report here
Private Sub ExecuteButton_Click()
    Dim t, pbarInc, cCount As Integer
    Dim stat As Boolean
    Dim curListItem As ListItem
    Dim exeStr As String
    
    Dim oSteam As CSTEAM

    If ListView.ListItems.Count = 0 Then
        Exit Sub
    End If

    For t = 1 To ListView.ListItems.Count
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

    
    Set oSteam = g_Steams.item(Combo_LineGroup.Text)
    
    Call AppendReport(oSteam, g_reportFileName)
    
    'update progress bar
    ProgressBar.Value = 100

    ProgressBar.Visible = False
    progressBarLabel.Visible = False


    Call WriteReportFooter(g_reportFileName)

    exeStr = "notepad.exe " & g_reportFileName
    Call Shell(exeStr, 1)

End Sub

Private Sub CancelButton_Click()

    Unload Me

End Sub

Private Sub UserForm_Terminate()

    ' clean up created objects
    Set g_oCivilApp = Nothing
    Set g_oAeccDoc = Nothing
    Set g_oAeccDb = Nothing
    Set g_Steams = Nothing
    Set g_Alignments = Nothing
    Set g_oDataToReport = Nothing
End Sub

Private Sub FileSelectionButton_Click()

    Dim fileDialog As New CFileDialog
    
    With fileDialog

        .DefaultExt = "geo"

        .DialogTitle = "Select a File"

        .Filter = "GEO Files|*.GEO"

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
    sCaptionArr(nBeginPtIndex) = "Begin Point"
    sCaptionArr(nEndPtIndex) = "End Point"
End Sub





