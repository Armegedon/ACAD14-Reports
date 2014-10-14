VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportForm_ParcelMapCheck 
   Caption         =   "Create Reports -Map Check  Report"
   ClientHeight    =   11805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   OleObjectBlob   =   "ReportForm_ParcelMapCheck.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReportForm_ParcelMapCheck"
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

Private Type ParcelInfo
    Obj As AeccParcel
    CounterWiseClock As Boolean
    AcrossChord As Boolean
End Type

Private g_sFileName As String
Private g_reportFileName As String
Private g_pobXCoord As Double
Private g_pobYCoord As Double
Private g_isCounter As Boolean
Private g_EnableCheck As Boolean

Private g_bCheckParcle As Boolean

Private g_parcelArr() As ParcelInfo
Private g_figureArr() As AeccSurveyFigure

Private Const nIncludeIndex = 0
Private Const nNameIndex = 1
Private Const nNumberIndex = 2
Private Const nDescriptionIndex = 3
Private Const nAreaIndex = 4
Private Const nPerimeterIndex = 5
Private Const nLastIndex = nPerimeterIndex
Private sCaptionArr() As String

Private Sub Button_DeSelAll_Click()
    Dim lItem As ListItem
    If g_bCheckParcle = True Then
        For Each lItem In ListView_Parcels.ListItems
            lItem.Checked = False
        Next
    Else
        For Each lItem In ListView_Figures.ListItems
            lItem.Checked = False
        Next
    End If
End Sub

Private Sub Button_SelAll_Click()
    Dim lItem As ListItem
    If g_bCheckParcle = True Then
        For Each lItem In ListView_Parcels.ListItems
            lItem.Checked = True
        Next
    Else
        For Each lItem In ListView_Figures.ListItems
            lItem.Checked = True
        Next
    End If
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub CheckAcrossChord_Click()
    Dim index As Integer
    index = ListView_Parcels.SelectedItem.index
    If index > 0 Then
        g_parcelArr(index - 1).AcrossChord = CheckAcrossChord.Value
    End If
End Sub

Private Sub CheckCounterClock_Click()
    Dim index As Integer
    index = ListView_Parcels.SelectedItem.index
    If index > 0 Then
        g_parcelArr(index - 1).CounterWiseClock = CheckCounterClock.Value
    End If
    
End Sub

Private Sub CommandButton1_Click()
    Me.Hide

    Dim selectPnt As Variant
    Dim prompt As String
    
    'Get AeccSettingCoordinate's decision
    Dim nDecision As Integer
    nDecision = g_oAeccDb.Settings.DrawingSettings.AmbientSettings.CoordinateSettings.precision
    On Error Resume Next

    prompt = vbCrLf & "Select Point of beginning: "
    selectPnt = ThisDrawing.Utility.GetPoint(, prompt)
    If Err Then
    Else
        Dim dXCoord As Double
        Dim dYCoord As Double
        dXCoord = CDbl(selectPnt(0))
        dYCoord = CDbl(selectPnt(1))
        PobXCoord.Text = RoundVal(dXCoord, nDecision)
        PobYCoord.Text = RoundVal(dYCoord, nDecision)
    End If
    
    Me.Show

End Sub

Private Sub ExecuteButton_Click()
    Dim t, cCount As Integer
    Dim pbarInc As Single
    Dim oParcel As AeccParcel
    Dim curListItem As ListItem
    Dim exeStr As String
    

    If ListView_Parcels.ListItems.count = 0 Then
        Exit Sub
    End If

    For t = 1 To ListView_Parcels.ListItems.count
        If ListView_Parcels.ListItems(t).Checked = True Then
            cCount = cCount + 1
        End If
    Next
'
    If cCount = 0 Then
        MsgBox "Please select at least one Parcel or Figure."
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
    
    Call WriteReportHeader(g_reportFileName)
'
    For t = 1 To ListView_Parcels.ListItems.count
        Set curListItem = ListView_Parcels.ListItems(t)
        If curListItem.Checked = True Then
            Set oParcel = g_parcelArr(t - 1).Obj
            If Not oParcel Is Nothing Then
                Call AppendReport(oParcel, g_parcelArr(t - 1).CounterWiseClock, g_parcelArr(t - 1).AcrossChord, g_reportFileName)
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

Private Sub FileSelectionButton_Click()
    Dim fileDialog As New CFileDialog
    Dim a  As Double
    a = ListView_Parcels.top
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

Private Sub HelpButton_Click()
    ReportsHH_hWnd = ReportsHtmlHelp(0, getReportsHelpFile, HH_HELP_CONTEXT, IDH_AECC_REPORTS_CREATE_PARCEL_MAPCHECK_REPORT)
End Sub







'Private Sub ListView_Parcels_AfterUpdate()
'    Dim selectedCount As Integer, i As Integer
'    selectedCount = 0
'    For i = 0 To ListView_Parcels.ListItems.count - 1 Step 1
'        If ListView_Parcels.ListItems(i).Selected = True Then
'            selectedCount = selectedCount + 1
'        End If
'    Next
'    If Not selectedCount = 1 Then
'        CheckCounterClock.Value = False
'        CheckCounterClock.Enabled = False
'        CheckAcrossChord.Value = False
'        CheckAcrossChord.Enabled = False
'        PobXCoord.Value = ""
'        PobYCoord.Value = ""
'    Else
'        Dim index As Integer
'        index = ListView_Parcels.SelectedItem.index
'        CheckCounterClock.Value = g_parcelArr(index).CounterWiseClock
'        CheckCounterClock.Enabled = True
'        CheckAcrossChord.Value = g_parcelArr(index).AcrossChord
'        CheckAcrossChord.Enabled = True
'    End If
'End Sub

Private Sub ListView_Parcels_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim selectedCount As Integer, i As Integer
    selectedCount = 0
    For i = 1 To ListView_Parcels.ListItems.count Step 1
        If ListView_Parcels.ListItems(i).Selected = True Then
            selectedCount = selectedCount + 1
        End If
    Next
    If Not selectedCount = 1 Then
        CheckCounterClock.Value = False
        CheckCounterClock.Enabled = False
        CheckAcrossChord.Value = False
        CheckAcrossChord.Enabled = False
        PobXCoord.Value = ""
        PobYCoord.Value = ""
    Else
        Dim index As Integer
        index = ListView_Parcels.SelectedItem.index
        CheckCounterClock.Value = g_parcelArr(index - 1).CounterWiseClock
        CheckCounterClock.Enabled = True
        CheckAcrossChord.Value = g_parcelArr(index - 1).AcrossChord
        CheckAcrossChord.Enabled = True
        PobXCoord.Value = g_parcelArr(index - 1).Obj.ParcelLoops(0).Item(0).StartX
        PobYCoord.Value = g_parcelArr(index - 1).Obj.ParcelLoops(0).Item(0).StartY
    End If
End Sub

'Private Sub OptionButton_Figure_Change()
'    ListView_Parcels.Visible = False
'    ListView_Figures.Visible = True
''    ListView_Figures.top = ListView_Parcels.top
''    ListView_Figures.height = ListView_Parcels.height
''    ListView_Figures.left = ListView_Parcels.left
''    ListView_Figures.width = ListView_Parcels.width
''    Dim left As Single
''    left = Me.left + 0.4
''    Me.Move left
'End Sub

Private Sub OptionButton_Figure_Click()
    ListView_Figures.top = ListView_Parcels.top
    ListView_Figures.height = ListView_Parcels.height
    ListView_Figures.left = ListView_Parcels.left
    ListView_Figures.width = ListView_Parcels.width
    
'    ListView_Parcels.Enabled = False
    ListView_Parcels.Visible = False
'    ListView_Figures.Enabled = True
    ListView_Figures.Visible = True
    g_bCheckParcle = False
   
End Sub

'
'End Sub

Private Sub OptionButton_Parcel_Click()
 '   ListView_Figures.Enabled = False
    ListView_Figures.Visible = False
  '  ListView_Parcels.Enabled = True
    ListView_Parcels.Visible = True
   
  

    g_bCheckParcle = True
    
End Sub

Private Sub UserForm_Initialize()
    Dim tempPath As String

    
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
    
    PobXCoord.Text = "0.000"
    PobYCoord.Text = "0.000"
    
    
    OptionButton_Parcel.Value = True
    'OptionButton_Figure.Value = True
    
    

        ' Get Captions
    GetCaptions

    ' list view grid setup
    ' setup columns
    Call ListView_Parcels.ColumnHeaders.Add(, , sCaptionArr(nIncludeIndex), 40, lvwColumnLeft)
    Call ListView_Parcels.ColumnHeaders.Add(, , sCaptionArr(nNameIndex), 100, lvwColumnRight)
    Call ListView_Parcels.ColumnHeaders.Add(, , sCaptionArr(nNumberIndex), 50, lvwColumnRight)
    Call ListView_Parcels.ColumnHeaders.Add(, , sCaptionArr(nDescriptionIndex), 80, lvwColumnRight)
    Call ListView_Parcels.ColumnHeaders.Add(, , sCaptionArr(nAreaIndex), 80, lvwColumnRight)
    Call ListView_Parcels.ColumnHeaders.Add(, , sCaptionArr(nPerimeterIndex), 80, lvwColumnRight)
    
    Call ListView_Figures.ColumnHeaders.Add(, , sCaptionArr(nIncludeIndex), 40, lvwColumnLeft)
    Call ListView_Figures.ColumnHeaders.Add(, , sCaptionArr(nNameIndex), 100, lvwColumnRight)
    Call ListView_Figures.ColumnHeaders.Add(, , sCaptionArr(nNumberIndex), 50, lvwColumnRight)
    Call ListView_Figures.ColumnHeaders.Add(, , sCaptionArr(nDescriptionIndex), 80, lvwColumnRight)
    Call ListView_Figures.ColumnHeaders.Add(, , sCaptionArr(nAreaIndex), 80, lvwColumnRight)
    Call ListView_Figures.ColumnHeaders.Add(, , sCaptionArr(nPerimeterIndex), 80, lvwColumnRight)
    
    FillSettings
    
    FillParcleGrid
    
'    If GetBaseSurveyObjects = True Then
'        FillFigureGrid
'    Else
'        OptionButton_Figure.Enabled = False
'    End If
    If ListView_Parcels.ListItems.count = 0 Then
        MsgBox "No parcels in " & ThisDrawing.Name & " exist."
        Err.Number = 1
    End If
    
End Sub

Private Sub GetCaptions()
    ReDim sCaptionArr(nLastIndex)
    sCaptionArr(nIncludeIndex) = "Include"
    sCaptionArr(nNameIndex) = "Name"
    sCaptionArr(nNumberIndex) = "Number"
    sCaptionArr(nDescriptionIndex) = "Description"
    sCaptionArr(nAreaIndex) = "Area"
    sCaptionArr(nPerimeterIndex) = "Perimeter"
End Sub


Private Sub FillParcleGrid()
    
    'Fill Parcel List View
    Dim oParcel As AeccParcel
    Dim oSite As AeccSite

    Dim li As ListItem
    Dim nCount As Long
    Dim nCurIndex As Long

    'get alignments from Site
    For Each oSite In g_oAeccDb.Sites
        nCount = nCount + oSite.Parcels.count
        ReDim Preserve g_parcelArr(nCount)
        For Each oParcel In oSite.Parcels
             Set li = ListView_Parcels.ListItems.Add
             li.SubItems(nNameIndex) = oParcel.Parent.Name + "-" + oParcel.Name
             li.SubItems(nNumberIndex) = oParcel.Number
             li.SubItems(nDescriptionIndex) = oParcel.Description
             li.SubItems(nAreaIndex) = FormatAreaSettings_Parcel(oParcel.Statistics.Area)
             li.SubItems(nPerimeterIndex) = FormatDistSettings_Parcel(oParcel.Statistics.Perimeter)
             li.Checked = True
             Set g_parcelArr(nCurIndex).Obj = oParcel
             g_parcelArr(nCurIndex).AcrossChord = False
             g_parcelArr(nCurIndex).CounterWiseClock = False
             nCurIndex = nCurIndex + 1
        Next
    Next

End Sub


Private Sub FillFigureGrid()
    Dim oFigure As AeccSurveyFigure
    Dim oProject As AeccSurveyProject
    
    Dim li As ListItem
    Dim nCount As Long
    Dim nCurIndex As Long

    'get alignments from Site
    For Each oProject In g_oSurveyDocument.Projects
        nCount = nCount + oProject.Figures.count
        ReDim Preserve g_figureArr(nCount)
        For Each oFigure In oProject.Figures
             Set li = ListView_Figures.ListItems.Add
             li.SubItems(nNameIndex) = oFigure.Name
             li.SubItems(nNumberIndex) = oFigure.Id
             li.SubItems(nDescriptionIndex) = ""
             li.SubItems(nAreaIndex) = FormatAreaSettings_Figure(0#)
             li.SubItems(nPerimeterIndex) = FormatDistSettings_Figure(0#)
             li.Checked = True
             Set g_figureArr(nCurIndex) = oFigure
             nCurIndex = nCurIndex + 1
        Next
    Next
    
End Sub

Private Sub FillSettings()
    Dim ambSettings As AeccSettingsAmbient
    Set ambSettings = g_oAeccDb.Settings.ParcelSettings.AmbientSettings

    Dim distFormat As String, dirFormat As String, coordFormat As String, areaFormat As String
    Select Case ambSettings.DirectionSettings.Format
        Case aeccFormatDecimal
            dirFormat = "(d)"
        Case aeccFormatDecimalDegreeMinuteSecond
            dirFormat = "DD.MMSSSS"
        Case aeccFormatDegreeMinuteSecond
            dirFormat = "DD" & Chr(176) & "MM" & Chr(176) & "SS"
        Case aeccFormatDegreeMinuteSecondSpaced
            dirFormat = "DD" & Chr(176) & " MM" & Chr(176) & " SS"
    End Select
    Label_DirecFormat.Caption = dirFormat
    Label_DistPrec.Caption = ambSettings.DistanceSettings.precision
    Label_DirePrec.Caption = ambSettings.DirectionSettings.precision
    Label_AreaPrec.Caption = ambSettings.areaSettings.precision
    Label_CloPrec.Caption = ambSettings.CoordinateSettings.precision
    Label_PrecPrec.Caption = ambSettings.DistanceSettings.precision
    
End Sub
