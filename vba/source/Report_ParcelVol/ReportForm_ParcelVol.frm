VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportForm_ParcelVol 
   Caption         =   "Create Reports - Parcel Volume Report"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9195
   OleObjectBlob   =   "ReportForm_ParcelVol.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReportForm_ParcelVol"
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
Const DefFilCorrection = "1"
Const DefCutCorrection = "1"
Const DefEleTolerance = "0.05"

Private g_parcelArr() As Variant
Private g_VolSurfaceArr() As Variant

Private g_sFileName As String
Private g_reportFileName As String
Private g_curParcel As AeccParcel
Private g_curVolSurface As AeccSurface

Private sCaptionArr() As String



Private Sub Combo_Surface_Change()
    Dim oSurface As AeccSurface
    Dim oTinVolSurface As AeccTinVolumeSurface
    Dim oGridVolSurface As AeccGridVolumeSurface
    
    Set oSurface = g_VolSurfaceArr(Combo_Surface.ListIndex)
    If oSurface.Type = aecckTinVolumeSurface Then
        Set g_curVolSurface = oSurface
        Set oTinVolSurface = oSurface
        Labe_Type.Caption = "Tin"
        Label_BaseSurface.Caption = oTinVolSurface.Statistics.BottomSurface.Name
        Label_CompSurface.Caption = oTinVolSurface.Statistics.TopSurface.Name
        TextBox_Ele.Enabled = False
    ElseIf oSurface.Type = aecckGridVolumeSurface Then
        Set g_curVolSurface = oSurface
        Set oGridVolSurface = oSurface
        Labe_Type.Caption = "Grid"
        Label_BaseSurface.Caption = oGridVolSurface.Statistics.BottomSurface.Name
        Label_CompSurface.Caption = oGridVolSurface.Statistics.TopSurface.Name
        TextBox_Ele.Enabled = True
    End If
          
End Sub

Private Sub Combo_Parcel_Change()
    Set g_curParcel = g_parcelArr(Combo_Parcel.ListIndex)
End Sub


Private Sub HelpButton_Click()
    ReportsHH_hWnd = ReportsHtmlHelp(0, getReportsHelpFile, HH_HELP_CONTEXT, IDH_AECC_REPORTS_CREATE_PARCEL_VOLUMES_REPORT)
End Sub

Private Sub SelParcelButton_Click()
    Dim sSet As AcadSelectionSet         'Define sset as a SelectionSet object
    Dim oParcel As AeccParcel
    
    'Set sset to a new selection set named SS1 (the name doesn't matter here)
    
    On Error Resume Next
    If Not IsNull(ThisDrawing.SelectionSets.Item("Parcel")) Then
        Set sSet = ThisDrawing.SelectionSets.Item("Parcel")
     sSet.Delete
    End If
  
    Set sSet = ThisDrawing.SelectionSets.Add("Parcel")
   
    
    Me.Hide  'Hide the current dialog
    
    sSet.SelectOnScreen       'Prompt user to select objects
    ' continue select on screen until select nothing or select a single point
    Do While Not sSet.count = 0
        If sSet.count = 1 Then
            If TypeOf sSet.Item(0) Is AeccParcel Then
                Exit Do
            End If
        End If
        MsgBox "Please select only one Parcel."
        sSet.Clear
        sSet.SelectOnScreen       'Prompt user to select objects
    Loop
    
    
    If sSet.count = 1 Then
        Dim iNum As Long
        Set oParcel = sSet.Item(0)
        Combo_Parcel.Text = oParcel.Parent.Name + " - " + oParcel.Name
    End If
         
    Me.Show ' restore the dialog
End Sub

Private Sub TextBox_Cut_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrHandle
    Dim correction As Double
    correction = val(TextBox_Cut.value)
    If correction >= 1# And correction <= 2# Then
        TextBox_Cut.value = str(correction)
        Exit Sub
    End If
ErrHandle:
    MsgBox "Please input the right number from 1.000 to 2.000"
    TextBox_Cut.value = DefCutCorrection
End Sub

Private Sub TextBox_Cut_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox_Ele_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrHandle
    Dim correction As Double
    correction = val(TextBox_Ele.value)
    If correction >= 0# And correction <= 1# Then
        TextBox_Ele.value = str(correction)
        Exit Sub
    End If
ErrHandle:
    MsgBox "Please input the right number from 0.000 to 1.000"
    TextBox_Ele.value = DefEleTolerance
End Sub

Private Sub TextBox_Fill_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrHandle
    Dim correction As Double
    correction = val(TextBox_Fill.value)
    If correction >= 1# And correction <= 2# Then
        TextBox_Fill.value = str(correction)
        Exit Sub
    End If
ErrHandle:
    MsgBox "Please input the right number from 1.000 to 2.000"
    TextBox_Fill.value = DefFilCorrection
End Sub

Private Sub TextBox_Fill_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Frame5_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Frame5.ActiveControl = TextBox_Fill Then
        TextBox_Fill_Exit Cancel
    ElseIf Frame5.ActiveControl = TextBox_Cut Then
        TextBox_Cut_Exit Cancel
    ElseIf Frame5.ActiveControl = TextBox_Ele Then
        TextBox_Ele_Exit Cancel
    End If
End Sub

' Initialize the main dialog
Private Sub UserForm_Initialize()

    Dim tempPath As String
    Dim oSite As AeccSite
    Dim oSurface As AeccSurface
    Dim oParcel As AeccParcel
    Dim nCurParcelIndex As Long
    Dim nCurVolSurfaceIndex As Long
    Dim nParcelCount As Long
    
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
  
    
    'parcels
    nParcelCount = 0
    nCurParcelIndex = 0
    If g_oAeccDb.Sites.count = 0 Then
       GoTo NoObject
    End If
    For Each oSite In g_oAeccDb.Sites
        If oSite.Parcels.count > 0 Then
            nParcelCount = oSite.Parcels.count + nParcelCount
            ReDim Preserve g_parcelArr(nParcelCount - 1)
            For Each oParcel In oSite.Parcels
                Set g_parcelArr(nCurParcelIndex) = oParcel
                Combo_Parcel.AddItem (oSite.Name + " - " + oParcel.Name)
                nCurParcelIndex = nCurParcelIndex + 1
            Next
        End If
    Next
    If nParcelCount > 0 Then
        Combo_Parcel.ListIndex = 0
    Else
        GoTo NoObject
    End If
    
    ' volume surfaces
    nCurVolSurfaceIndex = 0
    For Each oSurface In g_oAeccDb.Surfaces
        If oSurface.Type = aecckGridVolumeSurface Or oSurface.Type = aecckTinVolumeSurface Then
            ReDim Preserve g_VolSurfaceArr(nCurVolSurfaceIndex)
            Set g_VolSurfaceArr(nCurVolSurfaceIndex) = oSurface
            Combo_Surface.AddItem (oSurface.Name)
            nCurVolSurfaceIndex = nCurVolSurfaceIndex + 1
       End If
    Next
    If nCurVolSurfaceIndex > 0 Then
        Combo_Surface.ListIndex = 0
    Else
        GoTo NoObject
    End If
 
    TextBox_Fill.value = DefFilCorrection
    TextBox_Cut.value = DefCutCorrection
    TextBox_Ele.value = DefEleTolerance
   
    Exit Sub
NoObject:
        MsgBox "No Volume Surfaces or Parcels in " & ThisDrawing.Name & " exist."
        Err.Number = 1
        Exit Sub
End Sub


' Do the real work of gnerating the report here
Private Sub ExecuteButton_Click()
'    Dim t, pbarInc, cCount As Integer
'    Dim oAlignment As AeccAlignment
'    Dim curListItem As ListItem
    Dim exeStr As String
    Dim fillCorrection As Double
    Dim cutCorrection As Double
    Dim eleTolerance As Double
'
'

    If g_curParcel Is Nothing Or g_curVolSurface Is Nothing Then
        MsgBox "Please select a Parcel and a Volume Surface"
        Exit Sub
    End If
    
    progressBarLabel.Enabled = False
    progressBarLabel.Enabled = True

    progressBarLabel.Visible = True
    ProgressBar.Visible = True

    Call WriteReportHeader(g_reportFileName)
'
    'update progress bar
    ProgressBar.value = 10

    fillCorrection = val(TextBox_Fill.value)
    cutCorrection = val(TextBox_Cut.value)
    eleTolerance = val(TextBox_Ele.value)
    
    Call AppendReport(g_curParcel, g_curVolSurface, fillCorrection, cutCorrection, eleTolerance, g_reportFileName)
    'update progress bar
    ProgressBar.value = 100
'
    ProgressBar.Visible = False
    progressBarLabel.Visible = False
'
'
    Call WriteReportFooter(g_reportFileName)
'
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
    Set g_curParcel = Nothing
    Set g_curVolSurface = Nothing
  
    
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





Private Function GetDefaultInc() As Double
    Dim fInc As Double
    If GetUnitStr() = "Feet" Then
        fInc = 50#
    Else
        fInc = 25#
    End If
    GetDefaultInc = fInc
End Function
