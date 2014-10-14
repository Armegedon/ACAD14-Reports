' -----------------------------------------------------------------------------
' <copyright file="ReportForm_ParcelVol.vb" company="Autodesk">
' Copyright (C) Autodesk, Inc. All rights reserved.
'
' Permission to use, copy, modify, and distribute this software in
' object code form for any purpose and without fee is hereby granted,
' provided that the above copyright notice appears in all copies and
' that both that copyright notice and the limited warranty and
' restricted rights notice below appear in all supporting
' documentation.
'
' AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
' AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
' MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
' DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
' UNINTERRUPTED OR ERROR FREE.
'
' Use, duplication, or disclosure by the U.S. Government is subject to
' restrictions set forth in FAR 52.227-19 (Commercial Computer
' Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
' (Rights in Technical Data and Computer Software), as applicable.
'
' </copyright>
' -----------------------------------------------------------------------------

Option Explicit On
Option Strict On

Imports System
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports AecUIBase = Autodesk.AEC.Interop.UIBase
Imports AeccSurvey = Autodesk.AECC.Interop.Survey
Imports Autodesk.AutoCAD.EditorInput
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Friend Class ReportForm_ParcelVol

    Private m_Parcels As Collections.Generic.SortedList(Of String, AeccLandLib.AeccParcel)
    Private m_VolSurfaces As Collections.Generic.SortedList(Of String, AeccLandLib.AeccSurface)

    Private Const STR_DEF_FILL As String = "1"
    Private Const STR_DEF_CUT As String = "1"
    Private Const STR_DEF_ELE As Double = 0.05

    Private ctlProgressBar As CtrlProgressBar
    Private ctlSavePath As CtrlSaveReportFile

    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()
        Dim reportForm As New ReportForm_ParcelVol
        If reportForm.ReadyToOpen() = True Then
            ReportUtilities.RunModalDialog(reportForm)
        End If
    End Sub

    Private Function ReadyToOpen() As Boolean

        ReadyToOpen = True

        If ReportApplication.AeccXDatabase.Sites.Count = 0 Then
            ReadyToOpen = False
        End If

        Dim nParcelCount As Long
        nParcelCount = 0
        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            If oSite.Parcels.Count > 0 Then
                nParcelCount = oSite.Parcels.Count + nParcelCount
            End If
        Next

        If nParcelCount < 1 Then
            ReadyToOpen = False
        End If

        'Init Combo volume surfaces
        Dim nCurVolSurfaceIndex As Long
        nCurVolSurfaceIndex = 0
        For Each oSurface As AeccLandLib.AeccSurface In ReportApplication.AeccXDatabase.Surfaces
            If oSurface.Type = AeccLandLib.AeccSurfaceType.aecckGridVolumeSurface _
                Or oSurface.Type = AeccLandLib.AeccSurfaceType.aecckTinVolumeSurface Then
                nCurVolSurfaceIndex = nCurVolSurfaceIndex + 1
            End If
        Next
        If nCurVolSurfaceIndex < 1 Then
            ReadyToOpen = False
        End If

        If ReadyToOpen = False Then
            Dim sMsg As String = String.Format(LocalizedRes.Msg_No_VolSurface_Parcel_In_Drawing, ReportApplication.AeccXDocument.Name)
            ReportUtilities.ACADMsgBox(sMsg)
        End If

    End Function

    Private Sub ReportForm_ParcelVol_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)

        ctlSavePath = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)

        InitParcelListView()

        m_Parcels = New Collections.Generic.SortedList(Of String, AeccLandLib.AeccParcel)
        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            If oSite.Parcels.Count > 0 Then
                For Each oParcel As AeccLandLib.AeccParcel In oSite.Parcels
                    m_Parcels.Add(oSite.Name + " - " + oParcel.Name, oParcel)
                Next
            End If
        Next

        'Init Combo volume surfaces
        m_VolSurfaces = New Collections.Generic.SortedList(Of String, AeccLandLib.AeccSurface)
        For Each oSurface As AeccLandLib.AeccSurface In ReportApplication.AeccXDatabase.Surfaces
            If oSurface.Type = AeccLandLib.AeccSurfaceType.aecckGridVolumeSurface _
                Or oSurface.Type = AeccLandLib.AeccSurfaceType.aecckTinVolumeSurface Then
                m_VolSurfaces.Add(oSurface.Name, oSurface)
            End If
        Next

        For Each surfaceName As String In m_VolSurfaces.Keys
            Combo_Surface.Items.Add(surfaceName)
        Next

        If m_VolSurfaces.Count > 0 Then
            Combo_Surface.SelectedIndex = 0
        End If

        'Init Volume corrections
        TextBox_Fill.Text = STR_DEF_FILL
        TextBox_Cut.Text = STR_DEF_CUT
        TextBox_Ele.Text = STR_DEF_ELE.ToString
    End Sub

    Private Sub Combo_Surface_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Surface.SelectedIndexChanged

        If Combo_Surface.SelectedIndex < 0 Then
            Exit Sub
        End If

        Dim oSurface As AeccLandLib.AeccSurface

        oSurface = m_VolSurfaces.Values.Item(Combo_Surface.SelectedIndex)
        If oSurface.Type = AeccLandLib.AeccSurfaceType.aecckTinVolumeSurface Then
            Dim oTinVolSurface As AeccLandLib.AeccTinVolumeSurface
            'g_curVolSurface = oSurface
            oTinVolSurface = CType(oSurface, AeccLandLib.AeccTinVolumeSurface)
            Label_TypeValue.Text = LocalizedRes.ParcelVol_Form_Tin '"Tin"
            Label_BaseSurfaceValue.Text = oTinVolSurface.Statistics.BottomSurface.Name
            Label_CompSurfaceValue.Text = oTinVolSurface.Statistics.TopSurface.Name
            TextBox_Ele.Enabled = False
        ElseIf oSurface.Type = AeccLandLib.AeccSurfaceType.aecckGridVolumeSurface Then
            Dim oGridVolSurface As AeccLandLib.AeccGridVolumeSurface
            'g_curVolSurface = oSurface
            oGridVolSurface = CType(oSurface, AeccLandLib.AeccGridVolumeSurface)
            Label_TypeValue.Text = LocalizedRes.ParcelVol_Form_Grid '"Grid"
            Label_BaseSurfaceValue.Text = oGridVolSurface.Statistics.BottomSurface.Name
            Label_CompSurfaceValue.Text = oGridVolSurface.Statistics.TopSurface.Name
            TextBox_Ele.Enabled = True
        End If
    End Sub

    Private Sub Button_Done_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Done.Click
        Close()
    End Sub

    Private Sub Button_Help_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Help.Click
        openHelp()
    End Sub

    Private Sub ReportForm_HelpRequested(ByVal sender As System.Object, ByVal hlpevent As System.Windows.Forms.HelpEventArgs) Handles MyBase.HelpRequested
        openHelp()
    End Sub

    Private Sub TextBox_Fill_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Fill.Leave
        Try
            Dim correction As Double
            correction = CDbl(TextBox_Fill.Text)
            If correction >= 1.0 And correction <= 2.0 Then
                TextBox_Fill.Text = correction.ToString("N")
            Else
                Throw New ArgumentOutOfRangeException()
            End If
        Catch ex As Exception
            ReportUtilities.ACADMsgBox(LocalizedRes.ParcelVol_Msg_FillCutInput)
            TextBox_Fill.Text = STR_DEF_FILL
        End Try
    End Sub

    Private Sub TextBox_Cut_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Cut.Leave
        Try
            Dim correction As Double
            correction = CDbl(TextBox_Cut.Text)
            If correction >= 1.0 And correction <= 2.0 Then
                TextBox_Cut.Text = correction.ToString("N")
            Else
                Throw New ArgumentOutOfRangeException()
            End If
        Catch ex As Exception
            ReportUtilities.ACADMsgBox(LocalizedRes.ParcelVol_Msg_FillCutInput)
            TextBox_Cut.Text = STR_DEF_CUT
        End Try
    End Sub

    Private Sub TextBox_Ele_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_Ele.Leave
        Try
            Dim correction As Double
            correction = CDbl(TextBox_Ele.Text)
            If correction >= 0.0 And correction <= 1.0 Then
                TextBox_Ele.Text = correction.ToString("N")
            Else
                Throw New ArgumentOutOfRangeException()
            End If
        Catch ex As Exception
            ReportUtilities.ACADMsgBox(LocalizedRes.ParcelVol_Msg_EleInput)
            TextBox_Ele.Text = STR_DEF_ELE.ToString
        End Try
    End Sub

    Private Sub TextBox_Fill_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_Fill.KeyPress
        If IsInputValidKey(e.KeyChar, TextBox_Fill.Text) = False Then
            Beep()
            e.Handled = True
        End If
    End Sub

    Private Function IsInputValidKey(ByVal KeyChar As Char, ByVal inputed As String) As Boolean
        If KeyChar <> ChrW(Keys.Back) And _
                (KeyChar < ChrW(Keys.D0) Or KeyChar > ChrW(Keys.D9)) And _
                KeyChar <> ChrW(Keys.Delete) Then

            IsInputValidKey = False
        Else
            IsInputValidKey = True
        End If

        If KeyChar = ChrW(Keys.Delete) Then
            If inputed.Contains(ChrW(Keys.Delete).ToString()) Then
                IsInputValidKey = False
            End If
        End If
    End Function


    Private Sub TextBox_Cut_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_Cut.KeyPress
        If IsInputValidKey(e.KeyChar, TextBox_Cut.Text) = False Then
            Beep()
            e.Handled = True
        End If
    End Sub


    Private Sub TextBox_Ele_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox_Ele.KeyPress
        If IsInputValidKey(e.KeyChar, TextBox_Ele.Text) = False Then
            Beep()
            e.Handled = True
        End If
    End Sub

    Private Sub Button_CreateReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_CreateReport.Click
        'Get Current select parcel and surface
        If ListView_Parcels.Items.Count = 0 Or Combo_Surface.SelectedIndex < 0 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.ParcelVol_Msg_SelectParcelSurface)
            Exit Sub
        End If

        'reset progress bar
        ctlProgressBar.ProgressBegin(ListView_Parcels.CheckedItems.Count * 4)

        Try
            'write report
            Dim tempFile As String = ReportConverter.GetRandomTempHtmlPath()
            WriteHtmlReport(tempFile)

            'convert temp html report to target report
            ReportConverter.GenerateReportFileFromHTML(tempFile, _
                                                       ctlSavePath.SavePath, _
                                                       ctlSavePath.ReportFileType)
        Catch ex As Exception
            System.Diagnostics.Debug.Assert(False, ex.Message)
            Exit Sub
            ' met some errors, don't need to open
            ' maybe add some UI warning later
        End Try

        'update progress bar
        ctlProgressBar.ProgressEnd()

        ReportUtilities.OpenFileByDefaultBrowser(ctlSavePath.SavePath)
    End Sub

    Private Sub WriteHtmlReport(ByVal fileName As String)
        Dim oReportHtml As New ReportWriter(fileName)

        '<html>
        oReportHtml.HtmlBegin()
        oReportHtml.RenderHeader(LocalizedRes.ParcelVol_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.ParcelVol_Html_Title, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        Dim fillCorrection As Double
        Dim cutCorrection As Double
        Dim eleTolerance As Double

        fillCorrection = Double.Parse(TextBox_Fill.Text)
        cutCorrection = Double.Parse(TextBox_Cut.Text)
        eleTolerance = Double.Parse(TextBox_Ele.Text)

        Dim curVolSurface As AeccLandLib.AeccSurface
        curVolSurface = m_VolSurfaces.Values.Item(Combo_Surface.SelectedIndex)

        'step progress bar : 1/5
        ctlProgressBar.PerformStep()

        For Each item As ListViewItem In ListView_Parcels.CheckedItems
            Dim oParcel As AeccLandLib.AeccParcel = CType(item.Tag, AeccLandLib.AeccParcel)
            If Not oParcel Is Nothing Then
                AppendReport(oParcel, curVolSurface, fillCorrection, cutCorrection, eleTolerance, oReportHtml)
                ctlProgressBar.PerformStep()
            End If
        Next item

        '</body>
        oReportHtml.BodyEnd()

        '</html>
        oReportHtml.HtmlEnd()

        'step progress bar : 5/5
        ctlProgressBar.PerformStep()
    End Sub

    ' Append one item's report to html file
    Private Sub AppendReport(ByVal oParcel As AeccLandLib.AeccParcel, _
                             ByVal oVolumeSurface As AeccLandLib.AeccSurface, _
                             ByVal filCor As Double, _
                             ByVal cutCor As Double, _
                             ByVal eleTol As Double, _
                             ByVal oHtmlWriter As ReportWriter)

        'extract data
        ParcelVol_ExtractData.ExtractData(oParcel, oVolumeSurface, filCor, cutCor)

        'step progress bar : 2/5
        ctlProgressBar.PerformStep()

        '<hr>
        oHtmlWriter.RenderHr()

        ' Parcel info
        Dim str As String
        str = LocalizedRes.ParcelMapCheck_Html_ParcelName
        str += " " + oParcel.Parent.Name + " - " + oParcel.Name
        oHtmlWriter.RenderLine(str)

        str = String.Format(LocalizedRes.Common_Html_Description_One_Param, oParcel.Description)
        oHtmlWriter.RenderLine(str)

        str = LocalizedRes.ParcelVol_Html_VolSurface
        str += " " + oVolumeSurface.Name
        oHtmlWriter.RenderLine(str)

        str = LocalizedRes.ParcelVol_Html_Fill
        str += " " + filCor.ToString("N")
        oHtmlWriter.RenderLine(str)

        str = LocalizedRes.ParcelVol_Html_Cut
        str += " " + cutCor.ToString("N")
        oHtmlWriter.RenderLine(str)

        str = LocalizedRes.ParcelVol_Html_Ele
        str += " " + eleTol.ToString("N")
        oHtmlWriter.RenderLine(str)

        oHtmlWriter.RenderBr()
        oHtmlWriter.TableBegin("1")

        oHtmlWriter.TrBegin()
        oHtmlWriter.RenderTd(LocalizedRes.ParcelVol_Html_TblTitle_Fill, True)
        oHtmlWriter.RenderTd(LocalizedRes.ParcelVol_Html_TblTitle_Cut, True)
        oHtmlWriter.RenderTd(LocalizedRes.ParcelVol_Html_TblTitle_Net, True)

        oHtmlWriter.TrEnd()

        'step progress bar : 3/5
        ctlProgressBar.PerformStep()

        oHtmlWriter.TrBegin()
        oHtmlWriter.RenderTd(ParcelVol_ExtractData.ParcelVolumeData.Fill)
        oHtmlWriter.RenderTd(ParcelVol_ExtractData.ParcelVolumeData.Cut)
        oHtmlWriter.RenderTd(ParcelVol_ExtractData.ParcelVolumeData.Net)
        oHtmlWriter.TrEnd()

        oHtmlWriter.TableEnd()
        oHtmlWriter.RenderBr()

        'step progress bar : 4/5
        ctlProgressBar.PerformStep()
    End Sub

    Private Sub openHelp()
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_PARCEL_VOLUMES_REPORT)
    End Sub

    Private Sub InitParcelListView()
        ' setup columns
        ListView_Parcels.View = View.Details
        Call ListView_Parcels.Columns.Add(LocalizedRes.ParceMapCheck_Form_ColumnTitle_Include, _
            50, HorizontalAlignment.Left)
        Call ListView_Parcels.Columns.Add(LocalizedRes.ParceMapCheck_Form_ColumnTitle_Name, _
            150, HorizontalAlignment.Left)
        Call ListView_Parcels.Columns.Add(LocalizedRes.ParceMapCheck_Form_ColumnTitle_Number, _
            50, HorizontalAlignment.Left)
        Call ListView_Parcels.Columns.Add(LocalizedRes.ParceMapCheck_Form_ColumnTitle_Description, _
            80, HorizontalAlignment.Left)

        'Fill parcel grid. Set list view item's tag as parcel COM object.
        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            For Each oParcel As AeccLandLib.AeccParcel In oSite.Parcels
                Dim li As New ListViewItem
                li.SubItems.Add(oParcel.Parent.Name + "-" + oParcel.Name)
                li.SubItems.Add(oParcel.Number.ToString())
                li.SubItems.Add(oParcel.Description)
                li.Checked = True
                li.Tag = oParcel
                ListView_Parcels.Items.Add(li)
            Next
        Next
    End Sub
End Class
