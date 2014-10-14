' -----------------------------------------------------------------------------
' <copyright file="ReportForm_CrossSection.vb" company="Autodesk">
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
Imports System.IO
Imports System.Reflection ' For Missing.Value and BindingFlags
Imports System.Runtime.InteropServices ' For COMException
Imports System.Web.UI
Imports Microsoft.Office.Interop.Excel
Imports Autodesk.AECC.Interop.Land
Imports LocalizedRes = Report.My.Resources.LocalizedStrings
Imports Autodesk.AECC.Interop

Friend Class ReportForm_CrossSection

    Private ctlProgressBar As CtrlProgressBar
    Private WithEvents ctlCrossSectionListView As CtrlCrossSectionListView
    Private ctlSaveReport As CtrlSaveReportFile
    Private WithEvents ctlStartStation As CtrlStationTextBox
    Private WithEvents ctlEndStation As CtrlStationTextBox

    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()
        If CanOpen() Then
            ReportUtilities.RunModalDialog(New ReportForm_CrossSection)
        End If
    End Sub

    Private Shared Function CanOpen() As Boolean
        CanOpen = CtrlCrossSectionListView.CheckConditions()
    End Function

    Private Sub ReportForm_CrossSection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CheckBox_HTML.Checked = True

        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)

        ctlSaveReport = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)

        ctlStartStation = New CtrlStationTextBox
        ctlStartStation.Initialize(TextBox_StartStation)

        ctlEndStation = New CtrlStationTextBox
        ctlEndStation.Initialize(TextBox_EndStation)

        ctlCrossSectionListView = New CtrlCrossSectionListView
        ctlCrossSectionListView.Initialize(ListView_SLG, ctlStartStation, ctlEndStation)
    End Sub

    Private Sub BtnHelp_Click() Handles Button_Help.Click
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_EG_AND_LAYOUT_PROFILE_GRADES_REPORT)
    End Sub

    Private Sub ReportForm_HelpRequested(ByVal sender As System.Object, ByVal hlpevent As System.Windows.Forms.HelpEventArgs) Handles MyBase.HelpRequested
        BtnHelp_Click()
    End Sub

    Private Sub BtnDone_Click() Handles Button_Done.Click
        Close()
    End Sub

    Private Sub BtnExecute_Click() Handles Button_CreateReport.Click
        If ctlCrossSectionListView.CheckedCount = 0 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.CrossSection_Msg_SelectOneFirst)
            Exit Sub
        End If

        For Each item As ListViewItem In ctlCrossSectionListView.CheckedItems
            Dim sampleLineGroup As Land.AeccSampleLineGroup
            sampleLineGroup = ctlCrossSectionListView.SampleLineGroupeArr(item.Index)
            If Not sampleLineGroup Is Nothing Then
                Dim surfaceProfile As Land.AeccProfile
                surfaceProfile = Daylight_ExtractData.FindSurfaceProfile(sampleLineGroup.Parent)
                If surfaceProfile Is Nothing Then
                    Dim errorMessage As String
                    errorMessage = LocalizedRes.ReportForm_CrossSection_NoSurfaceProfile
                    Dim result As MsgBoxResult
                    result = MsgBox(String.Format(errorMessage, sampleLineGroup.Parent.Name, vbCrLf + vbCrLf), _
                                    CType(MsgBoxStyle.YesNo + MsgBoxStyle.Question, MsgBoxStyle), LocalizedRes.ReportUtilitie_Msg_Title)
                    If result <> MsgBoxResult.Yes Then
                        Exit Sub
                    End If
                End If
            End If
        Next item

        Try
            Dim tempFile As String = ReportConverter.GetRandomTempHtmlPath()

            ctlProgressBar.ProgressBegin(ctlCrossSectionListView.CheckedCount)

            WriteReportHTML(tempFile)

            'convert temp html report to target report
            ReportConverter.GenerateReportFileFromHTML(tempFile, _
                                                       ctlSaveReport.SavePath, _
                                                       ctlSaveReport.ReportFileType)
        Catch ex As Exception
            System.Diagnostics.Debug.Assert(False, ex.Message)
            Exit Sub
            ' met some errors, don't need to open
            ' maybe add some UI warning later
        End Try

        'update progress bar
        ctlProgressBar.ProgressEnd()

        ReportUtilities.OpenFileByDefaultBrowser(ctlSaveReport.SavePath)

    End Sub

    Private Sub WriteReportHTML(ByVal fileName As String)
        Dim oSampleLineGroup As Land.AeccSampleLineGroup
        Dim inStartStation As Double
        Dim inEndStation As Double
        ' number formatting in Slovenian language isn't right otherwise (decimal separator becomes "." instead of ",")
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Dim ci As System.Globalization.CultureInfo = New System.Globalization.CultureInfo(oldCI.Name)
        System.Threading.Thread.CurrentThread.CurrentCulture = ci

        Dim oReportHtml As New ReportWriter(fileName)

        '<html>
        oReportHtml.HtmlBegin()
        oReportHtml.RenderHeader(LocalizedRes.CrossSection_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.CrossSection_Html_Header, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        For Each item As ListViewItem In ctlCrossSectionListView.CheckedItems
            oSampleLineGroup = ctlCrossSectionListView.SampleLineGroupeArr(item.Index)
            If Not oSampleLineGroup Is Nothing Then
                inStartStation = ReportUtilities.GetRawStation(item.SubItems(CtrlCrossSectionListView.INDEX_STASTART).Text, oSampleLineGroup.Parent.StationIndexIncrement)
                inEndStation = ReportUtilities.GetRawStation(item.SubItems(CtrlCrossSectionListView.INDEX_STAEND).Text, oSampleLineGroup.Parent.StationIndexIncrement)
                AppendReportHTML(oSampleLineGroup, item.SubItems(CtrlCrossSectionListView.INDEX_CORRIDOR).Text, inStartStation, inEndStation, oReportHtml)
                ctlProgressBar.PerformStep()
            End If
        Next item

        '</body>
        oReportHtml.BodyEnd()

        '</html>
        oReportHtml.HtmlEnd()

        ' don't forget to switch back to original CI
        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI
    End Sub


    ' Append one item's report to html file
    Private Sub AppendReportHTML(ByVal oSampleLineGroup As Land.AeccSampleLineGroup, _
                                 ByVal sCorridor As String, _
                                 ByVal stationStart As Double, _
                                 ByVal stationEnd As Double, _
                                 ByVal oHtmlWriter As ReportWriter)

        If oSampleLineGroup Is Nothing Or oHtmlWriter Is Nothing Then
            Exit Sub
        End If

        '<hr>
        oHtmlWriter.RenderHr()

        ' alignmnent info
        Dim str As String

        ' render corridor name
        str = LocalizedRes.CrossSection_Html_CorridorName
        str &= " " & sCorridor
        oHtmlWriter.RenderLine(str)

        'render alignment description
        str = LocalizedRes.CrossSection_Html_AlignmentDescription
        str &= " " & oSampleLineGroup.Parent.Description
        oHtmlWriter.RenderLine(str)

        'render alignment name
        str = LocalizedRes.CrossSection_Html_AlignmentNameLabel
        str &= " " & oSampleLineGroup.Parent.Name
        oHtmlWriter.RenderLine(str)

        ' render sample line group name
        str = LocalizedRes.CrossSection_Html_SampleLineGroupName
        str &= " " & oSampleLineGroup.Name
        oHtmlWriter.RenderLine(str)

        '
        str = LocalizedRes.CrossSection_Html_Alignment_StaRange
        str += " " + LocalizedRes.CrossSection_Html_Alignment_StaStart
        str += " " + ReportUtilities.GetStationStringWithDerived(oSampleLineGroup.Parent, stationStart)
        str += LocalizedRes.CrossSection_Html_Alignment_StaEnd
        str += " " + ReportUtilities.GetStationStringWithDerived(oSampleLineGroup.Parent, stationEnd)
        oHtmlWriter.RenderLine(str)
        '<hr>
        oHtmlWriter.RenderHr()

        'extract data
        CrossSection_ExtractData.ExtractData(oSampleLineGroup, stationStart, stationEnd)

        oHtmlWriter.RenderBr()
        oHtmlWriter.TableBegin("1")

        'table title (head line)
        oHtmlWriter.TrBegin()
        oHtmlWriter.RenderTd(LocalizedRes.CrossSection_Html_TblTitle_Num, True)
        oHtmlWriter.RenderTd(LocalizedRes.CrossSection_Html_TblTitle_PK, True)
        oHtmlWriter.RenderTd(LocalizedRes.CrossSection_Html_TblTitle_Z_TN_LN1 & "<br>" & _
                             LocalizedRes.CrossSection_Html_TblTitle_Z_TN_LN2, True)
        oHtmlWriter.RenderTd(LocalizedRes.CrossSection_Html_TblTitle_Z_Proj_LN1 & "<br>" & _
                             LocalizedRes.CrossSection_Html_TblTitle_Z_Proj_LN2, True)
        oHtmlWriter.RenderTd(LocalizedRes.CrossSection_Html_TblTitle_X, True)
        oHtmlWriter.RenderTd(LocalizedRes.CrossSection_Html_TblTitle_Y, True)
        oHtmlWriter.RenderTd(LocalizedRes.CrossSection_Html_TblTitle_DevG, True)
        oHtmlWriter.RenderTd(LocalizedRes.CrossSection_Html_TblTitle_DevD, True)
        oHtmlWriter.TrEnd()


        For i As Integer = 0 To UBound(CrossSection_ExtractData.m_oSampleLineDataArr)
            ' If Not TypeName(CrossSection_ExtractData.m_oSampleLineDataArr(i)) = "Object()" Then
            If CrossSection_ExtractData.m_oSampleLineDataArr(i) Is Nothing Then
                Exit For
            End If
            'format string
            oHtmlWriter.TrBegin()
            oHtmlWriter.RenderTd(CrossSection_ExtractData.m_oSampleLineDataArr(i).Name)
            oHtmlWriter.RenderTd(CrossSection_ExtractData.m_oSampleLineDataArr(i).Station)
            oHtmlWriter.RenderTd(CrossSection_ExtractData.m_oSampleLineDataArr(i).ElevationEG)
            oHtmlWriter.RenderTd(CrossSection_ExtractData.m_oSampleLineDataArr(i).ElevationFG)
            oHtmlWriter.RenderTd(CrossSection_ExtractData.m_oSampleLineDataArr(i).East)
            oHtmlWriter.RenderTd(CrossSection_ExtractData.m_oSampleLineDataArr(i).North)
            oHtmlWriter.RenderTd(CrossSection_ExtractData.m_oSampleLineDataArr(i).CrossCrownLeft)
            oHtmlWriter.RenderTd(CrossSection_ExtractData.m_oSampleLineDataArr(i).CrossCrownRight)
            oHtmlWriter.TrEnd()
        Next i

        oHtmlWriter.TableEnd()
        oHtmlWriter.RenderBr()
    End Sub

End Class
