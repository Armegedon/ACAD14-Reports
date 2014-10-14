' -----------------------------------------------------------------------------
' <copyright file="ReportForm_Daylight.vb" company="Autodesk">
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

Friend Class ReportForm_Daylight

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
            ReportUtilities.RunModalDialog(New ReportForm_Daylight)
        End If
    End Sub

    Private Shared Function CanOpen() As Boolean
        CanOpen = CtrlCrossSectionListView.CheckConditions(True, True)
    End Function

    Private Sub ReportForm_Daylight_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CheckBox_HTML.Checked = True

        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)

        ctlSaveReport = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)

        ctlStartStation = New CtrlStationTextBox
        ctlStartStation.Initialize(TextBox_StartStation)

        ctlEndStation = New CtrlStationTextBox
        ctlEndStation.Initialize(TextBox_EndStation)

        ctlCrossSectionListView = New CtrlCrossSectionListView
        ctlCrossSectionListView.NeedSurfaceProfile = True
        ctlCrossSectionListView.Initialize(ListView_SLG, ctlStartStation, ctlEndStation)
    End Sub

    Private Sub BtnHelp_Click() Handles Button_Help.Click
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_DAYLIGHT_REPORT)
    End Sub

    Private Sub ReportForm_HelpRequested(ByVal sender As System.Object, ByVal hlpevent As System.Windows.Forms.HelpEventArgs) Handles MyBase.HelpRequested
        BtnHelp_Click()
    End Sub

    Private Sub BtnDone_Click() Handles Button_Done.Click
        Close()
    End Sub

    Private Sub BtnExecute_Click() Handles Button_CreateReport.Click
        If ctlCrossSectionListView.CheckedCount = 0 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.Daylight_Msg_SelectOneFirst)
            Exit Sub
        End If

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
        oReportHtml.RenderHeader(LocalizedRes.Daylight_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.Daylight_Html_Header, 1)
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

        Dim existingSurface As Land.AeccSurface
        existingSurface = Daylight_ExtractData.FindEGSurface(oSampleLineGroup.Parent)
        If existingSurface Is Nothing Then
            Exit Sub ' No static profile of the alignment.
        End If

        '<hr>
        oHtmlWriter.RenderHr()

        ' alignment info
        Dim str As String

        ' render corridor name
        str = LocalizedRes.Daylight_Html_CorridorName
        str &= " " & sCorridor
        oHtmlWriter.RenderLine(str)

        'render alignment description
        str = LocalizedRes.Daylight_Html_AlignmentDescription
        str &= " " & oSampleLineGroup.Parent.Description
        oHtmlWriter.RenderLine(str)

        'render alignment name
        str = LocalizedRes.Daylight_Html_AlignmentNameLabel
        str &= " " & oSampleLineGroup.Parent.Name
        oHtmlWriter.RenderLine(str)

        ' render sample line group name
        str = LocalizedRes.Daylight_Html_SampleLineGroupName
        str &= " " & oSampleLineGroup.Name
        oHtmlWriter.RenderLine(str)

        '
        str = LocalizedRes.Daylight_Html_Alignment_StaRange
        str += " " + LocalizedRes.Daylight_Html_Alignment_StaStart
        str += " " + ReportUtilities.GetStationStringWithDerived(oSampleLineGroup.Parent, stationStart)
        str += LocalizedRes.Daylight_Html_Alignment_StaEnd
        str += " " + ReportUtilities.GetStationStringWithDerived(oSampleLineGroup.Parent, stationEnd)
        oHtmlWriter.RenderLine(str)
        '<hr>
        oHtmlWriter.RenderHr()


        'extract data
        Daylight_ExtractData.ExtractData(oSampleLineGroup, existingSurface, stationStart, stationEnd)

        oHtmlWriter.RenderBr()
        oHtmlWriter.TableBegin("1")

        'table title (head line)
        oHtmlWriter.TrBegin()
        oHtmlWriter.RenderTd(LocalizedRes.Daylight_Html_TblTitle_Num, True, "center")
        oHtmlWriter.RenderTd(LocalizedRes.Daylight_Html_TblTitle_PK, True, "center")
        oHtmlWriter.RenderTd(LocalizedRes.Daylight_Html_TblTitle_EMPG_LN1 & "<br>" & _
                             LocalizedRes.Daylight_Html_TblTitle_EMPG_LN2, True, "center")
        oHtmlWriter.RenderTd(LocalizedRes.Daylight_Html_TblTitle_XG, True, "center")
        oHtmlWriter.RenderTd(LocalizedRes.Daylight_Html_TblTitle_YG, True, "center")
        oHtmlWriter.RenderTd(LocalizedRes.Daylight_Html_TblTitle_ZG, True, "center")
        oHtmlWriter.RenderTd(LocalizedRes.Daylight_Html_TblTitle_EMPD_LN1 & "<br>" & _
                             LocalizedRes.Daylight_Html_TblTitle_EMPD_LN2, True, "center")
        oHtmlWriter.RenderTd(LocalizedRes.Daylight_Html_TblTitle_XD, True, "center")
        oHtmlWriter.RenderTd(LocalizedRes.Daylight_Html_TblTitle_YD, True, "center")
        oHtmlWriter.RenderTd(LocalizedRes.Daylight_Html_TblTitle_ZD, True, "center")
        oHtmlWriter.TrEnd()


        For i As Integer = 0 To UBound(Daylight_ExtractData.m_oSampleLineDataArr)
            ' If Not TypeName(Daylight_ExtractData.m_oSampleLineDataArr(i)) = "Object()" Then
            If Daylight_ExtractData.m_oSampleLineDataArr(i) Is Nothing Then
                Exit For
            End If
            'format string
            oHtmlWriter.TrBegin()
            oHtmlWriter.RenderTd(Daylight_ExtractData.m_oSampleLineDataArr(i).Name, , "right")
            oHtmlWriter.RenderTd(Daylight_ExtractData.m_oSampleLineDataArr(i).Station, , "right")
            oHtmlWriter.RenderTd(Daylight_ExtractData.m_oSampleLineDataArr(i).LengthLeft, , "right")
            oHtmlWriter.RenderTd(Daylight_ExtractData.m_oSampleLineDataArr(i).EastLeft, , "right")
            oHtmlWriter.RenderTd(Daylight_ExtractData.m_oSampleLineDataArr(i).NorthLeft, , "right")
            oHtmlWriter.RenderTd(Daylight_ExtractData.m_oSampleLineDataArr(i).ElevLeft, , "right")
            oHtmlWriter.RenderTd(Daylight_ExtractData.m_oSampleLineDataArr(i).LengthRight, , "right")
            oHtmlWriter.RenderTd(Daylight_ExtractData.m_oSampleLineDataArr(i).EastRight, , "right")
            oHtmlWriter.RenderTd(Daylight_ExtractData.m_oSampleLineDataArr(i).NorthRight, , "right")
            oHtmlWriter.RenderTd(Daylight_ExtractData.m_oSampleLineDataArr(i).ElevRight, , "right")
            oHtmlWriter.TrEnd()
        Next i

        oHtmlWriter.TableEnd()
        oHtmlWriter.RenderBr()
    End Sub

End Class
