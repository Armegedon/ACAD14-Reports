' -----------------------------------------------------------------------------
' <copyright file="ReportForm_AlignStaInc.vb" company="Autodesk">
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
Imports AeccUiLandLib = Autodesk.AECC.Interop.UiLand
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Friend Class ReportForm_AlignStaInc

    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()
        If CtrlAlignmentListView.GetAlignmentCount() > 0 Then
            ReportUtilities.RunModalDialog(New ReportForm_AlignStaInc)
        End If
    End Sub

    Private Sub ReportForm_AlignStaInc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ctlStartStation = New CtrlStationTextBox
        ctlStartStation.Initialize(TextBox_StartStation)
        ctlEndStation = New CtrlStationTextBox
        ctlEndStation.Initialize(TextBox_EndStation)
        ctlAlignmentListView = New CtrlAlignmentListView
        ctlAlignmentListView.Initialize(ListView_Alignments, ctlStartStation, ctlEndStation)
        ctlSavePath = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)
        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)
        NumericUpDown_StationInc.Value = GetDefaultIncrement()
        NumericUpDown_StationInc.Maximum = Decimal.MaxValue
        NumericUpDown_StationInc.Minimum = 0.01D
    End Sub

    Private Sub Button_CreateReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_CreateReport.Click
        If ctlAlignmentListView.WinListView.CheckedItems.Count = 0 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.Alignment_Msg_SelectOneFirst)
            Exit Sub
        End If

        'reset progress bar
        ctlProgressBar.ProgressBegin(ctlAlignmentListView.WinListView.CheckedItems.Count)

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

        ReportUtilities.OpenFileByDefaultBrowser(ctlSavePath.SavePath) 'm_sReportFileName)
    End Sub

    Private Sub Button_Help_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Help.Click
        openHelp()
    End Sub

    Private Sub ReportForm_HelpRequested(ByVal sender As System.Object, ByVal hlpevent As System.Windows.Forms.HelpEventArgs) Handles MyBase.HelpRequested
        openHelp()
    End Sub

    Private Sub Button_Done_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Done.Click
        Close()
    End Sub

    Private Sub WriteHtmlReport(ByVal fileName As String)
        Dim oReportHtml As New ReportWriter(fileName)

        '<html>
        oReportHtml.HtmlBegin()
        oReportHtml.RenderHeader(LocalizedRes.AlignStaInc_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.AlignStaInc_Html_Title, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        For Each item As ListViewItem In ctlAlignmentListView.WinListView.CheckedItems
            Dim startRawStation, endRawStation As Double
            startRawStation = ctlAlignmentListView.StartRawStation(item)
            endRawStation = ctlAlignmentListView.EndRawStation(item)
            Try
                AppendReport(CType(item.Tag, AeccLandLib.AeccAlignment), startRawStation, endRawStation, oReportHtml)
                ctlProgressBar.PerformStep()
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
            End Try
        Next item

        '</body>
        oReportHtml.BodyEnd()

        '</html>
        oReportHtml.HtmlEnd()
    End Sub

    ' Append one item's report to html file
    Private Sub AppendReport(ByVal oAlignment As AeccLandLib.AeccAlignment, _
        ByVal stationStart As Double, ByVal stationEnd As Double, _
        ByVal oHtmlWriter As ReportWriter)

        If oAlignment Is Nothing Or oHtmlWriter Is Nothing Then
            Exit Sub
        End If

        '<hr>
        oHtmlWriter.RenderHr()

        ' alignment info
        Dim str As String

        'alignment name
        str = LocalizedRes.AlignSta_Html_Align_Name
        str += " " + oAlignment.Name
        oHtmlWriter.RenderLine(str)

        'alignment description
        str = String.Format(LocalizedRes.Common_Html_Description_One_Param, oAlignment.Description)
        oHtmlWriter.RenderLine(str)

        'start station, end station
        str = String.Format(LocalizedRes.Alignment_Html_Station_Range, ReportUtilities.GetStationStringWithDerived(oAlignment, stationStart), _
              ReportUtilities.GetStationStringWithDerived(oAlignment, stationEnd))
        oHtmlWriter.RenderLine(str)

        'station incremnt
        str = LocalizedRes.AlignStaInc_Html_Increment
        str += " " + NumericUpDown_StationInc.Value.ToString("N")
        oHtmlWriter.RenderLine(str)

        'begin to render alignment data table
        oHtmlWriter.RenderBr()
        oHtmlWriter.TableBegin("1")

        'title
        oHtmlWriter.TrBegin()
        oHtmlWriter.RenderTd(LocalizedRes.AlignStaInc_Html_TblTitle_Sta, True)
        oHtmlWriter.RenderTd(LocalizedRes.AlignStaInc_Html_TblTitle_Northing, True)
        oHtmlWriter.RenderTd(LocalizedRes.AlignStaInc_Html_TblTitle_Easting, True)
        oHtmlWriter.RenderTd(LocalizedRes.AlignStaInc_Html_TblTitle_Direction, True)
        oHtmlWriter.TrEnd()

        ' extract data
        AlignStaInc_ExtractData.ExtractData(oAlignment, stationStart, stationEnd, NumericUpDown_StationInc.Value)

        For Each data As AlignStaInc_ExtractData.DirectionData In AlignStaInc_ExtractData.AlignDataList
            Try
                oHtmlWriter.TrBegin()
                oHtmlWriter.RenderTd(data.station)
                oHtmlWriter.RenderTd(data.northing)
                oHtmlWriter.RenderTd(data.easting)
                oHtmlWriter.RenderTd(data.tangentialDirection)
                oHtmlWriter.TrEnd()
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
            End Try
        Next

        oHtmlWriter.TableEnd()
        oHtmlWriter.RenderBr()
    End Sub

    Private Function GetDefaultIncrement() As Decimal
        Dim unitType As AeccLandLib.AeccCoordinateUnitType = AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitMeter
        Try
            unitType = ReportApplication.AeccXDatabase.Settings.DrawingSettings.AmbientSettings.DistanceSettings.Unit.Value
        Catch ex As Exception
            Diagnostics.Debug.Assert(False, ex.Message)
        End Try

        Dim dIncrement As Decimal
        If unitType = AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitFoot Then
            dIncrement = 50
        Else
            dIncrement = 20
        End If
        Return dIncrement
    End Function

    Private Sub openHelp()
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_INCREMENTAL_STATIONING_REPORT_ALIGNMENTS)
    End Sub

    Private ctlStartStation As CtrlStationTextBox
    Private ctlEndStation As CtrlStationTextBox
    Private ctlAlignmentListView As CtrlAlignmentListView
    Private ctlSavePath As CtrlSaveReportFile
    Private ctlProgressBar As CtrlProgressBar
End Class
