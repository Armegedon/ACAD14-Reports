' -----------------------------------------------------------------------------
' <copyright file="ReportForm_ProfPVCurve.vb" company="Autodesk">
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
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Friend Class ReportForm_ProfStaInc

    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()
        If CanOpen() Then
            ' Create report dialog and display it
            ReportUtilities.RunModalDialog(New ReportForm_ProfStaInc)
        End If
    End Sub

    Private Shared Function CanOpen() As Boolean
        CanOpen = True

        If CtrlProfileListView.GetProfileCount(AeccLandLib.AeccProfileType.aeccFinishedGround) < 1 Then
            CanOpen = False
        End If
    End Function

    Private Sub ReportForm_ProfStaInc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ctlSavePath = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)

        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)

        ctlStartStation = New CtrlStationTextBox
        ctlStartStation.Initialize(TextBox_StartStation)

        ctlEndStation = New CtrlStationTextBox
        ctlEndStation.Initialize(TextBox_EndStation)

        ctlProfileListView = New CtrlProfileListView
        ctlProfileListView.Initialize(ListView_Profiles, ctlStartStation, ctlEndStation)

        NumericUpDown_StationInc.Value = GetDefaultIncrement()
        NumericUpDown_StationInc.Maximum = Decimal.MaxValue
        NumericUpDown_StationInc.Minimum = 0.01D
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

    Private Sub Button_CreateReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_CreateReport.Click
        If ctlProfileListView.WinListView.CheckedItems.Count = 0 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.Profile_Msg_SelectOneFirst)
            Exit Sub
        End If

        'reset progress bar
        ctlProgressBar.ProgressBegin(ctlProfileListView.WinListView.CheckedItems.Count)

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
        oReportHtml.RenderHeader(LocalizedRes.ProfStaInc_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.ProfStaInc_Html_Title, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        For Each item As ListViewItem In ctlProfileListView.WinListView.CheckedItems
            Try
                AppendReport(CType(item.Tag, AeccLandLib.AeccProfile), ctlProfileListView.StartRawStation(item), ctlProfileListView.EndRawStation(item), oReportHtml)
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
            End Try
            ctlProgressBar.PerformStep()
        Next item

        '</body>
        oReportHtml.BodyEnd()

        '</html>
        oReportHtml.HtmlEnd()
    End Sub

    Private Sub AppendReport(ByVal oProfile As AeccLandLib.AeccProfile, _
                            ByVal stationStart As Double, _
                            ByVal stationEnd As Double, _
                            ByVal oHtmlWriter As ReportWriter)

        If oProfile Is Nothing Or oHtmlWriter Is Nothing Then
            Exit Sub
        End If

        '<hr>
        oHtmlWriter.RenderHr()

        ' alignment info
        Dim str As String

        'render profile name
        str = LocalizedRes.ProfStaInc_Html_Profile_Name
        str += " " + oProfile.Name
        oHtmlWriter.RenderLine(str)

        'render profile description
        str = String.Format(LocalizedRes.Common_Html_Description_One_Param, oProfile.Description)
        oHtmlWriter.RenderLine(str)

        str = String.Format(LocalizedRes.Alignment_Html_Station_Range, _
                            ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, stationStart), _
                            ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, stationEnd))
        oHtmlWriter.RenderLine(str)

        str = LocalizedRes.ProfStaInc_Html_StationInc
        str += " " + NumericUpDown_StationInc.Value.ToString("N")
        oHtmlWriter.RenderLine(str)

        'extract data
        ProfStaInc_ExtractData.ExtractData(oProfile, stationStart, stationEnd, NumericUpDown_StationInc.Value)

        oHtmlWriter.RenderBr()
        oHtmlWriter.TableBegin("1")

        oHtmlWriter.TrBegin()
        oHtmlWriter.RenderTd(LocalizedRes.ProfStaInc_Html_TblTitle_Station, True)
        oHtmlWriter.RenderTd(LocalizedRes.ProfStaInc_Html_TblTitle_Elevation, True)
        oHtmlWriter.RenderTd(LocalizedRes.ProfStaInc_Html_TblTitle_GradePercent, True)
        oHtmlWriter.RenderTd(LocalizedRes.ProfStaInc_Html_TblTitle_Location, True)
        oHtmlWriter.TrEnd()

        For Each data As ProfStaInc_ExtractData.ProfileData In ProfStaInc_ExtractData.DataDictionary.Values
            'format string
            oHtmlWriter.TrBegin()
            oHtmlWriter.RenderTd(data.Station)
            oHtmlWriter.RenderTd(data.Elevation)
            oHtmlWriter.RenderTd(data.Grade)
            oHtmlWriter.RenderTd(data.Location)
            oHtmlWriter.TrEnd()
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
            dIncrement = 50D
        Else
            dIncrement = 20D
        End If
        Return dIncrement
    End Function

    Private Sub openHelp()
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_INCREMENTAL_STATIONING_REPORT_PROFILES)
    End Sub

    Private ctlSavePath As CtrlSaveReportFile
    Private ctlProgressBar As CtrlProgressBar
    Private ctlProfileListView As CtrlProfileListView
    Private ctlStartStation As CtrlStationTextBox
    Private ctlEndStation As CtrlStationTextBox
End Class
