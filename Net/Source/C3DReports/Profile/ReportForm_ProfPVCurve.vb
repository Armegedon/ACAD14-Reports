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

Imports System
Imports AutoCAD = Autodesk.AutoCAD.Interop
Imports AeccUiLandLib = Autodesk.AECC.Interop.UiLand
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports AcadCommonLib = Autodesk.AutoCAD.Interop.Common
Imports AecUIBase = Autodesk.AEC.Interop.UIBase
Imports LocalizedRes = Report.My.Resources.LocalizedStrings
Imports System.Web.UI


Friend Class ReportForm_ProfPVCurve
    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()
        Dim rptDlg As New ReportForm_ProfPVCurve
        If ReportForm_ProfPVCurve.CanOpen() = True Then
            ReportUtilities.RunModalDialog(rptDlg)
        End If
    End Sub

    Private Shared Function CanOpen() As Boolean
        CanOpen = True

        If CtrlProfileListView.GetProfileCount(AeccLandLib.AeccProfileType.aeccFinishedGround) < 1 Then
            CanOpen = False
        End If
    End Function

    Private Sub ReportForm_ProfPVCurve_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ctlSavePath = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)

        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)

        ctlStartStation = New CtrlStationTextBox
        ctlStartStation.Initialize(TextBox_StartStation)

        ctlEndStation = New CtrlStationTextBox
        ctlEndStation.Initialize(TextBox_EndStation)

        ctlProfileListView = New CtrlProfileListView
        ctlProfileListView.Initialize(ListView_Profiles, ctlStartStation, ctlEndStation)
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
        oReportHtml.RenderHeader(LocalizedRes.ProfPVCurve_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.ProfPVCurve_Html_Header, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        For Each item As ListViewItem In ctlProfileListView.WinListView.CheckedItems
            Try
                AppendReport(item.Tag, ctlProfileListView.StartRawStation(item), ctlProfileListView.EndRawStation(item), oReportHtml)
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
        str = LocalizedRes.ProfPVICurve_Html_ProfileNameLabel
        str += " " + oProfile.Name
        oHtmlWriter.RenderLine(str)

        'render profile description
        str = String.Format(LocalizedRes.Common_Html_Description_One_Param, oProfile.Description)
        oHtmlWriter.RenderLine(str)

        'station range
        str = String.Format(LocalizedRes.Alignment_Html_Station_Range, _
                    ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, stationStart), _
                    ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, stationEnd))
        oHtmlWriter.RenderLine(str)

        'extract data
        ProfPVCurve_ExtractData.ExtractData(oProfile, stationStart, stationEnd)

        oHtmlWriter.TableBegin(1, Nothing)

        For i As Integer = 0 To UBound(ProfPVCurve_ExtractData.ProfileDataArr)
            If ProfPVCurve_ExtractData.ProfileDataArr(i) Is Nothing Then
                Exit For
            End If
            WriteCurveInfoStr(i, oHtmlWriter)
        Next i

        oHtmlWriter.TableEnd()
        oHtmlWriter.RenderBr()
    End Sub


    Private Sub WriteCurveInfoStr(ByVal nNumber As Integer, ByVal oHtmlWriter As ReportWriter)
        On Error Resume Next

        oHtmlWriter.TrBegin()

        oHtmlWriter.AddAttribute(HtmlTextWriterAttribute.Colspan, 4.ToString())
        oHtmlWriter.TdBegin()

        writeCurveInfoHeader(nNumber, oHtmlWriter)

        WriteCurveInfo(nNumber, oHtmlWriter)

        oHtmlWriter.TdEnd()
        oHtmlWriter.TrEnd()
    End Sub

    Private Sub writeCurveInfoHeader(ByVal nNumber As Integer, ByVal oHtmlWriter As ReportWriter)
        Dim str As String

        Dim oCurveDataArr As Object
        oCurveDataArr = ProfPVCurve_ExtractData.ProfileDataArr(nNumber)

        If oCurveDataArr(ProfPVCurve_ExtractData.nCurvTypeIndex) = AeccLandLib.AeccProfileVerticalCurveType.aeccSag Then
            str = LocalizedRes.ProfPVICurve_Html_SagCurve
        Else
            str = LocalizedRes.ProfPVICurve_Html_CrestCurve
        End If
        oHtmlWriter.RenderDiv(str, , 10, , 4)
    End Sub

    Private Sub WriteCurveInfo(ByVal nNumber As Integer, ByVal oHtmlWriter As ReportWriter)

        Dim oCurveDataArr As Object
        oCurveDataArr = ProfPVCurve_ExtractData.ProfileDataArr(nNumber)

        oHtmlWriter.AddAttribute(HtmlTextWriterAttribute.Style, "BORDER-TOP: black 1px dashed")
        oHtmlWriter.TableBegin(, Nothing)
        oHtmlWriter.TBodyBegin()

        'station
        WriteTypeValueRecord2(LocalizedRes.ProfPVICurve_Html_PVCStation, _
                            oCurveDataArr(ProfPVCurve_ExtractData.nPVCStationIndex), _
                            LocalizedRes.ProfPVICurve_Html_Elevation, _
                            oCurveDataArr(ProfPVCurve_ExtractData.nPVCElevationIndex), _
                            oHtmlWriter, True)   '"PVC Station:","Elevation:"

        WriteTypeValueRecord2(LocalizedRes.ProfPVICurve_Html_PVIStation, _
                            oCurveDataArr(ProfPVCurve_ExtractData.nPVIStationIndex), _
                            LocalizedRes.ProfPVICurve_Html_Elevation, _
                            oCurveDataArr(ProfPVCurve_ExtractData.nPVIElevationIndex), _
                            oHtmlWriter) '"PVI Station:" "Elevation:"

        WriteTypeValueRecord2(LocalizedRes.ProfPVICurve_Html_PVTStation, _
                            oCurveDataArr(ProfPVCurve_ExtractData.nPVTStationIndex), _
                            LocalizedRes.ProfPVICurve_Html_Elevation, _
                            oCurveDataArr(ProfPVCurve_ExtractData.nPVTElevationIndex), _
                            oHtmlWriter) '"PVT Station:" "Elevation:"

        If TypeName(oCurveDataArr(ProfPVCurve_ExtractData.nHighStationIndex)) = "String" Then
            WriteTypeValueRecord2(LocalizedRes.ProfPVICurve_Html_HighPoint, _
                                oCurveDataArr(ProfPVCurve_ExtractData.nHighStationIndex), _
                                LocalizedRes.ProfPVICurve_Html_Elevation, _
                                oCurveDataArr(ProfPVCurve_ExtractData.nHightElevation), _
                                oHtmlWriter) '"High Point:"  "Elevation:"
        End If

        If TypeName(oCurveDataArr(ProfPVCurve_ExtractData.nLowStationIndex)) = "String" Then
            WriteTypeValueRecord2(LocalizedRes.ProfPVICurve_Html_LowPoint, _
                                oCurveDataArr(ProfPVCurve_ExtractData.nLowStationIndex), _
                                LocalizedRes.ProfPVICurve_Html_Elevation, _
                                oCurveDataArr(ProfPVCurve_ExtractData.nLowElevation), _
                                oHtmlWriter) '"Low Point:" "Elevation:"
        End If

        WriteTypeValueRecord2(LocalizedRes.ProfPVICurve_Html_GradeIn, _
                            oCurveDataArr(ProfPVCurve_ExtractData.nGradeInIndex), _
                            LocalizedRes.ProfPVICurve_Html_GradeOut, _
                            oCurveDataArr(ProfPVCurve_ExtractData.nGradeOutIndex), _
                            oHtmlWriter, True) '"Grade in(%):" "Grade out(%):"
        WriteTypeValueRecord2(LocalizedRes.ProfPVICurve_Html_Change, _
                            oCurveDataArr(ProfPVCurve_ExtractData.nGradeChange), _
                            LocalizedRes.ProfPVICurve_Html_K, _
                            oCurveDataArr(ProfPVCurve_ExtractData.nKIndex), _
                            oHtmlWriter) '"Change(%):" "K:"
        WriteTypeValueRecord2(LocalizedRes.ProfPVICurve_Html_CurveLength, _
                            oCurveDataArr(ProfPVCurve_ExtractData.nCurveLenIndex), _
                            LocalizedRes.ProfPVCurve_Html_CurveRadius, _
                            oCurveDataArr(ProfPVCurve_ExtractData.nCurveRadiusIndex), _
                            oHtmlWriter, True) '"Curve Length:"

        If oCurveDataArr(ProfPVCurve_ExtractData.nCurvTypeIndex) = AeccLandLib.AeccProfileVerticalCurveType.aeccSag Then
            WriteTypeValueRecord2(LocalizedRes.ProfPVICurve_Html_HeadlightDistance, _
                                oCurveDataArr(ProfPVCurve_ExtractData.nHeadlightDisIndex), _
                                "", _
                                "", _
                                oHtmlWriter) '"Headlight Distance:"
        Else
            WriteTypeValueRecord2(LocalizedRes.ProfPVICurve_Html_PassingDistance, _
                                oCurveDataArr(ProfPVCurve_ExtractData.nPassingDisIndex), _
                                LocalizedRes.ProfPVICurve_Html_StoppingDistance, _
                                oCurveDataArr(ProfPVCurve_ExtractData.nStoppingDisIndex), _
                                oHtmlWriter, True) '"Passing Distance:" "Stopping Distance:"
        End If
        oHtmlWriter.TBodyEnd()
        oHtmlWriter.TableEnd()
    End Sub

    Sub WriteTypeValueRecord2(ByVal sType1 As String, ByVal sValue1 As String, _
                                ByVal sType2 As String, ByVal sValue2 As String, _
                                ByVal oHtmlWriter As ReportWriter, Optional ByVal bInterval As Boolean = False)
        oHtmlWriter.TrBegin()

        If bInterval = True Then
            oHtmlWriter.RenderTd(sType1, , , , 6)
            oHtmlWriter.RenderTd(sValue1, , "right", 10, 6, 10)
            oHtmlWriter.RenderTd(sType2, , , , 6)
            oHtmlWriter.RenderTd(sValue2, , "right", 10, 6, 10)
        Else
            oHtmlWriter.RenderTd(sType1)
            oHtmlWriter.RenderTd(sValue1, , "right", 10, , 10)
            oHtmlWriter.RenderTd(sType2)
            oHtmlWriter.RenderTd(sValue2, , "right", 10, , 10)
        End If

        oHtmlWriter.TrEnd()
    End Sub

    Private Sub openHelp()
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_VERTICAL_CURVE_REPORT)
    End Sub

    Private ctlSavePath As CtrlSaveReportFile
    Private ctlProgressBar As CtrlProgressBar
    Private ctlProfileListView As CtrlProfileListView
    Private ctlStartStation As CtrlStationTextBox
    Private ctlEndStation As CtrlStationTextBox
End Class
