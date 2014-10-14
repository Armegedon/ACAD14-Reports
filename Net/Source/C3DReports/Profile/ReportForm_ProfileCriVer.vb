' -----------------------------------------------------------------------------
' <copyright file="ReportForm_ProfileCriVer.vb" company="Autodesk">
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
Imports Autodesk.AutoCAD.EditorInput
Imports System.Web.UI
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Friend Class ReportForm_ProfileCriVer

    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()
        Dim reportForm As New ReportForm_ProfileCriVer

        Try
            If reportForm.ReadyToOpen() = True Then
                ReportUtilities.RunModalDialog(reportForm)
            End If
        Catch ex As Exception
            reportForm.Dispose()
        End Try
    End Sub

    Private Sub ReportForm_ProfileCriVer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        oReportHtml.RenderHeader(LocalizedRes.ProfileCriVer_Html_Title)
        oReportHtml.Render("<div style=""width:7in"">")
        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.ProfileCriVer_Html_Header, 1)
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
        str = LocalizedRes.ProfileCriVer_Html_ProfileNameLabel
        str += " " + oProfile.Name
        oHtmlWriter.RenderLine(str)

        'render profile description
        str = String.Format(LocalizedRes.Common_Html_Description_One_Param, oProfile.Description)
        oHtmlWriter.RenderLine(str)

        'station range
        str = String.Format(LocalizedRes.Alignment_Html_Station_Range, _
                            ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, stationStart), _
                            ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, stationEnd))
        oHtmlWriter.Render(str)

        ProfileCriVer_ExtractData.ExtractData(oProfile, stationStart, stationEnd)
        str = "<TABLE id=ID_ProfileTable border='0' width=""100%"" >" 'style=""MARGIN-TOP: 4px"" cellSpacing=0 cellPadding=3>"
        oHtmlWriter.Render(str)

        Dim bCurve As Boolean
        Dim num As Integer
        num = 1

        For i As Integer = 0 To UBound(ProfileCriVer_ExtractData.ProfileDataArr)
            If ProfileCriVer_ExtractData.ProfileDataArr(i) Is Nothing Then
                Continue For
            End If
            If ProfileCriVer_ExtractData.ProfileDataArr(i)(ProfileCriVer_ExtractData.nCurveInfo) Is Nothing Then
                bCurve = False
            Else
                bCurve = True
            End If

            Dim sCurveLen As String
            If bCurve Then
                sCurveLen = CStr(ProfileCriVer_ExtractData.ProfileDataArr(i)(ProfileCriVer_ExtractData.nCurveInfo)(ProfPVICurve_ExtractData.nCurveLenIndex))
            Else
                sCurveLen = "&nbsp;"
            End If
            If bCurve Then
                WriteCurveInfoStr(i, num, oHtmlWriter)
                num = num + 1
            End If
        Next i
        oHtmlWriter.Render("</Table><HR/>")
        oHtmlWriter.RenderBr()
    End Sub


    Private Sub WriteCurveInfoStr(ByVal nNumber As Integer, ByVal num As Integer, ByVal oHtmlWriter As ReportWriter)
        On Error Resume Next

        writeCurveInfoHeader(nNumber, num, oHtmlWriter)

        WriteCurveInfo(nNumber, oHtmlWriter)

    End Sub

    Private Sub writeCurveInfoHeader(ByVal nNumber As Integer, ByVal num As Integer, ByVal oHtmlWriter As ReportWriter)
        Dim str As String

        Dim oCurveDataArr As Object
        oCurveDataArr = ProfileCriVer_ExtractData.ProfileDataArr(nNumber)(ProfileCriVer_ExtractData.nCurveInfo)
        oHtmlWriter.Render("<tr>")
        oHtmlWriter.Render("<td colspan=""4"" align=""center""><hr></td>")
        oHtmlWriter.Render("</tr>")
        oHtmlWriter.Render("<tr>")
        If oCurveDataArr(ProfileCriVer_ExtractData.nCurvTypeIndex) = AeccLandLib.AeccProfileVerticalCurveType.aeccSag Then
            str = "<td colspan=""3"" align=""left""><b>" & num
            str += " " + LocalizedRes.ProfileCriVer_Html_SagCurve + " </b></td>"

        Else
            str = "<td colspan=""3"" align=""left""><b>" & num
            str += " " + LocalizedRes.ProfileCriVer_Html_CrestCurve + " </b></td>"
        End If
        oHtmlWriter.Render(str)
        oHtmlWriter.Render("</tr>")
    End Sub

    Private Sub WriteCurveInfo(ByVal nNumber As Integer, ByVal oHtmlWriter As ReportWriter)

        Dim oCurveDataArr As Object
        oCurveDataArr = ProfileCriVer_ExtractData.ProfileDataArr(nNumber)(ProfileCriVer_ExtractData.nCurveInfo)
        'station
        WriteTypeValueRecord2("", LocalizedRes.ProfileCriVer_Html_PVCStation, _
                            oCurveDataArr(ProfileCriVer_ExtractData.nPVCStationIndex), "", _
                            oHtmlWriter, 1)   '"PVC Station:"
        WriteTypeValueRecord2("", LocalizedRes.ProfileCriVer_Html_PVIStation, _
                            ProfileCriVer_ExtractData.ProfileDataArr(nNumber)(ProfileCriVer_ExtractData.nStationIndex), "", oHtmlWriter, 1)  '"PVI Station:" 
        WriteTypeValueRecord2("", LocalizedRes.ProfileCriVer_Html_PVTStation, _
                            oCurveDataArr(ProfileCriVer_ExtractData.nPVTStationIndex), "", _
                            oHtmlWriter, 1) '"PVT Station:" 

        WriteTypeValueRecord2("", LocalizedRes.ProfileCriVer_Html_GradeIn, _
                            oCurveDataArr(ProfileCriVer_ExtractData.nGradeInIndex), "", _
                            oHtmlWriter, 1) '"Grade in(%):"
        WriteTypeValueRecord2("", LocalizedRes.ProfileCriVer_Html_GradeOut, _
                            ProfileCriVer_ExtractData.ProfileDataArr(nNumber)(ProfileCriVer_ExtractData.nGradeOutIndex), "", oHtmlWriter, 1) ' "Grade out(%):"

        WriteTypeValueRecord2("", LocalizedRes.ProfileCriVer_Html_CurveLength, _
                            oCurveDataArr(ProfileCriVer_ExtractData.nCurveLenIndex), "", _
                            oHtmlWriter, 1) '"Curve Length:"

        WriteTypeValueRecord2("", LocalizedRes.ProfileCriVer_Html_K, _
                            oCurveDataArr(ProfileCriVer_ExtractData.nKIndex), "", oHtmlWriter, 1) '"K:"
        WriteTypeValueRecord2("", LocalizedRes.ProfileCriVer_Html_DesignSpeed, oCurveDataArr(ProfileCriVer_ExtractData.nDesignSpeedIndex), "", oHtmlWriter, 1) '"Design Speed:"
        Dim temp As String
        temp = "<u>" + LocalizedRes.ProfileCriVer_Html_DesignCriteria + "</u>"
        WriteTypeValueRecord2("", temp, "", "", oHtmlWriter, 2)
        If oCurveDataArr(ProfileCriVer_ExtractData.nCurvTypeIndex) = AeccLandLib.AeccProfileVerticalCurveType.aeccCrest Then
            temp = "&nbsp;&nbsp;&nbsp;&nbsp;" + LocalizedRes.ProfileCriVer_Html_MinKForStoppingSightDis
            WriteTypeValueRecord2("", temp, _
                                oCurveDataArr(ProfileCriVer_ExtractData.nMinKForStoppingSightDisIndex), oCurveDataArr(ProfileCriVer_ExtractData.nMinKForStoppingSightDisResultIndex), oHtmlWriter, 2) '"Minimum K for Stopping Sight Distance::"
            temp = "&nbsp;&nbsp;&nbsp;&nbsp;" + LocalizedRes.ProfileCriVer_Html_MinKForPassingSightDis
            WriteTypeValueRecord2("", temp, _
                                oCurveDataArr(ProfileCriVer_ExtractData.nMinKForPassingSightDisIndex), oCurveDataArr(ProfileCriVer_ExtractData.nMinKForPassingSightDisResultIndex), oHtmlWriter, 2) '"Minimum K for Passing Sight Distance::"
            temp = "<u>" + LocalizedRes.ProfileCriVer_Html_DesignChecks + "</u>"
            WriteTypeValueRecord2("", temp, "", "", oHtmlWriter, 2) '"Design Checks:"            
        Else
            temp = "&nbsp;&nbsp;&nbsp;&nbsp;" + LocalizedRes.ProfileCriVer_Html_MinKForHeadlightSightDis
            WriteTypeValueRecord2("", temp, _
                                oCurveDataArr(ProfileCriVer_ExtractData.nMinKForHeadlightSightDisIndex), oCurveDataArr(ProfileCriVer_ExtractData.nMinKForHeadlightSightDisResultIndex), oHtmlWriter, 2) '"Minimum K for Headlight Sight Distance:"
            temp = "<u>" + LocalizedRes.ProfileCriVer_Html_DesignChecks + "</u>"
            WriteTypeValueRecord2("", temp, "", "", oHtmlWriter, 2) '"Design Checks:"    
        End If
        Dim arrDesignCheckNames()
        Dim arrDesignCheckResults()
        arrDesignCheckNames = oCurveDataArr(ProfileCriVer_ExtractData.nDesignCheckNamesIndex)
        arrDesignCheckResults = oCurveDataArr(ProfileCriVer_ExtractData.nDesignCheckResultsIndex)
        Dim i As Integer
        For i = 0 To UBound(arrDesignCheckNames)
            If Not TypeName(arrDesignCheckNames(i)) = "String" Then
                Continue For
            End If

            WriteTypeValueRecord2("", "&nbsp;&nbsp;&nbsp;&nbsp;" & arrDesignCheckNames(i), "", arrDesignCheckResults(i), oHtmlWriter, 2)

        Next i
    End Sub

    Private Sub WriteTypeValueRecord2(ByVal sType1 As String, ByVal sValue1 As String, ByVal sType2 As String, ByVal sValue2 As String, ByVal oHtmlWriter As ReportWriter, Optional ByVal nRowType As Long = 0)
        Dim str As String

        oHtmlWriter.Render("<TR>")
        If nRowType = 1 Then
            str = "<TD width=""5%"">" & sType1 & "</TD>"
            oHtmlWriter.Render(str)
            str = "<TD width=""25%"">" & sValue1 & "</TD>"
            oHtmlWriter.Render(str)
            str = "<TD width=""25%"">" & sType2 & "</TD>"
            oHtmlWriter.Render(str)
            str = "<TD>" & sValue2 & "</TD>"
            oHtmlWriter.Render(str)
        ElseIf nRowType = 2 Then
            str = "<TD width=""5%"">" & sType1 & "</TD>"
            oHtmlWriter.Render(str)
            str = "<TD width=""40%"">" & sValue1 & "</TD>"
            oHtmlWriter.Render(str)
            str = "<TD width=""30%"">" & sType2 & "</TD>"
            oHtmlWriter.Render(str)
            str = "<TD>" & sValue2 & "</TD>"
            oHtmlWriter.Render(str)
        Else
            str = "<TD>" & sType1 & "</TD>"
            oHtmlWriter.Render(str)
            str = "<TD>" & sValue1 & "</TD>"
            oHtmlWriter.Render(str)
            str = "<TD>" & sType2 & "</TD>"
            oHtmlWriter.Render(str)
            str = "<TD>" & sValue2 & "</TD>"
            oHtmlWriter.Render(str)
        End If
        oHtmlWriter.Render("</TR>")
    End Sub
    Private Function ReadyToOpen() As Boolean
        If CtrlProfileListView.GetProfileCount(AeccLandLib.AeccProfileType.aeccFinishedGround) < 1 Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub openHelp()
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_PROFILE_DESIGN_CRITERIA_VERIFICATION)
    End Sub

    Private ctlSavePath As CtrlSaveReportFile
    Private ctlProgressBar As CtrlProgressBar
    Private ctlProfileListView As CtrlProfileListView
    Private ctlStartStation As CtrlStationTextBox
    Private ctlEndStation As CtrlStationTextBox
End Class
