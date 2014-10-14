' -----------------------------------------------------------------------------
' <copyright file="ReportForm_AlignCriVer.vb" company="Autodesk">
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
Imports System.Web.UI
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Friend Class ReportForm_AlignCriVer

    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()
        If CtrlAlignmentListView.GetAlignmentCount() > 0 Then
            ReportUtilities.RunModalDialog(New ReportForm_AlignCriVer)
        End If
    End Sub

    Private Sub ReportForm_AlignCriVer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ctlStartStation = New CtrlStationTextBox
        ctlStartStation.Initialize(TextBox_StartStation)

        ctlEndStation = New CtrlStationTextBox
        ctlEndStation.Initialize(TextBox_EndStation)

        ctlAlignmentListView = New CtrlAlignmentListView
        ctlAlignmentListView.Initialize(ListView_Alignments, ctlStartStation, ctlEndStation)

        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)

        ctlSavePath = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)
    End Sub
    Private Sub Button_CreateReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_CreateReport.Click
        If ctlAlignmentListView.WinListView.CheckedItems.Count = 0 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.Alignment_Msg_SelectOneFirst)
            Exit Sub
        End If

        'init progress bar
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
        oReportHtml.HtmlBegin() 'change resource
        oReportHtml.RenderHeader(LocalizedRes.AlignCriVer_Html_Title)
        oReportHtml.Render("<div style=""width:7in"">")
        '<body>
        oReportHtml.BodyBegin()
        oReportHtml.RenderH(LocalizedRes.AlignCriVer_Html_Header, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        For Each item As ListViewItem In ctlAlignmentListView.WinListView.CheckedItems
            Dim startRawStation, endRawStation As Double
            startRawStation = ctlAlignmentListView.StartRawStation(item)
            endRawStation = ctlAlignmentListView.EndRawStation(item)
            Try
                AppendReport(item.Tag, startRawStation, endRawStation, oReportHtml)
                ctlProgressBar.PerformStep()
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
            End Try
        Next item

        '</body>
        oReportHtml.Render("</body>")

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

        'alignment name'change resource
        str = LocalizedRes.AlignCriVer_Html_Align_Name
        str += " " + oAlignment.Name
        oHtmlWriter.RenderLine(str)

        'alignment description'change resource
        str = String.Format(LocalizedRes.Common_Html_Description_One_Param, oAlignment.Description)
        oHtmlWriter.RenderLine(str)

        'start station, end station'change resource
        str = String.Format(LocalizedRes.Alignment_Html_Station_Range, ReportUtilities.GetStationStringWithDerived(oAlignment, stationStart), _
                      ReportUtilities.GetStationStringWithDerived(oAlignment, stationEnd))
        oHtmlWriter.Render(str)

        ' extract data
        AlignCriVer_ExtractData.ExtractData(oAlignment, stationStart, stationEnd)
        str = "<TABLE id=ID_ProfileTable border='0' width=""100%"" >" 'style=""MARGIN-TOP: 4px"" cellSpacing=0 cellPadding=3>"
        oHtmlWriter.Render(str)
        Dim i As Integer
        Dim nEntityPrefix As Integer
        nEntityPrefix = 0
        For i = 0 To UBound(AlignCriVer_ExtractData.AlignDataArr)
            If AlignCriVer_ExtractData.AlignDataArr(i) Is Nothing Then
                Exit For
            End If
            If AlignCriVer_ExtractData.AlignDataArr(i)(AlignCriVer_ExtractData.nEntityPrefixIndex) <= 1 Then
                nEntityPrefix = nEntityPrefix + 1
            End If
            WriteCurveInfoStr(i, nEntityPrefix, oHtmlWriter, stationStart, stationEnd, oAlignment)

        Next i
        oHtmlWriter.Render("</Table><HR/>")
        oHtmlWriter.RenderBr()
    End Sub
    Private Sub WriteCurveInfoStr(ByVal nNumber As Long, ByVal nEntityPrefix As Integer, _
                                  ByVal oHtmlWriter As ReportWriter, _
                                  ByVal startStation As Double, ByVal endStation As Double, _
                                  ByVal oAlignment As AeccLandLib.AeccAlignment)
        On Error Resume Next
        Dim strEntityPrefix As String
        strEntityPrefix = ""
        If AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nEntityPrefixIndex) = 0 Then
            strEntityPrefix = nEntityPrefix
            strEntityPrefix += " "
        Else
            strEntityPrefix = nEntityPrefix
            strEntityPrefix += "."
            strEntityPrefix = strEntityPrefix & AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nEntityPrefixIndex)
            strEntityPrefix += " "
        End If

        Dim prevStation As String = AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nPrevStationIndex)
        Dim nextStation As String = AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nNextStationIndex)
        Dim dPrevStation As Double = ReportUtilities.GetRawStation(prevStation, oAlignment.StationIndexIncrement)
        Dim dNextStation As Double = ReportUtilities.GetRawStation(nextStation, oAlignment.StationIndexIncrement)

        If dNextStation >= startStation And dPrevStation <= endStation Then

            Dim str As String
            'Line ---------------------------------------------------------
            If AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nTypeIndex) = AeccLandLib.AeccAlignmentEntityType.aeccTangent Then
                oHtmlWriter.Render("<tr>")
                oHtmlWriter.Render("<td colspan=""4"" align=""center""><hr></td>")
                oHtmlWriter.Render("</tr>")
                oHtmlWriter.Render("<tr>")
                str = "<td colspan=""3"" align=""left""><b>" & strEntityPrefix & LocalizedRes.AlignCriVer_Html_Align_Tangent & "</b></td>"
                oHtmlWriter.Render(str)
                oHtmlWriter.Render("</tr>")

                WriteTypeValueRecord2("", LocalizedRes.AlignCriVer_Form_ReportSettings_StartSta, _
                    AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nPrevStationIndex), _
                    "", oHtmlWriter, 1)
                WriteTypeValueRecord2("", LocalizedRes.AlignCriVer_Form_ReportSettings_EndSta, _
                    AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nNextStationIndex), _
                    "", oHtmlWriter, 1)
                WriteTypeValueRecord2("", LocalizedRes.AlignCriVer_Form_ReportSettings_Length, _
                    AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nLengthIndex), _
                    "", oHtmlWriter, 1)
                WriteTypeValueRecord2("", LocalizedRes.AlignCriVer_Form_ReportSettings_DesignSpeed, _
                    AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nDesignSpeedIndex), _
                    "", oHtmlWriter, 1)
                Dim temp As String
                temp = "<u>" + LocalizedRes.AlignCriVer_Html_Align_DesignChecks + "</u>"
                WriteTypeValueRecord2("", temp, "", "", oHtmlWriter)
                Dim arrDesignCheckNames As Object
                arrDesignCheckNames = AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nDesignCheckNamesIndex)
                Dim arrDesignCheckResults As Object
                arrDesignCheckResults = AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nDesignCheckResultsIndex)
                Dim i As Long
                For i = 0 To UBound(arrDesignCheckNames)
                    If Not TypeName(arrDesignCheckNames(i)) = "String" Then
                        Continue For
                    End If

                    WriteTypeValueRecord2("", "&nbsp;&nbsp;&nbsp;&nbsp;" & arrDesignCheckNames(i), "", arrDesignCheckResults(i), oHtmlWriter, 2)

                Next i
                'Arc ---------------------------------------------------------
            ElseIf AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nTypeIndex) = AeccLandLib.AeccAlignmentEntityType.aeccArc Then
                oHtmlWriter.Render("<tr>")
                oHtmlWriter.Render("<td colspan=""4"" align=""center""><hr></td>")
                oHtmlWriter.Render("</tr>")
                oHtmlWriter.Render("<tr>")
                str = "<td colspan=""3"" align=""left""><b>" & strEntityPrefix & _
                    LocalizedRes.AlignCriVer_Html_Align_CirCurve & "</b></td>"
                oHtmlWriter.Render(str)
                oHtmlWriter.Render("</tr>")
                WriteTypeValueRecord2("", LocalizedRes.AlignCriVer_Form_ReportSettings_StartSta, _
                    AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nPrevStationIndex), _
                    "", oHtmlWriter, 1)
                WriteTypeValueRecord2("", LocalizedRes.AlignCriVer_Form_ReportSettings_EndSta, _
                    AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nNextStationIndex), _
                    "", oHtmlWriter, 1)
                WriteTypeValueRecord2("", LocalizedRes.AlignCriVer_Form_ReportSettings_Radius, _
                    AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nRadiusIndex), _
                    "", oHtmlWriter, 1)
                WriteTypeValueRecord2("", LocalizedRes.AlignCriVer_Form_ReportSettings_DesignSpeed, _
                    AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nDesignSpeedIndex), _
                    "", oHtmlWriter, 1)
                Dim temp As String
                temp = "<u>" + LocalizedRes.AlignCriVer_Html_Align_DesignCriteria + "</u>"
                WriteTypeValueRecord2("", temp, "", "", oHtmlWriter)
                temp = "&nbsp;&nbsp;&nbsp;&nbsp;" + LocalizedRes.AlignCriVer_Html_Align_MinRadius
                WriteTypeValueRecord2("", temp, AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nMinRadiusIndex), AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nMinRadiusResultIndex), oHtmlWriter, 2)
                temp = "<u>" + LocalizedRes.AlignCriVer_Html_Align_DesignChecks + "</u>"
                WriteTypeValueRecord2("", temp, "", "", oHtmlWriter)
                Dim arrDesignCheckNames As Object
                arrDesignCheckNames = AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nDesignCheckNamesIndex)
                Dim arrDesignCheckResults As Object
                arrDesignCheckResults = AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nDesignCheckResultsIndex)
                Dim i As Long
                For i = 0 To UBound(arrDesignCheckNames)
                    If Not TypeName(arrDesignCheckNames(i)) = "String" Then
                        Continue For
                    End If

                    WriteTypeValueRecord2("", "&nbsp;&nbsp;&nbsp;&nbsp;" & arrDesignCheckNames(i), "", arrDesignCheckResults(i), oHtmlWriter, 2)

                Next i


                'Spiral ---------------------------------------------------------
            ElseIf AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nTypeIndex) = AeccLandLib.AeccAlignmentEntityType.aeccSpiral Then
                oHtmlWriter.Render("<tr>")
                oHtmlWriter.Render("<td colspan=""4"" align=""center""><hr></td>")
                oHtmlWriter.Render("</tr>")
                oHtmlWriter.Render("<tr>")
                str = "<td colspan=""3"" align=""left""><b>" & strEntityPrefix & _
                    LocalizedRes.AlignCriVer_Html_Align_SpiralCurve + "</b></td>"
                oHtmlWriter.Render(str)
                oHtmlWriter.Render("</tr>")
                WriteTypeValueRecord2("", LocalizedRes.AlignCriVer_Form_ReportSettings_StartSta, _
                    AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nPrevStationIndex), _
                    "", oHtmlWriter, 1)
                WriteTypeValueRecord2("", LocalizedRes.AlignCriVer_Form_ReportSettings_EndSta, _
                    AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nNextStationIndex), _
                    "", oHtmlWriter, 1)
                WriteTypeValueRecord2("", LocalizedRes.AlignCriVer_Form_ReportSettings_Length, _
                    AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nLengthIndex), _
                    "", oHtmlWriter, 1)
                WriteTypeValueRecord2("", LocalizedRes.AlignCriVer_Form_ReportSettings_A, _
                    AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nAValueIndex), _
                    "", oHtmlWriter, 1)
                WriteTypeValueRecord2("", LocalizedRes.AlignCriVer_Form_ReportSettings_DesignSpeed, _
                    AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nDesignSpeedIndex), _
                    "", oHtmlWriter, 1)
                Dim temp As String
                'don't output transition length for compound spirals
                If AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nCompoundSpiralIndex) = False Then
                    temp = "<u>" + LocalizedRes.AlignCriVer_Html_Align_DesignCriteria + "</u>"
                    WriteTypeValueRecord2("", temp, "", "", oHtmlWriter)
                    temp = "&nbsp;&nbsp;&nbsp;&nbsp;" + LocalizedRes.AlignCriVer_Form_ReportSettings_MinTransLen
                    WriteTypeValueRecord2("", temp, AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nMinTransLenIndex), AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nMinTransLenResultIndex), oHtmlWriter, 2)
                End If
                temp = "<u>" + LocalizedRes.AlignCriVer_Html_Align_DesignChecks + "</u>"
                WriteTypeValueRecord2("", temp, "", "", oHtmlWriter, 1)
                Dim arrDesignCheckNames()
                Dim arrDesignCheckResults()
                arrDesignCheckNames = AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nDesignCheckNamesIndex)
                arrDesignCheckResults = AlignCriVer_ExtractData.AlignDataArr(nNumber)(AlignCriVer_ExtractData.nDesignCheckResultsIndex)
                Dim i As Long
                For i = 0 To UBound(arrDesignCheckNames)
                    If Not TypeName(arrDesignCheckNames(i)) = "String" Then
                        Continue For
                    End If

                    WriteTypeValueRecord2("", "&nbsp;&nbsp;&nbsp;&nbsp;" & arrDesignCheckNames(i), "", arrDesignCheckResults(i), oHtmlWriter, 2)

                Next i

            End If
        End If

    End Sub
    Private Sub WriteTypeValueRecord2(ByVal sType1 As Object, ByVal sValue1 As Object, ByVal sType2 As Object, ByVal sValue2 As Object, ByVal oHtmlWriter As ReportWriter, Optional ByVal nRowType As Long = 0)
        Dim str

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
            str = "<TD width=""25%"">" & sType2 & "</TD>"
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

    Private Sub openHelp()
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_ALIGNMENT_DESIGN_CRITERIA_VERIFICATION)
    End Sub

    Private ctlStartStation As CtrlStationTextBox
    Private ctlEndStation As CtrlStationTextBox
    Private ctlAlignmentListView As CtrlAlignmentListView
    Private ctlProgressBar As CtrlProgressBar
    Private ctlSavePath As CtrlSaveReportFile
End Class
