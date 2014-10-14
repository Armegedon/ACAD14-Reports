' -----------------------------------------------------------------------------
' <copyright file="ReportForm_ExistingProfilePoints.vb" company="Autodesk">
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
Imports System.IO
Imports System.Reflection ' For Missing.Value and BindingFlags
Imports AutoCAD = Autodesk.AutoCAD.Interop
Imports AeccUiLandLib = Autodesk.AECC.Interop.UiLand
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports AcadCommonLib = Autodesk.AutoCAD.Interop.Common
Imports AecUIBase = Autodesk.Aec.Interop.UIBase
Imports Autodesk.AutoCAD.EditorInput
Imports System.Runtime.InteropServices ' For COMException
Imports System.Web.UI
Imports Microsoft.Office.Interop.Excel
Imports LocalizedRes = Report.My.Resources.LocalizedStrings
Imports Autodesk.AECC.Interop

Friend Class ReportForm_ExistingProfilePoints

    Public Shared m_bRegularIntervalCheck As Boolean
    Public Shared m_bHTPSCheck As Boolean
    Public Shared m_bVTPSCheck As Boolean
    Public Shared m_bExistingCheck As Boolean
    Public Shared m_bDecimalSeparator As Boolean
    'Public Shared m_dStaInc As Double

    Private ctlSaveReport As CtrlSaveReportFile
    Private ctlProgressBar As CtrlProgressBar
    Private ctlProfileListDesignView As CtrlProfileListView
    Private ctlProfileListExistingView As CtrlProfileListExistingView
    Private WithEvents ctlStartStation As CtrlStationTextBox
    Private WithEvents ctlEndStation As CtrlStationTextBox

    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()
        Dim rptDlg As New ReportForm_ExistingProfilePoints
        If rptDlg.ReadyToOpen() = True Then
            ReportUtilities.RunModalDialog(rptDlg)
        End If
    End Sub

    Private Sub ReportForm_ExistingProfilePoints_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        CheckBox_HTML.Checked = True

        ctlSaveReport = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)

        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)

        ctlStartStation = New CtrlStationTextBox
        ctlStartStation.Initialize(TextBox_StartStation)

        ctlEndStation = New CtrlStationTextBox
        ctlEndStation.Initialize(TextBox_EndStation)

        ctlProfileListDesignView = New CtrlProfileListView
        ctlProfileListDesignView.AllItemChecked = False
        ctlProfileListDesignView.Initialize(ListView_ProfilesDesign, ctlStartStation, ctlEndStation)

        ctlProfileListExistingView = New CtrlProfileListExistingView
        ctlProfileListExistingView.Initialize(ListView_ProfilesExist)

        ctlProfileListExistingView.FillListView_Profile(ctlProfileListDesignView)

        '
        'm_dStaInc = "25"
        ' CtlReportSettingsEPP1.StatInter.Value = CStr(m_dStaInc)
    End Sub

    Private Sub BtnExecute_Click() Handles Button_CreateReport.Click
        m_bRegularIntervalCheck = CheckBox_RegularInterval.Checked
        m_bHTPSCheck = CheckBox_HTPoints.Checked
        m_bVTPSCheck = CheckBox_VTPoints.Checked
        m_bExistingCheck = CheckBox_EGPoints.Checked

        Dim checkedCount As Integer = ctlProfileListDesignView.WinListView.CheckedItems.Count
        If checkedCount = 0 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.Profile_Msg_SelectOneFirst)
            Exit Sub
        End If

        If ListView_ProfilesExist.Items.Count = 0 And _
            ReportUtilities.ACADMsgBox(LocalizedRes.Profile_Msg_NoStaticProfiles, _
                                       LocalizedRes.ReportUtilitie_Msg_Title, _
                                       MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Exit Sub
        End If

        Try
            Dim tempFile As String = ReportConverter.GetRandomTempHtmlPath()

            ctlProgressBar.ProgressBegin(checkedCount)

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
        Dim oProfileDesign As Land.AeccProfile
        Dim oProfileExisting As Land.AeccProfile
        Dim curListItemDesign As ListViewItem
        Dim curListItemExisting As ListViewItem
        Dim inStartStation As Double
        Dim inEndStation As Double
        Dim oStationSettings As Land.AeccSettingsStation = ReportApplication.AeccXDocument.Settings.AlignmentSettings.AmbientSettings.StationSettings
        ' number formatting in Slovenian language isn't right otherwise (decimal separator becomes "." instead of ",")
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Dim ci As System.Globalization.CultureInfo = New System.Globalization.CultureInfo(oldCI.Name)
        System.Threading.Thread.CurrentThread.CurrentCulture = ci

        Dim oReportHtml As New ReportWriter(fileName)

        '<html>
        oReportHtml.HtmlBegin()
        oReportHtml.RenderHeader(LocalizedRes.ExistingProfilePoints_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.ExistingProfilePoints_Html_Header, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        For t As Integer = 1 To ctlProfileListDesignView.WinListView.Items.Count
            curListItemDesign = ctlProfileListDesignView.WinListView.Items(t - 1)
            If curListItemDesign.Checked = True Then
                oProfileDesign = CType(curListItemDesign.Tag, Land.AeccProfile)
                If Not oProfileDesign Is Nothing Then
                    inStartStation = ReportUtilities.GetRawStation(curListItemDesign.SubItems(CtrlProfileListView.ListViewColumn.kStartStation).Text, oProfileDesign.Alignment.StationIndexIncrement)
                    inEndStation = ReportUtilities.GetRawStation(curListItemDesign.SubItems(CtrlProfileListView.ListViewColumn.kEndStation).Text, oProfileDesign.Alignment.StationIndexIncrement)
                    If oStationSettings.Unit.Value = Land.AeccCoordinateUnitType.aeccCoordinateUnitKilometer Then
                        inStartStation = inStartStation * 1000
                        inEndStation = inEndStation * 1000
                    End If
                    For u As Integer = 1 To ListView_ProfilesExist.Items.Count
                        curListItemExisting = ListView_ProfilesExist.Items(u - 1)
                        If curListItemExisting.Checked = True Then
                            oProfileExisting = CType(curListItemExisting.Tag, Land.AeccProfile)
                            If Not oProfileExisting Is Nothing Then
                                If oProfileExisting.Alignment.DisplayName <> oProfileDesign.Alignment.DisplayName Then
                                    '                                MsgBox "Profiles must be referenced to the same Alignment!"
                                    ReportUtilities.ACADMsgBox(LocalizedRes.ExistingProfilePoints_CreatedFromSameAlignment)
                                    '</body>
                                    oReportHtml.BodyEnd()

                                    '</html>
                                    oReportHtml.HtmlEnd()

                                    ' don't forget to switch back to original CI
                                    System.Threading.Thread.CurrentThread.CurrentCulture = oldCI
                                    Exit Sub
                                End If
                                AppendReportHTML(oProfileDesign, oProfileExisting, inStartStation, inEndStation, oReportHtml)
                                ctlProgressBar.PerformStep()
                            End If
                        End If
                    Next
                End If
            End If
        Next

        '</body>
        oReportHtml.BodyEnd()

        '</html>
        oReportHtml.HtmlEnd()

        ' don't forget to switch back to original CI
        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI
    End Sub

    ' Append one item's report to html file
    Private Sub AppendReportHTML(ByVal oProfileDesign As Land.AeccProfile, _
                                 ByVal oProfileExisting As Land.AeccProfile, _
                                 ByVal stationStart As Double, _
                                 ByVal stationEnd As Double, _
                                 ByVal oHtmlWriter As ReportWriter)

        Dim StaniceniStart As String
        Dim StaniceniEnd As String
        Dim staniceni As String

        If oProfileDesign Is Nothing Or oProfileExisting Is Nothing Or oHtmlWriter Is Nothing Then
            Exit Sub
        End If

        If InStr(1, CStr(oProfileDesign.Alignment.EndingStation), ",", vbTextCompare) > 0 Then
            m_bDecimalSeparator = True
        Else
            m_bDecimalSeparator = False
        End If

        '<hr>
        oHtmlWriter.RenderHr()

        ' alignmnent info
        Dim str As String

        'render alignment name
        str = LocalizedRes.ExistingProfilePoints_Html_VertAlignName
        str &= " " & oProfileDesign.Name
        oHtmlWriter.RenderLine(str)

        ' existing profile name
        str = LocalizedRes.ExistingProfilePoints_Html_ExistingProfileName
        str &= " " & oProfileExisting.Name
        oHtmlWriter.RenderLine(str)

        'render alignment description
        str = LocalizedRes.ExistingProfilePoints_Html_Description
        str += " " + oProfileDesign.Description
        oHtmlWriter.RenderLine(str)

        '
        If m_bDecimalSeparator = False Then
            StaniceniStart = ReportUtilities.GetStationStringWithDerived(oProfileDesign.Alignment, stationStart)
            StaniceniEnd = ReportUtilities.GetStationStringWithDerived(oProfileDesign.Alignment, stationEnd)
        Else
            StaniceniStart = Replace(ReportUtilities.GetStationStringWithDerived(oProfileDesign.Alignment, stationStart), ".", ",", 1, 1)
            StaniceniEnd = Replace(ReportUtilities.GetStationStringWithDerived(oProfileDesign.Alignment, stationEnd), ".", ",", 1, 1)
        End If
        str = LocalizedRes.ExistingProfilePoints_Html_Alignment_StationRange
        str += " " + LocalizedRes.ExistingProfilePoints_Html_Alignment_Start
        str += " " + StaniceniStart
        str += LocalizedRes.ExistingProfilePoints_Html_Alignment_End
        str += " " + StaniceniEnd
        oHtmlWriter.RenderLine(str)
        '<hr>
        oHtmlWriter.RenderHr()

        If oProfileExisting.Type = Land.AeccProfileType.aeccExistingGround Then
            ' extract data
            ExistingPP_ExtractData.ExtractData(oProfileDesign.Alignment, oProfileDesign, oProfileExisting, stationStart, stationEnd, StationInterval)

            oHtmlWriter.RenderBr()
            oHtmlWriter.TableBegin("1")

            'table title (head line)

            'title
            oHtmlWriter.TrBegin()
            oHtmlWriter.RenderTd(LocalizedRes.ExistingProfilePoints_Html_TblTitle_PVI, True)
            oHtmlWriter.RenderTd(LocalizedRes.ExistingProfilePoints_Html_TblTitle_Station, True)
            oHtmlWriter.RenderTd(LocalizedRes.ExistingProfilePoints_Html_TblTitle_Easting, True)
            oHtmlWriter.RenderTd(LocalizedRes.ExistingProfilePoints_Html_TblTitle_Northing, True)
            oHtmlWriter.RenderTd(LocalizedRes.ExistingProfilePoints_Html_TblTitle_ElevationExisting, True)
            oHtmlWriter.RenderTd(LocalizedRes.ExistingProfilePoints_Html_TblTitle_ElevationDesign, True)
            oHtmlWriter.RenderTd(LocalizedRes.ExistingProfilePoints_Html_TblTitle_ElevationDifference, True)
            oHtmlWriter.RenderTd(LocalizedRes.ExistingProfilePoints_Html_TblTitle_PointType, True)
            oHtmlWriter.TrEnd()

            For i As Integer = 0 To UBound(ExistingPP_ExtractData.m_oProfileDataArr)
                If ExistingPP_ExtractData.m_oProfileDataArr(i) Is Nothing Then
                    Exit For
                End If

                If m_bDecimalSeparator = False Then
                    staniceni = ExistingPP_ExtractData.m_oProfileDataArr(i)(ExistingPP_ExtractData.nStationIndex)
                Else
                    staniceni = Replace(ExistingPP_ExtractData.m_oProfileDataArr(i)(ExistingPP_ExtractData.nStationIndex), ".", ",", 1, 1)
                End If
                'format string
                oHtmlWriter.TrBegin()
                oHtmlWriter.RenderTd(CStr(i))
                oHtmlWriter.RenderTd(staniceni)
                oHtmlWriter.RenderTd(ExistingPP_ExtractData.m_oProfileDataArr(i)(ExistingPP_ExtractData.nEastingIndex))
                oHtmlWriter.RenderTd(ExistingPP_ExtractData.m_oProfileDataArr(i)(ExistingPP_ExtractData.nNorthingIndex))
                oHtmlWriter.RenderTd(ExistingPP_ExtractData.m_oProfileDataArr(i)(ExistingPP_ExtractData.nElevationIndex))
                oHtmlWriter.RenderTd(ExistingPP_ExtractData.m_oProfileDataArr(i)(ExistingPP_ExtractData.nElevationDesignIndex))
                oHtmlWriter.RenderTd(ExistingPP_ExtractData.m_oProfileDataArr(i)(ExistingPP_ExtractData.nElevationDifferenceIndex))
                oHtmlWriter.RenderTd(ExistingPP_ExtractData.m_oProfileDataArr(i)(ExistingPP_ExtractData.nReferenceIndex))
                oHtmlWriter.TrEnd()
            Next i
            oHtmlWriter.TableEnd()
            oHtmlWriter.RenderBr()
        End If
    End Sub

    Private Sub BtnHelp_Click() Handles Button_Help.Click
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_EXISTING_PROFILE_POINTS_REPORT)
    End Sub

    Private Sub ReportForm_HelpRequested(ByVal sender As System.Object, ByVal hlpevent As System.Windows.Forms.HelpEventArgs) Handles MyBase.HelpRequested
        BtnHelp_Click()
    End Sub

    Private Sub BtnDone_Click() Handles Button_Done.Click
        Close()
    End Sub

    Private Function ReadyToOpen() As Boolean
        ReadyToOpen = True
        Dim nProfileCount As Long
        Dim oAlignment As AeccLandLib.AeccAlignment
        'get profiles from Document
        For Each oAlignment In ReportApplication.AeccXDocument.AlignmentsSiteless
            nProfileCount += oAlignment.Profiles.Count
        Next
        'get profiles from Site
        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            For Each oAlignment In oSite.Alignments
                nProfileCount += oAlignment.Profiles.Count
            Next
        Next
        If nProfileCount < 1 Then
            Dim str As String = LocalizedRes.ExistingProfilePoints_Msg_NoProfiles
            Dim sMsg As String = str.Replace("&DrawingName&", ReportApplication.AeccXDocument.Name)
            ReportUtilities.ACADMsgBox(sMsg)
            ReadyToOpen = False
        End If
    End Function

    Private Sub ListView_ProfilesDesign_ItemChecked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles ListView_ProfilesDesign.ItemChecked
        If e.Item().Checked Then 'uncheck all others
            For Each item As ListViewItem In ListView_ProfilesDesign.Items()
                If item.Checked And Not _
                   (item.SubItems(CtrlProfileListView.ListViewColumn.kName).Text = e.Item.SubItems(CtrlProfileListView.ListViewColumn.kName).Text And _
                    item.SubItems(CtrlProfileListView.ListViewColumn.kAlignmentName).Text = e.Item.SubItems(CtrlProfileListView.ListViewColumn.kAlignmentName).Text) Then
                    item.Checked = False
                End If
            Next
        End If
        If Not ctlProfileListExistingView Is Nothing Then
            ctlProfileListExistingView.FillListView_Profile(ctlProfileListDesignView)
        End If
    End Sub

    Private ReadOnly Property StationInterval() As Double
        Get
            Return NumericUpDown_StationInc.Value
        End Get
    End Property
End Class
