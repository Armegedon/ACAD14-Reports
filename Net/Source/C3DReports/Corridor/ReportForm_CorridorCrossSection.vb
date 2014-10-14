' -----------------------------------------------------------------------------
' <copyright file="ReportForm_CorridorCrossSection.vb" company="Autodesk">
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

Imports System.IO
Imports System.Reflection ' For Missing.Value and BindingFlags
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports AeccRoadLib = Autodesk.AECC.Interop.Roadway
Imports Microsoft.Office.Interop.Excel
Imports LocalizedRes = Report.My.Resources.LocalizedStrings
Imports Autodesk.AECC.Interop

Friend Class ReportForm_CorridorCrossSection

    Private ctlProgressBar As CtrlProgressBar
    Private ctlSaveReport As CtrlSaveReportFile
    Private WithEvents ctlStartStation As CtrlStationTextBox
    Private WithEvents ctlEndStation As CtrlStationTextBox

    Public Const INDEX_SELECT = 0
    Public Const INDEX_CODES = 1

    Private Const EXTRACT_STEPS = 30

    Private m_Corridors As New Dictionary(Of String, CCorridor_Emia)
    Private m_Alignments As New Dictionary(Of String, AeccLandLib.AeccAlignment)
    Private m_Surfaces As New Dictionary(Of String, AeccLandLib.AeccSurface)
    Private m_bEnableSurfaces As Boolean

    Private LastSelectedCorridorIndex As Integer = -1

    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()
        Dim rptDlg As New ReportForm_CorridorCrossSection
        If rptDlg.CanOpen() = True Then
            ReportUtilities.RunModalDialog(rptDlg)
        End If
    End Sub

    Private Function CanOpen() As Boolean
        Dim bNoCorridors As Boolean = False
        Dim nAlignmentCount As Integer = 0
        Dim nFeatureLineCount As Integer = 0
        Dim nSampleLineGroupsCount As Integer = 0

        Try
            If ReportApplication.AeccXRoadwayDocument.Corridors.Count < 1 Then
                bNoCorridors = True
            Else
                'initial Corridors
                Dim oCorridors As AeccRoadLib.AeccCorridors
                m_Corridors.Clear() '.removeAll()
                oCorridors = ReportApplication.AeccXRoadwayDocument.Corridors
                For Each oCorridor As AeccRoadLib.IAeccCorridor In oCorridors
                    If m_Corridors.ContainsKey(oCorridor.Name) = False Then
                        ' Block reference that hasn't been exploded should not be recognized in report dialog
                        ' Check whether its owner - that block record is layout, in-block corridor would not be in the layout block owner
                        If Not ReportUtilities.IsCorridorInLayout(oCorridor) Then
                            Continue For
                        End If

                        Dim oCorridorInfo As New CCorridor_Emia
                        BuildCorridorInfo(oCorridor, oCorridorInfo)

                        'count feature lines
                        nFeatureLineCount += oCorridor.FeatureLineCodeInfos.Count
                        m_Corridors.Add(oCorridor.Name, oCorridorInfo)
                    End If
                Next

                If m_Corridors.Count > 0 Then
                    bNoCorridors = False
                Else
                    bNoCorridors = True
                End If
            End If
        Catch ex As Exception
            bNoCorridors = True
        End Try

        'count alignments and sample line groups
        'get alignments from Document
        nAlignmentCount = ReportApplication.AeccXDocument.AlignmentsSiteless.Count()
        For Each oAlignment As Land.AeccAlignment In ReportApplication.AeccXDocument.AlignmentsSiteless
            nSampleLineGroupsCount += oAlignment.SampleLineGroups.Count()
        Next

        'get alignments from Site :
        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            nAlignmentCount += oSite.Alignments.Count()
            For Each oAlignment As Land.AeccAlignment In oSite.Alignments
                nSampleLineGroupsCount += oAlignment.SampleLineGroups.Count()
            Next
        Next

        CanOpen = False

        'If bNoCorridors = True Then
        '    Dim sNoCorridors As String = LocalizedRes.GetString("CorridorCrossSection_Msg_NoCorridors")
        '    Dim sMsg As String = sNoCorridors.Replace("&DrawingName&", ReportApplication.AeccXDocument.Name)
        '    ReportUtilities.ACADMsgBox(sMsg)
        'ElseIf nAlignmentCount < 1 Then
        '    ReportUtilities.ACADMsgBox(LocalizedRes.GetString("CorridorCrossSection_Msg_NoAlignment"))
        'ElseIf nFeatureLineCount < 1 Then
        '    ReportUtilities.ACADMsgBox(LocalizedRes.GetString("CorridorCrossSection_Msg_NoFeatureLine"))
        'Else
        '    CanOpen = True
        'End If

        m_bEnableSurfaces = False
        If (bNoCorridors = True Or nFeatureLineCount < 1) And ReportApplication.AeccXDocument.Surfaces.Count > 0 Then
            ' enable only Surfaces radio button and set it as default option in OnLoad Sub
            m_bEnableSurfaces = True
        End If

        If nAlignmentCount < 1 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorCrossSection_Msg_NoAlignment)
        ElseIf nSampleLineGroupsCount < 1 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorCrossSection_Msg_NoSLG)
        ElseIf bNoCorridors And ReportApplication.AeccXDocument.Surfaces.Count < 1 Then
            Dim sNoCorridors As String = LocalizedRes.CorridorCrossSection_Msg_NoCorridorsOrSurfaces
            Dim sMsg As String = sNoCorridors.Replace("&DrawingName&", ReportApplication.AeccXDocument.Name)
            ReportUtilities.ACADMsgBox(sMsg)
        Else
            CanOpen = True
        End If

    End Function

    Private ReadOnly Property StationInterval() As Decimal
        Get
            Return NumericUpDown_StationInc.Value
        End Get
    End Property


    Private Sub ReportForm_CorridorCrossSection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Init radio buttons
        If m_bEnableSurfaces Then
            radioSurfaces.Checked = True
            radioLinks.Enabled = False
            radioPoints.Enabled = False
            'enable and disable some control's depending on selected radio button
            Label_StationInc.Enabled = False
            NumericUpDown_StationInc.Enabled = False
            Combo_Corridor.Enabled = False
            Label_SelectCorridor.Enabled = False
            Label_SelectSLG.Enabled = True
            Combo_SLG.Enabled = True
        Else
            radioPoints.Checked = True
            'enable and disable some control's depending on selected radio button
            Label_StationInc.Enabled = True
            NumericUpDown_StationInc.Enabled = True
            Label_SelectSLG.Enabled = False
            Combo_SLG.Enabled = False
        End If

        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)

        ctlSaveReport = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)

        ctlStartStation = New CtrlStationTextBox
        ctlStartStation.Initialize(TextBox_StartStation)

        ctlEndStation = New CtrlStationTextBox
        ctlEndStation.Initialize(TextBox_EndStation)

        Dim bp1 As System.Drawing.Bitmap = My.Resources.Resources.ChkAll
        Dim bp2 As System.Drawing.Bitmap = My.Resources.Resources.UnchkAll
        bp1.MakeTransparent(Color.White)
        bp2.MakeTransparent(Color.White)
        Button_CheckAll.Image = bp1
        Button_CheckAll.ImageAlign = ContentAlignment.MiddleCenter
        Button_UncheckAll.Image = bp2
        Button_UncheckAll.ImageAlign = ContentAlignment.MiddleCenter

        'Init Feature Lines View
        InitFeatureLinesView()

        If Not m_bEnableSurfaces Then
            'initial Select report Components
            For Each corridorName As String In m_Corridors.Keys
                Combo_Corridor.Items.Add(corridorName)
            Next
        End If

        ' fill in private member with Alignments
        'get alignments from Document
        m_Alignments.Clear()
        For Each oAlignment As AeccLandLib.AeccAlignment In ReportApplication.AeccXDocument.AlignmentsSiteless
            m_Alignments.Add(oAlignment.Name, oAlignment)
        Next

        'get alignments from Site
        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            For Each oAlignment As AeccLandLib.AeccAlignment In oSite.Alignments
                m_Alignments.Add(oAlignment.Name, oAlignment)
            Next
        Next

        ' fill in private member with Surfaces
        m_Surfaces.Clear()
        For Each oSurface As AeccLandLib.AeccSurface In ReportApplication.AeccXDocument.Surfaces
            If UBound(oSurface.Points) > 2 Then 'include only surfaces with more than 2 points
                If (oSurface.Type <> Land.AeccSurfaceType.aecckGridVolumeSurface) Then
                    m_Surfaces.Add(oSurface.Name, oSurface)
                End If
            End If
        Next

        If m_bEnableSurfaces Then
            FillAlignmentComboAndFLL()
            FillInSLGAlign()
        Else
            'Corridor count should more then 0, because we have test it before dialog open
            ' it is an invariant of if statments in CanOpen(), when m_bEnableSurfaces = False
            If m_Corridors.Count > 0 Then
                Combo_Corridor.SelectedIndex = 0 'fills in data in controls in Combo_Corridor_SelectedIndexChanged()
            End If
        End If

        Exit Sub
    End Sub

    Private Sub InitFeatureLinesView()

        'set group title
        If m_bEnableSurfaces Then
            GroupBox_FeatureLines.Text = LocalizedRes.CorridorCrossSection_Form_List_Surfaces
        Else
            GroupBox_FeatureLines.Text = LocalizedRes.CorridorCrossSection_Form_List_FeatLns
        End If

        ' list view grid setup
        ' setup columns
        ListView_FeatureLines.View = View.Details
        ListView_FeatureLines.Columns.Add(LocalizedRes.CorridorCrossSection_Form_ColumnTitle_Check, _
            100, HorizontalAlignment.Left)
        If m_bEnableSurfaces Then
            ListView_FeatureLines.Columns.Add(LocalizedRes.CorridorCrossSection_Form_ColumnTitle_Surfaces, _
                    80, HorizontalAlignment.Left)
        Else
            ListView_FeatureLines.Columns.Add(LocalizedRes.CorridorCrossSection_Form_ColumnTitle_PointCodes, _
                80, HorizontalAlignment.Left)
        End If
    End Sub


    Private Sub BuildCorridorInfo(ByRef oCorridor As AeccRoadLib.IAeccCorridor, ByRef oCorridorInfo As CCorridor_Emia)
        oCorridorInfo.mObject = oCorridor
        oCorridorInfo.mBaselines.Clear() 'dicBaselines.Clear()

        For Each oBaseline As AeccRoadLib.AeccBaseline In oCorridor.Baselines
            Dim algnName As String
            algnName = oBaseline.Alignment.Name
            Dim vals As List(Of CBaseline)
            If oCorridorInfo.mBaselines.ContainsKey(algnName) = False Then
                vals = New List(Of CBaseline)
                oCorridorInfo.mBaselines.Add(algnName, vals)
            Else
                vals = oCorridorInfo.mBaselines(algnName)
            End If

            Dim oBaselineInfo As CBaseline
            oBaselineInfo = New CBaseline
            BuildBaselineInfo(oBaseline, oBaselineInfo)
            vals.Add(oBaselineInfo)
        Next
    End Sub

    Private Sub BuildBaselineInfo(ByRef oBaseline As AeccRoadLib.IAeccBaseline, ByRef oBaselineInfo As CBaseline)
        With oBaselineInfo
            .mObject = oBaseline
            .mSampleLineGroups.Clear()
            .mLinkCodesGroups.Clear()

            If oBaseline.Alignment.SampleLineGroups.Count > 0 And oBaseline.BaselineRegions.Count > 0 Then
                For Each oSampleLineGroup As AeccLandLib.AeccSampleLineGroup In oBaseline.Alignment.SampleLineGroups
                    If .mSampleLineGroups.ContainsKey(oSampleLineGroup.Name) = False Then
                        Dim oSampleLineGroupInfo As CSampleLineGroup
                        oSampleLineGroupInfo = New CSampleLineGroup
                        oSampleLineGroupInfo.mObject = oSampleLineGroup
                        .mSampleLineGroups.Add(oSampleLineGroup.Name, oSampleLineGroupInfo)
                    End If
                Next
            End If

            Dim oSubAssembly As AeccRoadLib.AeccSubassembly
            For Each oSubAssembly In ReportApplication.AeccXRoadwayDocument.Subassemblies
                Dim oRoadwayLinks As AeccRoadLib.AeccRoadwayLinks
                oRoadwayLinks = oSubAssembly.Links
                Dim oRoadwayLink As AeccRoadLib.AeccRoadwayLink
                For Each oRoadwayLink In oRoadwayLinks
                    Dim iIndex As Integer
                    For iIndex = 0 To oRoadwayLink.RoadwayCodes.Count - 1
                        Dim sCodeName As String
                        sCodeName = oRoadwayLink.RoadwayCodes.Item(iIndex)
                        If .mLinkCodesGroups.ContainsKey(sCodeName) = False Then
                            Dim oLinkCodesInfo As CLinkCodes
                            oLinkCodesInfo = New CLinkCodes
                            oLinkCodesInfo.mCodeName = sCodeName
                            .mLinkCodesGroups.Add(sCodeName, oLinkCodesInfo)
                        End If
                    Next iIndex
                Next
            Next
        End With
    End Sub

    Private Sub Combo_Align_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Align.SelectedIndexChanged
        'Dim oCorridorInfo As CCorridor
        'Dim oBaselineInfo As CBaseline
        Dim oStationSetting As AeccLandLib.AeccSettingsStation = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.StationSettings

        ctlStartStation.TextBox.Enabled = False
        ctlEndStation.TextBox.Enabled = False
        NumericUpDown_StationInc.Enabled = False
        ctlStartStation.Alignment = Nothing
        ctlEndStation.Alignment = Nothing

        ctlStartStation.EquationStation = ""
        ctlEndStation.EquationStation = ""
        'TextBox_StartStation.Text = ReportFormat.RoundVal(0.0, oStationSetting.Precision.Value, oStationSetting.Rounding.Value)
        'TextBox_EndStation.Text = ReportFormat.RoundVal(0.0, oStationSetting.Precision.Value, oStationSetting.Rounding.Value)

        If Combo_Align.Items.Count = 0 Then
            Exit Sub
        End If

        If Combo_Align.SelectedIndex < 0 Then
            Exit Sub
        End If

        'oCorridorInfo = m_Corridors.Item(Combo_Corridor.SelectedItem)
        'oBaselineInfo = oCorridorInfo.mBaselines.Item(Combo_Align.SelectedItem)
        TextBox_StartStation.Enabled = True
        TextBox_EndStation.Enabled = True
        If Not radioSurfaces.Checked And Not radioLinks.Checked Then NumericUpDown_StationInc.Enabled = True
        ctlStartStation.Alignment = m_Alignments(Combo_Align.SelectedItem)
        ctlEndStation.Alignment = ctlStartStation.Alignment
        ctlEndStation.RawStation = ReportFormat.RoundDouble(m_Alignments(Combo_Align.SelectedItem).EndingStation, oStationSetting.Precision.Value, oStationSetting.Rounding.Value)
        ctlStartStation.RawStation = ReportFormat.RoundDouble(m_Alignments(Combo_Align.SelectedItem).StartingStation, oStationSetting.Precision.Value, oStationSetting.Rounding.Value)
        ctlStartStation.EquationStation = ReportUtilities.GetStationString(m_Alignments(Combo_Align.SelectedItem), m_Alignments(Combo_Align.SelectedItem).StartingStation)
        ctlEndStation.EquationStation = ReportUtilities.GetStationString(m_Alignments(Combo_Align.SelectedItem), m_Alignments(Combo_Align.SelectedItem).EndingStation)
        If radioSurfaces.Checked Or radioLinks.Checked Then
            FillInSLGAlign()
            If radioLinks.Checked Then
                Dim oCorridorInfo As CCorridor_Emia = m_Corridors.Item(Combo_Corridor.SelectedItem)
                Dim oBaselineInfos As List(Of CBaseline) = oCorridorInfo.mBaselines.Item(Combo_Align.SelectedItem)
                If Combo_SLG.SelectedItem Is Nothing Then
                    ' just clear FeatureLineList
                    ListView_FeatureLines.Items.Clear()
                Else
                    For Each oBaselineInfo As CBaseline In oBaselineInfos
                        Dim oSampleLineGroupInfo As CSampleLineGroup = _
                                oBaselineInfo.mSampleLineGroups.Item(Combo_SLG.SelectedItem)

                        If oBaselineInfo.mIsLinkCodesFilterred = False Then
                            FilterLinkCode(oBaselineInfo, oSampleLineGroupInfo)
                        End If
                    Next
                    FillFeatureLineList(Nothing, oBaselineInfos)
                End If
            End If
        End If
    End Sub

    Private Sub Combo_Corridor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Corridor.SelectedIndexChanged
        If Combo_Corridor.SelectedIndex = Me.LastSelectedCorridorIndex Then
            Return
        Else
            If CheckAndFillFLL() = False Then
                Combo_Corridor.SelectedIndex = Me.LastSelectedCorridorIndex
            Else
                Me.LastSelectedCorridorIndex = Combo_Corridor.SelectedIndex
            End If
            FillInSLG()
        End If
    End Sub

    Private Sub FillInSLG()
        Combo_SLG.Items.Clear()

        Dim oCorridorInfo As CCorridor_Emia
        'Dim oBaselineInfo As CBaseline

        oCorridorInfo = m_Corridors.Item(Combo_Corridor.SelectedItem)

        Dim oBaselineInfos As List(Of CBaseline)
        oBaselineInfos = oCorridorInfo.mBaselines.Item(Combo_Align.SelectedItem)
        For Each oBaselineInfo As CBaseline In oBaselineInfos
            For Each sSLG As String In oBaselineInfo.mSampleLineGroups.Keys
                Combo_SLG.Items.Add(sSLG)
            Next

            If oBaselineInfo.mSampleLineGroups.Keys.Count > 0 Then
                Combo_SLG.SelectedIndex = 0
            End If
        Next
    End Sub

    Private Sub FillInSLGAlign()
        Combo_SLG.Items.Clear()
        If m_Alignments(Combo_Align.SelectedItem).SampleLineGroups.Count() > 0 Then
            For Each sSLG As Land.AeccSampleLineGroup In m_Alignments(Combo_Align.SelectedItem).SampleLineGroups
                Combo_SLG.Items.Add(sSLG.Name)
            Next
        End If

        If m_Alignments(Combo_Align.SelectedItem).SampleLineGroups.Count > 0 Then
            Combo_SLG.SelectedIndex = 0
        End If
    End Sub


    Private Sub FillFeatureLineList(ByVal oCorridor As Roadway.IAeccCorridor, _
                                    ByVal oBaselineInfos As List(Of CBaseline))

        If radioPoints.Checked Then
            ListView_FeatureLines.Items.Clear()
            For Each code As String In oCorridor.FeatureLineCodeInfos.CodeNames
                Dim li As New ListViewItem
                li.SubItems.Add(code)
                li.Checked = True
                ListView_FeatureLines.Items.Add(li)
            Next
        ElseIf radioLinks.Checked Then
            ListView_FeatureLines.Items.Clear()
            For Each oBaselineInfo As CBaseline In oBaselineInfos
                If Not Combo_SLG.SelectedItem Is Nothing Then
                    For Each sLinkCodeName As String In oBaselineInfo.mLinkCodesGroups.Keys
                        If Not ListView_FeatureLines.Items.ContainsKey(sLinkCodeName) Then
                            Dim li As ListViewItem
                            li = ListView_FeatureLines.Items.Add(sLinkCodeName, String.Empty, String.Empty)
                            li.Checked = True
                            li.SubItems.Add(sLinkCodeName)
                        End If
                    Next
                End If
            Next
        ElseIf radioSurfaces.Checked Then
            ListView_FeatureLines.Items.Clear()
            For Each oSurface As AeccLandLib.AeccSurface In ReportApplication.AeccXDocument.Surfaces
                If UBound(oSurface.Points) > 2 Then 'show only surfaces with more than 2 points
                    If (oSurface.Type <> Land.AeccSurfaceType.aecckGridVolumeSurface) Then
                        Dim li As New ListViewItem
                        li.SubItems.Add(oSurface.Name)
                        li.Checked = True
                        ListView_FeatureLines.Items.Add(li)
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub BtnDone_Clicked() Handles Button_Done.Click
        Close()
    End Sub

    Private Sub BtnHelp_Click() Handles Button_Help.Click
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_CORRIDOR_FEATURE_LINE_CROSS_SECTION_REPORT)
    End Sub

    Private Sub ReportForm_HelpRequested(ByVal sender As System.Object, ByVal hlpevent As System.Windows.Forms.HelpEventArgs) Handles MyBase.HelpRequested
        BtnHelp_Click()
    End Sub

    Private Sub ButtonExecute_Click() Handles Button_CreateReport.Click
        If ListView_FeatureLines.CheckedItems.Count = 0 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorCrossSection_Msg_NoFeatureLinesSelected)
            Exit Sub
        End If

        'If Not CheckBox_ASCII.Checked And Not CheckBox_HTML.Checked And Not CheckBox_XLS.Checked Then
        '    ReportUtilities.ACADMsgBox(LocalizedRes.Alignment_Msg_SelectOneCBFirst)
        '    Exit Sub
        'End If

        If radioPoints.Checked Then
            ' there is at least one output check box checked - ask user about rebuilding corridor
            Dim inStartStation As Double = ctlStartStation.RawStation
            Dim inEndStation As Double = ctlEndStation.RawStation
            Dim oCorridorInfo As CCorridor_Emia = m_Corridors.Item(Combo_Corridor.SelectedItem)
            Dim oCorridor As AeccRoadLib.IAeccCorridor = oCorridorInfo.mObject

            Dim oBaselineInfos As List(Of CBaseline)
            oBaselineInfos = oCorridorInfo.mBaselines.Item(Combo_Align.SelectedItem)
            For Each oBaselineInfo As CBaseline In oBaselineInfos
                If CheckCorridorChainage(inStartStation, inEndStation, oBaselineInfo.mObject) Then
                    RebuildCorridor(oCorridor, inStartStation, inEndStation, oBaselineInfo.mObject)
                End If
            Next
        ElseIf Combo_SLG.SelectedItem Is Nothing Then 'radioLinks or radioSurfaces Checked
            'no sample line groups error
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorCrossSection_Msg_NoSLG)
            Exit Sub
        End If

        Try
            Dim tempFile As String = ReportConverter.GetRandomTempHtmlPath()
            If radioPoints.Checked Then
                WriteReportHTML_Points(tempFile)
            ElseIf radioLinks.Checked Then
                WriteReportHTML_Links(tempFile)
            ElseIf radioSurfaces.Checked Then
                WriteReportHTML_Surfaces(tempFile)
            End If

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

    Private Sub WriteReportHTML_Points(ByVal fileName As String)
        Dim oReportHtml As New ReportWriter(fileName)
        '<html>
        oReportHtml.HtmlBegin()
        oReportHtml.RenderHeader(LocalizedRes.CorridorCrossSection_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.CorridorCrossSection_Html_Title, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        Dim inStartStation As Double
        Dim inEndStation As Double
        Dim oCorridorInfo As CCorridor_Emia = m_Corridors.Item(Combo_Corridor.SelectedItem)
        Dim oCorridor As AeccRoadLib.IAeccCorridor = oCorridorInfo.mObject
        Dim oStationSetting As AeccLandLib.AeccSettingsStation = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.StationSettings
        Dim str As String
        Dim oBaselineInfos As List(Of CBaseline)

        oBaselineInfos = oCorridorInfo.mBaselines.Item(Combo_Align.SelectedItem)

        ctlProgressBar.ProgressBegin(oBaselineInfos.Count * ListView_FeatureLines.CheckedItems.Count)

        For Each oBaselineInfo As CBaseline In oBaselineInfos
            inStartStation = ctlStartStation.RawStation
            inEndStation = ctlEndStation.RawStation

            Dim oAlignment As Land.AeccAlignment = oBaselineInfo.mObject.Alignment
            If Math.Abs(inStartStation - oAlignment.StartingStation) < 0.01 Then
                inStartStation = oAlignment.StartingStation
            End If
            If Math.Abs(inEndStation - oAlignment.EndingStation) < 0.01 Then
                inEndStation = oAlignment.EndingStation
            End If

            If oBaselineInfo.mObject.BaselineRegions.Count = 0 Then
                Continue For
            End If
            Dim region As Autodesk.AECC.Interop.Roadway.AeccBaselineRegion
            region = oBaselineInfo.mObject.BaselineRegions(0)
            Dim minStation As Double
            Dim maxStation As Double
            minStation = region.StartStation
            maxStation = region.EndStation
            For Each oBaselineRegion As Autodesk.AECC.Interop.Roadway.AeccBaselineRegion In oBaselineInfo.mObject.BaselineRegions
                If minStation > oBaselineRegion.StartStation Then
                    minStation = oBaselineRegion.StartStation
                End If
                If maxStation < oBaselineRegion.EndStation Then
                    maxStation = oBaselineRegion.EndStation
                End If
            Next

            If inStartStation < minStation Then
                inStartStation = minStation
            End If
            If inEndStation > maxStation Then
                inEndStation = maxStation
            End If


            oReportHtml.RenderHr()

            ' Alignment and Corridor common info
            ' render corridor name
            str = LocalizedRes.CorridorCrossSection_Html_CorridorName
            str &= " " & oCorridor.Name
            oReportHtml.RenderLine(str)

            'render alignment description
            str = LocalizedRes.CorridorCrossSection_Html_AlignmentDescription
            str &= " " & oAlignment.Description
            oReportHtml.RenderLine(str)

            'render alignment name
            str = LocalizedRes.CorridorCrossSection_Html_AlignmentNameLabel
            str &= " " & oAlignment.Name
            oReportHtml.RenderLine(str)

            '
            str = LocalizedRes.CorridorCrossSection_Html_Alignment_StaRange
            str &= " " & LocalizedRes.CorridorCrossSection_Html_Alignment_StaStart
            str &= " " & ReportUtilities.GetStationStringWithDerived(oAlignment, inStartStation)
            str &= LocalizedRes.CorridorCrossSection_Html_Alignment_StaEnd
            str &= " " & ReportUtilities.GetStationStringWithDerived(oAlignment, inEndStation)
            oReportHtml.RenderLine(str)
            '<hr>
            oReportHtml.RenderHr()

            Dim displayEmptyCrown As Object = Nothing

            For Each item As ListViewItem In ListView_FeatureLines.CheckedItems
                AppendReportHTML_Points(oCorridor, oAlignment, item.SubItems(INDEX_CODES).Text, _
                    inStartStation, inEndStation, _
                    oReportHtml, oBaselineInfo.mObject, _
                    displayEmptyCrown)
                ctlProgressBar.PerformStep()
            Next
        Next

        '</body>
        oReportHtml.BodyEnd()

        '</html>
        oReportHtml.HtmlEnd()
    End Sub

    ' check corridor at chainage interval - some station doesn't exist and therefore rebuilding of corridor is necessary (true - does not exist, rebuild it)
    Private Function CheckCorridorChainage(ByVal stationStart As Double, _
                                           ByVal stationEnd As Double, _
                                           ByVal oBaseLine As AeccRoadLib.IAeccBaseline) As Boolean
        Dim dStation As Double = stationStart
        ' get baseline stations
        Dim adStations() As Double = oBaseLine.GetSortedStations()

        If StationInterval <= 0.0 Then
            CheckCorridorChainage = False
            Exit Function
        End If

        For Each dStat As Double In adStations
            If dStat < dStation Then
                Continue For ' leave out additional stations, that already existed in the drawing before
            End If
            If dStat <> dStation Then
                CheckCorridorChainage = True ' not all the stations are contained in the SortedStations array : adStations
                Exit Function
            End If
            If dStation = stationEnd Then
                Exit For ' leave out the rest of stations, after end station
            End If
            dStation += StationInterval
            If dStation > stationEnd Then
                dStation = stationEnd
            End If
        Next
        CheckCorridorChainage = False
    End Function

    ' ask user about rebuilding the corridor
    Private Sub RebuildCorridor(ByRef oCorridor As AeccRoadLib.IAeccCorridor, _
                            ByVal stationStart As Double, _
                            ByVal stationEnd As Double, _
                            ByVal oBaseLine As AeccRoadLib.IAeccBaseline)
        Dim str As String
        Dim dStation As Double = stationStart
        str = LocalizedRes.CorridorCrossSection_Msg_CorridorRebuild
        If MsgBox(str, MsgBoxStyle.YesNo, Nothing) = MsgBoxResult.Yes Then

            If StationInterval <= 0.0 Then Exit Sub
            ' set baseline stations

            While dStation < stationEnd
                Try
                    oBaseLine.AddStation(dStation)
                Catch
                End Try
                dStation += StationInterval
            End While
            If stationStart < stationEnd Then
                Try
                    oBaseLine.AddStation(stationEnd)
                Catch
                End Try
            End If
            oCorridor.Rebuild()
        End If
    End Sub
    ' Append one item's report to html file
    Private Sub AppendReportHTML_Points(ByVal oCorridor As AeccRoadLib.IAeccCorridor, _
                            ByVal oAlignment As Land.AeccAlignment, _
                            ByVal sFeatureLineCodeName As String, _
                            ByVal stationStart As Double, _
                            ByVal stationEnd As Double, _
                            ByVal oHtmlWriter As ReportWriter, _
                            ByVal oBaseLine As AeccRoadLib.IAeccBaseline, _
                            ByRef displayEmptyCrown As Object)

        If oHtmlWriter Is Nothing Then
            Exit Sub
        End If


        Dim XYZ_Stations As List(Of Double) = New List(Of Double)
        Dim XYZ_Lefts As List(Of Object) = New List(Of Object)
        Dim XYZ_Rights As List(Of Object) = New List(Of Object)
        Dim XYZ_Crowns As List(Of Object) = New List(Of Object)

        Dim dStation As Double = stationStart
        While dStation <= stationEnd
            Dim XYZ_Left As Object = Nothing
            Dim XYZ_Right As Object = Nothing
            Dim XYZ_Crown As Object = Nothing
            If AppendReportHTML_Points_Sub(sFeatureLineCodeName, dStation, oBaseLine, XYZ_Left, XYZ_Right, XYZ_Crown) Then
                XYZ_Stations.Add(dStation)
                XYZ_Lefts.Add(XYZ_Left)
                XYZ_Rights.Add(XYZ_Right)
                XYZ_Crowns.Add(XYZ_Crown)
            End If
            If (dStation >= stationEnd) Then
                Exit While
            End If
            dStation += StationInterval
            If (dStation >= stationEnd) Then
                dStation = stationEnd           'for the last station write row with data
            End If
        End While

        Dim crownCount = 0
        For Each crown As Object In XYZ_Crowns
            If Not crown Is Nothing Then
                crownCount = crownCount + 1
            End If
        Next

        If (displayEmptyCrown Is Nothing And crownCount = 0) Then
            If MessageBox.Show(LocalizedRes.CrossSection_Msg_NoCrownData, _
                               LocalizedRes.WarningDialog_Title, _
                               MessageBoxButtons.YesNo, _
                               MessageBoxIcon.Warning, _
                               MessageBoxDefaultButton.Button2) = System.Windows.Forms.DialogResult.Yes Then
                displayEmptyCrown = True
            Else
                displayEmptyCrown = False
            End If
        End If

        Dim displayCrown As Boolean = crownCount <> 0 Or displayEmptyCrown

        Dim oStationSetting As AeccLandLib.AeccSettingsStation = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.StationSettings

        oHtmlWriter.TableBegin("1")
        'title - 2 rows
        oHtmlWriter.TrBegin()
        oHtmlWriter.RenderTd("&nbsp;")
        oHtmlWriter.RenderTd(sFeatureLineCodeName + " " + LocalizedRes.CorridorCrossSection_Html_TblTitle_Left, True, "center", , , , , "3")
        If displayCrown = True Then
            oHtmlWriter.RenderTd(LocalizedRes.CorridorCrossSection_Html_TblTitle_Crown, True, "center", , , , , "3")
        End If
        oHtmlWriter.RenderTd(sFeatureLineCodeName + " " + LocalizedRes.CorridorCrossSection_Html_TblTitle_Right, True, "center", , , , , "3")
        oHtmlWriter.TrEnd()

        oHtmlWriter.TrBegin()
        oHtmlWriter.RenderTd(LocalizedRes.CorridorCrossSection_Html_Chainage, True)
        oHtmlWriter.RenderTd(LocalizedRes.CorridorCrossSection_Html_TblTitleEasting, True)
        oHtmlWriter.RenderTd(LocalizedRes.CorridorCrossSection_Html_TblTitleNorthing, True)
        oHtmlWriter.RenderTd(LocalizedRes.CorridorCrossSection_Html_TblTitleLevel, True)
        If displayCrown = True Then
            oHtmlWriter.RenderTd(LocalizedRes.CorridorCrossSection_Html_TblTitleEasting, True)
            oHtmlWriter.RenderTd(LocalizedRes.CorridorCrossSection_Html_TblTitleNorthing, True)
            oHtmlWriter.RenderTd(LocalizedRes.CorridorCrossSection_Html_TblTitleLevel, True)
        End If
        oHtmlWriter.RenderTd(LocalizedRes.CorridorCrossSection_Html_TblTitleEasting, True)
        oHtmlWriter.RenderTd(LocalizedRes.CorridorCrossSection_Html_TblTitleNorthing, True)
        oHtmlWriter.RenderTd(LocalizedRes.CorridorCrossSection_Html_TblTitleLevel, True)
        oHtmlWriter.TrEnd()

        Dim index As Integer
        For index = 0 To XYZ_Stations.Count - 1
            dStation = XYZ_Stations(index)
            If Not XYZ_Lefts(index) Is Nothing Then
                oHtmlWriter.TrBegin()
                oHtmlWriter.RenderTd(ReportUtilities.GetStationStringWithDerived(oAlignment, dStation))
                oHtmlWriter.RenderTd(ReportFormat.RoundVal(CDbl(XYZ_Lefts(index)(0)), oStationSetting.Precision.Value, oStationSetting.Rounding.Value).ToString())
                oHtmlWriter.RenderTd(ReportFormat.RoundVal(CDbl(XYZ_Lefts(index)(1)), oStationSetting.Precision.Value, oStationSetting.Rounding.Value).ToString())
                oHtmlWriter.RenderTd(ReportFormat.RoundVal(CDbl(XYZ_Lefts(index)(2)), oStationSetting.Precision.Value, oStationSetting.Rounding.Value).ToString())
            Else
                oHtmlWriter.TrBegin()
                oHtmlWriter.RenderTd(ReportUtilities.GetStationStringWithDerived(oAlignment, dStation))
                oHtmlWriter.RenderTd("&nbsp;")
                oHtmlWriter.RenderTd("&nbsp;")
                oHtmlWriter.RenderTd("&nbsp;")
            End If
            If displayCrown = True Then
                If Not XYZ_Crowns(index) Is Nothing Then
                    oHtmlWriter.RenderTd(ReportFormat.RoundVal(CDbl(XYZ_Crowns(index)(0)), oStationSetting.Precision.Value, oStationSetting.Rounding.Value).ToString())
                    oHtmlWriter.RenderTd(ReportFormat.RoundVal(CDbl(XYZ_Crowns(index)(1)), oStationSetting.Precision.Value, oStationSetting.Rounding.Value).ToString())
                    oHtmlWriter.RenderTd(ReportFormat.RoundVal(CDbl(XYZ_Crowns(index)(2)), oStationSetting.Precision.Value, oStationSetting.Rounding.Value).ToString())
                Else
                    oHtmlWriter.RenderTd("&nbsp;")
                    oHtmlWriter.RenderTd("&nbsp;")
                    oHtmlWriter.RenderTd("&nbsp;")
                End If
            End If
            If Not XYZ_Rights(index) Is Nothing Then
                oHtmlWriter.RenderTd(ReportFormat.RoundVal(CDbl(XYZ_Rights(index)(0)), oStationSetting.Precision.Value, oStationSetting.Rounding.Value).ToString())
                oHtmlWriter.RenderTd(ReportFormat.RoundVal(CDbl(XYZ_Rights(index)(1)), oStationSetting.Precision.Value, oStationSetting.Rounding.Value).ToString())
                oHtmlWriter.RenderTd(ReportFormat.RoundVal(CDbl(XYZ_Rights(index)(2)), oStationSetting.Precision.Value, oStationSetting.Rounding.Value).ToString())
                oHtmlWriter.TrEnd()
            Else
                oHtmlWriter.RenderTd("&nbsp;")
                oHtmlWriter.RenderTd("&nbsp;")
                oHtmlWriter.RenderTd("&nbsp;")
                oHtmlWriter.TrEnd()
            End If
        Next

        ctlProgressBar.PerformStep()
        oHtmlWriter.TableEnd()
        oHtmlWriter.RenderBr()
    End Sub
    Private Function AppendReportHTML_Points_Sub( _
                        ByVal sFeatureLineCodeName As String, _
                        ByVal dStation As Double, _
                        ByVal oBaseLine As AeccRoadLib.IAeccBaseline, _
                        ByRef XYZ_Left As Object, _
                        ByRef XYZ_Right As Object, _
                        ByRef XYZ_Crown As Object) As Boolean

        Dim appliedAssembly As Roadway.IAeccAppliedAssembly
        Dim calcPoints As Roadway.AeccCalculatedPoints
        Dim calcPointsCrown As Roadway.AeccCalculatedPoints
        Dim oStation As Object

        Try
            oStation = dStation
            appliedAssembly = oBaseLine.AppliedAssembly(oStation)
            calcPoints = appliedAssembly.GetPointsByCode(sFeatureLineCodeName)
            calcPointsCrown = appliedAssembly.GetPointsByCode("Crown")
        Catch
            Return False
        End Try

        For Each calcPoint As Roadway.AeccCalculatedPoint In calcPoints 'at most 2 points expected
            Dim oSoe As Object = calcPoint.GetStationOffsetElevationToBaseline()
            Dim XYZ As Object = oBaseLine.StationOffsetElevationToXYZ(oSoe)

            If oSoe(1) < 0 Then
                XYZ_Left = XYZ
            End If
            If oSoe(1) > 0 Then
                XYZ_Right = XYZ
            End If
        Next
        For Each calcPointCrown As Roadway.AeccCalculatedPoint In calcPointsCrown
            Dim oSoe As Object = calcPointCrown.GetStationOffsetElevationToBaseline()
            Dim XYZ As Object = oBaseLine.StationOffsetElevationToXYZ(oSoe)

            XYZ_Crown = XYZ 'take just the first crown point, if more than one points
            Exit For
        Next

        AppendReportHTML_Points_Sub = True
    End Function

    Private Sub WriteReportHTML_Links(ByVal fileName As String)
        Dim oReportHtml As New ReportWriter(fileName)
        '<html>
        oReportHtml.HtmlBegin()
        oReportHtml.RenderHeader(LocalizedRes.CorridorCrossSection_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.CorridorCrossSection_Html_Title, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        Dim inStartStation As Double
        Dim inEndStation As Double
        Dim oCorridorInfo As CCorridor_Emia
        Dim oBaselineInfos As List(Of CBaseline)
        Dim oSampleLineGroupInfo As CSampleLineGroup
        Dim oLinkCodesInfo As CLinkCodes

        oCorridorInfo = m_Corridors.Item(Combo_Corridor.SelectedItem)
        oBaselineInfos = oCorridorInfo.mBaselines.Item(Combo_Align.SelectedItem)

        'ctlProgressBar.ProgressBegin(ListView_FeatureLines.CheckedItems.Count() * EXTRACT_STEPS)
        ctlProgressBar.ProgressBegin(oBaselineInfos.Count * ListView_FeatureLines.CheckedItems.Count())

        For Each oBaselineInfo As CBaseline In oBaselineInfos
            inStartStation = ctlStartStation.RawStation
            inEndStation = ctlEndStation.RawStation
            If oBaselineInfo.mObject.BaselineRegions.Count = 0 Then
                Continue For
            End If
            Dim region As Autodesk.AECC.Interop.Roadway.AeccBaselineRegion
            region = oBaselineInfo.mObject.BaselineRegions(0)
            Dim minStation As Double
            Dim maxStation As Double
            minStation = region.StartStation
            maxStation = region.EndStation
            For Each oBaselineRegion As Autodesk.AECC.Interop.Roadway.AeccBaselineRegion In oBaselineInfo.mObject.BaselineRegions
                If minStation > oBaselineRegion.StartStation Then
                    minStation = oBaselineRegion.StartStation
                End If
                If maxStation < oBaselineRegion.EndStation Then
                    maxStation = oBaselineRegion.EndStation
                End If
            Next

            If inStartStation < minStation Then
                inStartStation = minStation
            End If
            If inEndStation > maxStation Then
                inEndStation = maxStation
            End If

            For Each item As ListViewItem In ListView_FeatureLines.CheckedItems
                If Not oBaselineInfo.mSampleLineGroups.ContainsKey(Combo_SLG.SelectedItem) Then
                    Continue For
                End If
                If Not oBaselineInfo.mLinkCodesGroups.ContainsKey(item.SubItems(INDEX_CODES).Text) Then
                    Continue For
                End If

                oSampleLineGroupInfo = oBaselineInfo.mSampleLineGroups.Item(Combo_SLG.SelectedItem)
                oLinkCodesInfo = oBaselineInfo.mLinkCodesGroups.Item(item.SubItems(INDEX_CODES).Text)


                AppendReportHTML_Links(oCorridorInfo.mObject, oBaselineInfo.mObject, _
                    oSampleLineGroupInfo.mObject, inStartStation, inEndStation, _
                    oReportHtml, oLinkCodesInfo.mCodeName)
                ctlProgressBar.PerformStep()
            Next
        Next

        '</body>
        oReportHtml.BodyEnd()

        '</html>
        oReportHtml.HtmlEnd()
    End Sub

    Private Sub WriteReportHTML_Surfaces(ByVal fileName As String)
        Dim oReportHtml As New ReportWriter(fileName)
        '<html>
        oReportHtml.HtmlBegin()
        oReportHtml.RenderHeader(LocalizedRes.CorridorCrossSection_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.CorridorCrossSection_Html_Title, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        Dim inStartStation As Double
        Dim inEndStation As Double
        'Dim oCorridorInfo As CCorridor
        'Dim oBaselineInfo As CBaseline
        'Dim oSampleLineGroupInfo As CSampleLineGroup
        Dim oSampleLineGroup As Land.AeccSampleLineGroup
        Dim oSurface As AeccLandLib.AeccSurface

        ctlProgressBar.ProgressBegin(ListView_FeatureLines.CheckedItems.Count() * EXTRACT_STEPS)

        For Each item As ListViewItem In ListView_FeatureLines.CheckedItems
            inStartStation = ctlStartStation.RawStation
            inEndStation = ctlEndStation.RawStation
            ' oCorridorInfo = m_Corridors.Item(Combo_Corridor.SelectedItem)
            'oBaselineInfo = oCorridorInfo.mBaselines.Item(Combo_Align.SelectedItem)
            oSampleLineGroup = m_Alignments(Combo_Align.SelectedItem).SampleLineGroups.Item(Combo_SLG.SelectedItem)
            oSurface = m_Surfaces.Item(item.SubItems(INDEX_CODES).Text)
            AppendReportHTML_Surfaces(m_Alignments(Combo_Align.SelectedItem), _
                oSampleLineGroup, inStartStation, inEndStation, _
                oReportHtml, oSurface)
        Next

        '</body>
        oReportHtml.BodyEnd()

        '</html>
        oReportHtml.HtmlEnd()
    End Sub

    ' Append one item's report to html file
    Private Sub AppendReportHTML_Surfaces(ByVal oBaseAlignment As AeccLandLib.AeccAlignment, _
                            ByVal oSampleLineGroup As AeccLandLib.IAeccSampleLineGroup, _
                            ByVal stationStart As Double, _
                            ByVal stationEnd As Double, _
                            ByVal oHtmlWriter As ReportWriter, _
                            ByVal oSurface As AeccLandLib.AeccSurface)
        Dim str As String

        oHtmlWriter.RenderHr()

        '' Corridor info
        'str = LocalizedRes.GetString("CorridorSlopeStake_Html_CorridorName")
        'str += " " + oCorridor.Name
        'oHtmlWriter.RenderLine(str)

        'str = LocalizedRes.GetString("CorridorSlopeStake_Html_Description")
        'str += " " + oCorridor.Description
        'oHtmlWriter.RenderLine(str)

        str = LocalizedRes.CorridorSlopeStake_Html_BaseAlign
        str += " " + oBaseAlignment.Name
        oHtmlWriter.RenderLine(str)

        str = LocalizedRes.CorridorSlopeStake_Html_SampleLineGroupName
        str += " " + oSampleLineGroup.Name
        oHtmlWriter.RenderLine(str)

        str = LocalizedRes.CorridorSlopeStake_Html_Surface
        str += " " + oSurface.Name
        oHtmlWriter.RenderLine(str)

        str = LocalizedRes.CorridorSlopeStake_Html_Range
        str = str.Replace("&StartStation&", ReportUtilities.GetStationStringWithDerived(oBaseAlignment, stationStart))
        str = str.Replace("&EndStation&", ReportUtilities.GetStationStringWithDerived(oBaseAlignment, stationEnd))
        oHtmlWriter.RenderLine(str)

        '    extract data
        CorridorCrossSection_ExtractData.ExtractData(oBaseAlignment, _
            oSampleLineGroup, oSurface, stationStart, stationEnd, ProgressBar_Creating, EXTRACT_STEPS)

        For Each key As Double In CorridorCrossSection_ExtractData.SlopeStakeData.Keys
            AppendSingleTable_Surfaces(oBaseAlignment, key, oHtmlWriter)
        Next

        oHtmlWriter.RenderBr()
    End Sub

    Private Sub AppendSingleTable_Surfaces(ByVal baseAlignment As AeccLandLib.IAeccAlignment, _
                                           ByVal dStation As Double, ByVal oHtmlWriter As ReportWriter)
        Dim str As String
        str = baseAlignment.Name ' & "<br/>"
        oHtmlWriter.RenderLine(str, True)

        str = LocalizedRes.CorridorSlopeStake_Html_Station_Emia
        str &= " " & ReportUtilities.GetStationStringWithDerived(baseAlignment, dStation)
        oHtmlWriter.RenderLine(str, True)

        'table begin
        oHtmlWriter.TableBegin("1")

        Dim sData As CorridorCS_SectionData
        Try
            sData = CorridorCrossSection_ExtractData.SlopeStakeData.Item(dStation)
        Catch ex As Exception
            'can't be here
            Diagnostics.Debug.Assert(False, ex.Message)
            Exit Sub
        End Try

        'table size depence on count 
        Dim PtCount As Integer
        PtCount = sData.Datas.Count
        If PtCount < 1 Then
            Exit Sub
        End If

        'get enumerator of data dictionary
        Dim Keys(PtCount) As Double
        sData.Datas.Keys.CopyTo(Keys, 0)

        Dim numCols As Integer = sData.Datas.Keys.Count()

        'For line As Integer = 0 To lineNum - 1
        'as we mention before, we have 3 rows + header row
        For row As Integer = 0 To 3
            'there're numCols + 1 cells in one line include numCols points + "header cell"
            Dim DataArr(0 To numCols) As String

            Select Case row
                Case 0 'first row
                    DataArr(0) = LocalizedRes.CorridorSlopeStake_Html_Offset
                Case 1 '2ed row
                    DataArr(0) = LocalizedRes.CorridorSlopeStake_Html_Level
                    ' Case 2 '3rd row
                    '    DataArr(0) = LocalizedRes.GetString("CorridorSlopeStake_Html_Slope")
                Case 2 '4th row
                    DataArr(0) = LocalizedRes.CorridorSlopeStake_Html_Easting
                Case 3 '5th row
                    DataArr(0) = LocalizedRes.CorridorSlopeStake_Html_Northing
                Case Else 'impossible here
                    DataArr(0) = ""
            End Select
            'End If

            'fill cell column data, total numCols columns for surface points per line
            For col As Integer = 0 To numCols - 1
                Dim nowPtIndex As Integer
                nowPtIndex = col '10 * Line + col

                If nowPtIndex < PtCount Then
                    fillInPointColumn(row, sData.Datas.Item(Keys(nowPtIndex)), DataArr(col + 1))
                Else
                    DataArr(col + 1) = ""
                End If

            Next

            'data fill finished, write it
            WriteOneTableRow(DataArr, oHtmlWriter)
            'Next

        Next

        oHtmlWriter.TableEnd()
    End Sub

    ' Append one item's report to html file
    Private Sub AppendReportHTML_Links(ByVal oCorridor As AeccRoadLib.IAeccCorridor, _
                            ByVal oBaseline As AeccRoadLib.IAeccBaseline, _
                            ByVal oSampleLineGroup As AeccLandLib.IAeccSampleLineGroup, _
                            ByVal stationStart As Double, _
                            ByVal stationEnd As Double, _
                            ByVal oHtmlWriter As ReportWriter, _
                            ByVal sCodeName As String)
        Dim str As String

        oHtmlWriter.RenderHr()

        ' Corridor info
        str = LocalizedRes.CorridorSlopeStake_Html_CorridorName_Emia
        str += " " + oCorridor.Name
        oHtmlWriter.RenderLine(str)

        str = LocalizedRes.CorridorSlopeStake_Html_Description
        str += " " + oCorridor.Description
        oHtmlWriter.RenderLine(str)

        Dim oBaseAlignment As AeccLandLib.AeccAlignment
        oBaseAlignment = oBaseline.Alignment
        str = LocalizedRes.CorridorSlopeStake_Html_BaseAlign
        str += " " + oBaseAlignment.Name
        oHtmlWriter.RenderLine(str)

        str = LocalizedRes.CorridorSlopeStake_Html_SampleLineGroupName
        str += " " + oSampleLineGroup.Name
        oHtmlWriter.RenderLine(str)

        str = LocalizedRes.CorridorSlopeStake_Html_Link
        str += " " + sCodeName
        oHtmlWriter.RenderLine(str)

        str = LocalizedRes.CorridorSlopeStake_Html_Range
        str = str.Replace("&StartStation&", ReportUtilities.GetStationStringWithDerived(oBaseAlignment, stationStart))
        str = str.Replace("&EndStation&", ReportUtilities.GetStationStringWithDerived(oBaseAlignment, stationEnd))
        oHtmlWriter.RenderLine(str)

        '    extract data
        CorridorSlopeStake_ExtractData_Emia.ExtractData(oBaseline, _
            oSampleLineGroup, stationStart, stationEnd, sCodeName, ProgressBar_Creating, EXTRACT_STEPS)

        ctlProgressBar.PerformStep()

        For Each key As Double In CorridorSlopeStake_ExtractData_Emia.SlopeStakeData.Keys
            AppendSingleTable_Links(oBaseAlignment, key, oHtmlWriter)
        Next

        oHtmlWriter.RenderBr()
    End Sub

    Private Sub AppendSingleTable_Links(ByVal baseAlignment As AeccLandLib.IAeccAlignment, _
                                        ByVal dStation As Double, ByVal oHtmlWriter As ReportWriter)
        Dim str As String
        str = baseAlignment.Name ' & "<br/>"
        oHtmlWriter.RenderLine(str, True)

        str = LocalizedRes.CorridorSlopeStake_Html_Station_Emia
        str &= " " & ReportUtilities.GetStationStringWithDerived(baseAlignment, dStation)
        oHtmlWriter.RenderLine(str, True)

        'table begin
        oHtmlWriter.TableBegin("1")

        Dim sData As Slope_SectionData_Emia
        Try
            sData = CorridorSlopeStake_ExtractData_Emia.SlopeStakeData.Item(dStation)
        Catch ex As Exception
            'can't be here
            Diagnostics.Debug.Assert(False, ex.Message)
            Exit Sub
        End Try

        'table size depence on max count 
        Dim maxSidePtCount As Integer
        Dim leftPtCount As Integer
        Dim rightPtCount As Integer
        leftPtCount = sData.LeftDatas.Count
        rightPtCount = sData.RightDatas.Count
        maxSidePtCount = leftPtCount
        If maxSidePtCount < rightPtCount Then
            maxSidePtCount = rightPtCount
        End If
        If maxSidePtCount < 1 Then
            Exit Sub
        End If

        'get line number, every line we write 5 code points
        '---here one line means the whole data lines, actually holds 4 "rows", they're:
        '   Codes(row1), offset(row2), elevation(row3), slope(row4)
        Dim lineNum As Integer
        lineNum = Math.Ceiling(maxSidePtCount / 5)

        'get enumerator of left and right data dictionary
        Dim leftKeys(leftPtCount) As Double
        Dim rightKeys(rightPtCount) As Double
        sData.LeftDatas.Keys.CopyTo(leftKeys, 0)
        sData.RightDatas.Keys.CopyTo(rightKeys, 0)

        For line As Integer = 0 To lineNum - 1
            'as we mention before, we have 4 rows
            For row As Integer = 0 To 5
                'there're 13 cell in one line include 2 ends, 1 middle blank, 5 left points and 5 right points
                'left side should be reserve later after filled
                Dim leftDataArr(0 To 5) As String
                Dim rightDataArr(0 To 4) As String

                'set end data at the first line
                If line = 0 Then
                    'If row = 0 Then
                    '    If Not sData.LeftEndInfo Is Nothing Then leftDataArr(5) = sData.LeftEndInfo.EndType
                    '    If Not sData.RightEndInfo Is Nothing Then rightDataArr(5) = sData.RightEndInfo.EndType
                    'ElseIf row = 1 Then
                    '    If Not sData.LeftEndInfo Is Nothing Then leftDataArr(5) = sData.LeftEndInfo.DeltaOffset
                    '    If Not sData.RightEndInfo Is Nothing Then rightDataArr(5) = sData.RightEndInfo.DeltaOffset
                    'ElseIf row = 2 Then
                    '    If Not sData.LeftEndInfo Is Nothing Then leftDataArr(5) = sData.LeftEndInfo.EndSlope
                    '    If Not sData.RightEndInfo Is Nothing Then rightDataArr(5) = sData.RightEndInfo.EndSlope
                    'End If
                    Select Case row
                        Case 0 'first row
                            leftDataArr(5) = LocalizedRes.CorridorSlopeStake_Html_Code
                        Case 1 '2ed row
                            leftDataArr(5) = LocalizedRes.CorridorSlopeStake_Html_Offset
                        Case 2 '3rd row
                            leftDataArr(5) = LocalizedRes.CorridorSlopeStake_Html_Level
                        Case 3 '4th row
                            leftDataArr(5) = LocalizedRes.CorridorSlopeStake_Html_Slope
                        Case 4 '5th row
                            leftDataArr(5) = LocalizedRes.CorridorSlopeStake_Html_Easting
                        Case 5 '6th row
                            leftDataArr(5) = LocalizedRes.CorridorSlopeStake_Html_Northing
                        Case Else 'impossible here
                            leftDataArr(5) = ""
                    End Select
                End If

                'fill cell column data, total 5 columns for code points per line
                For col As Integer = 0 To 4
                    Dim nowPtIndex As Integer
                    nowPtIndex = 5 * line + col

                    If nowPtIndex < leftPtCount Then
                        fillPointColumn(row, sData.LeftDatas.Item(leftKeys(nowPtIndex)), leftDataArr(col))
                    Else
                        leftDataArr(col) = ""
                    End If

                    If nowPtIndex < rightPtCount Then
                        fillPointColumn(row, sData.RightDatas.Item(rightKeys(nowPtIndex)), rightDataArr(col))
                    Else
                        rightDataArr(col) = ""
                    End If
                Next

                'left and right data fill finished, write them
                Dim dataArr(0 To 11) As String
                Array.Reverse(leftDataArr)
                leftDataArr.CopyTo(dataArr, 0)
                dataArr(6) = ""
                rightDataArr.CopyTo(dataArr, 7)
                WriteOneTableRow(dataArr, oHtmlWriter)
            Next

            'separate row
            Dim arrBlank(0 To 11) As String
            WriteOneTableRow(arrBlank, oHtmlWriter)
        Next

        oHtmlWriter.TableEnd()
    End Sub

    Private Sub fillPointColumn(ByVal row As Integer, _
                                    ByVal ptData As Slope_PointData_Emia, _
                                    ByRef cell As String)
        Try
            Select Case row
                Case 0 'first row
                    cell = ptData.mCodes
                Case 1 '2ed row
                    cell = ptData.mOffsetString
                Case 2 '3rd row
                    cell = ptData.mElevationString
                Case 3 '4th row
                    cell = ptData.mSlope
                Case 4 '5th row
                    cell = ptData.mEastingString
                Case 5 '6th row
                    cell = ptData.mNorthingString
                Case Else 'impossible here
                    cell = ""
            End Select
        Catch ex As Exception
            Diagnostics.Debug.Assert(False, ex.Message)
        End Try
    End Sub

    Private Sub fillInPointColumn(ByVal row As Integer, _
                                  ByVal ptData As CorridorCS_PointData, _
                                  ByRef cell As String)
        Try
            Select Case row
                Case 0 'first row
                    cell = ptData.mOffsetString
                Case 1 '2ed row
                    cell = ptData.mElevationString
                    'Case 2 '3rd row
                    'cell = ptData.mSlope
                Case 2 '3rd row
                    cell = ptData.mEastingString
                Case 3 '4th row
                    cell = ptData.mNorthingString
                Case Else 'impossible here
                    cell = ""
            End Select
        Catch ex As Exception
            Diagnostics.Debug.Assert(False, ex.Message)
        End Try
    End Sub

    Private Sub WriteOneTableRow(ByVal arrData As String(), ByVal oHtmlWriter As ReportWriter)
        Dim i As Long
        'Dim str As String
        oHtmlWriter.TrBegin()
        For i = LBound(arrData) To UBound(arrData)
            If arrData(i) = "" Then
                oHtmlWriter.RenderTd("&nbsp;")
            Else
                oHtmlWriter.RenderTd(arrData(i))
            End If
        Next i
        oHtmlWriter.TrEnd()
    End Sub

    Private Sub Button_CheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_CheckAll.Click
        For Each li As ListViewItem In ListView_FeatureLines.Items
            li.Checked = True
        Next
    End Sub

    Private Sub Button_UncheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_UncheckAll.Click
        For Each li As ListViewItem In ListView_FeatureLines.CheckedItems
            li.Checked = False
        Next
    End Sub

    Private Sub radioPoints_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radioPoints.CheckedChanged
        If radioPoints.Checked Then
            GroupBox_FeatureLines.Text = LocalizedRes.CorridorCrossSection_Form_List_FeatLns
        End If
        If radioPoints.Checked Then
            If ListView_FeatureLines.Columns.Count > 0 Then
                ListView_FeatureLines.Columns.RemoveAt(1)
                ListView_FeatureLines.Columns.Add(LocalizedRes.CorridorCrossSection_Form_ColumnTitle_PointCodes, _
                    80, HorizontalAlignment.Left)
            End If

            CheckAndFillFLL()
            'enable and disable some control's depending on selected radio button
            Label_StationInc.Enabled = True
            NumericUpDown_StationInc.Enabled = True
            Combo_Corridor.Enabled = True
            Label_SelectCorridor.Enabled = True
            Label_SelectSLG.Enabled = False
            Combo_SLG.Enabled = False
        End If
    End Sub

    Private Sub radioLinks_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radioLinks.CheckedChanged
        If radioLinks.Checked Then
            GroupBox_FeatureLines.Text = LocalizedRes.CorridorCrossSection_Form_List_Links
        End If
        If radioLinks.Checked Then
            If ListView_FeatureLines.Columns.Count > 0 Then
                ListView_FeatureLines.Columns.RemoveAt(1)
                ListView_FeatureLines.Columns.Add(LocalizedRes.CorridorCrossSection_Form_ColumnTitle_LinkCodes, _
                    80, HorizontalAlignment.Left)
            End If

            CheckAndFillFLL()
            'enable and disable some control's depending on selected radio button
            Label_StationInc.Enabled = False
            NumericUpDown_StationInc.Enabled = False
            Combo_Corridor.Enabled = True
            Label_SelectCorridor.Enabled = True
            Label_SelectSLG.Enabled = True
            Combo_SLG.Enabled = True
        End If
    End Sub

    Private Sub radioSurfaces_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radioSurfaces.CheckedChanged
        If radioSurfaces.Checked Then
            GroupBox_FeatureLines.Text = LocalizedRes.CorridorCrossSection_Form_List_Surfaces
        End If
        If radioSurfaces.Checked Then
            If ListView_FeatureLines.Columns.Count > 0 Then
                ListView_FeatureLines.Columns.RemoveAt(1)
                ListView_FeatureLines.Columns.Add(LocalizedRes.CorridorCrossSection_Form_ColumnTitle_Surfaces, _
                    80, HorizontalAlignment.Left)
            End If

            FillAlignmentComboAndFLL()
            'enable and disable some control's depending on selected radio button
            Label_StationInc.Enabled = False
            NumericUpDown_StationInc.Enabled = False
            Combo_Corridor.Enabled = False
            Label_SelectCorridor.Enabled = False
            Label_SelectSLG.Enabled = True
            Combo_SLG.Enabled = True
        End If
    End Sub

    Private Sub FillAlignmentComboAndFLL()
        ' gets called only when Surfaces radio button is selected
        Combo_Align.Items.Clear()

        'get alignments from m_Alignment
        'For Each alig As String In m_Alignments.Keys()
        '    Combo_Align.Items.Add(alig)
        'Next

        Dim corridor As CCorridor_Emia = m_Corridors.Item(Combo_Corridor.SelectedItem.ToString())
        For Each alignName As String In corridor.mBaselines.Keys
            Combo_Align.Items.Add(alignName)
        Next

        If Combo_Align.Items.Count > 0 Then
            Combo_Align.SelectedIndex = 0
        Else
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorCrossSection_Msg_NoAlignment)
            Exit Sub
        End If

        FillFeatureLineList(Nothing, Nothing)
    End Sub

    Private Function CheckAndFillFLL() As Boolean

        Dim oCorridorInfo As CCorridor_Emia
        'Dim oBaselineInfo As CBaseline

        If Combo_Corridor.Items.Count = 0 Then
            Return False
        End If

        If Combo_Corridor.SelectedIndex < 0 Then
            Return False
        End If

        oCorridorInfo = m_Corridors.Item(Combo_Corridor.SelectedItem)
        If oCorridorInfo.mBaselines.Keys.Count = 0 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorCrossSection_Msg_NoAlignment)
            Return False
        Else
            Combo_Align.Items.Clear()
            For Each sAlignmentName As String In oCorridorInfo.mBaselines.Keys
                Combo_Align.Items.Add(sAlignmentName)
            Next
            Combo_Align.SelectedIndex = 0

            Dim oBaselineInfos As List(Of CBaseline)
            oBaselineInfos = oCorridorInfo.mBaselines.Item(Combo_Align.SelectedItem)
            'For Each oBaselineInfo As CBaseline In oBaselineInfos

            'If oBaselineInfo.mIsLinkCodesFilterred = False Then
            '    FilterLinkCode(oBaselineInfo)
            'End If
            'Next
            FillFeatureLineList(oCorridorInfo.mObject, oBaselineInfos)
            Return True
        End If
    End Function

    Private Sub FilterLinkCode(ByRef oBaseLineInfo As CBaseline, _
                          ByRef oSampleLineGroupInfo As CSampleLineGroup)

        Diagnostics.Debug.Assert(oBaseLineInfo.mIsLinkCodesFilterred = False)

        Dim oBaseLine As AeccRoadLib.AeccBaseline = oBaseLineInfo.mObject
        Dim oSampleLineGroup As AeccLandLib.AeccSampleLineGroup = oSampleLineGroupInfo.mObject

        Dim tempLinkCodesGroups As SortedDictionary(Of String, CLinkCodes) = _
            New SortedDictionary(Of String, CLinkCodes)(oBaseLineInfo.mLinkCodesGroups)

        For Each sLinkCodeName As String In oBaseLineInfo.mLinkCodesGroups.Keys
            Dim bDeleteLinkCode As Boolean = True

            For Each oSampleLine As AeccLandLib.AeccSampleLine In oSampleLineGroup.SampleLines
                Try
                    Dim oLinks As AeccRoadLib.AeccCalculatedLinks = _
                        oBaseLine.AppliedAssembly(oSampleLine.Station).GetLinksByCode(sLinkCodeName)

                    If oLinks.Count > 0 Then
                        bDeleteLinkCode = False
                        Exit For
                    End If

                Catch ex As Exception
                    'Diagnostics.Debug.Assert(False, ex.Message)
                    'error, try next point
                    Continue For
                End Try
            Next

            If bDeleteLinkCode = True Then
                tempLinkCodesGroups.Remove(sLinkCodeName)
            End If
        Next

        oBaseLineInfo.mIsLinkCodesFilterred = True
        oBaseLineInfo.mLinkCodesGroups = tempLinkCodesGroups

    End Sub

    Private Sub onStartStationChanging(ByVal rawStationValue As Double, ByVal equationStationValue As String, ByRef okToChange As Boolean) Handles ctlStartStation.StationChanging
        If rawStationValue > ctlEndStation.RawStation Then
            ReportUtilities.ACADMsgBox(LocalizedRes.Msg_StartStation_Less_EndStation)
            okToChange = False
        End If
    End Sub

    Private Sub onEndStationChanging(ByVal rawStationValue As Double, ByVal equationStationValue As String, ByRef okToChange As Boolean) Handles ctlEndStation.StationChanging
        If rawStationValue < ctlStartStation.RawStation Then
            ReportUtilities.ACADMsgBox(LocalizedRes.Msg_EndStation_Greater_StartStation)
            okToChange = False
        End If
    End Sub
End Class
