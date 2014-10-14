' -----------------------------------------------------------------------------
' <copyright file="ReportForm_CorridorMilling.vb" company="Autodesk">
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
'Option Strict On

Imports AeccUiLandLib = Autodesk.AECC.Interop.UiLand
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports AeccRoadLib = Autodesk.AECC.Interop.Roadway
Imports Autodesk.AutoCAD.EditorInput
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Friend Class ReportForm_CorridorMilling

    Public Const INDEX_NAME As Integer = 0
    Public Const INDEX_ALIGNMENT As Integer = 1
    Public Const INDEX_LINEGRP As Integer = 2
    Public Const INDEX_LINK As Integer = 3
    Public Const INDEX_STASTART As Integer = 4
    Public Const INDEX_STAEND As Integer = 5
    Public Const INDEX_COL_SUM As Integer = INDEX_ALIGNMENT

    Private Const EXTRACT_STEPS As Integer = 30

    Private m_Corridors As New SortedDictionary(Of String, CCorridor)


    Private WithEvents ctlStartStation As CtrlStationTextBox
    Private WithEvents ctlEndStation As CtrlStationTextBox
    Private ctlProgressBar As CtrlProgressBar
    Private ctlSaveReport As CtrlSaveReportFile
    Private ctlWidthEditBox As CtrlDistanceTextBox
    Private ctlThicknessEditBox As CtrlDistanceTextBox
    Private m_bStnChangeByComboIndexChange As Boolean = False

    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()
        Dim rptDlg As New ReportForm_CorridorMilling
        If rptDlg.CanOpen() = True Then
            ReportUtilities.RunModalDialog(rptDlg)
        End If
    End Sub

    Private Function CanOpen() As Boolean
        Dim bNoCorridors As Boolean
        bNoCorridors = False

        Dim nAlignmentCount As Integer
        nAlignmentCount = 0

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
                        ' Defect 1281817 - Block reference that hasn't been exploded should not be recognized in report dialog
                        ' Check whether its owner - that block record is layout, in-block corridor would not be in the layout block owner
                        If Not ReportUtilities.IsCorridorInLayout(oCorridor) Then
                            Continue For
                        End If

                        Dim oCorridorInfo As New CCorridor
                        BuildCorridorInfo(oCorridor, oCorridorInfo)
                        'count alignment
                        nAlignmentCount += oCorridorInfo.mBaselines.Values.Count
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

        CanOpen = False
        If bNoCorridors = True Then
            Dim sMsg As String = String.Format(LocalizedRes.Msg_No_Corridors_In_Drawing, ReportApplication.AeccXDocument.Name)
            ReportUtilities.ACADMsgBox(sMsg)
        ElseIf nAlignmentCount < 1 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorSlopeStake_Msg_NoAlignment)
        Else
            CanOpen = True
        End If

    End Function

    Private Sub ReportForm_CorridorGrinding_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Init report settings
        InitReportSettings()

        'initial Select report Components
        For Each corridorName As String In m_Corridors.Keys
            Combo_Corridor.Items.Add(corridorName)
        Next
        'Corridor count should more then 0, because we have test it before dialog open
        If m_Corridors.Count > 0 Then
            Combo_Corridor.SelectedIndex = 0
        End If

        'init button bmp
        Dim bp As Drawing.Bitmap = My.Resources.Resources.PickEntityBmp
        bp.MakeTransparent(Color.White)
        Button_SelectCorridor.Image = bp
        Button_SelectCorridor.ImageAlign = ContentAlignment.MiddleCenter

        bp = My.Resources.Resources.AddButton
        bp.MakeTransparent(Color.White)

        bp = My.Resources.Resources.Delete
        bp.MakeTransparent(Color.White)

    End Sub

    Private Sub InitReportSettings()

        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)

        ctlStartStation = New CtrlStationTextBox
        ctlStartStation.Initialize(TextBox_StartStation)

        ctlEndStation = New CtrlStationTextBox
        ctlEndStation.Initialize(TextBox_EndStation)

        ctlWidthEditBox = New CtrlDistanceTextBox
        ctlWidthEditBox.Initialize(TextBox_Width)

        ctlThicknessEditBox = New CtrlDistanceTextBox
        ctlThicknessEditBox.Initialize(TextBox_Thickness)

        ctlSaveReport = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)

        InitToleranceSettings()
    End Sub

    Private Sub InitToleranceSettings()
        Dim dWidth As Double = 0.0
        Dim dThickness As Double = 0.0
        Dim oDistSettings As AeccLandLib.IAeccSettingsDistance
        oDistSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.DistanceSettings
        If oDistSettings.Unit.Value <> AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitFoot And _
             oDistSettings.Unit.Value <> AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitInch And _
             oDistSettings.Unit.Value <> AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitMile And _
             oDistSettings.Unit.Value <> AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitYard Then
            dWidth = 0.15
            dThickness = 0.03
        Else
            dWidth = 0.5
            dThickness = 0.1
        End If

        ctlWidthEditBox.setValue(dWidth)
        ctlThicknessEditBox.setValue(dThickness)
    End Sub


    Private Sub BuildCorridorInfo(ByRef oCorridor As AeccRoadLib.IAeccCorridor, ByRef oCorridorInfo As CCorridor)
        oCorridorInfo.mObject = oCorridor
        oCorridorInfo.mBaselines.Clear() 'dicBaselines.Clear()

        For Each oBaseline As AeccRoadLib.IAeccBaseline In oCorridor.Baselines
            If oCorridorInfo.mBaselines.ContainsKey(oBaseline.Alignment.Name) = False Then
                Dim oBaselineInfo As CBaseline
                oBaselineInfo = New CBaseline
                BuildBaselineInfo(oBaseline, oBaselineInfo)
                oCorridorInfo.mBaselines.Add(oBaseline.Alignment.Name, oBaselineInfo)
            End If
        Next
    End Sub

    Private Sub BuildBaselineInfo(ByRef oBaseline As AeccRoadLib.IAeccBaseline, ByRef oBaselineInfo As CBaseline)
        oBaselineInfo.mObject = oBaseline
        oBaselineInfo.mSampleLineGroups.Clear()

        oBaselineInfo.mLinkCodesGroups.Clear()

        If oBaseline.Alignment.SampleLineGroups.Count > 0 And oBaseline.BaselineRegions.Count > 0 Then
            For Each oSampleLineGroup As AeccLandLib.AeccSampleLineGroup In oBaseline.Alignment.SampleLineGroups
                If oBaselineInfo.mSampleLineGroups.ContainsKey(oSampleLineGroup.Name) = False Then
                    Dim oSampleLineGroupInfo As CSampleLineGroup
                    oSampleLineGroupInfo = New CSampleLineGroup
                    oSampleLineGroupInfo.mObject = oSampleLineGroup
                    oBaselineInfo.mSampleLineGroups.Add(oSampleLineGroup.Name, oSampleLineGroupInfo)
                End If
            Next

            Dim oSubAssembly As AeccRoadLib.AeccSubassembly
            For Each oSubAssembly In ReportApplication.AeccXRoadwayDocument.Subassemblies 'g_oRoadwayDocument.Subassemblies
                Dim oRoadwayLinks As AeccRoadLib.AeccRoadwayLinks
                oRoadwayLinks = oSubAssembly.Links
                Dim oRoadwayLink As AeccRoadLib.AeccRoadwayLink
                For Each oRoadwayLink In oRoadwayLinks
                    Dim iIndex As Integer
                    For iIndex = 0 To oRoadwayLink.RoadwayCodes.Count - 1
                        Dim sCodeName As String
                        sCodeName = oRoadwayLink.RoadwayCodes.Item(iIndex)
                        If oBaselineInfo.mLinkCodesGroups.ContainsKey(sCodeName) = False Then
                            Dim oLinkCodesInfo As CLinkCodes
                            oLinkCodesInfo = New CLinkCodes
                            oLinkCodesInfo.mCodeName = sCodeName
                            oBaselineInfo.mLinkCodesGroups.Add(sCodeName, oLinkCodesInfo)
                        End If
                    Next iIndex
                Next
            Next
        End If
    End Sub

    Private Function FindSampleLineGroupInfo(ByVal sCorridor As String, _
                                            ByVal sAlignment As String, _
                                            ByVal sSampleLineGroupName As String) As CSampleLineGroup
        On Error GoTo ErrHandle
        Dim oCorridorInfo As CCorridor
        Dim oBaselineInfo As CBaseline

        oCorridorInfo = m_Corridors.Item(sCorridor)
        oBaselineInfo = oCorridorInfo.mBaselines.Item(sAlignment)
        FindSampleLineGroupInfo = oBaselineInfo.mSampleLineGroups.Item(sSampleLineGroupName)

        Exit Function
ErrHandle:
        FindSampleLineGroupInfo = Nothing
    End Function

    Private Function FindLinkCodesInfo(ByVal sCorridor As String, _
                                    ByVal sAlignment As String, _
                                    ByVal sCodeName As String) As CLinkCodes
        On Error GoTo ErrHandle
        Dim oCorridorInfo As CCorridor
        Dim oBaselineInfo As CBaseline

        oCorridorInfo = m_Corridors.Item(sCorridor)
        oBaselineInfo = oCorridorInfo.mBaselines.Item(sAlignment)
        FindLinkCodesInfo = oBaselineInfo.mLinkCodesGroups.Item(sCodeName)

        Exit Function
ErrHandle:
        FindLinkCodesInfo = Nothing
    End Function

    Private Sub Combo_Corridor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Corridor.SelectedIndexChanged
        Dim oCorridorInfo As CCorridor

        If Combo_Corridor.Items.Count = 0 Then
            Exit Sub
        End If

        If Combo_Corridor.SelectedIndex < 0 Then
            Exit Sub
        End If

        oCorridorInfo = m_Corridors.Item(Combo_Corridor.SelectedItem.ToString())
        Combo_Align.Items.Clear()
        For Each sAlignmentName As String In oCorridorInfo.mBaselines.Keys
            Combo_Align.Items.Add(sAlignmentName)
        Next

        If Combo_Align.Items.Count > 0 Then
            Combo_Align.SelectedIndex = 0
        Else
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorSlopeStake_Msg_NoAlignment)
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

    Private Sub Button_SelectCorridor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_SelectCorridor.Click
        Dim oCorridor As AeccRoadLib.AeccCorridor
        oCorridor = PickCorridor()

        If Not oCorridor Is Nothing Then
            Combo_Corridor.Text = oCorridor.Name
        End If
    End Sub

    Private Function PickCorridor() As AeccRoadLib.AeccCorridor
        PickCorridor = Nothing
        On Error Resume Next

        'Prompt user to select a point or block reference:
        Dim prEntityOpts As PromptEntityOptions
        prEntityOpts = New PromptEntityOptions(LocalizedRes.CorridorSlopeStake_Msg_SelectCorridor) '"Select a parcel")
        prEntityOpts.SetRejectMessage(vbNewLine + _
            LocalizedRes.CorridorSlopeStake_Msg_SelectReject + vbNewLine) '"Must Select a parcel")
        prEntityOpts.AddAllowedClass(GetType(Autodesk.Civil.DatabaseServices.Corridor), False)

        Dim prCorridorRes As PromptEntityResult
        prCorridorRes = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.GetEntity(prEntityOpts)

        If prCorridorRes.Status = PromptStatus.OK Then
            For Each oCorridor As AeccRoadLib.AeccCorridor In ReportApplication.AeccXRoadwayDocument.Corridors
                If ReportUtilities.CompareObjectID(oCorridor.ObjectID, prCorridorRes.ObjectId) Then
                    PickCorridor = oCorridor
                    Exit For
                End If
            Next
        End If

    End Function

    Private Sub Button_CreateReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_CreateReport.Click
        'If ListView_Corridors.Items.Count = 0 Then
        '    ReportUtilities.ACADMsgBox(LocalizedRes.CorridorSlopeStake_Msg_NoComponent)
        '    Exit Sub
        'End If

        'step count 13 per corridor, EXTRACT_STEPS for extract data
        'ctlProgressBar.ProgressBegin(ListView_Corridors.Items.Count * (EXTRACT_STEPS + 3))

        Try
            'write report
            Dim tempFile As String = ReportConverter.GetRandomTempHtmlPath()
            WriteHtmlReport(tempFile)

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

        'ctlProgressBar.ProgressEnd()

        ReportUtilities.OpenFileByDefaultBrowser(ctlSaveReport.SavePath)

    End Sub

    Private Sub WriteHtmlReport(ByVal fileName As String)
        Dim oReportHtml As New ReportWriter(fileName)
        '<html>
        oReportHtml.HtmlBegin()
        oReportHtml.RenderHeader(LocalizedRes.CorridorGrinding_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.CorridorGrinding_Html_Title, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        Dim oCorridorInfo As CCorridor
        Dim oBaselineInfo As CBaseline
        Dim inStartStation As Double
        Dim inEndStation As Double
        oCorridorInfo = m_Corridors.Item(Combo_Corridor.Text)
        If Not oCorridorInfo.mObject Is Nothing Then
            oBaselineInfo = oCorridorInfo.mBaselines.Item(Combo_Align.Text)
            If Not oBaselineInfo.mObject Is Nothing Then
                inStartStation = oBaselineInfo.mObject.StartStation
                inEndStation = oBaselineInfo.mObject.EndStation
                If inStartStation <= ctlStartStation.RawStation And ctlStartStation.RawStation < ctlEndStation.RawStation And ctlEndStation.RawStation <= inEndStation Then
                    inStartStation = ctlStartStation.RawStation
                    inEndStation = ctlEndStation.RawStation
                End If
                AppendReport(oCorridorInfo.mObject, oBaselineInfo.mObject, _
                    inStartStation, inEndStation, oReportHtml)
            End If
        End If

        'Dim inStartStation As Double
        'Dim inEndStation As Double
        'Dim oCorridorInfo As CCorridor
        'Dim oBaselineInfo As CBaseline
        'Dim oSampleLineGroupInfo As CSampleLineGroup
        'Dim oLinkCodesInfo As CLinkCodes

        'For Each item As ListViewItem In ctlCorridorListView.WinListView.Items
        '    inStartStation = ctlCorridorListView.StartRawStation(item)
        '    inEndStation = ctlCorridorListView.EndRawStation(item)
        '    oCorridorInfo = m_Corridors.Item(ctlCorridorListView.CorridorName(item))
        '    oBaselineInfo = oCorridorInfo.mBaselines.Item(ctlCorridorListView.AlignmentName(item))
        '    oSampleLineGroupInfo = oBaselineInfo.mSampleLineGroups.Item(ctlCorridorListView.SampleLineGroupName(item))
        '    oLinkCodesInfo = oBaselineInfo.mLinkCodesGroups.Item(ctlCorridorListView.ListCode(item))
        '    ctlProgressBar.PerformStep()
        '    AppendReport(oCorridorInfo.mObject, oBaselineInfo.mObject, _
        '        oSampleLineGroupInfo.mObject, inStartStation, inEndStation, _
        '        oReportHtml, oLinkCodesInfo.mCodeName)
        '    ctlProgressBar.PerformStep()
        'Next

        '</body>
        oReportHtml.BodyEnd()

        '</html>
        oReportHtml.HtmlEnd()
    End Sub

    ' Append one item's report to html file
    Private Sub AppendReport(ByVal oCorridor As AeccRoadLib.IAeccCorridor, _
                            ByVal oBaseline As AeccRoadLib.IAeccBaseline, _
                            ByVal stationStart As Double, _
                            ByVal stationEnd As Double, _
                            ByVal oHtmlWriter As ReportWriter)

        ' Progress begin
        ctlProgressBar.ProgressBegin(oBaseline.GetSortedStations().Length + 3)


        Dim str As String

        oHtmlWriter.RenderHr()

        ' Corridor info
        str = String.Format(LocalizedRes.CorridorSlopeStake_Html_CorridorName, oCorridor.Name)
        oHtmlWriter.RenderLine(str)

        str = String.Format(LocalizedRes.Common_Html_Description_One_Param, oCorridor.Description)
        oHtmlWriter.RenderLine(str)

        Dim oBaseAlignment As AeccLandLib.AeccAlignment
        oBaseAlignment = oBaseline.Alignment
        str = LocalizedRes.CorridorSlopeStake_Html_BaseAlign
        str += " " + oBaseAlignment.Name
        oHtmlWriter.RenderLine(str)

        'str = LocalizedRes.CorridorSlopeStake_Html_Link
        'str += " " + "XXX"
        'oHtmlWriter.RenderLine(str)

        str = String.Format(LocalizedRes.Alignment_Html_Station_Range, _
                            ReportUtilities.GetStationStringWithDerived(oBaseAlignment, stationStart), _
                            ReportUtilities.GetStationStringWithDerived(oBaseAlignment, stationEnd))
        oHtmlWriter.RenderLine(str)

        '    extract data
        CorridorMilling_ExtractData.m_ctlProgressBar = ctlProgressBar
        Dim dWidth As Double = 0.0
        Dim dThickness As Double = 0.0
        If ctlWidthEditBox.getDoubleValue(dWidth) Then
            CorridorMilling_ExtractData.m_dMillWidthTolerance = dWidth
        End If
        If ctlThicknessEditBox.getDoubleValue(dThickness) Then
            CorridorMilling_ExtractData.m_dMillHeightTolerance = dThickness
        End If
        CorridorMilling_ExtractData.ExtractData(oCorridor, oBaseline, stationStart, stationEnd) ', sCodeName, ctlProgressBar.ProgressBar, EXTRACT_STEPS)

        'ctlProgressBar.PerformStep()
        str = oBaseAlignment.Name ' & "<br/>"
        oHtmlWriter.RenderLine(str, True)

        Dim oDistSettings As AeccLandLib.IAeccSettingsDistance
        oDistSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.DistanceSettings
        Dim oAreaSettings As AeccLandLib.AeccSettingsArea
        oAreaSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.AreaSettings
        'Dim distUnit As String = oDistSettings.Unit.Value.ToString()
        'Dim areaUnit As String = oAreaSettings.Unit.Value.ToString()

        'table begin
        oHtmlWriter.TableBegin("1")

        ' Add Table Headers
        oHtmlWriter.TrBegin()
        ' TODO : Add to resource
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Colspan, "2")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorGrinding_Station, False, "center")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Colspan, "2")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorGrinding_OffsetAtBegin, False, "center")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Colspan, "2")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorGrinding_OffsetAtEnd, False, "center")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "20%")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Rowspan, "2")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorGrinding_Area, False, "center")
        oHtmlWriter.TrEnd()
        oHtmlWriter.TrBegin()
        ' TODO : Add to resource
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "13%")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorGrinding_StationBegin, False, "center")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "12%")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorGrinding_StationEnd, False, "center")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "12%")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorGrinding_OffsetFrom, False, "center")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "12%")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorGrinding_OffsetTo, False, "center")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "14%")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorGrinding_OffsetFrom, False, "center")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "13%")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorGrinding_OffsetTo, False, "center")
        oHtmlWriter.TrEnd()

        Dim dTotalArea As Double = 0.0
        For Each grindingSectionData As GrindingSectionData In CorridorMilling_ExtractData.GrindingData.mSectionDataList
            'separate row
            oHtmlWriter.TrBegin()
            oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "13%")
            oHtmlWriter.RenderTd(ReportUtilities.GetStationStringWithDerived(oBaseAlignment, grindingSectionData.mStartStation), False, "right")
            oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "12%")
            oHtmlWriter.RenderTd(ReportUtilities.GetStationStringWithDerived(oBaseAlignment, grindingSectionData.mEndStation), False, "right")
            oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "12%")
            oHtmlWriter.RenderTd(CorridorMilling_ExtractData.FormatDistSettings(grindingSectionData.mStartLeftOffset), False, "right")
            oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "12%")
            oHtmlWriter.RenderTd(CorridorMilling_ExtractData.FormatDistSettings(grindingSectionData.mStartRightOffset), False, "right")
            oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "14%")
            oHtmlWriter.RenderTd(CorridorMilling_ExtractData.FormatDistSettings(grindingSectionData.mEndLeftOffset), False, "right")
            oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "13%")
            oHtmlWriter.RenderTd(CorridorMilling_ExtractData.FormatDistSettings(grindingSectionData.mEndRightOffset), False, "right")
            Dim dArea As Double
            dArea = (grindingSectionData.mEndStation - grindingSectionData.mStartStation) * (grindingSectionData.mStartRightOffset - grindingSectionData.mStartLeftOffset + grindingSectionData.mEndRightOffset - grindingSectionData.mEndLeftOffset) / 2
            dTotalArea += dArea
            oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "20%")
            oHtmlWriter.RenderTd(CorridorMilling_ExtractData.FormatAreaSettings(dArea), False, "right")

            oHtmlWriter.TrEnd()

            'Dim arrData(0 To 7) As String
            'arrData(0) = oBaseAlignment.GetStationStringWithEquations(grindingSectionData.mStartStation)
            'arrData(1) = oBaseAlignment.GetStationStringWithEquations(grindingSectionData.mEndStation)
            'arrData(2) = grindingSectionData.mStartLeftOffset.ToString("N2")
            'arrData(3) = grindingSectionData.mStartRightOffset.ToString("N2")
            'arrData(4) = grindingSectionData.mEndLeftOffset.ToString("N2")
            'arrData(5) = grindingSectionData.mEndRightOffset.ToString("N2")
            'Dim dArea As Double
            'dArea = (grindingSectionData.mEndStation - grindingSectionData.mStartStation) * (grindingSectionData.mStartRightOffset - grindingSectionData.mStartLeftOffset + grindingSectionData.mEndRightOffset - grindingSectionData.mEndLeftOffset) / 2
            'arrData(6) = dArea.ToString("N2")

            'WriteOneTableRow(arrData, oHtmlWriter)

        Next

        ' Render the totalArea
        oHtmlWriter.TrBegin()
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "13%")
        oHtmlWriter.RenderTd("&nbsp;", False, "right")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "12%")
        oHtmlWriter.RenderTd("&nbsp;", False, "right")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "12%")
        oHtmlWriter.RenderTd("&nbsp;", False, "right")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "12%")
        oHtmlWriter.RenderTd("&nbsp;", False, "right")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "14%")
        oHtmlWriter.RenderTd("&nbsp;", False, "right")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "13%")
        oHtmlWriter.RenderTd("&nbsp;", False, "right")
        oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "20%")
        oHtmlWriter.RenderTd(CorridorMilling_ExtractData.FormatAreaSettings(dTotalArea), True, "right")

        oHtmlWriter.TrEnd()


        oHtmlWriter.TableEnd()


        'For Each key As Double In CorridorSlopeStake_ExtractData.SlopeStakeData.Keys
        '    ' Generate graph to file
        '    Dim graph As New SSRGraphics

        '    Dim tempSectData As Slope_SectionData = _
        '        CorridorSlopeStake_ExtractData.SlopeStakeData(key)
        '    Dim tempPtData As Slope_PointData
        '    For Each offsetKey As Double In tempSectData.LeftDatas.Keys()
        '        tempPtData = tempSectData.LeftDatas(offsetKey)
        '        graph.AddCodePoint(tempPtData)
        '    Next
        '    graph.AddCodePoint(tempSectData.CenterLineInfo)
        '    For Each offsetKey As Double In tempSectData.RightDatas.Keys()
        '        tempPtData = tempSectData.RightDatas(offsetKey)
        '        graph.AddCodePoint(tempPtData)
        '    Next
        '    graph.SetLinks(tempSectData.LinksOnGraph)
        '    graph.SetEGDatas(tempSectData.EGDatas)

        '    Dim tempPath As String = System.IO.Path.GetTempPath()

        '    Dim curStationString As String
        '    curStationString = oSampleLineGroup.Parent.GetStationStringWithEquations(key)

        '    Dim strImgFileName As String = tempPath + sCodeName + "_" + curStationString + "_" + DateTime.Now.ToFileTime().ToString()
        '    Dim strExt As String = "png"
        '    graph.CreateGraph(strImgFileName, _
        '                      strExt)

        '    ' Add graph above current table
        '    oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Border, "0")
        '    oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "100%")
        '    oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Src, "file:///" + strImgFileName + "." + strExt)
        '    oHtmlWriter.ImgBegin()
        '    oHtmlWriter.ImgEnd()

        '    AppendSingleTable(oBaseAlignment, key, oHtmlWriter)
        'Next

        oHtmlWriter.RenderBr()

        ' Progress end
        ctlProgressBar.ProgressEnd()

    End Sub


    Private Sub AppendSingleTable(ByVal baseAlignment As AeccLandLib.IAeccAlignment, _
                                ByVal dStation As Double, ByVal oHtmlWriter As ReportWriter)
        Dim str As String
        str = baseAlignment.Name ' & "<br/>"
        oHtmlWriter.RenderLine(str, True)

        str = String.Format(LocalizedRes.CorridorSlopeStake_Html_Station, _
                            ReportUtilities.GetStationStringWithDerived(baseAlignment, dStation))
        oHtmlWriter.RenderLine(str, True)

        'table begin
        oHtmlWriter.TableBegin("1")

        Dim sData As Slope_SectionData
        Try
            sData = CorridorSlopeStake_ExtractData.SlopeStakeData.Item(dStation)
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
        lineNum = CInt(Math.Ceiling(maxSidePtCount / 5))

        'get enumerator of left and right data dictionary
        Dim leftKeys(leftPtCount) As Double
        Dim rightKeys(rightPtCount) As Double
        sData.LeftDatas.Keys.CopyTo(leftKeys, 0)
        sData.RightDatas.Keys.CopyTo(rightKeys, 0)

        For line As Integer = 0 To lineNum - 1
            'as we mention before, we have 4 rows
            For row As Integer = 0 To 3
                'there're 13 cell in one line include 2 ends, 1 middle CL, 5 left points and 5 right points
                'left side should be reserve later after filled
                Dim leftDataArr(0 To 5) As String
                Dim rightDataArr(0 To 5) As String
                Dim clData As String = ""

                'set end data at the first line
                If line = 0 Then
                    If row = 0 Then
                        leftDataArr(5) = sData.LeftEndInfo.EndType
                        rightDataArr(5) = sData.RightEndInfo.EndType
                        clData = sData.CenterLineInfo.mCodes
                    ElseIf row = 1 Then
                        leftDataArr(5) = sData.LeftEndInfo.DeltaOffset
                        rightDataArr(5) = sData.RightEndInfo.DeltaOffset
                        clData = sData.CenterLineInfo.mOffsetString
                    ElseIf row = 2 Then
                        leftDataArr(5) = sData.LeftEndInfo.EndSlope
                        rightDataArr(5) = sData.RightEndInfo.EndSlope
                        clData = sData.CenterLineInfo.mElevationString
                    End If
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
                Dim dataArr(0 To 12) As String
                Array.Reverse(leftDataArr)
                leftDataArr.CopyTo(dataArr, 0)
                dataArr(6) = clData
                rightDataArr.CopyTo(dataArr, 7)
                WriteOneTableRow(dataArr, oHtmlWriter)
            Next

            'separate row
            Dim arrBlank(0 To 12) As String
            WriteOneTableRow(arrBlank, oHtmlWriter)
        Next

        oHtmlWriter.TableEnd()
    End Sub

    Private Sub fillPointColumn(ByVal row As Integer, _
                                    ByVal ptData As Slope_PointData, _
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
                Case Else 'impossible here
                    cell = ""
            End Select
        Catch ex As Exception
            Diagnostics.Debug.Assert(False, ex.Message)
        End Try
    End Sub

    Private Sub WriteOneTableRow(ByVal arrData As String(), ByVal oHtmlWriter As ReportWriter)
        Dim i As Integer
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


    Private Sub FilterLinkCode(ByRef oBaseLineInfo As CBaseline, _
                            ByRef oSampleLineGroupInfo As CSampleLineGroup)

        Diagnostics.Debug.Assert(oBaseLineInfo.mIsLinkCodesFilterred = False)

        Dim oBaseLine As AeccRoadLib.IAeccBaseline = oBaseLineInfo.mObject
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
                    Diagnostics.Debug.Assert(False, ex.Message)
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

    Private Sub openHelp()
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_MILLING_REPORT)
    End Sub

    Private Sub ctlEndStation_StationChanging(ByVal rawStation As Double, ByVal equationStation As String, ByRef okToChange As Boolean) Handles ctlEndStation.StationChanging
        If rawStation < ctlStartStation.RawStation AndAlso m_bStnChangeByComboIndexChange = False Then
            ReportUtilities.ACADMsgBox(LocalizedRes.Msg_EndStation_Greater_StartStation)
            okToChange = False
        End If
    End Sub

    Private Sub ctlStartStation_StationChanging(ByVal rawStation As Double, ByVal equationStation As String, ByRef okToChange As Boolean) Handles ctlStartStation.StationChanging
        If rawStation > ctlEndStation.RawStation AndAlso m_bStnChangeByComboIndexChange = False Then
            ReportUtilities.ACADMsgBox(LocalizedRes.Msg_StartStation_Less_EndStation)
            okToChange = False
        End If
    End Sub

    Private Sub Combo_Align_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Align.SelectedIndexChanged
        m_bStnChangeByComboIndexChange = True
        updateStartEndStations()
        m_bStnChangeByComboIndexChange = False
    End Sub

    Private Sub updateStartEndStations()

        Dim oCorridorInfo As CCorridor
        Dim oBaselineInfo As CBaseline
        Dim inStartStation As Double
        Dim inEndStation As Double
        Dim oStationSetting As AeccLandLib.AeccSettingsStation = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.StationSettings
        oCorridorInfo = m_Corridors.Item(Combo_Corridor.Text)

        If Not oCorridorInfo.mObject Is Nothing Then
            oBaselineInfo = oCorridorInfo.mBaselines.Item(Combo_Align.Text)
            If Not oBaselineInfo.mObject Is Nothing Then
                inStartStation = oBaselineInfo.mObject.StartStation
                inEndStation = oBaselineInfo.mObject.EndStation
                Dim oBaseAlignment As AeccLandLib.AeccAlignment
                oBaseAlignment = oBaselineInfo.mObject.Alignment
                If Not oBaseAlignment Is Nothing Then
                    ctlStartStation.Alignment = oBaseAlignment
                    ctlEndStation.Alignment = oBaseAlignment
                    ctlStartStation.EquationStation = ReportUtilities.GetStationString(oBaseAlignment, inStartStation)
                    ctlEndStation.EquationStation = ReportUtilities.GetStationString(oBaseAlignment, inEndStation)
                    ctlStartStation.RawStation = ReportFormat.RoundDouble(inStartStation, oStationSetting.Precision.Value, oStationSetting.Rounding.Value)
                    ctlEndStation.RawStation = ReportFormat.RoundDouble(inEndStation, oStationSetting.Precision.Value, oStationSetting.Rounding.Value)

                End If
            End If
        End If

    End Sub
End Class
