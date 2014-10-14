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
Imports System.Linq
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports AeccRoadLib = Autodesk.AECC.Interop.Roadway
Imports Microsoft.Office.Interop.Excel
Imports LocalizedRes = Report.My.Resources.LocalizedStrings
Imports Autodesk.AECC.Interop

Friend Class ReportForm_CorridorSectionPoints

    Private ctlProgressBar As CtrlProgressBar
    Private ctlSaveReport As CtrlSaveReportFile
    Private WithEvents ctlStartStation As CtrlStationTextBox
    Private WithEvents ctlEndStation As CtrlStationTextBox

    Private m_Corridors As New Dictionary(Of String, CCorridor_Emia)
    Private m_Alignments As New Dictionary(Of String, AeccLandLib.AeccAlignment)

    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()
        Dim rptDlg As New ReportForm_CorridorSectionPoints
        If rptDlg.CanOpen() = True Then
            ReportUtilities.RunModalDialog(rptDlg)
        End If
    End Sub

    Private Function CanOpen() As Boolean
        Dim bNoCorridors As Boolean = False
        Dim nAlignmentCount As Integer = 0

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

        'get alignments from Site :
        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            nAlignmentCount += oSite.Alignments.Count()
        Next

        CanOpen = False

        If nAlignmentCount < 1 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorSectionPoints_Msg_NoAlignment)
        ElseIf bNoCorridors Then
            'TODO
            Dim sNoCorridors As String = LocalizedRes.CorridorSectionPoints_Msg_NoCorridors
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

    Private Sub InitReportSettings()
        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)

        ctlStartStation = New CtrlStationTextBox
        ctlStartStation.Initialize(TextBox_StartStation)

        ctlEndStation = New CtrlStationTextBox
        ctlEndStation.Initialize(TextBox_EndStation)

        ctlSaveReport = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)

        ' listview check button
        Dim bp1 As System.Drawing.Bitmap = My.Resources.Resources.ChkAll
        Dim bp2 As System.Drawing.Bitmap = My.Resources.Resources.UnchkAll
        bp1.MakeTransparent(Color.White)
        bp2.MakeTransparent(Color.White)
        Button_CheckAll.Image = bp1
        Button_CheckAll.ImageAlign = ContentAlignment.MiddleCenter
        Button_UncheckAll.Image = bp2
        Button_UncheckAll.ImageAlign = ContentAlignment.MiddleCenter

        ' listview initialization
        ListView_FeatureLines.View = View.Details
        ListView_FeatureLines.Columns.Add(LocalizedRes.CorridorCrossSection_Form_ColumnTitle_Check, _
            100, HorizontalAlignment.Left)

        ListView_FeatureLines.Columns.Add(LocalizedRes.CorridorCrossSection_Form_ColumnTitle_PointCodes, _
            80, HorizontalAlignment.Left)

    End Sub

    Private Sub ReportForm_CorridorCrossSection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InitReportSettings()

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

        'initial Select report Components
        For Each corridorName As String In m_Corridors.Keys
            Combo_Corridor.Items.Add(corridorName)
        Next

        'Corridor count should more then 0, because we have test it before dialog open
        If m_Corridors.Count > 0 Then
            Combo_Corridor.SelectedIndex = 0
        End If

        Exit Sub
    End Sub

    Private Sub BuildCorridorInfo(ByRef oCorridor As AeccRoadLib.IAeccCorridor, ByRef oCorridorInfo As CCorridor_Emia)
        oCorridorInfo.mObject = oCorridor
        oCorridorInfo.mBaselines.Clear() 'dicBaselines.Clear()

        For Each oBaseline As AeccRoadLib.AeccBaseline In oCorridor.Baselines
            'If oCorridorInfo.mBaselines.ContainsKey(oBaseline.Alignment.Name) = False Then
            '    Dim oBaselineInfo As CBaseline
            '    oBaselineInfo = New CBaseline
            '    BuildBaselineInfo(oBaseline, oBaselineInfo)
            '    oCorridorInfo.mBaselines.Add(oBaseline.Alignment.Name, oBaselineInfo)
            'End If
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

    Private Function GetPointCodesForLink(ByRef link As AeccRoadLib.AeccRoadwayLink) As List(Of String)

        'Dim pointCodesStrings = From pt In link.Points Select pt


        Dim pointCodesList As New List(Of String)
        ''Dim points As AeccRoadLib.AeccRoadwayPoints = link.Points
        Dim points As AeccRoadLib.AeccRoadwayPoints = link.Points
        For i As Integer = 0 To points.Count - 1
            Dim point As AeccRoadLib.AeccRoadwayPoint = points.Item(i)
            For Each code As String In point.RoadwayCodes
                pointCodesList.Add(code)
            Next
        Next
        'For Each pt As AeccRoadLib.AeccRoadwayPoint In link.Points
        '    For Each code As String In pt.RoadwayCodes
        '        pointCodesList.Add(code)
        '    Next
        'Next
        Return pointCodesList
    End Function

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
                    Dim ptCodeStrings As List(Of String) = GetPointCodesForLink(oRoadwayLink)
                    Dim iIndex As Integer
                    For iIndex = 0 To oRoadwayLink.RoadwayCodes.Count - 1
                        Dim sCodeName As String
                        sCodeName = oRoadwayLink.RoadwayCodes.Item(iIndex)
                        If .mLinkCodesGroups.ContainsKey(sCodeName) = False Then
                            Dim oLinkCodesInfo As CLinkCodes
                            oLinkCodesInfo = New CLinkCodes
                            oLinkCodesInfo.mCodeName = sCodeName
                            oLinkCodesInfo.mPointCodes = ptCodeStrings
                            .mLinkCodesGroups.Add(sCodeName, oLinkCodesInfo)
                        Else
                            .mLinkCodesGroups(sCodeName).mPointCodes.AddRange(ptCodeStrings)
                        End If
                    Next iIndex
                Next
            Next
        End With
    End Sub

    Private Sub Combo_Align_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Align.SelectedIndexChanged
        If Combo_Align.Items.Count = 0 Then
            Exit Sub
        End If

        If Combo_Align.SelectedIndex < 0 Then
            Exit Sub
        End If

        Dim oStationSetting As AeccLandLib.AeccSettingsStation = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.StationSettings

        ctlStartStation.TextBox.Enabled = False
        ctlEndStation.TextBox.Enabled = False

        ctlStartStation.Alignment = Nothing
        ctlEndStation.Alignment = Nothing

        ctlStartStation.EquationStation = ""
        ctlEndStation.EquationStation = ""
        'TextBox_StartStation.Text = ReportFormat.RoundVal(0.0, oStationSetting.Precision.Value, oStationSetting.Rounding.Value)
        'TextBox_EndStation.Text = ReportFormat.RoundVal(0.0, oStationSetting.Precision.Value, oStationSetting.Rounding.Value)

        'oCorridorInfo = m_Corridors.Item(Combo_Corridor.SelectedItem)
        'oBaselineInfo = oCorridorInfo.mBaselines.Item(Combo_Align.SelectedItem)
        TextBox_StartStation.Enabled = True
        TextBox_EndStation.Enabled = True

        Dim alignment As AeccLandLib.IAeccAlignment = m_Alignments(Combo_Align.SelectedItem)
        ctlStartStation.Alignment = alignment
        ctlEndStation.Alignment = ctlStartStation.Alignment
        ctlEndStation.RawStation = ReportFormat.RoundVal(alignment.EndingStation, oStationSetting.Precision.Value, oStationSetting.Rounding.Value)
        ctlStartStation.RawStation = ReportFormat.RoundVal(alignment.StartingStation, oStationSetting.Precision.Value, oStationSetting.Rounding.Value)
        ctlStartStation.EquationStation = ReportUtilities.GetStationString(alignment, alignment.StartingStation)
        ctlEndStation.EquationStation = ReportUtilities.GetStationString(alignment, alignment.EndingStation)
        FillLinkcodesCombo()
    End Sub

    Private Sub Combo_Corridor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Corridor.SelectedIndexChanged
        If Combo_Corridor.Items.Count = 0 Then
            Exit Sub
        End If

        If Combo_Corridor.SelectedIndex < 0 Then
            Exit Sub
        End If

        FillAlignmentCombo()
    End Sub

    Private Sub BtnDone_Clicked() Handles Button_Done.Click
        Close()
    End Sub

    Private Sub BtnHelp_Click() Handles Button_Help.Click
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CORRIDOR_SECTION_POINTS_REPORT)
    End Sub

    Private Sub ReportForm_HelpRequested(ByVal sender As System.Object, ByVal hlpevent As System.Windows.Forms.HelpEventArgs) Handles MyBase.HelpRequested
        BtnHelp_Click()
    End Sub

    Private Sub ButtonExecute_Click() Handles Button_CreateReport.Click

        ' there is at least one output check box checked - ask user about rebuilding corridor
        Dim inStartStation As Double = ctlStartStation.RawStation
        Dim inEndStation As Double = ctlEndStation.RawStation
        Dim oCorridorInfo As CCorridor_Emia = m_Corridors.Item(Combo_Corridor.SelectedItem)
        Dim oCorridor As AeccRoadLib.IAeccCorridor = oCorridorInfo.mObject
        'If CheckCorridorChainage(inStartStation, inEndStation, oCorridorInfo.mBaselines.Item(Combo_Align.SelectedItem).mObject) Then
        '    RebuildCorridor(oCorridor, inStartStation, inEndStation, oCorridorInfo.mBaselines.Item(Combo_Align.SelectedItem).mObject)
        'End If

        Try
            Dim tempFile As String = ReportConverter.GetRandomTempHtmlPath()

            WriteReportHTML_Points(tempFile)

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

        ReportUtilities.OpenFileByDefaultBrowser(ctlSaveReport.SavePath)

    End Sub

    Private Sub WriteReportHTML_Points(ByVal fileName As String)
        Dim oReportHtml As New ReportWriter(fileName)
        '<html>
        oReportHtml.HtmlBegin()
        oReportHtml.RenderHeader(LocalizedRes.CorridorSectionPoints_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.CorridorSectionPoints_Html_Title, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        Dim inStartStation As Double
        Dim inEndStation As Double
        Dim inStationInterval As Double
        Dim oCorridorInfo As CCorridor_Emia
        Dim oBaselineInfos As List(Of CBaseline)

        inStartStation = ctlStartStation.RawStation
        inEndStation = ctlEndStation.RawStation
        inStationInterval = StationInterval
        oCorridorInfo = m_Corridors.Item(Combo_Corridor.SelectedItem)
        'oBaselineInfo = oCorridorInfo.mBaselines.Item(Combo_Align.SelectedItem)
        oBaselineInfos = oCorridorInfo.mBaselines.Item(Combo_Align.SelectedItem)

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

            AppendReport(oCorridorInfo.mObject, oBaselineInfo.mObject, _
                         inStartStation, inEndStation, inStationInterval, _
                         oReportHtml)
        Next

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
                            ByVal stationInterval As Double, _
                            ByVal oHtmlWriter As ReportWriter)

        Dim str As String

        oHtmlWriter.RenderHr()

        ' Corridor info
        str = String.Format(LocalizedRes.CorridorSectionPoints_Html_CorridorName, oCorridor.Name)
        oHtmlWriter.RenderLine(str)

        str = String.Format(LocalizedRes.Common_Html_Description_One_Param, oCorridor.Description)
        oHtmlWriter.RenderLine(str)

        Dim oBaseAlignment As AeccLandLib.AeccAlignment
        oBaseAlignment = oBaseline.Alignment
        str = LocalizedRes.CorridorSlopeStake_Html_BaseAlign
        str += " " + oBaseAlignment.Name
        oHtmlWriter.RenderLine(str)

        str = String.Format(LocalizedRes.Alignment_Html_Station_Range, _
                            ReportUtilities.GetStationStringWithDerived(oBaseAlignment, stationStart), _
                            ReportUtilities.GetStationStringWithDerived(oBaseAlignment, stationEnd))
        oHtmlWriter.RenderLine(str)

        'init progress bar
        Dim count As UInteger = (stationEnd - stationStart) / stationInterval + 1
        ctlProgressBar.ProgressBegin(count)

        Dim dStation As Double = stationStart

        While dStation < stationEnd
            AppendSingleTable(oBaseline, dStation, oHtmlWriter)
            dStation += stationInterval
            ctlProgressBar.PerformStep()
        End While

        ctlProgressBar.ProgressEnd()
        oHtmlWriter.RenderBr()
    End Sub

    Friend Class PointData_Comparer
        Implements IComparer(Of PointData)
        Function Compare(ByVal x As PointData, ByVal y As PointData) As Integer _
            Implements IComparer(Of PointData).Compare
            If x.mOffset < y.mOffset Then
                Return -1
            ElseIf x.mOffset = y.mOffset Then
                Return 0
            Else
                Return 1
            End If
        End Function 'IComparer.Compare
    End Class

    Friend Class PointData
        Public mOffset As Double = 0
        Public mEasting As Double = 0
        Public mNorthing As Double = 0
        Public mElevation As Double = 0
        Public mCode As String = ""
    End Class

    Private Sub RemoveDuplicatedPoint(ByRef pointsData As List(Of PointData))
        Dim listPD As New List(Of PointData)
        Dim comparer As New PointData_Comparer

        For i As Integer = 0 To pointsData.Count - 1
            Dim bExist As Boolean = False
            For j As Integer = 0 To listPD.Count - 1
                If pointsData.Item(i).mOffset = listPD.Item(j).mOffset And pointsData.Item(i).mElevation = listPD.Item(j).mElevation Then
                    bExist = True
                    Exit For
                End If
            Next

            If bExist = False And Not pointsData.Item(i).mCode = "" Then
                listPD.Add(pointsData.Item(i))
            End If
        Next

        pointsData = listPD
    End Sub

    Private Sub AppendSingleTable(ByVal oBaseline As AeccRoadLib.IAeccBaseline, _
                                  ByVal dStation As Double, ByVal oHtmlWriter As ReportWriter)
        Dim str As String
        oHtmlWriter.RenderLine("", False)
        str = String.Format(LocalizedRes.CorridorSectionPoints_Html_Chainage, _
                            ReportUtilities.GetStationStringWithDerived(oBaseline.Alignment, dStation))
        oHtmlWriter.RenderLine(str, False)

        'table begin
        oHtmlWriter.TableBegin("0")

        'get Calculated Points for selected Feature Line
        Dim appliedAssembly As Roadway.IAeccAppliedAssembly
        Dim calcPoints As Object = Nothing
        'Dim calcPointsCrown As Roadway.AeccCalculatedPoints = Nothing
        Dim oStation As Object = dStation

        Try
            appliedAssembly = oBaseline.AppliedAssembly(oStation)
            calcPoints = From item In ListView_FeatureLines.CheckedItems, _
                              calcPnt In appliedAssembly.GetPointsByCode(item.SubItems(1).Text) _
                              Select calcPnt
            'calcPoints = appliedAssembly.GetPoints()
        Catch
            Exit Sub
        End Try

        'Dim nPointsCount As Integer = calcPoints.Count

        Dim pointsData As New List(Of PointData)

        For Each calcPoint As Roadway.AeccCalculatedPoint In calcPoints
            'For idx As Integer = 0 To nPointsCount - 1
            Dim dataArr As New PointData

            'Dim calcPoint As Roadway.AeccCalculatedPoint = calcPoints.Item(idx)

            Dim oSoe As Object = calcPoint.GetStationOffsetElevationToBaseline()
            Dim XYZ As Object = oBaseline.StationOffsetElevationToXYZ(oSoe)

            dataArr.mEasting = XYZ(0)
            dataArr.mNorthing = XYZ(1)
            dataArr.mElevation = XYZ(2)
            dataArr.mOffset = oSoe(1)
            dataArr.mCode = calcPoint.CorridorCodes.Item(0)

            pointsData.Add(dataArr)
        Next

        pointsData.Sort(New PointData_Comparer)
        RemoveDuplicatedPoint(pointsData)

        WriteTableTitle(oHtmlWriter)
        For row As Integer = 0 To pointsData.Count - 1
            WriteOneTableRow(row + 1, pointsData.Item(row), oHtmlWriter)
        Next

        oHtmlWriter.TableEnd()
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
        str = LocalizedRes.CorridorSectionPoints_Msg_CorridorRebuild
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

    Public Shared Function FormatDistSettings(ByVal dDis As Double, ByRef dDisRounded As Double) As String
        Dim oDistSettings As AeccLandLib.IAeccSettingsDistance
        oDistSettings = ReportApplication.AeccXRoadwayDatabase.Settings.SectionSettings.AmbientSettings.DistanceSettings

        'oDistSettings.Precision.Value
        dDisRounded = ReportFormat.RoundDouble(dDis, oDistSettings.Precision.Value, oDistSettings.Rounding.Value)

        FormatDistSettings = ReportFormat.FormatDistance(dDis, oDistSettings.Unit.Value, _
            oDistSettings.Precision.Value, oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
    End Function

    Private Sub WriteOneTableRow(ByVal row As Integer, ByVal data As PointData, ByVal oHtmlWriter As ReportWriter)

        Dim desting As Double
        Dim dnorthing As Double
        Dim delevation As Double
        Dim doffset As Double

        Dim esting As String = CorridorCrossSection_ExtractData.FormatCoordSettings(data.mEasting, desting)
        Dim northing As String = CorridorCrossSection_ExtractData.FormatCoordSettings(data.mNorthing, dnorthing)
        Dim elevation As String = CorridorCrossSection_ExtractData.FormatCoordSettings(data.mElevation, delevation)
        Dim offset As String = FormatDistSettings(data.mOffset, doffset)

        oHtmlWriter.TrBegin()

        oHtmlWriter.RenderTd(row.ToString(), False, "Center")
        oHtmlWriter.RenderTd(esting, False, "Center")
        oHtmlWriter.RenderTd(northing, False, "Center")
        oHtmlWriter.RenderTd(elevation, False, "Center")
        oHtmlWriter.RenderTd(offset, False, "Center")
        oHtmlWriter.RenderTd(data.mCode, False, "Center")

        oHtmlWriter.TrEnd()
    End Sub

    Private Sub WriteTableTitle(ByVal oHtmlWriter As ReportWriter)

        oHtmlWriter.TrBegin()

        oHtmlWriter.RenderTd(LocalizedRes.CorridorSectionPoints_Html_Point, True, "Center")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorSectionPoints_Html_X, True, "Center")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorSectionPoints_Html_Y, True, "Center")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorSectionPoints_Html_Z, True, "Center")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorSectionPoints_Html_Offset, True, "Center")
        oHtmlWriter.RenderTd(LocalizedRes.CorridorSectionPoints_Html_StringCut, True, "Center")

        oHtmlWriter.TrEnd()

    End Sub

    Private Sub WriteOneBlankTableRow(ByVal row As Integer, ByVal data As PointData, ByVal oHtmlWriter As ReportWriter)
        oHtmlWriter.TrBegin()

        oHtmlWriter.RenderTd("", False, "Center")
        oHtmlWriter.RenderTd("", False, "Center")
        oHtmlWriter.RenderTd("", False, "Center")
        oHtmlWriter.RenderTd("", False, "Center")
        oHtmlWriter.RenderTd("", False, "Center")
        oHtmlWriter.RenderTd("", False, "Center")

        oHtmlWriter.TrEnd()
    End Sub

    Private Sub FillLinkcodesCombo()

        Combo_Link.Items.Clear()
        Dim selectAll = "All"
        Combo_Link.Items.Add(selectAll)

        'Dim baseLine As CBaseline = _
        '        m_Corridors.Item(Combo_Corridor.SelectedItem).mBaselines.Item(Combo_Align.SelectedItem)
        Dim corridor As CCorridor_Emia = m_Corridors.Item(Combo_Corridor.SelectedItem)
        Dim baseLines As List(Of CBaseline) = corridor.mBaselines.Item(Combo_Align.SelectedItem)

        'Dim codes = (From link In baseline.mLinkCodesGroups _
        '                 Select link.Key).Distinct
        Dim codes As New List(Of String)
        For Each baseline As CBaseline In baseLines
            For Each link As String In baseline.mLinkCodesGroups.Keys
                If Not codes.Contains(link) Then
                    codes.Add(link)
                End If
            Next
        Next
        For Each code As String In codes
            Combo_Link.Items.Add(code)
        Next

        If Combo_Link.Items.Count > 1 Then ' "All" is always there
            Combo_Link.SelectedIndex = 0
        Else
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorCrossSection_Msg_NoLinkCode)
            Exit Sub
        End If

    End Sub

    Private Sub FillAlignmentCombo()
        ' gets called only when Surfaces radio button is selected
        Combo_Align.Items.Clear()

        'For Each align As String In m_Alignments.Keys()
        '    Combo_Align.Items.Add(align)
        'Next

        Dim corridor As CCorridor_Emia = m_Corridors.Item(Combo_Corridor.SelectedItem.ToString())
        For Each alignName As String In corridor.mBaselines.Keys
            Combo_Align.Items.Add(alignName)
        Next

        If Combo_Align.Items.Count > 0 Then
            Combo_Align.SelectedIndex = 0
        Else
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorSectionPoints_Msg_NoAlignment)
            Exit Sub
        End If
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

    Private Sub Combo_Link_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Link.SelectedIndexChanged
        ListView_FeatureLines.Items.Clear()
        Dim linkCode As String = Combo_Link.SelectedItem
        Dim corridor As CCorridor_Emia = m_Corridors.Item(Combo_Corridor.SelectedItem)
        Dim baseLines As List(Of CBaseline) = corridor.mBaselines.Item(Combo_Align.SelectedItem)

        Dim selectedFls As Object = Nothing
        If linkCode = "All" Then
            selectedFls = corridor.mObject.FeatureLineCodeInfos.CodeNames
        Else
            selectedFls = New List(Of String)
            For Each baseline As CBaseline In baseLines
                Dim selectedFlsTmp As Object = (From link In baseLines(0).mLinkCodesGroups _
                                   Where link.Value.mCodeName = linkCode _
                                   From ptCodeByLink In link.Value.mPointCodes, _
                                        ptCode In corridor.mObject.FeatureLineCodeInfos.CodeNames _
                                        Where ptCode = ptCodeByLink _
                                        Select ptCode).Distinct
                For Each code As String In selectedFlsTmp
                    selectedFls.Add(code)
                Next
                'For Each link As CLinkCodes In baseline.mLinkCodesGroups.Values
                '    If link.mCodeName = linkCode Then
                '        For Each ptCodeByLink As String In link.mPointCodes
                '            For Each ptCode As String In corridor.mObject.FeatureLineCodeInfos.CodeNames
                '                If ptCode = ptCodeByLink Then
                '                    selectedFls.Add(ptCode)
                '                End If
                '            Next
                '        Next
                '    End If
                'Next
            Next
            selectedFls.Sort()
        End If
        For Each code As String In selectedFls
            If ListView_FeatureLines.Items.ContainsKey(code) Then
                Continue For
            End If
            Dim li As New ListViewItem
            li = ListView_FeatureLines.Items.Add(code, String.Empty, String.Empty)
            li.SubItems.Add(code)
            li.Checked = True
        Next
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
End Class
