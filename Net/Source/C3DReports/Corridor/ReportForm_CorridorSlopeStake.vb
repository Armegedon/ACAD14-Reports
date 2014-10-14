' -----------------------------------------------------------------------------
' <copyright file="ReportForm_CorridorSlopeStake.vb" company="Autodesk">
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
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.Civil.DatabaseServices

Friend Class ReportForm_CorridorSlopeStake

    Private Const EXTRACT_STEPS As Integer = 30

    Private m_Corridors As New SortedDictionary(Of String, CCorridor)

    Private ctlCorridorListView As CtrlCorridorListView
    Private ctlStartStation As CtrlStationTextBox
    Private ctlEndtStation As CtrlStationTextBox
    Private ctlProgressBar As CtrlProgressBar
    Private ctlSaveReport As CtrlSaveReportFile
    Private oVolumeSettings As AeccLandLib.AeccSettingsVolume
    Private oAreaSettings As AeccLandLib.AeccSettingsArea

    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()
        Dim reportForm As New ReportForm_CorridorSlopeStake
        If reportForm.CanOpen() = True Then
            ReportUtilities.RunModalDialog(reportForm)
        End If
    End Sub

    Private Function CanOpen() As Boolean
        Dim bNoCorridors As Boolean = False

        m_Corridors.Clear()
        If ReportApplication.AeccXRoadwayDocument.Corridors.Count > 0 Then
            'initial Corridors
            Dim oCorridors As AeccRoadLib.AeccCorridors
            oCorridors = ReportApplication.AeccXRoadwayDocument.Corridors
            For Each oCorridor As AeccRoadLib.IAeccCorridor In oCorridors
                If m_Corridors.ContainsKey(oCorridor.Name) = False Then
                    Dim bSuccess As Boolean = False
                    ' Block reference that hasn't been exploded should not be recognized in report dialog
                    ' Check whether its owner - that block record is layout, in-block corridor would not be in the layout block owner
                    Try
                        If Not ReportUtilities.IsCorridorInLayout(oCorridor) Then
                            Continue For
                        End If

                        Dim oCorridorInfo As New CCorridor

                        Dim currentBaseCount As Integer = 0
                        Dim currentSamplelineGroupCount As Integer = 0
                        Dim currentSamplelineCount As Integer = 0
                        Dim currentLinkCodesGroupCount As Integer = 0

                        If buildCorridorInfo(oCorridor, _
                                             oCorridorInfo, _
                                             currentBaseCount, _
                                             currentSamplelineGroupCount, _
                                             currentSamplelineCount, _
                                             currentLinkCodesGroupCount) Then
                            m_Corridors.Add(oCorridor.Name, oCorridorInfo)
                        End If

                        Dim sMsg As String = String.Format(LocalizedRes.Msg_No_Corridors_In_Drawing, ReportApplication.AeccXDocument.Name)
                        If currentBaseCount < 1 Then
                            ReportUtilities.ACADMsgBox(String.Format(LocalizedRes.CorridorSlopeStake_Msg_NoAlignment_In_Corridor, oCorridor.Name))
                        ElseIf currentSamplelineGroupCount < 1 Then
                            ReportUtilities.ACADMsgBox(String.Format(LocalizedRes.CorridorSlopeStake_Msg_NoLineGrp_In_Corridor, oCorridor.Name))
                        ElseIf currentLinkCodesGroupCount < 1 Then
                            ReportUtilities.ACADMsgBox(String.Format(LocalizedRes.CorridorSlopeStake_Msg_NoLine_In_Corridor, oCorridor.Name))
                        ElseIf currentSamplelineCount < 1 Then
                            ReportUtilities.ACADMsgBox(String.Format(LocalizedRes.CorridorSlopeStake_Msg_NoSampleLine_In_Corridor, oCorridor.Name))
                        Else
                            bSuccess = True
                        End If

                    Catch ex As Exception
                        Diagnostics.Debug.Assert(False, ex.Message)
                    End Try

                    If bSuccess = False And m_Corridors.ContainsKey(oCorridor.Name) = True Then
                        m_Corridors.Remove(oCorridor.Name)
                    End If
                End If
            Next
        End If
        CanOpen = False
        If bNoCorridors = True Or m_Corridors.Count = 0 Then
            Dim sMsg As String = String.Format(LocalizedRes.Msg_No_Corridors_In_Drawing, ReportApplication.AeccXDocument.Name)
            ReportUtilities.ACADMsgBox(sMsg)
        Else
            CanOpen = True
        End If

    End Function

    Private Sub ReportForm_CorridorSlopeStake_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Button_AddLink.Image = bp
        Button_AddLink.ImageAlign = ContentAlignment.MiddleCenter

        bp = My.Resources.Resources.Delete
        bp.MakeTransparent(Color.White)
        Button_Delete.Image = bp
        Button_Delete.ImageAlign = ContentAlignment.MiddleCenter

        'Init report settings
        InitReportSettings()

        Exit Sub
    End Sub

    Private Sub InitReportSettings()

        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)

        ctlStartStation = New CtrlStationTextBox
        ctlStartStation.Initialize(TextBox_StartStation)

        ctlEndtStation = New CtrlStationTextBox
        ctlEndtStation.Initialize(TextBox_EndStation)

        ctlCorridorListView = New CtrlCorridorListView
        ctlCorridorListView.haveMaterialList = True
        ctlCorridorListView.Initialize(ListView_Corridors, ctlStartStation, ctlEndtStation)

        ctlSaveReport = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)

        'Area/Volume Unit settings
        oAreaSettings = ReportApplication.AeccXDatabase.Settings.SampleLineSettings.AmbientSettings.AreaSettings
        oVolumeSettings = ReportApplication.AeccXDatabase.Settings.SampleLineSettings.AmbientSettings.VolumeSettings
    End Sub

    Private Function buildCorridorInfo(ByRef oCorridor As AeccRoadLib.IAeccCorridor, _
                                       ByRef oCorridorInfo As CCorridor, _
                                       ByRef baseLineCount As Integer, _
                                       ByRef samplelineGroupCount As Integer, _
                                       ByRef samplelineCount As Integer, _
                                       ByRef linkCodesGroupCount As Integer) As Boolean
        oCorridorInfo.mObject = oCorridor
        oCorridorInfo.mBaselines.Clear()

        baseLineCount = oCorridor.Baselines.Count

        For Each oBaseline As AeccRoadLib.IAeccBaseline In oCorridor.Baselines
            If oCorridorInfo.mBaselines.ContainsKey(oBaseline.Alignment.Name) = False Then
                Dim oBaselineInfo As CBaseline
                oBaselineInfo = New CBaseline

                Dim currentSamplelineGroupCount As Integer = 0
                Dim currentSamplelineCount As Integer = 0
                Dim currentLinkCodesGroupCount As Integer = 0

                If buildBaselineInfo(oBaseline, oBaselineInfo, currentSamplelineGroupCount, currentSamplelineCount, currentLinkCodesGroupCount) Then
                    oCorridorInfo.mBaselines.Add(oBaseline.Alignment.Name, oBaselineInfo)
                End If

                samplelineGroupCount += currentSamplelineGroupCount
                samplelineCount += currentSamplelineCount
                linkCodesGroupCount += currentLinkCodesGroupCount
            End If
        Next

        Return (oCorridorInfo.mBaselines.Count > 0)
    End Function

    Private Function buildBaselineInfo(ByRef oBaseline As AeccRoadLib.IAeccBaseline, _
                                       ByRef oBaselineInfo As CBaseline, _
                                       ByRef samplelineGroupCount As Integer, _
                                       ByRef samplelineCount As Integer, _
                                       ByRef linkCodesGroupCount As Integer) As Boolean
        oBaselineInfo.mObject = oBaseline
        oBaselineInfo.mSampleLineGroups.Clear()

        oBaselineInfo.mLinkCodesGroups.Clear()

        samplelineGroupCount = oBaseline.Alignment.SampleLineGroups.Count
        samplelineCount = 0
        linkCodesGroupCount = 0

        If samplelineGroupCount > 0 And oBaseline.BaselineRegions.Count > 0 Then
            For Each oSampleLineGroup As AeccLandLib.AeccSampleLineGroup In oBaseline.Alignment.SampleLineGroups
                Dim childSamplelineCount As Integer = oSampleLineGroup.SampleLines.Count
                samplelineCount += childSamplelineCount
                If childSamplelineCount > 0 Then
                    If oBaselineInfo.mSampleLineGroups.ContainsKey(oSampleLineGroup.Name) = False Then
                        Dim oSampleLineGroupInfo As CSampleLineGroup
                        oSampleLineGroupInfo = New CSampleLineGroup
                        oSampleLineGroupInfo.mObject = oSampleLineGroup
                        oBaselineInfo.mSampleLineGroups.Add(oSampleLineGroup.Name, oSampleLineGroupInfo)
                    End If
                End If
            Next

            Dim oSubAssembly As AeccRoadLib.AeccSubassembly
            For Each oSubAssembly In ReportApplication.AeccXRoadwayDocument.Subassemblies 'g_oRoadwayDocument.Subassemblies
                ' collect the link codes info
                Dim oRoadwayLinks As AeccRoadLib.AeccRoadwayLinks
                oRoadwayLinks = oSubAssembly.Links
                Dim oRoadwayLink As AeccRoadLib.AeccRoadwayLink
                For Each oRoadwayLink In oRoadwayLinks

                    Dim roadwayCodesCount As Integer = oRoadwayLink.RoadwayCodes.Count
                    linkCodesGroupCount += roadwayCodesCount

                    Dim iIndex As Integer
                    For iIndex = 0 To roadwayCodesCount - 1
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

                ' collect point codes info
                Dim oRoadwayPoints As AeccRoadLib.AeccRoadwayPoints = oSubAssembly.Points
                Dim oRoadwayPoint As AeccRoadLib.AeccRoadwayPoint

                For Each oRoadwayPoint In oRoadwayPoints

                    Dim oPointCodes As AeccRoadLib.AeccRoadwayCodes = oRoadwayPoint.RoadwayCodes
                    Dim sCode As String

                    For Each sCode In oPointCodes
                        If oBaselineInfo.mPointCodesGroups.Contains(sCode) = False Then
                            oBaselineInfo.mPointCodesGroups.Add(sCode)
                        End If
                    Next
                Next

            Next
        End If

        Return (oBaselineInfo.mSampleLineGroups.Count > 0 And oBaselineInfo.mLinkCodesGroups.Count > 0)
    End Function

    Private Sub Button_Delete_Click() Handles Button_Delete.Click

        If ListView_Corridors.SelectedItems.Count < 1 Then
            Exit Sub
        End If

        For Each item As ListViewItem In ctlCorridorListView.WinListView.SelectedItems
            Dim oCorridorInfo As CCorridor
            Dim oBaselineInfo As CBaseline
            Dim oLinkCodeInfo As CLinkCodes

            oCorridorInfo = m_Corridors.Item(item.Text)
            oBaselineInfo = oCorridorInfo.mBaselines.Item(ctlCorridorListView.AlignmentName(item))
            oLinkCodeInfo = oBaselineInfo.mLinkCodesGroups.Item(ctlCorridorListView.ListCode(item))
            oLinkCodeInfo.mAdded = False
            item.Remove()
        Next
    End Sub

    Private Sub Button_AddLink_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_AddLink.Click
        If Combo_LineGroup.Items.Count = 0 Or Combo_Link.Items.Count = 0 Then
            Exit Sub
        End If

        Dim oSampleLineGroupInfo As CSampleLineGroup
        Dim oSampleLineGroup As AeccLandLib.AeccSampleLineGroup
        oSampleLineGroupInfo = FindSampleLineGroupInfo(Combo_Corridor.SelectedItem.ToString(), _
                                                   Combo_Align.SelectedItem.ToString(), _
                                                   Combo_LineGroup.SelectedItem.ToString())
        oSampleLineGroup = oSampleLineGroupInfo.mObject

        Dim oLinkCodesInfo As CLinkCodes
        Dim sCodeName As String
        oLinkCodesInfo = FindLinkCodesInfo(Combo_Corridor.SelectedItem.ToString(), _
                                        Combo_Align.SelectedItem.ToString(), _
                                        Combo_Link.SelectedItem.ToString())
        sCodeName = oLinkCodesInfo.mCodeName

        If oLinkCodesInfo.mAdded = False Then
            m_Corridors.Item(Combo_Corridor.SelectedItem.ToString()).mMaterialList = Combo_Material_List.SelectedItem.ToString()
            ctlCorridorListView.AddListItem(oSampleLineGroup, Combo_Corridor.SelectedItem.ToString(), Combo_Align.SelectedItem.ToString(), Combo_Link.SelectedItem.ToString(), Combo_Material_List.SelectedItem.ToString())
            oLinkCodesInfo.mAdded = True
        Else
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorSlopeStake_Msg_Added)
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

    Private Sub Combo_LineGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_LineGroup.SelectedIndexChanged
        'clear material list
        Combo_Material_List.Items.Clear()
        'add None
        Combo_Material_List.Items.Add(LocalizedRes.CorridorSlopeStake_None)
        Combo_Material_List.SelectedIndex = 0

        Dim oSampleLineGroupInfo As CSampleLineGroup
        If Combo_LineGroup.Items.Count = 0 Then
            Exit Sub
        End If

        If Combo_LineGroup.SelectedIndex < 0 Then
            Exit Sub
        End If

        oSampleLineGroupInfo = FindSampleLineGroupInfo(Combo_Corridor.SelectedItem.ToString(), _
                                                   Combo_Align.SelectedItem.ToString(), _
                                                   Combo_LineGroup.SelectedItem.ToString())

        If oSampleLineGroupInfo.mObject.SampleLines.Count = 0 Then
            Button_AddLink.Enabled = False
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorSlopeStake_Msg_NoSampleLine)

        Else
            Button_AddLink.Enabled = True

            'get mapping names
            Dim sampleLineGroupId As ObjectId = Autodesk.AutoCAD.DatabaseServices.DBObject.FromAcadObject(oSampleLineGroupInfo.mObject)
            Using trans As Transaction = sampleLineGroupId.Database.TransactionManager.StartTransaction
                Dim oSampleLineGroupExt As SampleLineGroup = trans.GetObject(sampleLineGroupId, OpenMode.ForRead)
                Dim mappingNames() As String = oSampleLineGroupExt.GetQTOMappingNames()

                For Each name As String In mappingNames
                    Combo_Material_List.Items.Add(name)
                Next

                If mappingNames.Length > 0 Then
                    Combo_Material_List.SelectedIndex = 1
                End If
            End Using
        End If
    End Sub

    Private Sub Combo_Align_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Align.SelectedIndexChanged
        Dim oCorridorInfo As CCorridor
        Dim oBaselineInfo As CBaseline

        If Combo_Align.Items.Count = 0 Then
            Exit Sub
        End If

        If Combo_Align.SelectedIndex < 0 Then
            Exit Sub
        End If

        oCorridorInfo = m_Corridors.Item(Combo_Corridor.SelectedItem.ToString())
        oBaselineInfo = oCorridorInfo.mBaselines.Item(Combo_Align.SelectedItem.ToString())

        Combo_LineGroup.Items.Clear()
        For Each sSampleLineGroupName As String In oBaselineInfo.mSampleLineGroups.Keys
            Combo_LineGroup.Items.Add(sSampleLineGroupName)
        Next

        If Combo_LineGroup.Items.Count > 0 Then
            Combo_LineGroup.SelectedIndex = 0
        Else
            Button_AddLink.Enabled = False
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorSlopeStake_Msg_NoLineGrp)
        End If



        'Defect 1028432: Filter those Linkcodes which dosen't exist
        If oBaselineInfo.mIsLinkCodesFilterred = False Then
            Dim oSampleLineGroupInfo As CSampleLineGroup = _
                                    oBaselineInfo.mSampleLineGroups.Item(Combo_LineGroup.SelectedItem.ToString())
            FilterLinkCode(oBaselineInfo, oSampleLineGroupInfo)
        End If

        ' update the list of available links
        Combo_Link.Items.Clear()
        For Each sLinkCodeName As String In oBaselineInfo.mLinkCodesGroups.Keys
            Combo_Link.Items.Add(sLinkCodeName)
        Next

        If Combo_Link.Items.Count > 0 Then
            Combo_Link.SelectedIndex = 0
        Else
            Button_AddLink.Enabled = False
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorSlopeStake_Msg_NoLine)
        End If

        ' update the list of available points
        ComboBox_ROW_PointCode.Items.Clear()

        For Each sPointCode As String In oBaselineInfo.mPointCodesGroups
            ComboBox_ROW_PointCode.Items.Add(sPointCode)
        Next

        If ComboBox_ROW_PointCode.Items.Count > 0 Then
            ComboBox_ROW_PointCode.SelectedIndex = 0
        End If

    End Sub

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
            Button_AddLink.Enabled = False
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
        If ListView_Corridors.Items.Count = 0 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorSlopeStake_Msg_NoComponent)
            Exit Sub
        End If

        'step count 13 per corridor, EXTRACT_STEPS for extract data
        ctlProgressBar.ProgressBegin(ListView_Corridors.Items.Count * (EXTRACT_STEPS + 3))

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

        ctlProgressBar.ProgressEnd()

        ReportUtilities.OpenFileByDefaultBrowser(ctlSaveReport.SavePath)

    End Sub

    Private Sub WriteHtmlReport(ByVal fileName As String)
        Dim oReportHtml As New ReportWriter(fileName)
        '<html>
        oReportHtml.HtmlBegin()
        oReportHtml.RenderHeader(LocalizedRes.CorridorSlopeStake_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.CorridorSlopeStake_Html_Title, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        Dim inStartStation As Double
        Dim inEndStation As Double
        Dim oCorridorInfo As CCorridor
        Dim oBaselineInfo As CBaseline
        Dim oSampleLineGroupInfo As CSampleLineGroup
        Dim oLinkCodesInfo As CLinkCodes
        Dim sROWPointCode As String
        Dim materialListName As String

        For Each item As ListViewItem In ctlCorridorListView.WinListView.Items
            inStartStation = ctlCorridorListView.StartRawStation(item)
            inEndStation = ctlCorridorListView.EndRawStation(item)
            oCorridorInfo = m_Corridors.Item(ctlCorridorListView.CorridorName(item))
            oBaselineInfo = oCorridorInfo.mBaselines.Item(ctlCorridorListView.AlignmentName(item))
            oSampleLineGroupInfo = oBaselineInfo.mSampleLineGroups.Item(ctlCorridorListView.SampleLineGroupName(item))
            oLinkCodesInfo = oBaselineInfo.mLinkCodesGroups.Item(ctlCorridorListView.ListCode(item))
            materialListName = oCorridorInfo.mMaterialList

            Dim selectedIndex As Integer = ComboBox_ROW_PointCode.SelectedIndex
            sROWPointCode = ComboBox_ROW_PointCode.Items.Item(selectedIndex).ToString

            ctlProgressBar.PerformStep()
            AppendReport(oCorridorInfo.mObject, oBaselineInfo.mObject, _
                oSampleLineGroupInfo.mObject, inStartStation, inEndStation, _
                oReportHtml, oLinkCodesInfo.mCodeName, materialListName, sROWPointCode)
            ctlProgressBar.PerformStep()
        Next

        '</body>
        oReportHtml.BodyEnd()

        '</html>
        oReportHtml.HtmlEnd()
    End Sub

    ' Append one item's report to html file
    Private Sub AppendReport(ByVal oCorridor As AeccRoadLib.IAeccCorridor, _
                            ByVal oBaseline As AeccRoadLib.IAeccBaseline, _
                            ByVal oSampleLineGroup As AeccLandLib.IAeccSampleLineGroup, _
                            ByVal stationStart As Double, _
                            ByVal stationEnd As Double, _
                            ByVal oHtmlWriter As ReportWriter, _
                            ByVal sLinkCode As String, _
                            ByVal sMaterialListName As String, _
                            ByVal sROWPointCode As String)
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

        str = LocalizedRes.CorridorSlopeStake_Html_LineGrp
        str += " " + oSampleLineGroup.Name
        oHtmlWriter.RenderLine(str)

        str = LocalizedRes.CorridorSlopeStake_Html_Link
        str += " " + sLinkCode
        oHtmlWriter.RenderLine(str)

        str = String.Format(LocalizedRes.Alignment_Html_Station_Range, _
                            ReportUtilities.GetStationStringWithDerived(oBaseAlignment, stationStart), _
                            ReportUtilities.GetStationStringWithDerived(oBaseAlignment, stationEnd))
        oHtmlWriter.RenderLine(str)

        '    extract data
        CorridorSlopeStake_ExtractData.ExtractData(oBaseline, _
            oSampleLineGroup, stationStart, stationEnd, sLinkCode, sROWPointCode, sMaterialListName, ctlProgressBar.ProgressBar, EXTRACT_STEPS)

        ctlProgressBar.PerformStep()

        ' turn on/off the graphics
        Dim bDisplayImage As Boolean
        bDisplayImage = CheckBox_AllowGraphics.Checked

        For Each key As Double In CorridorSlopeStake_ExtractData.SlopeStakeData.Keys

            If bDisplayImage Then

                ' Generate graph to file
                Dim graph As New SSRGraphics

                Dim tempSectData As Slope_SectionData = _
                    CorridorSlopeStake_ExtractData.SlopeStakeData(key)
                Dim tempPtData As Slope_PointData
                For Each offsetKey As Double In tempSectData.LeftDatas.Keys()
                    tempPtData = tempSectData.LeftDatas(offsetKey)
                    graph.AddCodePoint(tempPtData)
                Next
                graph.AddCodePoint(tempSectData.CenterLineInfo)
                For Each offsetKey As Double In tempSectData.RightDatas.Keys()
                    tempPtData = tempSectData.RightDatas(offsetKey)
                    graph.AddCodePoint(tempPtData)
                Next
                For Each tempPtData In tempSectData.ROWPoints
                    graph.AddCodePoint(tempPtData)
                Next

                graph.SetLinks(tempSectData.LinksOnGraph)
                graph.SetEGDatas(tempSectData.EGDatas)

                Dim tempPath As String = System.IO.Path.GetTempPath()

                Dim curStationString As String
                curStationString = oSampleLineGroup.Parent.GetStationStringWithEquations(key)

                Dim strImgFileName As String = tempPath + sLinkCode + "_" + curStationString + "_" + DateTime.Now.ToFileTime().ToString()
                Dim strExt As String = "png"
                graph.CreateGraph(strImgFileName, _
                                  strExt)

                ' Add graph above current table
                oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Border, "0")
                oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "100%")
                oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Src, "file:///" + strImgFileName + "." + strExt)
                oHtmlWriter.ImgBegin()
                oHtmlWriter.ImgEnd()

            End If

            AppendSingleTable(oBaseAlignment, key, oHtmlWriter)
        Next

        oHtmlWriter.RenderBr()
    End Sub

    Public Function FormatAreaSettings(ByVal dArea As Double) As String


        FormatAreaSettings = ReportFormat.FormatArea(dArea, _
                                            oAreaSettings.Unit.Value, _
                                            oAreaSettings.Precision.Value, _
                                            oAreaSettings.Rounding.Value, _
                                            oAreaSettings.Sign.Value)
    End Function

    Private Function FormatVolumeSettings(ByVal dVolume As Double) As String

        FormatVolumeSettings = ReportFormat.FormatVolume(dVolume, oVolumeSettings.Unit.Value, _
                oVolumeSettings.Precision.Value, oVolumeSettings.Rounding.Value, oVolumeSettings.Sign.Value)
    End Function

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

        Dim bDisplayCutFillArea As Boolean
        bDisplayCutFillArea = CheckBox_DisplayCutFill.Checked
        If sData.HaveMaterialInfo And bDisplayCutFillArea Then
            str = String.Format(LocalizedRes.CorridorSlopeStake_Html_Cut_Area, FormatAreaSettings(sData.CutArea))
            oHtmlWriter.RenderLine(str, True)
            str = String.Format(LocalizedRes.CorridorSlopeStake_Html_Fill_Area, FormatAreaSettings(sData.FillArea))
            oHtmlWriter.RenderLine(str, True)
            str = String.Format(LocalizedRes.CorridorSlopeStake_Html_Cumulative_Net_Volume, FormatVolumeSettings(sData.CumulativeNetVolume))
            oHtmlWriter.RenderLine(str, True)
        End If

        'table begin
        oHtmlWriter.TableBegin("1")


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

        Dim nLeft As Integer
        If leftPtCount < 5 Then
            nLeft = leftPtCount
        Else
            nLeft = 5
        End If
        Dim nRight As Integer
        If rightPtCount < 5 Then
            nRight = rightPtCount
        Else
            nRight = 5
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
                Dim leftDataArr(0 To nLeft) As String
                Dim rightDataArr(0 To nRight) As String
                Dim clData As String = ""

                'set end data at the first line
                If line = 0 Then
                    If row = 0 Then
                        leftDataArr(nLeft) = sData.LeftEndInfo.EndType
                        rightDataArr(nRight) = sData.RightEndInfo.EndType
                        clData = sData.CenterLineInfo.mCodes
                    ElseIf row = 1 Then
                        leftDataArr(nLeft) = sData.LeftEndInfo.DeltaOffset
                        rightDataArr(nRight) = sData.RightEndInfo.DeltaOffset
                        clData = sData.CenterLineInfo.mOffsetString
                    ElseIf row = 2 Then
                        leftDataArr(nLeft) = sData.LeftEndInfo.EndSlope
                        rightDataArr(nRight) = sData.RightEndInfo.EndSlope
                        clData = sData.CenterLineInfo.mElevationString
                    End If
                End If

                'fill cell column data, total 5 columns for code points per line
                For col As Integer = 0 To 4
                    Dim nowPtIndex As Integer
                    nowPtIndex = 5 * line + col

                    If col < nLeft Then
                        If nowPtIndex < leftPtCount Then
                            fillPointColumn(row, sData.LeftDatas.Item(leftKeys(nowPtIndex)), leftDataArr(col))
                        Else
                            leftDataArr(col) = ""
                        End If
                    End If

                    If col < nRight Then
                        If nowPtIndex < rightPtCount Then
                            fillPointColumn(row, sData.RightDatas.Item(rightKeys(nowPtIndex)), rightDataArr(col))
                        Else
                            rightDataArr(col) = ""
                        End If
                    End If
                Next

                'left and right data fill finished, write them
                Dim dataArr(0 To 2 + nLeft + nRight) As String
                Array.Reverse(leftDataArr)
                leftDataArr.CopyTo(dataArr, 0)
                dataArr(nLeft + 1) = clData
                rightDataArr.CopyTo(dataArr, nLeft + 2)
                WriteOneTableRow(dataArr, oHtmlWriter)
            Next

            'separate row
            Dim arrBlank(0 To 2 + nLeft + nRight) As String
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
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_SLOPE_GRADE_REPORT)
    End Sub

    Private Sub CheckBox_AllowGraphics_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox_AllowGraphics.CheckedChanged

        GroupBox_ROW_Settings.Enabled = CheckBox_AllowGraphics.Checked

    End Sub

    Private Sub CheckBox_DisplayROW_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox_DisplayROW.CheckedChanged

        ComboBox_ROW_PointCode.Enabled = CheckBox_DisplayROW.Checked

    End Sub

End Class
