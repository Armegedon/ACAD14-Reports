' -----------------------------------------------------------------------------
' <copyright file="ReportForm_CorridorVolume.vb" company="Autodesk">
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
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.Civil.DatabaseServices
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

'Imports Autodesk.Civil.AeccTransExt

Friend Class ReportForm_CorridorVolume

    Private Const EXTRACT_STEPS As Integer = 30

    'member variable
    Private m_alignments As Collections.Generic.SortedList(Of String, AeccLandLib.IAeccAlignment)

    Private m_sampleLineGroups As Collections.Generic.SortedList(Of String, AeccLandLib.IAeccSampleLineGroup)

    Private m_massHaulIds As Collections.Generic.List(Of System.Collections.Generic.KeyValuePair(Of ObjectId, ObjectId))

    'Private m_Corridors As New SortedDictionary(Of String, CCorridor)

    'Private ctlCorridorListView As CtrlCorridorListView
    Private ctlStartStation As CtrlStationTextBox
    Private ctlEndStation As CtrlStationTextBox
    Private ctlProgressBar As CtrlProgressBar
    Private ctlSaveReport As CtrlSaveReportFile
    Private oVolumeSettings As AeccLandLib.AeccSettingsVolume
    Private oAreaSettings As AeccLandLib.AeccSettingsArea

    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()
        Dim reportForm As New ReportForm_CorridorVolume
        If reportForm.CanOpen() = True Then
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(reportForm)
        End If
    End Sub

    Private Function CanOpen() As Boolean
        'get alignments
        m_alignments = New Collections.Generic.SortedList(Of String, AeccLandLib.IAeccAlignment)
        'add siteliess alignments
        For Each oAlignment As AeccLandLib.AeccAlignment In ReportApplication.AeccXDocument.AlignmentsSiteless
            m_alignments.Add(oAlignment.Name, oAlignment)
        Next
        'add alignments from Site
        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            For Each oAlignment As AeccLandLib.AeccAlignment In oSite.Alignments
                m_alignments.Add(oAlignment.Name, oAlignment)
            Next
        Next

        m_sampleLineGroups = New Collections.Generic.SortedList(Of String, AeccLandLib.IAeccSampleLineGroup)

        Dim bNoAlignments As Boolean = False

        Dim nAlignmentCount As Integer = 0
        Dim nSamplelineGroupCount As Integer = 0
        Dim nMaterialListCount As Integer = 0

        Try
            nAlignmentCount = m_alignments.Count
            If nAlignmentCount < 1 Then
                bNoAlignments = True
            Else
                'initial alignments
                For Each iAlignment As AeccLandLib.IAeccAlignment In m_alignments.Values
                    Try
                        Dim slgs As Autodesk.AECC.Interop.Land.IAeccSampleLineGroups = iAlignment.SampleLineGroups
                        nSamplelineGroupCount += slgs.Count

                        For Each slg As Autodesk.AECC.Interop.Land.IAeccSampleLineGroup In slgs
                            'calculate material list count

                            Dim sampleLineGroupId As ObjectId = Autodesk.AutoCAD.DatabaseServices.DBObject.FromAcadObject(slg)
                            Using trans As Transaction = sampleLineGroupId.Database.TransactionManager.StartTransaction
                                Dim oSampleLineGroup As SampleLineGroup
                                oSampleLineGroup = trans.GetObject(sampleLineGroupId, OpenMode.ForRead)
                                'Dim oSampleLineGroupExt As SampleLineGroupExt = New SampleLineGroupExt(oSampleLineGroup)

                                'Init Material lists
                                Dim mappingNames() As String = oSampleLineGroup.GetQTOMappingNames()
                                nMaterialListCount += mappingNames.Length
                            End Using
                        Next
                    Catch ex As Exception
                        'skip the exception of sample line groups
                    End Try
                Next
            End If
        Catch ex As Exception
            bNoAlignments = True
        End Try

        CanOpen = False
        If bNoAlignments = True Then
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorVolume_Msg_NoAlignment)
        ElseIf nSamplelineGroupCount < 1 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorVolume_Msg_NoSampleLineGroup)
        ElseIf nMaterialListCount < 1 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorVolume_Msg_NoMaterialList)
        Else
            CanOpen = True
        End If

        m_massHaulIds = New Collections.Generic.List(Of System.Collections.Generic.KeyValuePair(Of ObjectId, ObjectId))

    End Function

    Private Sub ReportForm_CorridorVolume_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'initial Select report Components
        'init button bmp
        Dim bp As Drawing.Bitmap = My.Resources.Resources.PickEntityBmp
        bp.MakeTransparent(Color.White)
        Button_SelectAlign.Image = bp
        Button_SelectAlign.ImageAlign = ContentAlignment.MiddleCenter

        'Init report settings
        InitReportSettings()

        ' add alignments to combo
        For Each iAlignment As AeccLandLib.IAeccAlignment In m_alignments.Values
            Combo_Align.Items.Add(iAlignment.Name)
        Next
        'Corridor count should more then 0, because we have test it before dialog open
        If m_alignments.Count > 0 Then
            Combo_Align.SelectedIndex = 0
        End If

        Exit Sub
    End Sub

    Private Sub InitReportSettings()

        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)

        ctlStartStation = New CtrlStationTextBox
        ctlStartStation.Initialize(TextBox_StartStation)

        ctlEndStation = New CtrlStationTextBox
        ctlEndStation.Initialize(TextBox_EndStation)

        ctlSaveReport = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)

        'Area/Volume Unit settings
        oAreaSettings = ReportApplication.AeccXDatabase.Settings.SampleLineSettings.AmbientSettings.AreaSettings
        oVolumeSettings = ReportApplication.AeccXDatabase.Settings.SampleLineSettings.AmbientSettings.VolumeSettings

    End Sub


    Private Sub Combo_LineGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_LineGroup.SelectedIndexChanged
        ClearMaterialListAndMassHaulView()

        'Dim oSampleLineGroupInfo As CSampleLineGroup
        If Combo_LineGroup.Items.Count = 0 Then
            Exit Sub
        Else

            Dim oLineGroup As AeccLandLib.AeccSampleLineGroup = m_sampleLineGroups.Values(Combo_LineGroup.SelectedIndex)
            Dim sampleLineGroupId As ObjectId = Autodesk.AutoCAD.DatabaseServices.DBObject.FromAcadObject(oLineGroup)
            Using trans As Transaction = sampleLineGroupId.Database.TransactionManager.StartTransaction
                Dim oSampleLineGroup As SampleLineGroup
                oSampleLineGroup = trans.GetObject(sampleLineGroupId, OpenMode.ForRead)


                'Init Material lists
                Dim mappingNames() As String = oSampleLineGroup.GetQTOMappingNames()
                For Each name As String In mappingNames
                    Combo_Material_List.Items.Add(name)
                Next
                If mappingNames.Length > 0 Then
                    Combo_Material_List.SelectedIndex = 0
                End If

                'Init mass haul views
                Dim idsMassHaulView As Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection = oSampleLineGroup.GetMassHaulViewIds()
                For Each idMassHaulView As ObjectId In idsMassHaulView
                    Dim idMassHaulLine As ObjectId
                    Using transInner As Transaction = sampleLineGroupId.Database.TransactionManager.StartTransaction
                        Dim oMassHaulView As MassHaulView
                        oMassHaulView = transInner.GetObject(idMassHaulView, OpenMode.ForRead)
                        idMassHaulLine = oMassHaulView.MassHaulLineId
                        If Not idMassHaulLine.IsErased And idMassHaulLine.IsValid Then
                            m_massHaulIds.Add(New System.Collections.Generic.KeyValuePair(Of ObjectId, ObjectId)(idMassHaulView, idMassHaulLine))
                            Combo_MassHaulView.Items.Add(oMassHaulView.Name)
                        End If
                    End Using

                Next

                If m_massHaulIds.Count > 0 Then
                    Combo_MassHaulView.SelectedIndex = 1
                End If

            End Using

        End If

    End Sub

    Private Sub stationsOfSampleLineGroup(ByVal oSampleLineGroup As AeccLandLib.IAeccSampleLineGroup, _
                                      ByRef startStation As String, ByRef endStation As String)
        Try
            Dim dStartStation, dEndStation As Double
            dStartStation = oSampleLineGroup.SampleLines.Item(0).Station
            dEndStation = dStartStation
            For Each oSampleLine As AeccLandLib.IAeccSampleLine In oSampleLineGroup.SampleLines
                If oSampleLine.Station > dEndStation Then
                    dEndStation = oSampleLine.Station
                End If
                If oSampleLine.Station < dStartStation Then
                    dStartStation = oSampleLine.Station
                End If
            Next
            startStation = ReportUtilities.GetStationStringWithDerived(oSampleLineGroup.Parent, dStartStation)
            endStation = ReportUtilities.GetStationStringWithDerived(oSampleLineGroup.Parent, dEndStation)
        Catch ex As Exception
            Diagnostics.Debug.Assert(False, ex.Message)
        End Try
    End Sub

    Private Sub ClearMaterialListAndMassHaulView()
        'clear material list
        Combo_Material_List.Items.Clear()


        'clear mass haul views
        Combo_MassHaulView.Items.Clear()
        m_massHaulIds.Clear()
        'add None
        Combo_MassHaulView.Items.Add(LocalizedRes.CorridorVolume_None)
        Combo_MassHaulView.SelectedIndex = 0
    End Sub

    Private Sub Combo_Align_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Align.SelectedIndexChanged
        Dim iAlignment As AeccLandLib.IAeccAlignment = m_alignments.Values(Combo_Align.SelectedIndex)

        '--------------------------------
        ' sample line groups

        'clear old items
        Combo_LineGroup.Items.Clear()
        m_sampleLineGroups.Clear()

        'add new items
        Try
            For Each oSampleLineGroup As AeccLandLib.AeccSampleLineGroup In iAlignment.SampleLineGroups
                m_sampleLineGroups.Add(oSampleLineGroup.Name, oSampleLineGroup)
            Next
            'Add the sorted names to combo box
            For Each oSampleLineGroup As AeccLandLib.AeccSampleLineGroup In m_sampleLineGroups.Values
                Combo_LineGroup.Items.Add(oSampleLineGroup.Name)
            Next
        Catch ex As Exception

        End Try

        If m_sampleLineGroups.Count > 0 Then
            Combo_LineGroup.SelectedIndex = 0
        Else
            ClearMaterialListAndMassHaulView()
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorVolume_Msg_NoSampleLineGroup)
        End If

        Dim oStationSetting As AeccLandLib.AeccSettingsStation = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.StationSettings

        '------------------------------------
        ' init stations controls

        ctlStartStation.TextBox.Enabled = False
        ctlEndStation.TextBox.Enabled = False
        'NumericUpDown_StationInc.Enabled = False
        ctlStartStation.Alignment = Nothing
        ctlEndStation.Alignment = Nothing

        ctlStartStation.EquationStation = ""
        ctlEndStation.EquationStation = ""
        TextBox_StartStation.Enabled = True
        TextBox_EndStation.Enabled = True

        ctlStartStation.Alignment = iAlignment
        ctlEndStation.Alignment = iAlignment
        ctlStartStation.EquationStation = ReportUtilities.GetStationStringWithDerived(iAlignment, iAlignment.StartingStation)
        ctlEndStation.EquationStation = ReportUtilities.GetStationStringWithDerived(iAlignment, iAlignment.EndingStation)
        ctlStartStation.RawStation = ReportFormat.RoundVal(iAlignment.StartingStation, oStationSetting.Precision.Value, oStationSetting.Rounding.Value)
        ctlEndStation.RawStation = ReportFormat.RoundVal(iAlignment.EndingStation, oStationSetting.Precision.Value, oStationSetting.Rounding.Value)

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

    Private Sub Button_SelectAlign_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_SelectAlign.Click
        Dim oAlignment As AeccLandLib.AeccAlignment
        oAlignment = PickAlignment()
        If Not oAlignment Is Nothing And Combo_Align.Items.Count > 0 Then
            Combo_Align.Text = oAlignment.Name
        End If
    End Sub

    Private Function PickAlignment() As AeccLandLib.AeccAlignment
        'do not hide form, autoCAD will hide the form automatically.

        PickAlignment = Nothing
        On Error Resume Next

        'Prompt user to select a point or block reference:
        Dim prEntityOpts As PromptEntityOptions
        prEntityOpts = New PromptEntityOptions(LocalizedRes.PointsStaOffset_Form_SelectAlign) '"Select a Alignment")
        prEntityOpts.SetRejectMessage(vbNewLine + _
            LocalizedRes.PointsStaOffset_Form_Reject + vbNewLine) '"Please select an alignment.
        prEntityOpts.AddAllowedClass(GetType(Autodesk.Civil.DatabaseServices.Alignment), False)

        Dim prEntityRes As PromptEntityResult
        prEntityRes = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.GetEntity(prEntityOpts)

        Dim founded As Boolean = False
        If (prEntityRes.Status = PromptStatus.OK) Then
            For Each oAlign As AeccLandLib.AeccAlignment In ReportApplication.AeccXDocument.AlignmentsSiteless
                If ReportUtilities.CompareObjectID(oAlign.ObjectID, prEntityRes.ObjectId) Then
                    PickAlignment = oAlign
                    founded = True
                    Exit For
                End If
            Next

            If Not founded Then
                For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
                    For Each oAlign As AeccLandLib.AeccAlignment In oSite.Alignments
                        If ReportUtilities.CompareObjectID(oAlign.ObjectID, prEntityRes.ObjectId) Then
                            PickAlignment = oAlign
                            Exit For
                        End If
                    Next
                Next
            End If
        End If
    End Function

    Private Sub Button_CreateReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_CreateReport.Click
        If Combo_Material_List.Items.Count = 0 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.CorridorVolume_Msg_NoMaterialList)
            Exit Sub
        End If

        'step count 13 , EXTRACT_STEPS for extract data
        ctlProgressBar.ProgressBegin(EXTRACT_STEPS + 3)

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
        oReportHtml.RenderHeader(LocalizedRes.CorridorVolume_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.CorridorVolume_Html_Title, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        Dim oAlignment As AeccLandLib.IAeccAlignment
        Dim oSampleLineGroup As AeccLandLib.IAeccSampleLineGroup
        Dim sMaterialListName As String
        Dim sMassHaulViewName As String = String.Empty
        Dim oidMassHaulView As ObjectId
        Dim oidMassHaulLine As ObjectId
        Dim inStartStation As Double
        Dim inEndStation As Double
        Dim bExtractImage As Boolean = True

        'Dim oBaselineInfo As CBaseline
        'Dim oSampleLineGroupInfo As CSampleLineGroup
        'Dim oLinkCodesInfo As CLinkCodes

        inStartStation = ctlStartStation.RawStation
        inEndStation = ctlEndStation.RawStation
        oAlignment = m_alignments.Item(Combo_Align.SelectedItem)
        oSampleLineGroup = m_sampleLineGroups.Item(Combo_LineGroup.SelectedItem)
        sMaterialListName = Combo_Material_List.SelectedItem
        If Combo_MassHaulView.SelectedIndex = 0 Then
            'no mass haul image
            bExtractImage = False
        Else
            sMassHaulViewName = Combo_MassHaulView.SelectedItem
            oidMassHaulView = m_massHaulIds.Item(Combo_MassHaulView.SelectedIndex - 1).Key
            oidMassHaulLine = m_massHaulIds.Item(Combo_MassHaulView.SelectedIndex - 1).Value
        End If

        ctlProgressBar.PerformStep()
        AppendReport(oAlignment, oSampleLineGroup, sMaterialListName, _
                     sMassHaulViewName, oidMassHaulView, oidMassHaulLine, inStartStation, _
                     inEndStation, bExtractImage, oReportHtml)
        ctlProgressBar.PerformStep()

        '</body>
        oReportHtml.BodyEnd()

        '</html>
        oReportHtml.HtmlEnd()
    End Sub

    ' Append one item's report to html file
    Private Sub AppendReport(ByVal oAlignment As AeccLandLib.IAeccAlignment, _
                            ByVal oSampleLineGroup As AeccLandLib.IAeccSampleLineGroup, _
                            ByVal sMaterialListName As String, _
                            ByVal sMassHaulViewName As String, _
                            ByVal oidMassHaulView As ObjectId, _
                            ByVal oidMassHaulLine As ObjectId, _
                            ByVal stationStart As Double, _
                            ByVal stationEnd As Double, _
                            ByVal bExtractImage As Boolean, _
                            ByVal oHtmlWriter As ReportWriter)
        Dim str As String

        ctlProgressBar.PerformStep()

        oHtmlWriter.RenderHr()

        ' Alignment
        str = String.Format(LocalizedRes.CorridorVolume_Html_Alignment, oAlignment.Name)
        oHtmlWriter.RenderLine(str)

        ' Sample line group
        str = String.Format(LocalizedRes.CorridorVolume_Html_SampleLineGroup, oSampleLineGroup.Name)
        oHtmlWriter.RenderLine(str)

        ' start station
        str = String.Format(LocalizedRes.CorridorVolume_Html_StartStation, ReportUtilities.GetStationStringWithDerived(oAlignment, stationStart))
        oHtmlWriter.RenderLine(str)

        ' end station
        str = String.Format(LocalizedRes.CorridorVolume_Html_EndStation, ReportUtilities.GetStationStringWithDerived(oAlignment, stationEnd))
        oHtmlWriter.RenderLine(str)

        If bExtractImage Then
            ' mass haul diagram
            str = String.Format(LocalizedRes.CorridorVolume_Html_MassHaulDiagram, sMassHaulViewName)
            oHtmlWriter.RenderLine(str)

            Dim tempPath As String = System.IO.Path.GetTempPath()
            Dim strImgFileName As String = tempPath + "_" + DateTime.Now.ToFileTime().ToString()

            Dim entExtents As Extents3d = Nothing
            Using trans As Transaction = oidMassHaulView.Database.TransactionManager.StartTransaction
                Dim oMassHaulView As MassHaulView
                oMassHaulView = trans.GetObject(oidMassHaulView, OpenMode.ForRead)
                entExtents = oMassHaulView.GeometricExtents
            End Using

            Dim oids As ObjectIdCollection = New ObjectIdCollection
            oids.Add(oidMassHaulView)
            oids.Add(oidMassHaulLine)

            Dim capture As Autodesk.Civil.AeccUiMgd.EntityImageCapture = New Autodesk.Civil.AeccUiMgd.EntityImageCapture(oids)
            capture.SetImageFormat(Autodesk.Civil.AeccUiMgd.ImageFormatType.PNG)

            Dim dWidth As Double = entExtents.MaxPoint.X - entExtents.MinPoint.X
            Dim dHeight As Double = entExtents.MaxPoint.Y - entExtents.MinPoint.Y

            'capture.AdjustExtents(-5, -10, -5, -10)
            capture.CaptureImage(strImgFileName, dWidth * 2, dHeight * 2)

            ' Add graph above current table
            oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Border, "0")
            oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Width, "100%")
            oHtmlWriter.AddAttribute(System.Web.UI.HtmlTextWriterAttribute.Src, "file:///" + strImgFileName + ".PNG")
            oHtmlWriter.ImgBegin()
            oHtmlWriter.ImgEnd()
        End If

        Dim sAreaUnit As String = ReportFormat.AreaUnitString(oAreaSettings.Unit.Value)
        Dim sVolumeUnit As String = ReportFormat.VolumeUnitString(oVolumeSettings.Unit.Value)

        'table begin
        oHtmlWriter.TableBegin("1")

        oHtmlWriter.TrBegin()
        oHtmlWriter.RenderTd(LocalizedRes.CorridorVolume_Html_Station, True)
        oHtmlWriter.RenderTd(String.Format(LocalizedRes.CorridorVolume_Html_CutArea, sAreaUnit), True)
        oHtmlWriter.RenderTd(String.Format(LocalizedRes.CorridorVolume_Html_CutVolume, sVolumeUnit), True)
        oHtmlWriter.RenderTd(String.Format(LocalizedRes.CorridorVolume_Html_ReusableVolume, sVolumeUnit), True)
        oHtmlWriter.RenderTd(String.Format(LocalizedRes.CorridorVolume_Html_FillArea, sAreaUnit), True)
        oHtmlWriter.RenderTd(String.Format(LocalizedRes.CorridorVolume_Html_FillVolume, sVolumeUnit), True)
        oHtmlWriter.RenderTd(String.Format(LocalizedRes.CorridorVolume_Html_CumCutVolume, sVolumeUnit), True)
        oHtmlWriter.RenderTd(String.Format(LocalizedRes.CorridorVolume_Html_CumReusableVolume, sVolumeUnit), True)
        oHtmlWriter.RenderTd(String.Format(LocalizedRes.CorridorVolume_Html_CumFillVolume, sVolumeUnit), True)
        oHtmlWriter.RenderTd(String.Format(LocalizedRes.CorridorVolume_Html_CumNetVol, sVolumeUnit), True)
        oHtmlWriter.TrEnd()

        Dim sampleLineGroupId As ObjectId = Autodesk.AutoCAD.DatabaseServices.DBObject.FromAcadObject(oSampleLineGroup)

        Dim qtoSectionalResult() As Autodesk.Civil.DatabaseServices.QTOSectionalResult = Nothing

        Dim oSampleLines As AeccLandLib.AeccSampleLines = oSampleLineGroup.SampleLines

        Using trans As Transaction = sampleLineGroupId.Database.TransactionManager.StartTransaction
            'Open SampleLineGroup
            Dim omgdSampleLineGroup As SampleLineGroup
            omgdSampleLineGroup = trans.GetObject(sampleLineGroupId, OpenMode.ForRead)

            'Get volume info
            Dim mappingGuid As System.Guid = omgdSampleLineGroup.GetMappingGuid(sMaterialListName)
            Dim qtoResult As Autodesk.Civil.DatabaseServices.QuantityTakeoffResult = Nothing
            qtoResult = omgdSampleLineGroup.GetTotalVolumeResultDataForMaterialList(mappingGuid)
            qtoSectionalResult = qtoResult.GetResultsAlongSampleLines()
            Debug.Assert(qtoSectionalResult.Length = oSampleLines.Count)
        End Using

        Dim index As Integer = 0
        For Each oSampleLine As AeccLandLib.AeccSampleLine In oSampleLines
            If index >= qtoSectionalResult.Length Then
                Exit For
            End If

            Dim qtoAtStatoin As Autodesk.Civil.DatabaseServices.QTOSectionalResult = qtoSectionalResult(index)
            Dim areaAtStation As Autodesk.Civil.DatabaseServices.QTOAreaResult = qtoAtStatoin.AreaResult
            Dim volumeAtStation As Autodesk.Civil.DatabaseServices.QTOVolumeResult = qtoAtStatoin.VolumeResult

            oHtmlWriter.TrBegin()
            oHtmlWriter.RenderTd(ReportUtilities.GetStationStringWithDerived(oAlignment, oSampleLine.Station))

            oHtmlWriter.RenderTd(FormatAreaSettings(areaAtStation.CutArea)) 'cut area
            oHtmlWriter.RenderTd(FormatVolumeSettings(volumeAtStation.IncrementalCutVolume)) ' Cut volume
            oHtmlWriter.RenderTd(FormatVolumeSettings(volumeAtStation.IncrementalUsableVolume)) ' reusable volume

            oHtmlWriter.RenderTd(FormatAreaSettings(areaAtStation.FillArea)) 'fill area
            oHtmlWriter.RenderTd(FormatVolumeSettings(volumeAtStation.IncrementalFillVolume)) 'fill volume
            oHtmlWriter.RenderTd(FormatVolumeSettings(volumeAtStation.CumulativeCutVolume)) 'cumulative cut vol
            oHtmlWriter.RenderTd(FormatVolumeSettings(volumeAtStation.CumulativeUsableVolume)) 'cumulative resuable vol

            oHtmlWriter.RenderTd(FormatVolumeSettings(volumeAtStation.CumulativeFillVolume)) 'cumulative fill vol
            oHtmlWriter.RenderTd(FormatVolumeSettings(volumeAtStation.CumulativeCutVolume - volumeAtStation.CumulativeFillVolume)) 'net volume

            oHtmlWriter.TrEnd()

            index = index + 1
        Next

        oHtmlWriter.TableEnd()

        oHtmlWriter.RenderBr()
    End Sub

    Public Function FormatAreaSettings(ByVal dArea As Double) As String


        FormatAreaSettings = ReportFormat.FormatArea(dArea, _
                                            oAreaSettings.Unit.Value, _
                                            oAreaSettings.Precision.Value, _
                                            oAreaSettings.Rounding.Value, _
                                            oAreaSettings.Sign.Value, _
                                            False)
    End Function

    Private Function FormatVolumeSettings(ByVal dVolume As Double) As String

        FormatVolumeSettings = ReportFormat.FormatVolume(dVolume, _
                                                         oVolumeSettings.Unit.Value, _
                                                         oVolumeSettings.Precision.Value, _
                                                         oVolumeSettings.Rounding.Value, _
                                                         oVolumeSettings.Sign.Value, _
                                                         False)
    End Function


    Private Sub openHelp()
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_SLOPE_GRADE_REPORT)
    End Sub


End Class
