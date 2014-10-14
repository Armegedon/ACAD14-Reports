' -----------------------------------------------------------------------------
' <copyright file="ReportForm_ParcelMapCheck.vb" company="Autodesk">
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
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports AeccSurvey = Autodesk.AECC.Interop.Survey

Imports Autodesk.AutoCAD.EditorInput
Imports LocalizedRes = Report.My.Resources.LocalizedStrings


Friend Class ReportForm_ParcelMapCheck

    Private Class ParcelInfo
        Public Obj As AeccLandLib.AeccParcel
        Public CounterWiseClock As Boolean
        Public AcrossChord As Boolean
    End Class

    Private m_oFigureArr() As AeccSurvey.AeccSurveyFigure
    Private m_bCheckParcle As Boolean

    Private ctlProgressBar As CtrlProgressBar
    Private ctlSavePath As CtrlSaveReportFile

    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()

        Dim rptDlg As New ReportForm_ParcelMapCheck
        If rptDlg.ReadyToOpen() = True Then
            ReportUtilities.RunModalDialog(rptDlg)
        End If
    End Sub

    Private Function ReadyToOpen() As Boolean

        ReadyToOpen = True

        Dim nCount As Long
        nCount = 0
        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            nCount = nCount + oSite.Parcels.Count
        Next

        If nCount < 1 Then
            Dim sMsg As String = String.Format(LocalizedRes.Msg_No_Parcels_In_Drawing, ReportApplication.AeccXDocument.Name)
            ReportUtilities.ACADMsgBox(sMsg)
            ReadyToOpen = False
        End If
    End Function

    Private Sub ReportForm_ParcelMapCheck_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'init list of parcels
        InitParcelListView()

        'init analysis
        Text_PobXCoord.Text = "0.000"
        Text_PobYCoord.Text = "0.000"

        'init settings
        InitSettings()

        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)
        ctlSavePath = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)

    End Sub

    Private Sub InitSettings()
        Dim ambSettings As AeccLandLib.AeccSettingsAmbient
        ambSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings

        Dim dirFormat As String
        Select Case ambSettings.DirectionSettings.Format.Value
            Case AeccLandLib.AeccFormatType.aeccFormatDecimal
                dirFormat = LocalizedRes.ParceMapCheck_Form_Settings_Format_D '"(d)"
            Case AeccLandLib.AeccFormatType.aeccFormatDecimalDegreeMinuteSecond
                dirFormat = LocalizedRes.ParceMapCheck_Form_Settings_Format_DDMS '"DD.MMSSSS"
            Case AeccLandLib.AeccFormatType.aeccFormatDegreeMinuteSecond
                dirFormat = LocalizedRes.ParceMapCheck_Form_Settings_Format_DMS '"DD" & Chr(176) & "MM" & Chr(39) & "SS"
            Case AeccLandLib.AeccFormatType.aeccFormatDegreeMinuteSecondSpaced
                dirFormat = LocalizedRes.ParceMapCheck_Form_Settings_Format_DMSS '"DD" & Chr(176) & " MM" & Chr(39) & " SS"
            Case Else
                dirFormat = ""
        End Select

        Label_DirecFormat.Text = dirFormat
        Label_DistPrec.Text = ambSettings.DistanceSettings.Precision.Value
        Label_DirePrec.Text = ambSettings.DirectionSettings.Precision.Value
        Label_AreaPrec.Text = ambSettings.AreaSettings.Precision.Value
        Label_CloPrec.Text = ambSettings.CoordinateSettings.Precision.Value
        Label_PrecPrec.Text = ambSettings.DistanceSettings.Precision.Value
    End Sub

    Private Sub InitParcelListView()
        Radio_Parcels.Checked = True
        ' Get Captions
        'GetCaptions()

        ' list view grid setup
        ' setup columns
        ListView_Parcels.View = View.Details
        Call ListView_Parcels.Columns.Add(LocalizedRes.ParceMapCheck_Form_ColumnTitle_Include, _
            50, HorizontalAlignment.Left)
        Call ListView_Parcels.Columns.Add(LocalizedRes.ParceMapCheck_Form_ColumnTitle_Name, _
            100, HorizontalAlignment.Left)
        Call ListView_Parcels.Columns.Add(LocalizedRes.ParceMapCheck_Form_ColumnTitle_Number, _
            50, HorizontalAlignment.Left)
        Call ListView_Parcels.Columns.Add(LocalizedRes.ParceMapCheck_Form_ColumnTitle_Description, _
            80, HorizontalAlignment.Left)
        Call ListView_Parcels.Columns.Add(LocalizedRes.ParceMapCheck_Form_ColumnTitle_Area, _
            80, HorizontalAlignment.Left)
        Call ListView_Parcels.Columns.Add(LocalizedRes.ParceMapCheck_Form_ColumnTitle_Perimeter, _
            80, HorizontalAlignment.Left)

        'get alignments from Site
        Dim parcels As New SortedList(Of String, ParcelInfo)
        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            For Each oParcel As AeccLandLib.AeccParcel In oSite.Parcels
                Dim sParcelInfo As New ParcelInfo
                sParcelInfo.Obj = oParcel
                sParcelInfo.AcrossChord = False
                sParcelInfo.CounterWiseClock = False
                parcels.Add(oParcel.Parent.Name + "-" + oParcel.Name, sParcelInfo)
            Next
        Next

        'fill parcel grid
        For Each sParcelInfo As ParcelInfo In parcels.Values
            Dim oParcel As AeccLandLib.AeccParcel
            oParcel = sParcelInfo.Obj
            Dim li As New ListViewItem
            li.SubItems.Add(oParcel.Parent.Name + "-" + oParcel.Name)
            li.SubItems.Add(oParcel.Number)
            li.SubItems.Add(oParcel.Description)
            li.SubItems.Add(ParcelMapCheck_ExtractData.FormatAreaSettings_Parcel( _
                oParcel.Statistics.Area))
            li.SubItems.Add(ParcelMapCheck_ExtractData.FormatDistSettings_Parcel( _
                oParcel.Statistics.Perimeter))
            li.Checked = True
            li.Tag = sParcelInfo
            ListView_Parcels.Items.Add(li)
        Next

        If ListView_Parcels.Items.Count = 0 Then
            Dim sMsg As String = String.Format(LocalizedRes.Msg_No_Parcels_In_Drawing, ReportApplication.AeccXDocument.Name)
            ReportUtilities.ACADMsgBox(sMsg)
            Err.Number = 1
        End If
    End Sub


    Private Sub FillFigureGrid()
        Dim oFigure As AeccSurvey.AeccSurveyFigure
        Dim oProject As AeccSurvey.AeccSurveyProject

        Dim nCount As Long
        Dim nCurIndex As Long
        'get alignments from Site
        For Each oProject In ParcelMapCheck_ExtractData.SurveyDocument.Projects
            nCount = nCount + oProject.Figures.Count
            ReDim Preserve m_oFigureArr(nCount)
            For Each oFigure In oProject.Figures
                Dim li As New ListViewItem
                li.SubItems.Add(oFigure.Name)
                li.SubItems.Add(oFigure.ID)
                li.SubItems.Add("")
                li.SubItems.Add(ParcelMapCheck_ExtractData.FormatAreaSettings_Figure(0.0#))
                li.SubItems.Add(ParcelMapCheck_ExtractData.FormatDistSettings_Figure(0.0#))
                li.Checked = True
                m_oFigureArr(nCurIndex) = oFigure
                nCurIndex = nCurIndex + 1
                ListView_Figures.Items.Add(li)
            Next
        Next
    End Sub

    Private Sub Button_Deselect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Deselect.Click
        If m_bCheckParcle = True Then
            For Each lItem As ListViewItem In ListView_Parcels.Items
                lItem.Checked = False
            Next
        Else
            For Each lItem As ListViewItem In ListView_Figures.Items
                lItem.Checked = False
            Next
        End If
    End Sub

    Private Sub Button_Select_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Select.Click
        If m_bCheckParcle = True Then
            For Each lItem As ListViewItem In ListView_Parcels.Items
                lItem.Checked = True
            Next
        Else
            For Each lItem As ListViewItem In ListView_Figures.Items
                lItem.Checked = True
            Next
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

    Private Sub Check_AcrossChord_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Check_AcrossChord.CheckedChanged
        If ListView_Parcels.SelectedItems.Count = 0 Then
            Exit Sub
        End If

        Dim oParcelInfo As ParcelInfo
        oParcelInfo = ListView_Parcels.SelectedItems.Item(0).Tag
        oParcelInfo.AcrossChord = Check_AcrossChord.Checked
    End Sub

    Private Sub Check_CounterClock_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Check_CounterClock.CheckedChanged
        If ListView_Parcels.SelectedItems.Count = 0 Then
            Exit Sub
        End If

        Dim oParcelInfo As ParcelInfo
        oParcelInfo = ListView_Parcels.SelectedItems.Item(0).Tag
        oParcelInfo.CounterWiseClock = Check_CounterClock.Checked
    End Sub

    Private Sub Button_Pt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Pt.Click
        Me.Hide()

        Dim selectPnt As Object 'AeccLandLib.AeccPoint
        Dim prompt As String

        'Get AeccSettingCoordinate's decision
        Dim nDecision As Integer
        nDecision = ReportApplication.AeccXDatabase.Settings.DrawingSettings.AmbientSettings.CoordinateSettings.Precision.Value
        On Error Resume Next
        selectPnt = PickPoint()
        If Not selectPnt Is Nothing Then
            Dim dXCoord As Double
            Dim dYCoord As Double
            dXCoord = CDbl(selectPnt(0))
            dYCoord = CDbl(selectPnt(1))
            Text_PobXCoord.Text = ReportFormat.RoundVal(dXCoord, nDecision)
            Text_PobYCoord.Text = ReportFormat.RoundVal(dYCoord, nDecision)
        End If

        Me.Show()
    End Sub

    Private Function PickPoint() As AeccLandLib.AeccPoint
        'do not hide form here, AutoCAD will hide it.
        'move away report form
        Dim oldLocation As Drawing.Point
        oldLocation = Me.Location
        Me.Location = New Drawing.Point(-10000, -10000)

        PickPoint = Nothing
        On Error Resume Next

        'Prompt user to select a point or block reference:
        Dim prPointOpts As PromptEntityOptions
        prPointOpts = New PromptEntityOptions(LocalizedRes.AlignStakeout_Form_SelectPoint) '"Select a point")
        prPointOpts.SetRejectMessage(LocalizedRes.AlignStakeout_Form_SelectReject) '"Must Select a point")
        'should use civil .net entity type. but there's no Autodesk.Civil.XXX.Point yet.
        prPointOpts.AddAllowedClass(GetType(AeccLandLib.AeccPoint), False)

        Dim prPointRes As PromptEntityResult
        Dim docEditor As Autodesk.AutoCAD.EditorInput.Editor
        docEditor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
        prPointRes = docEditor.GetEntity(prPointOpts)
        Do While prPointRes.Status = PromptStatus.OK
            Dim currentPoint As AeccLandLib.AeccPoint
            For Each currentPoint In ReportApplication.AeccXDatabase.Points
                If ReportUtilities.CompareObjectID(currentPoint.ObjectID, prPointRes.ObjectId) Then
                    PickPoint = currentPoint
                    Exit For
                End If
            Next currentPoint
            If PickPoint Is Nothing Then
                docEditor.WriteMessage(vbNewLine + _
                    LocalizedRes.AlignStakeout_Form_SelectReject + vbNewLine)
                prPointRes = docEditor.GetEntity(prPointOpts)
            Else
                Exit Do
            End If
        Loop

        ' restore the dialog
        Me.Location = oldLocation
        Me.Focus()
    End Function

    Private Sub Button_CreateReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_CreateReport.Click

        If ListView_Parcels.Items.Count = 0 Then
            Exit Sub
        End If

        If ListView_Parcels.CheckedItems.Count = 0 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.ParcelMapCheck_Msg_SelectParcel)
            Exit Sub
        End If

        'reset progress bar
        ctlProgressBar.ProgressBegin(ListView_Parcels.CheckedItems.Count)

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

    Private Sub WriteHtmlReport(ByVal fileName As String)

        Dim oParcel As AeccLandLib.AeccParcel

        Dim oReportHtml As New ReportWriter(fileName)

        '<html>
        oReportHtml.HtmlBegin()
        oReportHtml.RenderHeader(LocalizedRes.ParcelMapCheck_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.ParcelMapCheck_Html_Title, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        For Each item As ListViewItem In ListView_Parcels.CheckedItems
            oParcel = item.Tag.Obj
            If Not oParcel Is Nothing Then
                AppendReport(oParcel, item.Tag.CounterWiseClock, _
                    item.Tag.AcrossChord, oReportHtml)
                ctlProgressBar.PerformStep()
            End If
        Next item

        '</body>
        oReportHtml.BodyEnd()

        '</html>
        oReportHtml.HtmlEnd()
    End Sub

    ' Append one item's report to html file
    Private Function AppendReport(ByVal oParcel As AeccLandLib.AeccParcel, _
                            ByVal bCounterClockWise As Boolean, _
                            ByVal bAcrossChord As Boolean, _
                            ByVal oHtmlWriter As ReportWriter) As Boolean
        Dim str As String

        If ParcelMapCheck_ExtractData.ExtractData(oParcel, bCounterClockWise, bAcrossChord) = False Then
            AppendReport = False
            Exit Function
        End If

        '<hr>
        oHtmlWriter.RenderHr()

        ' Parcel info
        str = LocalizedRes.ParcelMapCheck_Html_ParcelName
        str += " " + oParcel.Parent.Name + " - " + oParcel.Name
        oHtmlWriter.RenderLine(str)

        str = String.Format(LocalizedRes.Common_Html_Description_One_Param, oParcel.Description)
        oHtmlWriter.RenderLine(str)

        If bCounterClockWise = True Then
            str = LocalizedRes.ParcelMapCheck_Html_CounterClockwise_True
        Else
            str = LocalizedRes.ParcelMapCheck_Html_CounterClockwise_False
        End If
        oHtmlWriter.RenderLine(str)

        If bAcrossChord = True Then
            str = LocalizedRes.ParcelMapCheck_Html_MapCheck_True
        Else
            str = LocalizedRes.ParcelMapCheck_Html_MapCheck_False
        End If
        oHtmlWriter.RenderLine(str)

        'POB
        oHtmlWriter.TableBegin(, 550.ToString())

        AppendOneRecordInTable(oHtmlWriter, _
                    LocalizedRes.ParcelMapCheck_Html_North + _
                        ParcelMapCheck_ExtractData.FormateCoordSettings_Parcel( _
                            ParcelMapCheck_ExtractData.ParcelCheckData.POB_North), _
                    LocalizedRes.ParcelMapCheck_Html_East + _
                        ParcelMapCheck_ExtractData.FormateCoordSettings_Parcel( _
                            ParcelMapCheck_ExtractData.ParcelCheckData.POB_East))

        BreakLine(oHtmlWriter)

        'Segments
        AppendSegments(oHtmlWriter)

        'Check Error
        AppendCheckErrorInfo(oHtmlWriter)


        oHtmlWriter.TableEnd()
        oHtmlWriter.RenderBr()

        AppendReport = True
    End Function


    Private Sub AppendCheckErrorInfo(ByVal oHtmlWriter As ReportWriter)
        Dim ErrorClosure As Double
        Dim ErrorCourse As Double
        Dim ErrorNorth As Double, ErrorEast As Double
        Dim ErrorPrecision As Double

        ParcelMapCheck_ExtractData.ParcelCheckData.GetErrorInfo(ErrorClosure, _
                ErrorCourse, ErrorNorth, ErrorEast, ErrorPrecision)


        Call AppendOneRecordInTable(oHtmlWriter, _
                LocalizedRes.ParcelMapCheck_Html_Perimeter + " " + _
                    ParcelMapCheck_ExtractData.FormatDistSettings_Parcel(ParcelMapCheck_ExtractData.ParcelCheckData.Perimeter), _
                LocalizedRes.ParcelMapCheck_Html_Area + " " + _
                    ParcelMapCheck_ExtractData.FormatAreaSettings_Parcel(ParcelMapCheck_ExtractData.ParcelCheckData.Area))

        'get Closure
        Dim coordPrecision As Integer
        coordPrecision = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.CoordinateSettings.Precision.Value
        Dim formatString As String
        formatString = "N" + coordPrecision.ToString()
        Call AppendOneRecordInTable(oHtmlWriter, _
                                    LocalizedRes.ParcelMapCheck_Html_ErrClosure + " " + _
                                        ErrorClosure.ToString(formatString), _
                                    LocalizedRes.ParcelMapCheck_Html_Course + " " + _
                                        ParcelMapCheck_ExtractData.FormatDirSettings_Parcel(ErrorCourse))
        Call AppendOneRecordInTable(oHtmlWriter, _
                                    LocalizedRes.ParcelMapCheck_Html_ErrNorth + " " + _
                                        ReportFormat.RoundVal(ErrorNorth, coordPrecision + 1), _
                                    LocalizedRes.ParcelMapCheck_Html_East + " " + _
                                        ReportFormat.RoundVal(ErrorEast, coordPrecision + 1))

        BreakLine(oHtmlWriter)

        'Precision
        Dim distPrecision As Integer
        distPrecision = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.DistanceSettings.Precision.Value
        Call AppendOneRecordInTable(oHtmlWriter, LocalizedRes.ParcelMapCheck_Html_Precision1 + _
            " " + ReportFormat.RoundVal(ErrorPrecision, distPrecision), "&nbsp")
    End Sub

    Private Sub AppendSegments(ByVal oHtmlWriter As ReportWriter)
        Dim i As Long
        Dim segCount As Long
        Dim seg As Object
        Dim segLine As SegmentLine
        Dim segCurve As SegmentCurve
        segCount = ParcelMapCheck_ExtractData.ParcelCheckData.SegmentCount

        For i = 1 To segCount Step 1
            seg = ParcelMapCheck_ExtractData.ParcelCheckData.Item(i)
            If TypeOf seg Is SegmentLine Then
                segLine = seg
                AppendOneRecordInTable(oHtmlWriter, _
                            String.Format(LocalizedRes.ParcelMapCheck_Html_Line, i.ToString()), "&nbsp")
                AppendOneRecordInTable(oHtmlWriter, _
                            LocalizedRes.ParcelMapCheck_Html_Course + " " + _
                                ParcelMapCheck_ExtractData.FormatDirSettings_Parcel(segLine.Course), _
                            LocalizedRes.ParcelMapCheck_Html_Length + " " + _
                                ParcelMapCheck_ExtractData.FormatDistSettings_Parcel(segLine.Length))
                AppendOneRecordInTable(oHtmlWriter, _
                            LocalizedRes.ParcelMapCheck_Html_North + " " + _
                                ParcelMapCheck_ExtractData.FormateCoordSettings_Parcel(segLine.End_North), _
                            LocalizedRes.ParcelMapCheck_Html_East + " " + _
                                ParcelMapCheck_ExtractData.FormateCoordSettings_Parcel(segLine.End_East))
            Else
                segCurve = seg
                AppendOneRecordInTable(oHtmlWriter, _
                            String.Format(LocalizedRes.ParcelMapCheck_Html_Curve, i.ToString()), "&nbsp")
                AppendOneRecordInTable(oHtmlWriter, _
                            LocalizedRes.ParcelMapCheck_Html_Length + " " + _
                                ParcelMapCheck_ExtractData.FormatDistSettings_Parcel(segCurve.Length), _
                            LocalizedRes.ParcelMapCheck_Html_Radius + " " + _
                                ParcelMapCheck_ExtractData.FormatDistSettings_Parcel(segCurve.Radius))
                AppendOneRecordInTable(oHtmlWriter, _
                            LocalizedRes.ParcelMapCheck_Html_Delta + " " + _
                                ParcelMapCheck_ExtractData.FormatAngleSettings_Parcel(segCurve.Delta), _
                            LocalizedRes.ParcelMapCheck_Html_Tangent + " " + _
                                ParcelMapCheck_ExtractData.FormatDistSettings_Parcel(segCurve.Tangent))
                AppendOneRecordInTable(oHtmlWriter, _
                            LocalizedRes.ParcelMapCheck_Html_Chord + " " + _
                                ParcelMapCheck_ExtractData.FormatDistSettings_Parcel(segCurve.Chord), _
                            LocalizedRes.ParcelMapCheck_Html_Course + " " + _
                                ParcelMapCheck_ExtractData.FormatDirSettings_Parcel(segCurve.Course))
                AppendOneRecordInTable(oHtmlWriter, _
                            LocalizedRes.ParcelMapCheck_Html_CourseIn + " " + _
                                ParcelMapCheck_ExtractData.FormatDirSettings_Parcel(segCurve.CourseIn), _
                            LocalizedRes.ParcelMapCheck_Html_CourseOut + " " + _
                                ParcelMapCheck_ExtractData.FormatDirSettings_Parcel(segCurve.CourseOut))
                AppendOneRecordInTable(oHtmlWriter, _
                            LocalizedRes.ParcelMapCheck_Html_RPNorth + " " + _
                                ParcelMapCheck_ExtractData.FormateCoordSettings_Parcel(segCurve.RP_North), _
                            LocalizedRes.ParcelMapCheck_Html_East + " " + _
                                ParcelMapCheck_ExtractData.FormateCoordSettings_Parcel(segCurve.RP_East))
                AppendOneRecordInTable(oHtmlWriter, _
                            LocalizedRes.ParcelMapCheck_Html_EndNorth + " " + _
                                ParcelMapCheck_ExtractData.FormateCoordSettings_Parcel(segCurve.End_North), _
                            LocalizedRes.ParcelMapCheck_Html_East + " " + _
                                ParcelMapCheck_ExtractData.FormateCoordSettings_Parcel(segCurve.End_East))
            End If

            BreakLine(oHtmlWriter)
        Next
    End Sub

    Private Sub BreakLine(ByVal oHtmlWriter As ReportWriter)
        AppendOneRecordInTable(oHtmlWriter, "&nbsp", "&nbsp")
    End Sub


    Private Sub AppendOneRecordInTable(ByVal oHtmlWriter As ReportWriter, _
                                        ByVal strTd1 As String, ByVal strTd2 As String)
        oHtmlWriter.TrBegin()
        oHtmlWriter.RenderTd(strTd1)
        oHtmlWriter.RenderTd(strTd2)
        oHtmlWriter.TrEnd()
    End Sub

    Private Sub ListView_Parcels_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView_Parcels.SelectedIndexChanged
        If ListView_Parcels.SelectedItems.Count = 0 Then
            Check_CounterClock.Checked = False
            Check_CounterClock.Enabled = False
            Check_AcrossChord.Checked = False
            Check_AcrossChord.Enabled = False
            Exit Sub
        End If

        Dim selItem As ListViewItem
        Dim oParcelInfo As ParcelInfo
        selItem = ListView_Parcels.SelectedItems.Item(0)
        oParcelInfo = selItem.Tag
        Check_CounterClock.Checked = oParcelInfo.CounterWiseClock
        Check_CounterClock.Enabled = True
        Check_AcrossChord.Checked = oParcelInfo.AcrossChord
        Check_AcrossChord.Enabled = True
        Text_PobXCoord.Text = oParcelInfo.Obj.ParcelLoops(0).Item(0).StartX
        Text_PobYCoord.Text = oParcelInfo.Obj.ParcelLoops(0).Item(0).StartY
    End Sub

    Private Sub Radio_Figures_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Radio_Figures.CheckedChanged
        ListView_Figures.Top = ListView_Parcels.Top
        ListView_Figures.Height = ListView_Parcels.Height
        ListView_Figures.Left = ListView_Parcels.Left
        ListView_Figures.Width = ListView_Parcels.Width

        ListView_Parcels.Visible = False
        ListView_Figures.Visible = True
        m_bCheckParcle = False
    End Sub

    Private Sub Radio_Parcels_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Radio_Parcels.CheckedChanged
        ListView_Figures.Visible = False
        ListView_Parcels.Visible = True

        m_bCheckParcle = True
    End Sub

    Private Sub openHelp()
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_PARCEL_MAPCHECK_REPORT)
    End Sub
End Class
