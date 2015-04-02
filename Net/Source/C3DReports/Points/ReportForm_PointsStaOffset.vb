' -----------------------------------------------------------------------------
' <copyright file="ReportForm_PointsStaOffset.vb" company="Autodesk">
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
Imports Autodesk.AutoCAD.EditorInput
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Friend Class ReportForm_PointsStaOffset

    Private m_Alignments As Collections.Generic.SortedList(Of String, AeccLandLib.IAeccAlignment)
    Private m_Pointgroups As Collections.Generic.SortedList(Of String, AeccLandLib.IAeccPointGroup)

    Private m_oCurAlign As AeccLandLib.IAeccAlignment
    Private m_oCurPgroup As AeccLandLib.IAeccPointGroup

    Private ctlProgressBar As CtrlProgressBar
    Private ctlSavePath As CtrlSaveReportFile

    '
    'Main entry point for this reports application
    '
    Public Shared Sub Run()

        Dim rptDlg As New ReportForm_PointsStaOffset
        If rptDlg.ReadyToOpen() = True Then
            ReportUtilities.RunModalDialog(rptDlg)
        End If
    End Sub

    Private Function ReadyToOpen() As Boolean
        'Alignments
        Dim nAlignCount As Long
        nAlignCount = ReportApplication.AeccXDocument.AlignmentsSiteless.Count

        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            nAlignCount = oSite.Alignments.Count + nAlignCount
        Next
        If nAlignCount < 1 Then
            Dim sMsg As String = String.Format(LocalizedRes.Msg_No_Alignments_In_Drawing, ReportApplication.AeccXDocument.Name)
            ReportUtilities.ACADMsgBox(sMsg)

            Return False
        End If

        If ReportApplication.AeccXDatabase.Points.Count = 0 Then
            Dim sMsg As String = String.Format(LocalizedRes.Msg_No_Points_In_Drawing, ReportApplication.AeccXDocument.Name)
            ReportUtilities.ACADMsgBox(sMsg)

            Return False
        End If

        Return True
    End Function

    Private Sub ReportForm_PointsStaOffset_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)

        ctlSavePath = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)

        'button bmp
        Dim bp As Drawing.Bitmap = My.Resources.Resources.PickEntityBmp
        bp.MakeTransparent(Color.White)
        Button_SelectAlign.Image = bp
        Button_SelectAlign.ImageAlign = ContentAlignment.MiddleCenter

        If ReportApplication.AeccXDatabase.Points.Count = 0 Then
            Dim sMsg As String = String.Format(LocalizedRes.Msg_No_Points_In_Drawing, ReportApplication.AeccXDocument.Name)
            ReportUtilities.ACADMsgBox(sMsg)
            Err.Number = 1
            Exit Sub
        End If

        'get Point Groups
        m_Pointgroups = New Collections.Generic.SortedList(Of String, AeccLandLib.IAeccPointGroup)
        For Each oPgroup As AeccLandLib.AeccPointGroup In ReportApplication.AeccXDocument.PointGroups
            m_Pointgroups.Add(oPgroup.Name, oPgroup)
        Next

        'fill Point Group combo box
        For Each PgroupName As String In m_Pointgroups.Keys()
            Combo_Pgroup.Items.Add(PgroupName)
        Next

        Combo_Pgroup.SelectedIndex = 0

        'get Alignments
        m_Alignments = New Collections.Generic.SortedList(Of String, AeccLandLib.IAeccAlignment)
        For Each oAlign As AeccLandLib.AeccAlignment In ReportApplication.AeccXDocument.AlignmentsSiteless
            m_Alignments.Add(oAlign.Name, oAlign)
        Next

        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            For Each oAlign As AeccLandLib.AeccAlignment In oSite.Alignments
                m_Alignments.Add(oAlign.Name, oAlign)
            Next
        Next

        'fill alignment combo box
        For Each alignName As String In m_Alignments.Keys
            Combo_Align.Items.Add(alignName)
        Next

        Combo_Align.SelectedIndex = 0

        ' list view grid setup
        InitPointsView()
    End Sub

    Private Sub InitPointsView()
        'setup columns
        ListView_Aligns.View = View.Details
        ListView_Aligns.Columns.Add(LocalizedRes.PointsStaOffset_Form_ColumnTitle_Include, _
            50, HorizontalAlignment.Left)
        ListView_Aligns.Columns.Add(LocalizedRes.PointsStaOffset_Form_ColumnTitle_PointNum, _
            80, HorizontalAlignment.Left)
        ListView_Aligns.Columns.Add(LocalizedRes.PointsStaOffset_Form_ColumnTitle_Northing, _
            60, HorizontalAlignment.Left)
        ListView_Aligns.Columns.Add(LocalizedRes.PointsStaOffset_Form_ColumnTitle_Easting, _
            60, HorizontalAlignment.Left)
        ListView_Aligns.Columns.Add(LocalizedRes.PointsStaOffset_Form_ColumnTitle_Elevation, _
            80, HorizontalAlignment.Left)
        ListView_Aligns.Columns.Add(LocalizedRes.PointsStaOffset_Form_ColumnTitle_Name, _
            40, HorizontalAlignment.Left)
        'Don't need RawDesc with FullDesc 
        'ListView_Aligns.Columns.Add(LocalizedRes.PointsStaOffset_Form_ColumnTitle_RawDesc, _
        '    100, HorizontalAlignment.Left)
        ListView_Aligns.Columns.Add(LocalizedRes.PointsStaOffset_Form_ColumnTitle_FullDesc, _
            100, HorizontalAlignment.Left)

        'fill grid <Need to figure out how to list the points according to the selected point group>
        For Each oPoint As AeccLandLib.AeccPoint In ReportApplication.AeccXDatabase.Points
            Dim li As New ListViewItem
            li.SubItems.Add(oPoint.Number.ToString())
            li.SubItems.Add(PointsStaOffset_ExtractData.FormatPtCoordSettings(oPoint.Northing))
            li.SubItems.Add(PointsStaOffset_ExtractData.FormatPtCoordSettings(oPoint.Easting))
            li.SubItems.Add(PointsStaOffset_ExtractData.FormatPtElevationSettings(oPoint.Elevation))
            On Error Resume Next
            Dim name As String
            name = ""
            name = oPoint.Name
            On Error GoTo 0
            li.SubItems.Add(name)
            'Don't need RawDesc with FullDesc
            'li.SubItems.Add(oPoint.RawDescription)
            li.SubItems.Add(oPoint.FullDescription)
            li.Checked = True
            ListView_Aligns.Items.Add(li)
        Next
    End Sub

    Private Sub Combo_Align_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Align.SelectedIndexChanged
        If Combo_Align.SelectedIndex < 0 Then
            Exit Sub
        End If

        m_oCurAlign = m_Alignments(Combo_Align.SelectedItem.ToString())

        Label_AlignNameValue.Text = m_oCurAlign.Name

        Dim str As String
        str = String.Format(LocalizedRes.PointsStaOffset_Form_StaRange, _
                            ReportUtilities.GetStationStringWithDerived(m_oCurAlign, m_oCurAlign.StartingStation), _
                            ReportUtilities.GetStationStringWithDerived(m_oCurAlign, m_oCurAlign.EndingStation))
        Label_StaRangeValue.Text = str

        If m_oCurAlign.StationEquations.Count = 0 Then
            Label_StaEquationValue.Text = LocalizedRes.PointsStaOffset_Form_EquationNone '"None"
        Else
            Label_StaEquationValue.Text = LocalizedRes.PointsStaOffset_Form_EquationApp '"Applied"
        End If
    End Sub

    Private Sub Button_Deselect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Deselect.Click
        For Each item As ListViewItem In ListView_Aligns.Items
            item.Checked = False
        Next
    End Sub

    Private Sub Button_Select_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Select.Click
        For Each item As ListViewItem In ListView_Aligns.Items
            item.Checked = True
        Next
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
        If Not oAlignment Is Nothing Then
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
        If ListView_Aligns.Items.Count = 0 Then
            Exit Sub
        End If

        If ListView_Aligns.CheckedItems.Count = 0 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.PointsStaOffset_Msg_SelectPoint)
            Exit Sub
        End If

        'Init Progress bar
        ctlProgressBar.ProgressBegin(ListView_Aligns.CheckedItems.Count)

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

        ctlProgressBar.ProgressEnd()

        ReportUtilities.OpenFileByDefaultBrowser(ctlSavePath.SavePath)
    End Sub

    Private Sub Button_Done_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Done.Click
        Close()
    End Sub

    Private Sub WriteHtmlReport(ByVal fileName As String)

        Dim oReportHtml As New ReportWriter(fileName) 'm_sReportFileName)

        '<html>
        oReportHtml.HtmlBegin()
        oReportHtml.RenderHeader(LocalizedRes.PointsStaOffset_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.PointsStaOffset_Html_Title, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        '<hr>
        oReportHtml.RenderHr()

        ' alignment info
        Dim str As String
        str = LocalizedRes.PointsStaOffset_Html_AlignName
        str += " " + m_oCurAlign.Name
        oReportHtml.RenderLine(str)

        str = String.Format(LocalizedRes.Common_Html_Description_One_Param, m_oCurAlign.Description)
        oReportHtml.RenderLine(str)

        str = String.Format(LocalizedRes.Alignment_Html_Station_Range, _
                    ReportUtilities.GetStationStringWithDerived(m_oCurAlign, m_oCurAlign.StartingStation), _
                    ReportUtilities.GetStationStringWithDerived(m_oCurAlign, m_oCurAlign.EndingStation))
        oReportHtml.RenderLine(str)

        'br
        oReportHtml.RenderBr()

        'render data table
        '<table>
        oReportHtml.TableBegin("1")

        'table header
        oReportHtml.TrBegin()
        oReportHtml.RenderTd(LocalizedRes.PointsStaOffset_Html_TblTitle_Point, True)
        oReportHtml.RenderTd(LocalizedRes.PointsStaOffset_Html_TblTitle_Station, True)
        oReportHtml.RenderTd(LocalizedRes.PointsStaOffset_Html_TblTitle_Offset, True)
        oReportHtml.RenderTd(LocalizedRes.PointsStaOffset_Html_TblTitle_Ele, True)
        oReportHtml.RenderTd(LocalizedRes.PointsStaOffset_Html_TblTitle_Desc, True)
        oReportHtml.TrEnd()

        For Each item As ListViewItem In ListView_Aligns.CheckedItems
            Dim oPoint As AeccLandLib.AeccPoint
            oPoint = ReportApplication.AeccXDatabase.Points.Find(CInt(item.SubItems.Item(1).Text))
            If Not oPoint Is Nothing Then
                AppendReport(m_oCurAlign, oPoint, oReportHtml)
                ctlProgressBar.PerformStep()
            End If
        Next item

        '</Table>
        oReportHtml.TableEnd()

        oReportHtml.RenderBr()
        oReportHtml.RenderLine(LocalizedRes.PointsStaOffset_Html_TblTitle_OutRange)

        '</body>
        oReportHtml.BodyEnd()

        '</html>
        oReportHtml.HtmlEnd()
    End Sub

    Private Sub AppendReport(ByVal oAlignment As AeccLandLib.IAeccAlignment, _
                            ByVal oPoint As AeccLandLib.AeccPoint, _
                            ByVal oHtmlWriter As ReportWriter)

        ' extract data
        PointsStaOffset_ExtractData.ExtractData(oAlignment, oPoint)

        Dim description As String
        description = ""
        If Not PointsStaOffset_ExtractData.PointDataArr(PointsStaOffset_ExtractData.nDescriptionRaw) = "" Then
            description = "(" + _
                PointsStaOffset_ExtractData.PointDataArr(PointsStaOffset_ExtractData.nDescriptionRaw) + ")"
        End If
        description = description + _
            PointsStaOffset_ExtractData.PointDataArr(PointsStaOffset_ExtractData.nDescriptionFull)
        If description = "" Then
            description = "&nbsp"
        End If

        oHtmlWriter.TrBegin()
        oHtmlWriter.RenderTd(PointsStaOffset_ExtractData.PointDataArr(PointsStaOffset_ExtractData.nPointIndex))
        oHtmlWriter.RenderTd(PointsStaOffset_ExtractData.PointDataArr(PointsStaOffset_ExtractData.nStationIndex))
        oHtmlWriter.RenderTd(PointsStaOffset_ExtractData.PointDataArr(PointsStaOffset_ExtractData.nOffsetIndex))
        oHtmlWriter.RenderTd(PointsStaOffset_ExtractData.PointDataArr(PointsStaOffset_ExtractData.nElevation))
        oHtmlWriter.RenderTd(description)
        oHtmlWriter.TrEnd()
    End Sub

    Private Sub openHelp()
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_STATION_OFFSET_TO_POINTS_REPORT)
    End Sub

    Private Sub ListView_Aligns_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView_Aligns.SelectedIndexChanged

    End Sub

    Private Sub Combo_Pgroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_Pgroup.SelectedIndexChanged
        If Combo_Pgroup.SelectedIndex < 0 Then
            Exit Sub
        End If

        m_oCurPgroup = m_Pointgroups(Combo_Pgroup.SelectedItem.ToString())

        Label_PgroupNameValue.Text = m_oCurPgroup.Name

    End Sub
End Class
