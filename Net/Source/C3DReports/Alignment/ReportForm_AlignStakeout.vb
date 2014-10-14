' -----------------------------------------------------------------------------
' <copyright file="ReportForm_AlignStakeout.vb" company="Autodesk">
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
Option Strict On

Imports System
Imports AeccUiLandLib = Autodesk.AECC.Interop.UiLand
Imports AeccLandLib = Autodesk.AECC.Interop.Land

'Imports Autodesk.Civil
Imports Autodesk.AutoCAD.EditorInput
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Friend Class ReportForm_AlignStakeout

    Enum StakeAngleType
        TURNED_PLUS
        TURNED_MINUS
        DEFLECT_PLUS
        DEFLECT_MINUS
        DIRECTION
    End Enum

    Private m_PointDict As Collections.Generic.Dictionary(Of Integer, AeccLandLib.AeccPoint)
    Private m_eAngleType As StakeAngleType
    Private m_OccupiedPt As AeccLandLib.AeccPoint
    Private m_BacksightPt As AeccLandLib.AeccPoint

    Public Shared Sub Run()
        If CtrlAlignmentListView.GetAlignmentCount() > 0 Then
            ' Create report dialog
            Dim rptDlg As New ReportForm_AlignStakeout

            If rptDlg.CanOpen() = True Then
                ReportUtilities.RunModalDialog(rptDlg)
            End If
        End If
    End Sub

    Private Function CanOpen() As Boolean
        CanOpen = False
        ' get the point dictionary from current drawing database
        Try
            If ReportApplication.AeccXDatabase.Points.Count >= 2 Then
                CanOpen = True
            End If
        Catch ex As Exception

        End Try

        If CanOpen = False Then
            ReportUtilities.ACADMsgBox(LocalizedRes.AlignStakeout_Msg_2Points)
        End If
    End Function

    Private Sub ReportForm_AlignStakeout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'add points to point dictionary
        Dim currentPoint As AeccLandLib.AeccPoint
        m_PointDict = New Collections.Generic.Dictionary(Of Integer, AeccLandLib.AeccPoint)

        For Each currentPoint In ReportApplication.AeccXDatabase.Points
            If Not currentPoint Is Nothing And Not m_PointDict.ContainsKey(currentPoint.Number) Then
                m_PointDict.Add(currentPoint.Number, currentPoint)
            End If
        Next currentPoint

        If m_PointDict.Count < 2 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.AlignStakeout_Msg_2Points)
            Err.Number = 1
            Exit Sub
        End If

        ctlStartStation = New CtrlStationTextBox
        ctlStartStation.Initialize(TextBox_StartStation)
        ctlEndStation = New CtrlStationTextBox
        ctlEndStation.Initialize(TextBox_EndStation)
        ctlAlignmentListView = New CtrlAlignmentListView
        ctlAlignmentListView.Initialize(ListView_Alignments, ctlStartStation, ctlEndStation)
        ctlSavePath = New CtrlSaveReportFile(TextBox_SaveReport, Button_Save)
        ctlProgressBar = New CtrlProgressBar
        ctlProgressBar.Initialize(ProgressBar_Creating, GroupBox_Progress)

        'Init Stakeout options
        ' set angle type option here
        Radio_TurnedPlus.Checked = True
        Radio_TurnedMinus.Checked = False
        Radio_DeflectPlus.Checked = False
        Radio_DeflectMinus.Checked = False
        Radio_Direction.Checked = False
        m_eAngleType = StakeAngleType.TURNED_PLUS

        'button bmp
        Dim bp As Drawing.Bitmap = My.Resources.Resources.PickEntityBmp
        bp.MakeTransparent(Color.White)
        Button_Occupied.Image = bp
        Button_Occupied.ImageAlign = ContentAlignment.MiddleCenter
        Button_Backsight.Image = bp
        Button_Backsight.ImageAlign = ContentAlignment.MiddleCenter

        'up&down station increment
        NumericUpDown_StationInc.Value = GetDefaultIncrement()
        NumericUpDown_StationInc.Maximum = Decimal.MaxValue
        NumericUpDown_StationInc.Minimum = 0.01D

        'up&down offset
        NumericUpDown_Offset.Value = 0
        NumericUpDown_Offset.Maximum = Decimal.MaxValue
        NumericUpDown_Offset.Minimum = Decimal.MinValue
    End Sub

    Private Sub Button_CreateReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_CreateReport.Click
        If ctlAlignmentListView.WinListView.CheckedItems.Count = 0 Then
            ReportUtilities.ACADMsgBox(LocalizedRes.Alignment_Msg_SelectOneFirst)
            Exit Sub
        End If

        If VerifyBeforeExe() = False Then
            Exit Sub
        End If

        'init progress bar
        ctlProgressBar.ProgressBegin(ctlAlignmentListView.WinListView.CheckedItems.Count * 3)

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

        'Dim oReportHtml As New ReportWriter(ctlSavePath.SavePath)
        Dim oReportHtml As New ReportWriter(fileName)
        '<html>
        oReportHtml.HtmlBegin()
        oReportHtml.RenderHeader(LocalizedRes.AlignStakeout_Html_Title)

        '<body>
        oReportHtml.BodyBegin()
        '<h1>
        oReportHtml.RenderH(LocalizedRes.AlignStakeout_Html_Title, 1)
        oReportHtml.RenderUserSetting()

        'data info
        oReportHtml.RenderDateInfo()

        'point format
        Dim ptPrecision As Integer
        Dim ptRoundType As AeccLandLib.AeccRoundingType
        Dim ptCoordSettings As AeccLandLib.IAeccSettingsCoordinate
        ptCoordSettings = ReportApplication.AeccXDatabase.Settings.PointSettings.AmbientSettings.CoordinateSettings
        ptPrecision = ptCoordSettings.Precision.Value
        ptRoundType = ptCoordSettings.Rounding.Value

        For Each item As ListViewItem In ctlAlignmentListView.WinListView.CheckedItems
            Dim startRawStation, endRawStation As Double
            startRawStation = ctlAlignmentListView.StartRawStation(item)
            endRawStation = ctlAlignmentListView.EndRawStation(item)
            Try
                AppendReport(CType(item.Tag, AeccLandLib.AeccAlignment), startRawStation, endRawStation, ptPrecision, ptRoundType, oReportHtml)
                ctlProgressBar.PerformStep()
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
            End Try
        Next item

        '</body>
        oReportHtml.BodyEnd()

        '</html>
        oReportHtml.HtmlEnd()
    End Sub

    ' Append one item's report to html file
    Private Sub AppendReport(ByVal oAlignment As AeccLandLib.AeccAlignment, _
        ByVal stationStart As Double, ByVal stationEnd As Double, _
        ByVal ptPrecision As Integer, _
        ByVal ptRoundType As AeccLandLib.AeccRoundingType, _
        ByVal oHtmlWriter As ReportWriter)

        If oAlignment Is Nothing Or oHtmlWriter Is Nothing Then
            Exit Sub
        End If

        '<hr>
        oHtmlWriter.RenderHr()

        ' alignment info
        Dim str As String

        'alignment name
        str = LocalizedRes.AlignSta_Html_Align_Name
        str += " " + oAlignment.Name
        oHtmlWriter.RenderLine(str)

        'alignment description
        str = String.Format(LocalizedRes.Common_Html_Description_One_Param, oAlignment.Description)
        oHtmlWriter.RenderLine(str)

        'start station, end station
        str = String.Format(LocalizedRes.Alignment_Html_Station_Range, ReportUtilities.GetStationStringWithDerived(oAlignment, stationStart), _
              ReportUtilities.GetStationStringWithDerived(oAlignment, stationEnd))
        oHtmlWriter.RenderLine(str)

        'angle type
        str = LocalizedRes.AlignStakeout_Html_StakeoutAngleType
        str += " " + GetStakeAngleTypeStr(m_eAngleType)
        oHtmlWriter.RenderLine(str)

        'Get occupied and backsight pt info
        str = String.Format(LocalizedRes.AlignStakeout_Html_OccupiedPt, _
                            ReportFormat.RoundVal(m_OccupiedPt.Northing, ptPrecision, ptRoundType), _
                            ReportFormat.RoundVal(m_OccupiedPt.Easting, ptPrecision, ptRoundType))
        oHtmlWriter.RenderLine(str)
        'str = str & "Occupied Pt: " & "<b>Northing </b>" & g_oOccupiedPt.Northing & "," & "<b>Easting </b>" & g_oOccupiedPt.Easting & "<br/>"
        If Not m_eAngleType = StakeAngleType.DIRECTION Then
            str = String.Format(LocalizedRes.AlignStakeout_Html_BacksightPt, _
                    ReportFormat.RoundVal(m_BacksightPt.Northing, ptPrecision, ptRoundType), _
                    ReportFormat.RoundVal(m_BacksightPt.Easting, ptPrecision, ptRoundType))
            oHtmlWriter.RenderLine(str)
        End If

        'station increment
        'NOTICE: here, use "Station Increment" instead of "Station interval: "
        str = LocalizedRes.AlignStaInc_Html_Increment
        str += " " + NumericUpDown_StationInc.Value.ToString("N2")
        oHtmlWriter.RenderLine(str)

        ' offset
        str = LocalizedRes.AlignStakeout_Html_Offset
        str += " " + NumericUpDown_Offset.Value.ToString("N2")
        oHtmlWriter.RenderLine(str)

        'begin to render alignment data table
        oHtmlWriter.RenderBr()
        oHtmlWriter.TableBegin("1")

        'title
        oHtmlWriter.TrBegin()
        oHtmlWriter.RenderTd(LocalizedRes.AlignStakeout_Html_TblTitle_Station, True)
        Select Case m_eAngleType
            Case StakeAngleType.TURNED_PLUS '"Turned.Right"
                oHtmlWriter.RenderTd(LocalizedRes.AlignStakeout_Html_TblTitle_TurnedR, True)
            Case StakeAngleType.TURNED_MINUS '"Turned.Left"
                oHtmlWriter.RenderTd(LocalizedRes.AlignStakeout_Html_TblTitle_TurnedL, True)
            Case StakeAngleType.DEFLECT_PLUS '"Defl.Right"
                oHtmlWriter.RenderTd(LocalizedRes.AlignStakeout_Html_TblTitle_DeflR, True)
            Case StakeAngleType.DEFLECT_MINUS '"Defl.Left"
                oHtmlWriter.RenderTd(LocalizedRes.AlignStakeout_Html_TblTitle_DeflL, True)
            Case Else 'direction
                oHtmlWriter.RenderTd(LocalizedRes.AlignStakeout_Html_TblTitle_Direction, True)
        End Select
        oHtmlWriter.RenderTd(LocalizedRes.AlignStakeout_Html_TblTitle_Distance, True)
        oHtmlWriter.RenderTd(LocalizedRes.AlignStakeout_Html_TblTitle_CoordinateN, True)
        oHtmlWriter.RenderTd(LocalizedRes.AlignStakeout_Html_TblTitle_CoordinateE, True)
        oHtmlWriter.TrEnd()

        ctlProgressBar.PerformStep()

        ' extract data
        AlignStakeout_ExtractData.ExtractData(oAlignment, stationStart, stationEnd, _
            NumericUpDown_StationInc.Value, NumericUpDown_Offset.Value, _
            m_eAngleType, m_OccupiedPt, m_BacksightPt)

        ctlProgressBar.PerformStep()

        For Each data As AlignStakeout_ExtractData.AlignStakeoutData In AlignStakeout_ExtractData.DataDictionary.Values
            Try
                'format string
                oHtmlWriter.TrBegin()
                oHtmlWriter.RenderTd(data.Station)
                oHtmlWriter.RenderTd(data.Direction)
                oHtmlWriter.RenderTd(data.Distance)
                oHtmlWriter.RenderTd(data.Northing)
                oHtmlWriter.RenderTd(data.Easting)
                oHtmlWriter.TrEnd()
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
                Exit For
            End Try
        Next

        oHtmlWriter.TableEnd()
        oHtmlWriter.RenderBr()
    End Sub

    Private Sub Button_Occupied_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Occupied.Click
        Dim pt As AeccLandLib.AeccPoint
        pt = PickPoint()

        If Not pt Is Nothing Then
            m_OccupiedPt = m_PointDict.Item(pt.Number)
            Edit_PointOccupied.Text = CStr(pt.Number)
        End If
    End Sub

    Private Sub Button_Backsight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Backsight.Click
        Dim pt As AeccLandLib.AeccPoint
        pt = PickPoint()

        If Not pt Is Nothing Then
            m_BacksightPt = m_PointDict.Item(pt.Number)
            Edit_BacksightPoint.Text = CStr(pt.Number)
        End If
    End Sub

    Private Function PickPoint() As AeccLandLib.AeccPoint
        'move away report form to 
        Dim oldLocation As Drawing.Point
        oldLocation = Me.Location
        Me.Location = New Drawing.Point(-10000, -10000)

        PickPoint = Nothing
        On Error Resume Next

        'Prompt user to select a point or block reference:
        Dim prPointOpts As PromptEntityOptions
        prPointOpts = New PromptEntityOptions(LocalizedRes.AlignStakeout_Form_SelectPoint)
        prPointOpts.SetRejectMessage(LocalizedRes.AlignStakeout_Form_SelectReject)
        'should use civil .net entity type. but there's no Autodesk.Civil.XXX.Point yet.
        prPointOpts.AddAllowedClass(GetType(AeccLandLib.AeccPoint), False)

        Dim prPointRes As PromptEntityResult
        Dim docEditor As Autodesk.AutoCAD.EditorInput.Editor
        docEditor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
        prPointRes = docEditor.GetEntity(prPointOpts)
        Do While prPointRes.Status = PromptStatus.OK
            Dim currentPoint As AeccLandLib.AeccPoint
            For Each currentPoint In ReportApplication.AeccXDatabase.Points
                If currentPoint.ObjectID.ToString().CompareTo(prPointRes.ObjectId.ToString().Replace("(", "").Replace(")", "")) = 0 Then
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

    Private Sub Edit_PointOccupied_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Edit_PointOccupied.KeyPress
        If IsInputValidKey(e.KeyChar) = False Then
            Beep()
            e.Handled = True
        End If
    End Sub

    Private Sub Edit_PointOccupied_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Edit_PointOccupied.Leave
        If Edit_PointOccupied.Text = "" Then
            m_OccupiedPt = Nothing
            Exit Sub
        End If

        Try
            Dim pointNumber As Integer = CInt(Edit_PointOccupied.Text)
            If m_PointDict.ContainsKey(pointNumber) Then
                m_OccupiedPt = m_PointDict.Item(pointNumber)
            Else
                Throw New ArgumentOutOfRangeException("Invalid input")
            End If
        Catch ex As Exception
            ReportUtilities.ACADMsgBox(LocalizedRes.AlignStakeout_Msg_InputPoint)
            If m_OccupiedPt Is Nothing Then
                Edit_PointOccupied.Text = ""
            Else
                Edit_PointOccupied.Text = CStr(m_OccupiedPt.Number)
            End If
        End Try
    End Sub

    Private Sub Edit_BacksightPoint_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Edit_BacksightPoint.KeyPress
        If IsInputValidKey(e.KeyChar) = False Then
            Beep()
            e.Handled = True
        End If
    End Sub

    Private Sub Edit_BacksightPoint_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Edit_BacksightPoint.Leave
        If Edit_BacksightPoint.Text = "" Then
            m_BacksightPt = Nothing
            Exit Sub
        End If

        Try
            Dim pointNumber As Integer = CInt(Edit_BacksightPoint.Text)
            If m_PointDict.ContainsKey(pointNumber) Then
                m_BacksightPt = m_PointDict.Item(pointNumber)
            Else
                Throw New ArgumentOutOfRangeException("Invalid input")
            End If
        Catch ex As Exception
            ReportUtilities.ACADMsgBox(LocalizedRes.AlignStakeout_Msg_InputPoint)
            If m_BacksightPt Is Nothing Then
                Edit_BacksightPoint.Text = ""
            Else
                Edit_BacksightPoint.Text = CStr(m_BacksightPt.Number)
            End If
        End Try
    End Sub

    Private Function IsInputValidKey(ByVal KeyChar As Char) As Boolean
        If KeyChar <> ChrW(Keys.Back) And _
                (KeyChar < ChrW(Keys.D0) Or KeyChar > ChrW(Keys.D9)) Then
            IsInputValidKey = False
        Else
            IsInputValidKey = True
        End If
    End Function

    Private Sub Radio_TurnedPlus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Radio_TurnedPlus.Click
        m_eAngleType = StakeAngleType.TURNED_PLUS
    End Sub

    Private Sub Radio_TurnedMinus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Radio_TurnedMinus.Click
        m_eAngleType = StakeAngleType.TURNED_MINUS
    End Sub

    Private Sub Radio_DeflectPlus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Radio_DeflectPlus.Click
        m_eAngleType = StakeAngleType.DEFLECT_PLUS
    End Sub

    Private Sub Radio_DeflectMinus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Radio_DeflectMinus.Click
        m_eAngleType = StakeAngleType.DEFLECT_MINUS
    End Sub

    Private Sub Radio_Direction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Radio_Direction.Click
        m_eAngleType = StakeAngleType.DIRECTION
    End Sub

    Private Function VerifyBeforeExe() As Boolean
        VerifyBeforeExe = False
        If m_OccupiedPt Is Nothing Then
            ReportUtilities.ACADMsgBox(LocalizedRes.AlignStakeout_Msg_SelectOccupiedPoint)
        ElseIf Not m_eAngleType = StakeAngleType.DIRECTION Then
            If m_BacksightPt Is Nothing Then
                ReportUtilities.ACADMsgBox(LocalizedRes.AlignStakeout_Msg_SelectBacksightPoint)
            ElseIf m_OccupiedPt.Northing = m_BacksightPt.Northing And m_OccupiedPt.Easting = m_BacksightPt.Easting Then
                ReportUtilities.ACADMsgBox(LocalizedRes.AlignStakeout_Msg_SamePoint)
            Else
                VerifyBeforeExe = True
            End If
        Else
            VerifyBeforeExe = True
        End If
    End Function

    Private Function GetStakeAngleTypeStr(ByVal aType As StakeAngleType) As String
        If aType = StakeAngleType.DEFLECT_MINUS Then
            GetStakeAngleTypeStr = LocalizedRes.AlignStakeout_Html_AngleType_DeflectMinus '"DeflectMinus"
        ElseIf aType = StakeAngleType.DEFLECT_PLUS Then
            GetStakeAngleTypeStr = LocalizedRes.AlignStakeout_Html_AngleType_DeflectPlus '"DeflectPlus"
        ElseIf aType = StakeAngleType.DIRECTION Then
            GetStakeAngleTypeStr = LocalizedRes.AlignStakeout_Html_AngleType_Direction '"Direction"
        ElseIf aType = StakeAngleType.TURNED_MINUS Then
            GetStakeAngleTypeStr = LocalizedRes.AlignStakeout_Html_AngleType_TurnedMinus '"TurnedMinus"
        Else
            GetStakeAngleTypeStr = LocalizedRes.AlignStakeout_Html_AngleType_TurnedPlus '"TurnedPlus "
        End If
    End Function

    Private Function GetDefaultIncrement() As Decimal
        Dim unitType As AeccLandLib.AeccCoordinateUnitType = AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitMeter
        Try
            unitType = ReportApplication.AeccXDatabase.Settings.DrawingSettings.AmbientSettings.DistanceSettings.Unit.Value
        Catch ex As Exception
            Diagnostics.Debug.Assert(False, ex.Message)
        End Try

        Dim dIncrement As Decimal
        If unitType = AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitFoot Then
            dIncrement = 50D
        Else
            dIncrement = 20D
        End If
        Return dIncrement
    End Function

    Private Sub openHelp()
        ReportApplication.InvokeHelp(ReportHelpID.IDH_AECC_REPORTS_CREATE_STAKEOUT_ALIGNMENT_REPORT)
    End Sub

    '-----------------------------------------------------------------------
    '   calc the 3 point angle, return a string formated based on the current drawing setup angular settings
    '-----------------------------------------------------------------------
    ' returned radAngle is the North-Based radian format
    ' Parameter: oNorth the Northing value of occupied point
    ' Parameter: oEast the Easting value of occupied point
    ' Parameter: sNorth the Northing value of Station point
    ' Parameter: sNorth the Easting value of Station point
    ' Parameter: bNorth the Northing value of backsight point
    ' Parameter: bNorth the Easting value of backsight point
    ' Return value: in [0,2PI)
    Public Shared Function Calc3PointAngle(ByVal oNorth As Double, ByVal oEast As Double, _
    ByVal sNorth As Double, ByVal sEast As Double, _
    ByVal bNorth As Double, ByVal bEast As Double, _
    ByVal anType As StakeAngleType) As Double
        Dim dirOtoS As Double
        Dim dirOtoB As Double

        If oNorth = bNorth And oEast = bEast Then
            Calc3PointAngle = 0.0#
            Exit Function
        End If

        dirOtoS = ReportUtilities.CalcDirRad(oNorth, oEast, sNorth, sEast)
        dirOtoB = ReportUtilities.CalcDirRad(oNorth, oEast, bNorth, bEast)

        Calc3PointAngle = dirOtoS - dirOtoB ' for TurnedPlus
        If anType = StakeAngleType.TURNED_PLUS Then
            Calc3PointAngle = dirOtoS - dirOtoB
        ElseIf anType = StakeAngleType.TURNED_MINUS Then
            Calc3PointAngle = dirOtoB - dirOtoS
        ElseIf anType = StakeAngleType.DEFLECT_PLUS Then
            Calc3PointAngle = dirOtoS + dirOtoB
        ElseIf anType = StakeAngleType.DEFLECT_MINUS Then
            Calc3PointAngle = -dirOtoS - dirOtoB
        End If
        Calc3PointAngle = ReportFormat.AdjustAngle2PI(Calc3PointAngle)
    End Function


    Private ctlStartStation As CtrlStationTextBox
    Private ctlEndStation As CtrlStationTextBox
    Private ctlAlignmentListView As CtrlAlignmentListView
    Private ctlProgressBar As CtrlProgressBar
    Private ctlSavePath As CtrlSaveReportFile
End Class
