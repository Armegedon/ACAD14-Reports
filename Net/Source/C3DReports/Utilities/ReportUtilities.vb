' -----------------------------------------------------------------------------
' <copyright file="ReportUtilities.vb" company="Autodesk">
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
Imports System.Globalization
Imports System.Threading.Thread
Imports AutoCAD = Autodesk.AutoCAD.Interop
Imports AeccUiLandLib = Autodesk.AECC.Interop.UiLand
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports AeccRoadLib = Autodesk.AECC.Interop.Roadway
Imports AeccUiRoadLib = Autodesk.AECC.Interop.UiRoadway
Imports LocalizedRes = Report.My.Resources.LocalizedStrings
Imports Autodesk.AECC.Interop
Imports System.Runtime.InteropServices

Public Class ReportUtilities
    <DllImport("kernel32.dll", ExactSpelling:=True)> _
    Public Shared Function GetUserDefaultLCID() As Integer
    End Function

    Public Shared oldThreadLocal As Integer

    Public Shared Sub UpdateUserLocalSetting()
        Dim localeCur As Integer = GetUserDefaultLCID()
        Dim thread As System.Threading.Thread = System.Threading.Thread.CurrentThread
        oldThreadLocal = thread.CurrentCulture.LCID
        If (oldThreadLocal <> localeCur) Then
            thread.CurrentCulture = New System.Globalization.CultureInfo(localeCur)
        End If
    End Sub
    Public Shared Sub RestoreThreadLocalSetting()
        Dim localeCur As Integer = GetUserDefaultLCID()
        Dim thread As System.Threading.Thread = System.Threading.Thread.CurrentThread
        If (oldThreadLocal <> thread.CurrentCulture.LCID) Then
            thread.CurrentCulture = New System.Globalization.CultureInfo(oldThreadLocal)
        End If
    End Sub



    'NOTICE: In C3D, RAW means original value of the station, it's relative to Equation.
    '        But here, RAW means the double value of the station, it's relative to the station formated string - equaSationValue.

    '        So to avoid confusion, we call the Raw Value in C3D as Original Value, and the Equation value in C3D as Derived Value.
    '        and we call the double value as RAW Value, and the formated station string as Equation value.

    'IMPORTANT: Generally, we can not a original station from a derived station because derived stations are not unique in a drawing.
    '           So, we have to show the original value to user on the dialog until a better solution was found - eg. let user select a station from the drawing.


    Public Shared Function GetStationStringWithDerived(ByVal oAlignment As AeccLandLib.IAeccAlignment, ByVal orginialStation As Double) As String
        Dim staString As String = oAlignment.GetStationStringWithEquations(orginialStation)

        'staString is always formated like 1+263.9558 or 1263.9558, the dot seperator should be changed here according local setting.
        Dim decSeparaor As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator
        GetStationStringWithDerived = staString.Replace(".", decSeparaor)
    End Function

    'Raw station to string
    'station could be the station with or without Derived
    Public Shared Function GetStationString(ByVal oAlignment As AeccLandLib.IAeccAlignment, ByVal station As Double, _
                                        Optional ByVal settingsStation As AeccLandLib.IAeccSettingsStation = Nothing) As String
        If settingsStation Is Nothing Then
            settingsStation = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.StationSettings
        End If

        Dim delimiter As String
        delimiter = "+"
        Select Case settingsStation.StationDelimiterCharacter.Value
            Case AeccLandLib.AeccStationDelimiterCharacterType.aeccStationDelimiterPlusSign
                delimiter = "+"
            Case AeccLandLib.AeccStationDelimiterCharacterType.aeccStationDelimiterMinusSign
                delimiter = "-"
            Case AeccLandLib.AeccStationDelimiterCharacterType.aeccStationDelimiterAutomatic
                delimiter = "+"
            Case AeccLandLib.AeccStationDelimiterCharacterType.aeccStationDelimiterUnderscore
                delimiter = "_"
            Case AeccLandLib.AeccStationDelimiterCharacterType.aeccStationDelimiterNone
                delimiter = " "
            Case Else
                Diagnostics.Debug.Assert(False)
                delimiter = "+"
        End Select

        Dim bStationIndexFormat As Boolean = settingsStation.Format().Value = AeccLandLib.AeccStationFormatType.aeccStationIndexFormat
        Dim stationIndexInc As Double = oAlignment.StationIndexIncrement

        Dim precision As Integer = settingsStation.Precision.Value
        Dim positionType As AeccLandLib.AeccStationDelimiterPositionType = settingsStation.StationDelimiterPosition.Value
        Dim dropDecimal As Boolean = settingsStation.DropDecimalsForWholeNumbers.Value

        Dim position As Integer = positionType

        Dim strFormat As String
        If dropDecimal Then
            strFormat = "{0:F0}"
        Else
            strFormat = String.Format("{{0:F{0}}}", precision)
        End If

        Dim staString As String = String.Format(strFormat, station)

        Dim dotSeparator As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
        Dim dotIndex As Long = staString.IndexOf(dotSeparator)
        If (dotIndex < 0) Then
            dotIndex = staString.Length
        End If

        If (bStationIndexFormat = True) Then
            Dim strBuilder As System.Text.StringBuilder = New System.Text.StringBuilder()
            Dim intValue As Integer
            If (Math.Sign(station) >= 0) Then
                intValue = Math.Floor(station)
            Else
                intValue = Math.Ceiling(station)
            End If

            strBuilder.AppendFormat("{0}{1}{2}", intValue \ stationIndexInc, delimiter, intValue Mod stationIndexInc)
            If (dotIndex > 0) Then
                staString = staString.Substring(dotIndex)
                strBuilder.Append(staString)
            End If
            GetStationString = strBuilder.ToString()
        Else
            Dim insertPos = dotIndex - position
            If (insertPos <= 0) Then
                Dim strBuilder As System.Text.StringBuilder = New System.Text.StringBuilder()
                strBuilder.AppendFormat("0{0}", delimiter)
                Dim count As Integer = -insertPos
                While (count > 0)
                    strBuilder.Append("0")
                    count -= 1
                End While
                strBuilder.Append(staString)
                staString = strBuilder.ToString()
            Else
                staString = staString.Insert(insertPos, delimiter)
            End If
            GetStationString = staString
        End If
    End Function

    'String to double
    'station could be the station with or without Derived
    Public Shared Function GetRawStation(ByVal station As String, _
                                         ByVal stationIndexInc As Double, _
                                        Optional ByVal settingsStation As AeccLandLib.IAeccSettingsStation = Nothing) As Double
        If settingsStation Is Nothing Then
            settingsStation = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.StationSettings
        End If

        Dim delimiter As String
        delimiter = "+"
        Select Case settingsStation.StationDelimiterCharacter.Value
            Case AeccLandLib.AeccStationDelimiterCharacterType.aeccStationDelimiterPlusSign
                delimiter = "+"
            Case AeccLandLib.AeccStationDelimiterCharacterType.aeccStationDelimiterMinusSign
                delimiter = "-"
            Case AeccLandLib.AeccStationDelimiterCharacterType.aeccStationDelimiterAutomatic
                delimiter = "+"
            Case AeccLandLib.AeccStationDelimiterCharacterType.aeccStationDelimiterUnderscore
                delimiter = "_"
            Case AeccLandLib.AeccStationDelimiterCharacterType.aeccStationDelimiterNone
                delimiter = " "
            Case Else
                Diagnostics.Debug.Assert(False)
                delimiter = "+"
        End Select

        'DID 1197617, if station uses station index format, we should re-calculate station value
        If (settingsStation.Format().Value = AeccLandLib.AeccStationFormatType.aeccStationIndexFormat) Then
            Dim index As Long = station.IndexOf(delimiter)
            Dim iStationIndex As Long = 0
            Dim dReminder As Double

            If (index > 0) Then
                iStationIndex = Long.Parse(Left(station, index))
                dReminder = Double.Parse(Right(station, Len(station) - index - 1))
            Else
                iStationIndex = 0
                dReminder = Double.Parse(station)
            End If

            GetRawStation = iStationIndex * stationIndexInc + dReminder
        Else
            Dim tmpStr As String
            'replace from index 1, not replace index 0
            tmpStr = Replace(station, delimiter, "", 1, 1)
            GetRawStation = Double.Parse(tmpStr)
        End If

    End Function

    Public Shared Function ACADMsgBox(ByVal msg As String, Optional ByVal title As String = Nothing, _
                                 Optional ByVal style As MsgBoxStyle = MsgBoxStyle.OkOnly) As MsgBoxResult
        If title Is Nothing Then
            title = LocalizedRes.ReportUtilitie_Msg_Title
        End If
        ACADMsgBox = MsgBox(msg, style, title)
    End Function

    Public Shared Sub OpenFileByDefaultBrowser(ByVal fileName As String)
        Try
            System.Diagnostics.Process.Start(fileName)
        Catch ex As Exception
            'file can't open
        End Try
    End Sub

    '-----------------------------------------------------------------------
    '   calc the direction in radians between two points
    '-----------------------------------------------------------------------
    ' returned radAngle is the North-Azimuth radian format
    Public Shared Function CalcDirRad(ByVal sNorth As Double, ByVal sEast As Double, ByVal eNorth As Double, ByVal eEast As Double) As Double

        Dim nDiff As Double
        Dim eDiff As Double

        nDiff = eNorth - sNorth
        eDiff = eEast - sEast

        If eDiff = 0.0# And nDiff = 0.0# Then
            ' the points are at the same location. This should never happened, if it does return a 0.00 direction
            CalcDirRad = 0.0#
            Exit Function

            ' return the direction for horizontal/vertical directions
        ElseIf nDiff = 0.0# Then
            If eDiff > 0 Then
                CalcDirRad = 0.5 * Math.PI
            Else
                CalcDirRad = 1.5 * Math.PI
            End If
            Exit Function

        ElseIf eDiff = 0.0# Then
            If nDiff > 0 Then
                CalcDirRad = 0.0#
            Else
                CalcDirRad = Math.PI
            End If
            Exit Function

        End If

        'calc the angle from the x axis
        CalcDirRad = Math.Atan2(Math.Abs(nDiff), Math.Abs(eDiff))

        ' calc the direction
        If nDiff > 0 And eDiff < 0 Then
            ' 270 to 360
            CalcDirRad = 1.5 * Math.PI + CalcDirRad
        ElseIf nDiff < 0 And eDiff < 0 Then
            ' 180 to 270
            CalcDirRad = 1.5 * Math.PI - CalcDirRad
        ElseIf nDiff < 0 And eDiff > 0 Then
            '90 to 180
            CalcDirRad = 0.5 * Math.PI + CalcDirRad
        Else
            '0-90
            CalcDirRad = 0.5 * Math.PI - CalcDirRad
        End If

    End Function

    Public Shared Function CalcDist(ByVal sNorth As Double, ByVal sEast As Double, ByVal eNorth As Double, ByVal eEast As Double) As Double
        CalcDist = (Math.Abs(sEast - eEast) ^ 2 + Math.Abs(sNorth - eNorth) ^ 2) ^ 0.5
    End Function

    Public Shared Function IsCorridorInLayout(ByVal oCorridor As AeccRoadLib.IAeccCorridor) As Boolean
        ' Block reference that hasn't been exploded should not be recognized in report dialog
        ' Check whether its owner - that block record is layout, in-block corridor would not be in the layout block owner
        Dim bRet As Boolean = True
        If oCorridor Is Nothing Then
            Return bRet
        End If
        Try
            Dim ownerObj As Object = ReportApplication.AeccXDatabase.ObjectIdToObject(oCorridor.OwnerID)
            Dim block As Autodesk.AutoCAD.Interop.Common.IAcadBlock = Nothing
            block = CType(ownerObj, Autodesk.AutoCAD.Interop.Common.IAcadBlock)
            If Not block Is Nothing And Not block.IsLayout() Then
                bRet = False
            End If
        Catch ex As Exception
            bRet = False
        End Try

        Return bRet
    End Function

    Public Shared Sub SetCultureInfo()
        OldCultureInfo = CultureInfo.CurrentCulture
        CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(OldCultureInfo.Name)
    End Sub

    Public Shared Sub RestoreCultureInfo()
        If Not OldCultureInfo Is Nothing Then
            CurrentThread.CurrentCulture = OldCultureInfo
        End If
    End Sub

    Public Shared Function RunModalDialog(ByVal form As System.Windows.Forms.Form) As System.Windows.Forms.DialogResult
        ReportUtilities.SetCultureInfo()
        'Dim result As System.Windows.Forms.DialogResult = _
        'Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(form)
        Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(form)
        ReportUtilities.RestoreCultureInfo()
        'Return result
    End Function

    Public Shared Function CompareObjectID(ByVal iObjectID As IntPtr, _
                                    ByVal oObjectID As Autodesk.AutoCAD.DatabaseServices.ObjectId) As Boolean
        If iObjectID = oObjectID.OldIdPtr Then
            CompareObjectID = True
        Else
            CompareObjectID = False
        End If
    End Function

    Private Shared OldCultureInfo As Globalization.CultureInfo = Nothing
End Class
