''//
''// (C) Copyright 2008 by Autodesk, Inc.
''//
''// Permission to use, copy, modify, and distribute this software in
''// object code form for any purpose and without fee is hereby granted,
''// provided that the above copyright notice appears in all copies and
''// that both that copyright notice and the limited warranty and
''// restricted rights notice below appear in all supporting
''// documentation.
''//
''// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
''// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
''// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
''// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
''// UNINTERRUPTED OR ERROR FREE.
''//
''// Use, duplication, or disclosure by the U.S. Government is subject to
''// restrictions set forth in FAR 52.227-19 (Commercial Computer
''// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
''// (Rights in Technical Data and Computer Software), as applicable.
'
'
'
Option Explicit On

Imports System
Imports AeccLandLib = Autodesk.AECC.Interop.Land


Friend Class ParcelMapCheck_ExtractData

    Private Shared m_oParcelCheckData As New ParcelCheckData
    Public Shared SurveyDocument As Autodesk.AECC.Interop.UiSurvey.AeccSurveyDocument
    Public Shared ReadOnly Property ParcelCheckData() As ParcelCheckData
        Get
            Return m_oParcelCheckData
        End Get
    End Property

    Public Shared Function FormatDistSettings_Parcel(ByVal dDis As Double) As String
        Dim oDistSettings As AeccLandLib.AeccSettingsDistance
        oDistSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.DistanceSettings

        FormatDistSettings_Parcel = ReportFormat.FormatDistance(dDis, oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                    oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
    End Function

    Public Shared Function FormateCoordSettings_Parcel(ByVal dCoord As Double) As String
        Dim oCoordSettings As AeccLandLib.AeccSettingsCoordinate
        oCoordSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.CoordinateSettings

        FormateCoordSettings_Parcel = ReportFormat.FormatDistance(dCoord, oCoordSettings.Unit.Value, oCoordSettings.Precision.Value, _
                    oCoordSettings.Rounding.Value, oCoordSettings.Sign.Value)
    End Function

    Public Shared Function FormatDirSettings_Parcel(ByVal dDirection As Double) As String
        Dim oDirSettings As AeccLandLib.AeccSettingsDirection
        oDirSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.DirectionSettings

        FormatDirSettings_Parcel = ReportFormat.FormatDirection(dDirection, _
                                            oDirSettings.Unit.Value, _
                                            oDirSettings.Precision.Value, _
                                            oDirSettings.Rounding.Value, _
                                            oDirSettings.Format.Value, _
                                            oDirSettings.Direction.Value, _
                                            oDirSettings.Capitalization.Value, _
                                            oDirSettings.Sign.Value, _
                                            oDirSettings.MeasurementType.Value, _
                                            oDirSettings.BearingQuadrant.Value)
    End Function

    Public Shared Function FormatAngleSettings_Parcel(ByVal dAngle As Double) As String
        Dim oAngleSettings As AeccLandLib.AeccSettingsAngle
        oAngleSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.AngleSettings

        FormatAngleSettings_Parcel = ReportFormat.FormatAngle(dAngle, _
                                            oAngleSettings.Unit.Value, _
                                            oAngleSettings.Precision.Value, _
                                            oAngleSettings.Rounding.Value, _
                                            oAngleSettings.Format.Value, _
                                            oAngleSettings.Sign.Value)
    End Function

    Public Shared Function FormatAreaSettings_Parcel(ByVal dArea As Double) As String
        Dim oAreaSettings As AeccLandLib.AeccSettingsArea
        oAreaSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.AreaSettings

        FormatAreaSettings_Parcel = ReportFormat.FormatArea(dArea, _
                                            oAreaSettings.Unit.Value, _
                                            oAreaSettings.Precision.Value, _
                                            oAreaSettings.Rounding.Value, _
                                            oAreaSettings.Sign.Value)
    End Function

    Public Shared Function FormatDistSettings_Figure(ByVal dDis As Double) As String
        Try
            Dim oDistSettings As AeccLandLib.AeccSettingsDistance
            Dim oSurveySettings As Autodesk.AECC.Interop.Survey.AeccSurveySettingsRoot
            oSurveySettings = SurveyDocument.Settings
            oDistSettings = oSurveySettings.SurveySettings.AmbientSettings.DistanceSettings

            FormatDistSettings_Figure = ReportFormat.FormatDistance(dDis, _
                        oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                        oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
        Catch ex As Exception
            FormatDistSettings_Figure = ""
        End Try
    End Function
    Public Shared Function FormatAreaSettings_Figure(ByVal dArea As Double) As String
        Dim oAreaSettings As AeccLandLib.AeccSettingsArea
        oAreaSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.AreaSettings

        FormatAreaSettings_Figure = ReportFormat.FormatArea(dArea, _
                    oAreaSettings.Unit.Value, oAreaSettings.Precision.Value, _
                    oAreaSettings.Rounding.Value, oAreaSettings.Sign.Value)
    End Function

    'the return angle/direction is x axis based and counterclockwise
    Public Shared Function RoundDirection_Parcel(ByVal sNorth As Double, ByVal sEast As Double, ByVal eNorth As Double, ByVal eEast As Double) As Double
        Dim dirSettings As AeccLandLib.AeccSettingsDirection
        Dim angleValInRadians As Double
        dirSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.DirectionSettings
        angleValInRadians = Math.PI / 2 - ReportUtilities.CalcDirRad(sNorth, sEast, eNorth, eEast)

        Dim tmpVal As Double
        tmpVal = ReportFormat.convertUnits(angleValInRadians, AeccLandLib.AeccAngleUnitType.aeccAngleUnitRadian, _
            dirSettings.Unit.Value)

        If (dirSettings.Format.Value = AeccLandLib.AeccFormatType.aeccFormatDecimal) Then
            tmpVal = ReportFormat.RoundDouble(tmpVal, _
                                            dirSettings.Precision.Value, dirSettings.Rounding.Value)
        Else
            tmpVal = ReportFormat.RoundDMS_2(tmpVal, _
                                            dirSettings.Precision.Value, dirSettings.Rounding.Value)
        End If

        angleValInRadians = ReportFormat.convertUnits(tmpVal, dirSettings.Unit.Value, _
            AeccLandLib.AeccAngleUnitType.aeccAngleUnitRadian)

        Return angleValInRadians

    End Function

    Public Shared Function RoundLength_Parcel(ByVal Length As Double) As Double
        Dim distSettings As AeccLandLib.AeccSettingsDistance
        distSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.DistanceSettings
        RoundLength_Parcel = ReportFormat.RoundDouble(Length, distSettings.Precision.Value, _
                                distSettings.Rounding.Value)
    End Function

    ' input angle is x-axis based
    Public Shared Sub CalculateNextPoint(ByVal curNorth As Double, ByVal curEast As Double, _
                                   ByRef nextNorth As Double, ByRef nextEast As Double, _
                                   ByVal dAngle As Double, ByVal dLength As Double)
        Dim polarsin As Double
        Dim polarcos As Double

        polarsin = Math.Sin(dAngle)
        polarcos = Math.Cos(dAngle)

        nextEast = curEast + dLength * polarcos
        nextNorth = curNorth + dLength * polarsin
    End Sub


    Public Shared Function ExtractData(ByVal oParcel As AeccLandLib.AeccParcel, _
                                ByVal bCounterClockWise As Boolean, _
                                ByVal bAcrossChord As Boolean) As Boolean
        If oParcel.ParcelLoops.Count <= 0 Then
            ExtractData = False
            Exit Function
        End If
        Dim oLoop As AeccLandLib.AeccParcelLoop
        oLoop = oParcel.ParcelLoops(0)
        If oParcel.ParcelLoops(0).Count <= 0 Then
            ExtractData = False
            Exit Function
        End If

        'set segments data
        Dim segCount As Long
        Dim curIndex As Long

        'clean the segmentsData
        m_oParcelCheckData.Reset(bCounterClockWise, bAcrossChord)

        segCount = oLoop.Count
        curIndex = 0
        Dim oCurSeg As AeccLandLib.AeccParcelSegmentElement
        'Dim curNorth As Double, curEast As Double
        While curIndex < segCount
            If bCounterClockWise = False Then
                oCurSeg = oLoop(curIndex)
            Else
                oCurSeg = oLoop(segCount - curIndex - 1)
            End If

            If TypeOf oCurSeg Is AeccLandLib.AeccParcelSegmentLine Then
                m_oParcelCheckData.AddLine(oCurSeg)
            ElseIf TypeOf oCurSeg Is AeccLandLib.AeccParcelSegmentCurve Then
                m_oParcelCheckData.AddCurve(oCurSeg)
            Else
            End If

            curIndex = curIndex + 1
        End While

        'Fill the left Datas
        m_oParcelCheckData.Area = oParcel.Statistics.Area
        m_oParcelCheckData.Perimeter = oParcel.Statistics.Perimeter

        ExtractData = True

    End Function

End Class
