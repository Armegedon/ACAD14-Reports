' -----------------------------------------------------------------------------
' <copyright file="AlignStakeout_ExtractData.vb" company="Autodesk">
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

Friend Class AlignStakeout_ExtractData

    Public Class AlignStakeoutData
        Public Station As String
        Public Direction As String
        Public Distance As String
        Public Northing As String
        Public Easting As String
    End Class

    Private Shared m_DataDictionary As New Collections.Generic.Dictionary(Of String, AlignStakeoutData)
    Public Shared ReadOnly Property DataDictionary() As Collections.Generic.Dictionary(Of String, AlignStakeoutData)
        Get
            Return m_DataDictionary
        End Get
    End Property

    Private Shared Function FormatCoordSettings(ByVal dDis As Double) As String
        Dim oCoordSettings As AeccLandLib.AeccSettingsCoordinate
        oCoordSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.CoordinateSettings

        FormatCoordSettings = ReportFormat.FormatDistance(dDis, oCoordSettings.Unit.Value, oCoordSettings.Precision.Value, _
                   oCoordSettings.Rounding.Value, oCoordSettings.Sign.Value)
    End Function
    Private Shared Function FormatDistSettings(ByVal dDis As Double) As String
        Dim oDistSettings As AeccLandLib.IAeccSettingsDistance
        oDistSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.DistanceSettings

        FormatDistSettings = ReportFormat.FormatDistance(dDis, oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                    oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
    End Function
    Private Shared Function FormatAngleSettings(ByVal dAngle As Double) As String
        Dim oAngleSettings As AeccLandLib.AeccSettingsAngle
        oAngleSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.AngleSettings

        FormatAngleSettings = ReportFormat.FormatAngle(dAngle, oAngleSettings.Unit.Value, oAngleSettings.Precision.Value, _
                                         oAngleSettings.Rounding.Value, oAngleSettings.Format.Value, oAngleSettings.Sign.Value)
    End Function
    Private Shared Function FormatDirSettings(ByVal dDirection As Double) As String
        Dim oDirSettings As AeccLandLib.AeccSettingsDirection
        oDirSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.DirectionSettings

        FormatDirSettings = ReportFormat.FormatDirection(dDirection, oDirSettings.Unit.Value, oDirSettings.Precision.Value, _
                                           oDirSettings.Rounding.Value, oDirSettings.Format.Value, _
                                           oDirSettings.Direction.Value, oDirSettings.Capitalization.Value, _
                                          oDirSettings.Sign.Value, oDirSettings.MeasurementType.Value, oDirSettings.BearingQuadrant.Value)
    End Function

    Public Shared Function ExtractData(ByVal oAlignment As AeccLandLib.AeccAlignment, _
                                        ByVal stationStart As Double, _
                                        ByVal stationEnd As Double, _
                                        ByVal stationIncrement As Double, _
                                        ByVal stationOffset As Double, _
                                        ByVal eAngleType As ReportForm_AlignStakeout.StakeAngleType, _
                                        ByVal OccupiedPt As AeccLandLib.AeccPoint, _
                                        ByVal BacksigntPt As AeccLandLib.AeccPoint) As Boolean

        m_DataDictionary.Clear()

        Dim sStations As AeccLandLib.AeccAlignmentStations
        sStations = oAlignment.GetStations(Autodesk.AECC.Interop.Land.AeccStationType.aeccAll, stationIncrement, stationIncrement)

        Dim settngsStation As AeccLandLib.AeccSettingsStation
        settngsStation = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.StationSettings

        For Each oStation As AeccLandLib.AeccAlignmentStation In sStations

            If oStation.Type = AeccLandLib.AeccStationType.aeccMajor Or _
            oStation.Type = AeccLandLib.AeccStationType.aeccMinor Or _
            oStation.Station = oAlignment.StartingStation Then

                Dim stationRounded As Double
                Dim stationString As String
                stationString = ReportUtilities.GetStationString(oAlignment, oStation.Station)
                stationRounded = ReportUtilities.GetRawStation(stationString, oAlignment.StationIndexIncrement, settngsStation)

                If stationRounded >= stationStart And stationRounded <= stationEnd Then

                    Dim oData As New AlignStakeoutData
                    Dim reportPtNorthing As Double, reportPtEasting As Double

                    ' Get reported Point using offset
                    oAlignment.PointLocation(oStation.Station, stationOffset, reportPtEasting, reportPtNorthing)

                    oData.Station = stationString
                    oData.Direction = GetDirectAngleInfo(reportPtNorthing, _
                        reportPtEasting, eAngleType, OccupiedPt, BacksigntPt)
                    oData.Distance = FormatDistSettings(ReportUtilities.CalcDist(reportPtNorthing, _
                        reportPtEasting, OccupiedPt.Northing, OccupiedPt.Easting))
                    oData.Northing = FormatCoordSettings(reportPtNorthing)
                    oData.Easting = FormatCoordSettings(reportPtEasting)

                    If Not m_DataDictionary.ContainsKey(oData.Station) Then
                        m_DataDictionary.Add(oData.Station, oData)
                    End If
                End If
            End If
        Next

        ExtractData = True
    End Function

    Private Shared Function GetDirectAngleInfo(ByVal dPtNorthing As Double, _
                                                ByVal dPtEasting As Double, _
                                                ByVal eAngleType As ReportForm_AlignStakeout.StakeAngleType, _
                                                ByVal OccupiedPt As AeccLandLib.AeccPoint, _
                                                ByVal BacksigntPt As AeccLandLib.AeccPoint) As String
        Dim sResult As String
        If eAngleType = ReportForm_AlignStakeout.StakeAngleType.DIRECTION Then
            Dim rawDirNAzimuth As Double
            rawDirNAzimuth = ReportUtilities.CalcDirRad(OccupiedPt.Northing, OccupiedPt.Easting, dPtNorthing, dPtEasting)
            sResult = FormatDirSettings(2 * Math.Atan(1) - rawDirNAzimuth) ' frome N Azimuth to x axis based.

        Else
            Dim rawAngle As Double
            rawAngle = ReportForm_AlignStakeout.Calc3PointAngle(OccupiedPt.Northing, OccupiedPt.Easting, _
                                       dPtNorthing, dPtEasting, _
                                       BacksigntPt.Northing, BacksigntPt.Easting, eAngleType)
            sResult = FormatAngleSettings(rawAngle)
        End If
        GetDirectAngleInfo = sResult
    End Function

End Class
