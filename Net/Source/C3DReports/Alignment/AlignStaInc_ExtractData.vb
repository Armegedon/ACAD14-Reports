' -----------------------------------------------------------------------------
' <copyright file="AlignStaInc_ExtractData.vb" company="Autodesk">
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
Imports AutoCAD = Autodesk.AutoCAD.Interop
Imports AeccUiLandLib = Autodesk.AECC.Interop.UiLand
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AEC.Interop.UIBase

Friend Class AlignStaInc_ExtractData

    Public Class DirectionData
        Public station As String
        Public northing As String
        Public easting As String
        Public tangentialDirection As String
    End Class

    Public Shared m_AlignDataList As New Collections.Generic.List(Of DirectionData)

    Public Shared ReadOnly Property AlignDataList() As Collections.Generic.List(Of DirectionData)
        Get
            Return m_AlignDataList
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
                                ByVal staIncment As Double) As Boolean

        m_AlignDataList.Clear()

        Dim settingsStation As AeccLandLib.AeccSettingsStation
        settingsStation = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.StationSettings
        Dim tolerance As Double
        tolerance = Math.Pow(10, -settingsStation.Precision.Value)

        Dim curStation As Double
        For curStation = stationStart To stationEnd Step staIncment
            Dim curNorthing As Double
            Dim curEasting As Double
            Dim curTangent As Double

            Try
                If oAlignment.PointLocationEx(curStation, 0.0, tolerance, curEasting, curNorthing, curTangent) _
            = AeccLandLib.AeccAlignmentReturnValue.aeccOK Then
                    Dim stationString As String
                    stationString = ReportUtilities.GetStationStringWithDerived(oAlignment, curStation)

                    Dim alignData As New DirectionData
                    alignData.station = stationString
                    alignData.northing = FormatCoordSettings(curNorthing)
                    alignData.easting = FormatCoordSettings(curEasting)
                    alignData.tangentialDirection = FormatDirSettings(Math.PI / 2 - curTangent)

                    m_AlignDataList.Add(alignData)
                End If
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
            End Try
        Next

        If stationEnd - (curStation - staIncment) > tolerance Then
            Dim endNorthing As Double
            Dim endEasting As Double
            Dim endTangent As Double
            Try
                If oAlignment.PointLocationEx(stationEnd, 0.0, tolerance, endEasting, endNorthing, endTangent) _
            = AeccLandLib.AeccAlignmentReturnValue.aeccOK Then
                    Dim stationString As String
                    stationString = ReportUtilities.GetStationStringWithDerived(oAlignment, stationEnd)

                    Dim alignData As New DirectionData
                    alignData.station = stationString
                    alignData.northing = FormatCoordSettings(endNorthing)
                    alignData.easting = FormatCoordSettings(endEasting)
                    alignData.tangentialDirection = FormatDirSettings(Math.PI / 2 - endTangent)

                    m_AlignDataList.Add(alignData)
                End If
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
            End Try
        End If

        ExtractData = True
    End Function

End Class
