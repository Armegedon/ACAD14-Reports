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

Option Explicit On

Imports System
Imports AeccLandLib = Autodesk.AECC.Interop.Land

Friend Class AlignSta_ExtractData
    Public Const INDEX_STATION As UInteger = 0
    Public Const INDEX_NORTHING As UInteger = 1
    Public Const INDEX_EASTING As UInteger = 2
    Public Const INDEX_LENGTHVALUE As UInteger = 3
    Public Const INDEX_DIRECTION As UInteger = 4
    Public Const INDEX_NORTHING_VAL As UInteger = 5
    Public Const INDEX_EASTING_VAL As UInteger = 6
    Public Const INDEX_LAST As UInteger = INDEX_EASTING_VAL

    Public Shared AlignDataArr As New SortedList

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

    Public Shared Function ExtractData(ByVal oAlignment As AeccLandLib.AeccAlignment, ByVal stationStart As Double, ByVal stationEnd As Double) As Boolean

        Dim oStations As AeccLandLib.AeccAlignmentStations
        Dim oStation As AeccLandLib.AeccAlignmentStation
        oStations = oAlignment.GetStations(AeccLandLib.AeccStationType.aeccAll, 0.0, 0.0)
        Dim oStationSetting As AeccLandLib.AeccSettingsStation
        oStationSetting = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.StationSettings
        Dim nCur As Integer
        AlignDataArr.Clear()
        nCur = 0
        Dim dLastStation As Double
        dLastStation = 0.0

        For Each oStation In oStations
            Dim nCurStationValue As Double
            nCurStationValue = ReportFormat.RoundDouble(oStation.Station, oStationSetting.Precision.Value, oStationSetting.Rounding.Value)
            If nCurStationValue >= stationStart And _
               nCurStationValue <= stationEnd Then
                With oStation
                    If nCur > 0 And ReportFormat.RoundDouble(dLastStation, oStationSetting.Precision.Value, oStationSetting.Rounding.Value) _
                    = nCurStationValue Then
                        nCur = nCur - 1
                    ElseIf .Type = AeccLandLib.AeccStationType.aeccMajor Then
                        nCur = nCur - 1
                    Else
                        If .GeometryPointType = AeccLandLib.AeccGeometryPointType.aeccBegOfAlign Or _
                        .GeometryPointType = AeccLandLib.AeccGeometryPointType.aeccEndOfAlign Or _
                        .GeometryPointType = AeccLandLib.AeccGeometryPointType.aeccCPI Or _
                        .GeometryPointType = AeccLandLib.AeccGeometryPointType.aeccPI Or _
                        .GeometryPointType = AeccLandLib.AeccGeometryPointType.aeccSPI Then

                            Dim oDataArr(INDEX_LAST) As Object
                            oDataArr(INDEX_STATION) = ReportUtilities.GetStationStringWithDerived(oAlignment, .Station)
                            oDataArr(INDEX_NORTHING) = FormatCoordSettings(.Northing)
                            oDataArr(INDEX_EASTING) = FormatCoordSettings(.Easting)
                            oDataArr(INDEX_NORTHING_VAL) = oStation.Northing
                            oDataArr(INDEX_EASTING_VAL) = oStation.Easting
                            dLastStation = .Station
                            AlignDataArr(.Station) = oDataArr
                        Else
                            nCur = nCur - 1
                        End If
                    End If
                End With

                nCur = nCur + 1
            End If
        Next

        If AlignDataArr.Count >= 2 Then
            Dim lastNorthing As Double
            Dim lastEasting As Double

            Dim firstData As Boolean = True
            For Each dataArr() As Object In AlignDataArr.Values
                If firstData Then
                    lastEasting = dataArr(INDEX_EASTING_VAL)
                    lastNorthing = dataArr(INDEX_NORTHING_VAL)
                    dataArr(INDEX_LENGTHVALUE) = ""
                    dataArr(INDEX_DIRECTION) = ""
                    firstData = False
                Else
                    dataArr(INDEX_LENGTHVALUE) = FormatDistSettings(ReportUtilities.CalcDist(dataArr(INDEX_NORTHING_VAL), _
                        dataArr(INDEX_EASTING_VAL), lastNorthing, lastEasting))
                    dataArr(INDEX_DIRECTION) = FormatDirSettings(Math.PI / 2 - _
                        ReportUtilities.CalcDirRad(lastNorthing, lastEasting, dataArr(INDEX_NORTHING_VAL), dataArr(INDEX_EASTING_VAL)))
                    lastEasting = dataArr(INDEX_EASTING_VAL)
                    lastNorthing = dataArr(INDEX_NORTHING_VAL)
                End If
            Next
        End If

        ExtractData = True
    End Function
End Class
