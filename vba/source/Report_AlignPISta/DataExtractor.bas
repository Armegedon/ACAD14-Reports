Attribute VB_Name = "DataExtractor"
''//
''// (C) Copyright 2005 by Autodesk, Inc.
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
Option Explicit

Public Const nStationIndex = 0
Public Const nNorthingIndex = 1
Public Const nEastingIndex = 2
Public Const nLengthValueIndex = 3
Public Const nDirectionIndex = 4
Public Const nLastIndex = nDirectionIndex

Public g_oAlignDataArr()

Function FormatCoordSettings(ByVal dDis As Double) As String
    Dim oCoordSettings As AeccSettingsCoordinate
    Set oCoordSettings = g_oAeccDb.Settings.AlignmentSettings.AmbientSettings.CoordinateSettings
        
    FormatCoordSettings = FormatDistance(dDis, oCoordSettings.Unit.Value, oCoordSettings.Precision.Value, _
                oCoordSettings.Rounding.Value, oCoordSettings.Sign.Value)
End Function
Function FormatDistSettings(ByVal dDis As Double) As String
    Dim oDistSettings As AeccSettingsDistance
    Set oDistSettings = g_oAeccDb.Settings.AlignmentSettings.AmbientSettings.DistanceSettings
    
    FormatDistSettings = FormatDistance(dDis, oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
End Function
Function FormatDirSettings(ByVal dDirection As Double) As String
    Dim oDirSettings As AeccSettingsDirection
    Set oDirSettings = g_oAeccDb.Settings.AlignmentSettings.AmbientSettings.DirectionSettings
    
    FormatDirSettings = FormatDirection(dDirection, oDirSettings.Unit.Value, oDirSettings.Precision.Value, _
                                        oDirSettings.Rounding.Value, oDirSettings.Format.Value, _
                                        oDirSettings.Direction.Value, oDirSettings.Capitalization.Value, _
                                        oDirSettings.Sign.Value, oDirSettings.MeasurementType.Value, oDirSettings.BearingQuadrant)
End Function

Public Function ExtractData(oAlignment As AeccAlignment, stationStart As Double, stationEnd As Double) As Boolean
    Dim oDataArr(nLastIndex) As Variant
    Dim oStations As AeccAlignmentStations
    Dim oStation As AeccAlignmentStation
    Set oStations = oAlignment.GetStations(AeccStationType.aeccAll, 0#, 0#)
    Dim oStationSetting As AeccSettingsStation
    Set oStationSetting = g_oAeccDb.Settings.AlignmentSettings.AmbientSettings.StationSettings
    Dim nCur As Integer
    ReDim g_oAlignDataArr(oStations.Count)
    nCur = 0
    Dim dLastStation As Double, dLastNorthing As Double, dLastEasting As Double
    dLastStation = 0#
    dLastNorthing = 0#
    dLastEasting = 0#
    
    For Each oStation In oStations
        If RoundVal(oStation.Station, oStationSetting.Precision, oStationSetting.Rounding) >= stationStart And _
           RoundVal(oStation.Station, oStationSetting.Precision, oStationSetting.Rounding) <= stationEnd Then
            With oStation
                If nCur > 0 And RoundVal(dLastStation, oStationSetting.Precision, oStationSetting.Rounding) = RoundVal(.Station, oStationSetting.Precision, oStationSetting.Rounding) Then
                    nCur = nCur - 1
                ElseIf .Type = aeccMajor Then
                    nCur = nCur - 1
                Else
                    If .GeometryPointType = aeccBegOfAlign Or .GeometryPointType = aeccEndOfAlign Or _
                    .GeometryPointType = aeccCPI Or .GeometryPointType = aeccPI Or .GeometryPointType = aeccSPI Then
                        oDataArr(nStationIndex) = oAlignment.GetStationStringWithEquations(.Station)
                        oDataArr(nNorthingIndex) = FormatCoordSettings(.Northing)
                        oDataArr(nEastingIndex) = FormatCoordSettings(.Easting)
                        
                        If nCur > 0 Then
                            oDataArr(nLengthValueIndex) = FormatDistSettings(CalcDist(.Northing, .Easting, dLastNorthing, dLastEasting))
                            oDataArr(nDirectionIndex) = FormatDirSettings(2 * Atn(1) - CalcDirRad(dLastNorthing, dLastEasting, .Northing, .Easting))
                        Else
                            oDataArr(nLengthValueIndex) = ""
                            oDataArr(nDirectionIndex) = ""
                        End If
                        
                        dLastStation = .Station
                        dLastNorthing = .Northing
                        dLastEasting = .Easting
                        
                        g_oAlignDataArr(nCur) = oDataArr
                    Else
                        nCur = nCur - 1
                    End If
                End If
            End With
        
            nCur = nCur + 1
        End If
    Next
    
    ExtractData = True
End Function

