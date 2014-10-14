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

Enum StakeAngleType
    TurnedPlus
    TurnedMinus
    DeflectPlus
    DeflectMinus
    Direction
End Enum

Public Const nStationIndex = 0
Public Const nDirectionIndex = 1
Public Const nDistance = 2
Public Const nNorthingIndex = 3
Public Const nEastingIndex = 4
Public Const nLastIndex = nEastingIndex

Public g_angleType As StakeAngleType
Public g_staInc As Double
Public g_offSet As Double
Public g_oOccupiedPt As AeccPoint
Public g_oBacksightPt As AeccPoint

Public g_oAlignDataArr()
Function FormatCoordSettings(ByVal dDis As Double) As String
    Dim oCoordSettings As AeccSettingsCoordinate
    Set oCoordSettings = g_oAeccDb.settings.AlignmentSettings.AmbientSettings.CoordinateSettings
        
    FormatCoordSettings = FormatDistance(dDis, oCoordSettings.Unit.Value, oCoordSettings.Precision.Value, _
                oCoordSettings.Rounding.Value, oCoordSettings.Sign.Value)
End Function
Function FormatDistSettings(ByVal dDis As Double) As String
    Dim oDistSettings As AeccSettingsDistance
    Set oDistSettings = g_oAeccDb.settings.AlignmentSettings.AmbientSettings.DistanceSettings
    
    FormatDistSettings = FormatDistance(dDis, oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
End Function
Function FormatAngleSettings(ByVal dAngle As Double) As String
    Dim oAngleSettings As AeccSettingsAngle
    Set oAngleSettings = g_oAeccDb.settings.AlignmentSettings.AmbientSettings.AngleSettings

    FormatAngleSettings = FormatAngle(dAngle, oAngleSettings.Unit, oAngleSettings.Precision, _
                                      oAngleSettings.Rounding, oAngleSettings.Format, oAngleSettings.Sign)
End Function
Function FormatDirSettings(ByVal dDirection As Double) As String
    Dim oDirSettings As AeccSettingsDirection
    Set oDirSettings = g_oAeccDb.settings.AlignmentSettings.AmbientSettings.DirectionSettings
    
    FormatDirSettings = FormatDirection(dDirection, oDirSettings.Unit.Value, oDirSettings.Precision.Value, _
                                        oDirSettings.Rounding.Value, oDirSettings.Format.Value, _
                                        oDirSettings.Direction.Value, oDirSettings.Capitalization.Value, _
                                        oDirSettings.Sign.Value, oDirSettings.MeasurementType.Value, oDirSettings.BearingQuadrant)
End Function

Public Function ExtractData(oAlignment As AeccAlignment, stationStart As Double, stationEnd As Double) As Boolean
    Dim oDataArr(nLastIndex) As Variant
    Dim sStations As AeccAlignmentStations
    Dim oStation As AeccAlignmentStation
    Dim nCur As Integer
    
    
    Set sStations = oAlignment.GetStations(aeccAll, g_staInc, g_staInc)
    ReDim g_oAlignDataArr(sStations.Count)
    nCur = 0
    
    For Each oStation In sStations
            With oStation
                If .Type = aeccMajor Or .Type = aeccMinor Or .Station = oAlignment.StartingStation Then
                    If oStation.Station >= stationStart And oStation.Station <= stationEnd Then
                        Dim reportPtNorthing As Double, reportPtEasting As Double
                        ' Get reported Point using offset
                        oAlignment.PointLocation .Station, g_offSet, reportPtEasting, reportPtNorthing
                        
                        oDataArr(nStationIndex) = oAlignment.GetStationStringWithEquations(.Station)
                        oDataArr(nDirectionIndex) = GetDirectAngleInfo(reportPtNorthing, reportPtEasting)
                        oDataArr(nDistance) = FormatDistSettings(CalcDist(reportPtNorthing, reportPtEasting, g_oOccupiedPt.Northing, g_oOccupiedPt.Easting))
                        oDataArr(nNorthingIndex) = FormatCoordSettings(reportPtNorthing)
                        oDataArr(nEastingIndex) = FormatCoordSettings(reportPtEasting)
                        g_oAlignDataArr(nCur) = oDataArr
                        nCur = nCur + 1
                    End If
                End If
            End With
    Next
    

    ExtractData = True
End Function

Private Function GetDirectAngleInfo(dPtNorthing As Double, dPtEasting As Double) As String
    Dim sResult As String
    If g_angleType = Direction Then
        Dim rawDirNAzimuth As Double
        rawDirNAzimuth = CalcDirRad(g_oOccupiedPt.Northing, g_oOccupiedPt.Easting, dPtNorthing, dPtEasting)
        sResult = FormatDirSettings(2 * Atn(1) - rawDirNAzimuth) ' frome N Azimuth to x axis based.
 '       sResult = Get2PointDirData(g_oOccupiedPt.Northing, g_oOccupiedPt.Easting, dPtNorthing, dPtEasting, 2, cDMS, Bearing)
    Else
        Dim rawAngle As Double
        rawAngle = Calc3PointAngle(g_oOccupiedPt.Northing, g_oOccupiedPt.Easting, _
                                   dPtNorthing, dPtEasting, _
                                   g_oBacksightPt.Northing, g_oBacksightPt.Easting, g_angleType)
        sResult = FormatAngleSettings(rawAngle)
'        sResult = Report_Utilities.Utilities.FormatAngle(rawAngle, cDMS)
    End If
    GetDirectAngleInfo = sResult
End Function


Public Function GetUnitStr() As String
    Dim settings
    Dim unitsStr As String
    
    Set settings = g_oAeccDb.settings.DrawingSettings.AmbientSettings.DistanceSettings.Unit
    If settings = AeccCoordinateUnitType.aeccCoordinateUnitFoot Then
        unitsStr = "Feet"
    ElseIf settings = AeccCoordinateUnitType.aeccCoordinateUnitMeter Then
        unitsStr = "Meters"
    End If
    GetUnitStr = unitsStr
End Function

