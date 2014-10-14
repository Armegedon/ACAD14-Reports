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

Public g_staInc As Double

Public Const nStationIndex = 0
Public Const nNorthingIndex = 1
Public Const nEastingIndex = 2
Public Const nDirectionIndex = 3
Public Const nLastIndex = nDirectionIndex

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
    Dim oStations As AeccAlignmentStations
    Dim oStation As AeccAlignmentStation
    Dim curStation As Double
    Dim curNorthing As Double
    Dim curEasting As Double
    Dim curTangent As Double
    Dim nCur As Integer
    
    Set oStations = oAlignment.GetStations(aeccAll, 0#, 5#)
    ReDim g_oAlignDataArr(Int((oAlignment.EndingStation - oAlignment.StartingStation) / g_staInc))
    nCur = 0
    
    For curStation = g_staInc * Int(oAlignment.StartingStation / g_staInc) To oAlignment.EndingStation Step g_staInc
        If curStation >= stationStart And curStation <= stationEnd Then
            If oAlignment.PointLocationEx(curStation, 0#, 0#, curEasting, curNorthing, curTangent) = aeccOK Then
                oDataArr(nStationIndex) = oAlignment.GetStationStringWithEquations(curStation)
                oDataArr(nNorthingIndex) = FormatCoordSettings(curNorthing)
                oDataArr(nEastingIndex) = FormatCoordSettings(curEasting)
                oDataArr(nDirectionIndex) = FormatDirSettings(2 * Atn(1) - curTangent)
                g_oAlignDataArr(nCur) = oDataArr
            nCur = nCur + 1
            End If
        End If
    
    Next
   
    ExtractData = True
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
