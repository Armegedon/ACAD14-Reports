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

Public Const nPointIndex = 0
Public Const nStationIndex = 1
Public Const nOffsetIndex = 2
Public Const nElevation = 3
Public Const nDescriptionRaw = 4
Public Const nDescriptionFull = 5
Public Const nLastIndex = nDescriptionFull

Public g_oPointDataArr(nLastIndex)


Function FormatPtElevationSettings(ByVal dDis As Double) As String
    Dim oElevationSettings As AeccSettingsElevation
    Set oElevationSettings = g_oAeccDb.settings.PointSettings.AmbientSettings.ElevationSettings
        
    FormatPtElevationSettings = FormatDistance(dDis, oElevationSettings.Unit.Value, oElevationSettings.Precision.Value, _
                oElevationSettings.Rounding.Value, oElevationSettings.Sign.Value)
End Function

Function FormatPtCoordSettings(ByVal dDis As Double) As String
    Dim oCoordSettings As AeccSettingsCoordinate
    Set oCoordSettings = g_oAeccDb.settings.PointSettings.AmbientSettings.CoordinateSettings
        
    FormatPtCoordSettings = FormatDistance(dDis, oCoordSettings.Unit.Value, oCoordSettings.Precision.Value, _
                oCoordSettings.Rounding.Value, oCoordSettings.Sign.Value)
End Function

Function FormatDistSettings(ByVal dDis As Double) As String
    Dim oDistSettings As AeccSettingsDistance
    Set oDistSettings = g_oAeccDb.settings.AlignmentSettings.AmbientSettings.DistanceSettings

    FormatDistSettings = FormatDistance(dDis, oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
End Function

Public Function ExtractData(oAlignment As AeccAlignment, oPoint As AeccPoint) As Boolean
    Dim ptNorthing As Double
    Dim ptEasting As Double
    Dim station As Double
    Dim offSet As Double
    
    ptNorthing = oPoint.Northing
    ptEasting = oPoint.Easting
    
    g_oPointDataArr(nPointIndex) = oPoint.Number
    g_oPointDataArr(nElevation) = FormatDistSettings(oPoint.elevation)
    g_oPointDataArr(nDescriptionRaw) = oPoint.RawDescription
    g_oPointDataArr(nDescriptionFull) = oPoint.FullDescription
    
    If oAlignment.StationOffsetEx(ptEasting, ptNorthing, 0#, station, offSet) = aeccStationOutOfRange Then
        g_oPointDataArr(nStationIndex) = "Out of range"
        g_oPointDataArr(nOffsetIndex) = "Out of range"
    Else
        g_oPointDataArr(nStationIndex) = oAlignment.GetStationStringWithEquations(station)
        g_oPointDataArr(nOffsetIndex) = FormatDistSettings(offSet)
    End If

    ExtractData = True
End Function
