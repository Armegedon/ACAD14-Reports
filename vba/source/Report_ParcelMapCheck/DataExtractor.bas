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

Public g_ParcelCheckData As New ParcelCheckData

Function FormatDistSettings_Parcel(ByVal dDis As Double) As String
    Dim oDistSettings As AeccSettingsDistance
    Set oDistSettings = g_oAeccDb.Settings.ParcelSettings.AmbientSettings.DistanceSettings
    
    FormatDistSettings_Parcel = FormatDistance(dDis, oDistSettings.Unit.Value, oDistSettings.precision.Value, _
                oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
End Function

Function FormateCoordSettings_Parcel(ByVal dCoord As Double) As String
    Dim oCoordSettings As AeccSettingsCoordinate
    Set oCoordSettings = g_oAeccDb.Settings.ParcelSettings.AmbientSettings.CoordinateSettings
    
    FormateCoordSettings_Parcel = FormatDistance(dCoord, oCoordSettings.Unit.Value, oCoordSettings.precision.Value, _
                oCoordSettings.Rounding.Value, oCoordSettings.Sign.Value)
End Function

Function FormatDirSettings_Parcel(ByVal dDirection As Double) As String
    Dim oDirSettings As AeccSettingsDirection
    Set oDirSettings = g_oAeccDb.Settings.ParcelSettings.AmbientSettings.DirectionSettings
    
    FormatDirSettings_Parcel = FormatDirection(dDirection, oDirSettings.Unit.Value, oDirSettings.precision.Value, _
                                        oDirSettings.Rounding.Value, oDirSettings.Format.Value, _
                                        oDirSettings.Direction.Value, oDirSettings.Capitalization.Value, _
                                        oDirSettings.Sign.Value, oDirSettings.MeasurementType.Value, oDirSettings.BearingQuadrant)
End Function

Function FormatAngleSettings_Parcel(ByVal dAngle As Double) As String
    Dim oAngleSettings As AeccSettingsAngle
    Set oAngleSettings = g_oAeccDb.Settings.ParcelSettings.AmbientSettings.AngleSettings

    FormatAngleSettings_Parcel = FormatAngle(dAngle, oAngleSettings.Unit, oAngleSettings.precision, _
                                      oAngleSettings.Rounding, oAngleSettings.Format, oAngleSettings.Sign)
End Function

Function FormatAreaSettings_Parcel(ByVal dArea As Double) As String
    Dim oAreaSettings As AeccSettingsArea
    Set oAreaSettings = g_oAeccDb.Settings.ParcelSettings.AmbientSettings.areaSettings
    
    FormatAreaSettings_Parcel = FormatArea(dArea, oAreaSettings.Unit, oAreaSettings.precision.Value, _
                oAreaSettings.Rounding.Value, oAreaSettings.Sign.Value)
End Function

Function FormatDistSettings_Figure(ByVal dDis As Double) As String
    Dim oDistSettings As AeccSettingsDistance
    Dim oSurveySettings As AeccSurveySettingsRoot
    Set oSurveySettings = g_oSurveyDocument.Settings
    Set oDistSettings = oSurveySettings.SurveySettings.AmbientSettings.DistanceSettings
    
    FormatDistSettings_Figure = FormatDistance(dDis, oDistSettings.Unit.Value, oDistSettings.precision.Value, _
                oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
End Function
Function FormatAreaSettings_Figure(ByVal dArea As Double) As String
    Dim oAreaSettings As AeccSettingsArea
    Set oAreaSettings = g_oAeccDb.Settings.ParcelSettings.AmbientSettings.areaSettings
    
    FormatAreaSettings_Figure = FormatArea(dArea, oAreaSettings.Unit, oAreaSettings.precision.Value, _
                oAreaSettings.Rounding.Value, oAreaSettings.Sign.Value)
End Function

'the return angle/direction is x axis based and counterclockwise
Function RoundDirection_Parcel(ByVal sNorth As Double, ByVal sEast As Double, ByVal eNorth As Double, ByVal eEast As Double) As Double
    Dim dirSettings As AeccSettingsDirection
    Dim tmpVal As Double
    Set dirSettings = g_oAeccDb.Settings.ParcelSettings.AmbientSettings.DirectionSettings
    tmpVal = 2 * Atn(1) - CalcDirRad(sNorth, sEast, eNorth, eEast)
    RoundDirection_Parcel = RoundVal(tmpVal, dirSettings.precision, dirSettings.Rounding)
End Function

Function RoundLength_Parcel(ByVal Length As Double) As Double
    Dim distSettings As AeccSettingsDistance
    Set distSettings = g_oAeccDb.Settings.ParcelSettings.AmbientSettings.DistanceSettings
    RoundLength_Parcel = RoundVal(Length, distSettings.precision, distSettings.Rounding)
End Function

' input angle is x-axis based
Sub CalculateNextPoint(ByVal curNorth As Double, ByVal curEast As Double, _
                               ByRef nextNorth As Double, ByRef nextEast As Double, _
                               ByVal dAngle As Double, ByVal dLength As Double)
    Dim polarsin As Double
    Dim polarcos As Double
    
    polarsin = Sin(dAngle)
    polarcos = Cos(dAngle)
    
    nextEast = curEast + dLength * polarcos
    nextNorth = curNorth + dLength * polarsin
End Sub


Public Function ExtractData(oParcel As AeccParcel, bCounterClockWise As Boolean, bAcrossChord As Boolean) As Boolean
    If oParcel.ParcelLoops.count <= 0 Then
        ExtractData = False
        Exit Function
    End If
    Dim oLoop As AeccParcelLoop
    Set oLoop = oParcel.ParcelLoops(0)
    If oParcel.ParcelLoops(0).count <= 0 Then
        ExtractData = False
        Exit Function
    End If
    
    'set segments data
    Dim segCount As Long
    Dim curIndex As Long
    
    'clean the segmentsData
    g_ParcelCheckData.Reset bCounterClockWise, bAcrossChord
    
    segCount = oLoop.count
    curIndex = 0
    Dim oCurSeg As AeccParcelSegmentElement
    Dim curNorth As Double, curEast As Double
    While curIndex < segCount
        If bCounterClockWise = False Then
            Set oCurSeg = oLoop(curIndex)
        Else
            Set oCurSeg = oLoop(segCount - curIndex - 1)
        End If
        
        If TypeOf oCurSeg Is AeccParcelSegmentLine Then
            g_ParcelCheckData.AddLine oCurSeg
        ElseIf TypeOf oCurSeg Is AeccParcelSegmentCurve Then
            g_ParcelCheckData.AddCurve oCurSeg
        Else
        End If
        
        curIndex = curIndex + 1
    Wend
    
    'Fill the left Datas
    g_ParcelCheckData.Area = oParcel.Statistics.Area
    g_ParcelCheckData.Perimeter = oParcel.Statistics.Perimeter
    
    ExtractData = True
    
End Function

         
                       



