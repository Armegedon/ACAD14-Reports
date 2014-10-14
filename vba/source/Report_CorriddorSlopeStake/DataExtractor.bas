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

Public Const nStartStationIndex = 0
Public Const nStartNorthingIndex = 1
Public Const nStartEastingIndex = 2
Public Const nEndStationIndex = 3
Public Const nEndNorthingIndex = 4
Public Const nEndEastingIndex = 5
Public Const nLengthValueIndex = 6
Public Const nDirectionIndex = 7
Public Const nLastIndex = nDirectionIndex

Public Const nLeftIndex = 0
Public Const nRightIndex = 1
Public Const nLeftEndIndex = 2
Public Const nRightEndIndex = 3

Public Const nCodeIndex = 0
Public Const nOffsetIndex = 1
Public Const nElevIndex = 2
Public Const nSlopeIndex = 3

Public g_oSlopeStakeData As New Dictionary

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

Function FormatElevSettings(ByVal dDis As Double) As String
    Dim oElevSettings As AeccSettingsElevation
    Set oElevSettings = g_oAeccDb.Settings.AlignmentSettings.AmbientSettings.ElevationSettings
    
    FormatElevSettings = FormatDistance(dDis, oElevSettings.Unit.Value, oElevSettings.Precision.Value, _
                oElevSettings.Rounding.Value, oElevSettings.Sign.Value)
End Function


Public Function ExtractData(oCorridor As AeccCorridor, oBaseline As AeccBaseline, oSampleLineGroups As AeccSampleLineGroup, stationStart As Double, stationEnd As Double, sCodeName As String) As Boolean
    Dim oSampleLine As AeccSampleLine
    Dim oLinks As AeccCalculatedLinks
    Dim curStation As Double
    Dim datumElevation As Double
    Dim cLeftDatas As Dictionary
    Dim cRightDatas As Dictionary
    ExtractData = False

    g_oSlopeStakeData.removeAll
    
    Dim dLastStation As Double
    dLastStation = -1122323.323

    For Each oSampleLine In oSampleLineGroups.SampleLines
        Dim aDataBySta(0 To 3) As Variant

        Dim oLink As AeccCalculatedLink
        Dim leftOff As Double
        Dim leftElev As Double
        Dim rightOff As Double
        Dim rightElev As Double
        Dim dLastLeftOff As Double
        Dim dLastLeftElev As Double
        Dim dLastRightOff As Double
        Dim dLastRightElev As Double
        curStation = oSampleLine.Station
        curStation = GetRawStation(oSampleLineGroups.Parent.GetStationStringWithEquations(curStation))
        If curStation < stationStart Or curStation > stationEnd Then
            GoTo ErrHandler
        End If
        If curStation - dLastStation < 0.001 Then
            GoTo ErrHandler
        Else
            dLastStation = curStation
        End If
        
        datumElevation = oBaseline.profile.ElevationAt(curStation)
        On Error GoTo ErrHandler
        Set oLinks = oBaseline.AppliedAssembly(curStation).GetLinksByCode(sCodeName)
        Set cLeftDatas = New Dictionary
        Set cRightDatas = New Dictionary
        leftElev = datumElevation
        leftOff = 0
        rightElev = datumElevation
        rightOff = 0
        For Each oLink In oLinks
            Dim oPoint As AeccCalculatedPoint
            Dim vx As Variant
            Dim off As Double
            For Each oPoint In oLink.CalculatedPoints
                Dim oDataArr(0 To 3)

                vx = oPoint.GetStationOffsetElevationToBaseline
                off = vx(1)
                If off < 0 Then
                    If Not cLeftDatas.Exists(off) Then
                        dLastLeftOff = leftOff
                        dLastLeftElev = leftElev
                        leftOff = off: leftElev = vx(2) + datumElevation
                        oDataArr(nOffsetIndex) = leftOff
                        oDataArr(nElevIndex) = leftElev
                        oDataArr(nCodeIndex) = oPoint.CorridorCodes(0)
                        oDataArr(nSlopeIndex) = GetSlopeData(leftOff, dLastLeftOff, leftElev, dLastLeftElev)
                        dLastLeftElev = leftElev
                        cLeftDatas.Add leftOff, oDataArr
                End If
                Else
                    If Not cRightDatas.Exists(off) Then
                        dLastRightOff = rightOff
                        dLastRightElev = rightElev
                        rightOff = off: rightElev = vx(2) + datumElevation
                        oDataArr(nOffsetIndex) = rightOff
                        oDataArr(nElevIndex) = rightElev
                        oDataArr(nCodeIndex) = oPoint.CorridorCodes(0)
                        oDataArr(nSlopeIndex) = GetSlopeData(rightOff, dLastRightOff, rightElev, dLastRightElev)
                        cRightDatas.Add rightOff, oDataArr
                    End If
                End If

            Next


        Next

        'now sort the keys
        Dim varLeftKeyArray As Variant
        varLeftKeyArray = cLeftDatas.Keys
        quicksort varLeftKeyArray
        Dim varRightKeyArray As Variant
        varRightKeyArray = cRightDatas.Keys
        quicksort (varRightKeyArray)

        Dim cNewLeftDatas
        Dim cNewRightDatas
        Set cNewLeftDatas = New Dictionary
        Set cNewRightDatas = New Dictionary

        Dim dPreKey As Variant
        Dim dKey As Variant
        Dim arrLeftData As Variant
        Dim arrRightData As Variant
        Dim iIndex As Integer
        dPreKey = 1#
        Dim arrLeftDataStr(0 To 3) As String
        Dim arrRightDataStr(0 To 3) As String
        For iIndex = UBound(varLeftKeyArray) To 0 Step -1
            dKey = varLeftKeyArray(iIndex)
            If Abs(dKey - dPreKey) > 0.01 Then
                dLastLeftOff = leftOff
                dLastLeftElev = leftElev
                arrLeftData = cLeftDatas.item(dKey)
                arrLeftDataStr(nOffsetIndex) = FormatDistSettings(arrLeftData(nOffsetIndex))
                arrLeftDataStr(nElevIndex) = FormatElevSettings(arrLeftData(nElevIndex))
                arrLeftDataStr(nCodeIndex) = arrLeftData(nCodeIndex)
                arrLeftDataStr(nSlopeIndex) = arrLeftData(nSlopeIndex)
                cNewLeftDatas.Add dKey, arrLeftDataStr
                leftOff = arrLeftData(nOffsetIndex)
                leftElev = arrLeftData(nElevIndex)
            End If
            dPreKey = dKey
        Next iIndex
        dPreKey = -1#
        For Each dKey In varRightKeyArray
            If Abs(dKey - dPreKey) > 0.01 Then
                dLastRightOff = rightOff
                dLastRightElev = rightElev
                arrRightData = cRightDatas.item(dKey)
                arrRightDataStr(nOffsetIndex) = FormatDistSettings(arrRightData(nOffsetIndex))
                arrRightDataStr(nElevIndex) = FormatElevSettings(arrRightData(nElevIndex))
                arrRightDataStr(nCodeIndex) = arrRightData(nCodeIndex)
                arrRightDataStr(nSlopeIndex) = arrRightData(nSlopeIndex)
                cNewRightDatas.Add dKey, arrRightDataStr
                rightOff = arrRightData(nOffsetIndex)
                rightElev = arrRightData(nElevIndex)
            End If
            dPreKey = dKey
        Next


        aDataBySta(nLeftEndIndex) = GetEndColumInfo(dLastLeftOff, dLastLeftElev, leftOff, leftElev)
        aDataBySta(nRightEndIndex) = GetEndColumInfo(dLastRightOff, dLastRightElev, rightOff, rightElev)
        Set aDataBySta(nLeftIndex) = cNewLeftDatas
        Set aDataBySta(nRightIndex) = cNewRightDatas

        If False = g_oSlopeStakeData.Exists(curStation) Then
            g_oSlopeStakeData.Add curStation, aDataBySta
        End If
ErrHandler:
        If Err.Number <> 0 Then
            Resume NextLoop
        End If
NextLoop:
    Next
    
    
    ExtractData = True
    
End Function


Private Function GetSlopeData(curOff As Double, lastOff As Double, curElev As Double, lastElev As Double) As String
    Dim deltaX As Double
    Dim deltaY As Double
    Dim slope As Double
    Const percent2Slope = 0.1
    deltaX = Abs(curOff - lastOff)
    deltaY = curElev - lastElev
    
    If (deltaX <> 0) Then
        slope = deltaY / deltaX
        If (Abs(slope) <= percent2Slope) Then
            GetSlopeData = VBA.Format(slope, "percent")
        Else
            GetSlopeData = "1:" & VBA.Format(slope, "0.00")
        End If
    Else
        GetSlopeData = VBA.Format(0, "0.00")
    End If
   

End Function

Private Function GetEndColumInfo(last2Off As Double, last2Elev As Double, lastOff As Double, lastElev As Double) As Variant
    Dim tmpArr(0 To 2)
    Dim deltaY As Double
    
    deltaY = lastElev - last2Elev
    If deltaY < 0 Then
        tmpArr(0) = "F" & VBA.Format(Abs(deltaY), "0.00")
    Else
        tmpArr(0) = "C" & VBA.Format(deltaY, "0.00")
    End If
    tmpArr(1) = "@" & VBA.Format(Abs(lastOff - last2Off), "0.00")
    tmpArr(2) = GetSlopeData(last2Off, lastOff, last2Elev, lastElev)

    GetEndColumInfo = tmpArr
End Function
