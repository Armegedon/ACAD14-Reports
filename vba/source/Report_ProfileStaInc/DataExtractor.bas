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
Public Const nElevationIndex = 1
Public Const nGradeIndex = 2
Public Const nLocIndex = 3
Public Const nLastIndex = nLocIndex



Public g_oProfileDataArr()

Function FormatElevSettings(ByVal dDis As Double) As String
    Dim oElevSettings As AeccSettingsElevation
    Set oElevSettings = g_oAeccDb.settings.ProfileSettings.AmbientSettings.ElevationSettings
    
    FormatElevSettings = FormatDistance(dDis, oElevSettings.Unit.Value, oElevSettings.Precision.Value, _
                oElevSettings.Rounding.Value, oElevSettings.Sign.Value)
End Function
Function FormatDistSettings(ByVal dDis As Double) As String
    Dim oDistSettings As AeccSettingsDistance
    Set oDistSettings = g_oAeccDb.settings.ProfileSettings.AmbientSettings.DistanceSettings
    
    FormatDistSettings = FormatDistance(dDis, oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
End Function
Function FormatGradSettings(ByVal dGrad As Double) As String
    Dim oGradSettings As AeccSettingsGradeSlope
    Set oGradSettings = g_oAeccDb.settings.ProfileSettings.AmbientSettings.GradeSlopeSettings
    
    FormatGradSettings = FormatGrade(dGrad, oGradSettings.Format.Value, oGradSettings.Precision.Value, _
                oGradSettings.Rounding.Value, oGradSettings.Sign.Value)
End Function

Public Function ExtractData(oProfile As AeccProfile, stationStart As Double, stationEnd As Double) As Boolean
    Dim oDataArr(nLastIndex) As Variant
    Dim oPVI As AeccProfilePVI
    Dim oPVIC As AeccProfilePVICurve
    Dim nCur As Integer
    nCur = 0


    Dim arrSize, stationCount As Long
    Dim specCount As Long
    Dim curSta As Double
    Dim rptArr(3) As Variant
    Dim rptStrArr() As Variant

    stationCount = 0
    arrSize = Int((oProfile.EndingStation - oProfile.StartingStation) / g_staInc)
    arrSize = arrSize + (oProfile.PVIs.Count * 3) + 1
    ReDim rptStrArr(arrSize)
    rptArr(2) = 0#
    
    For Each oPVI In oProfile.PVIs
        
        If Not TypeOf oPVI Is AeccProfilePVICurve Then
            If oPVI.Station >= stationStart And oPVI.Station <= stationEnd Then
                rptArr(0) = oPVI.Station
                rptArr(1) = oPVI.Elevation
    
                rptArr(3) = "PVI"
    
                rptStrArr(stationCount) = rptArr
                stationCount = stationCount + 1
            End If
        Else
            Set oPVIC = oPVI
            With oPVIC
                'PVC
                If .BeginStation >= stationStart And .BeginStation <= stationEnd Then
                
                    rptArr(0) = .BeginStation
                    rptArr(1) = .BeginElevation
                    rptArr(3) = "PVC"
        
                    rptStrArr(stationCount) = rptArr
                    stationCount = stationCount + 1
                End If
                ' PVI
                If .Station >= stationStart And .Station <= stationEnd Then
                    rptArr(0) = .Station
                    rptArr(1) = .Elevation
        
                    If .VerticalCurveType = AeccProfileVerticalCurveType.aeccCrest Then
                        rptArr(3) = "Crest"
                    Else
                        rptArr(3) = "Sag"
                    End If
    
                    rptStrArr(stationCount) = rptArr
                    stationCount = stationCount + 1
                End If
                'PVT
                If .EndStation >= stationStart And .EndStation <= stationEnd Then
                    rptArr(0) = .EndStation
                    rptArr(1) = .EndElevation
                    rptArr(3) = "PVT"
        
                    rptStrArr(stationCount) = rptArr
                    stationCount = stationCount + 1
                End If
            End With
        End If
    Next
    
    specCount = stationCount - 1
     ' there will be a PVI at the start and end station, so
    ' don't add duplicate stations there
    Dim noAdd As Boolean
    Dim xx As Long
    For curSta = stationStart To stationEnd Step g_staInc
        noAdd = False

        For xx = 0 To specCount
            If rptStrArr(xx)(0) = curSta Then
                noAdd = True
                xx = specCount
            End If
        Next

       If noAdd = False Then
           
            rptArr(0) = curSta
            rptArr(1) = oProfile.ElevationAt(curSta)
            rptArr(2) = 0#

            rptArr(3) = vbNullString

            rptStrArr(stationCount) = rptArr
            stationCount = stationCount + 1
        End If
    Next
    
    'resort and clean up PVI, inc data

    ReDim oProfileDataArr(stationCount - 1)
    Dim trArr(3)

    Dim i As Long
    Dim ii As Long
    Dim ins As Long
    Dim movedCount As Long
    Dim movedIncCount As Long
    ins = 0
    movedCount = 0
    movedIncCount = 0

    For ii = movedCount To specCount
        For i = (specCount + movedIncCount + 1) To stationCount - 1
            If rptStrArr(i)(0) > rptStrArr(ii)(0) Then
                oProfileDataArr(ins) = trArr
                oProfileDataArr(ins)(0) = rptStrArr(ii)(0)
                oProfileDataArr(ins)(1) = rptStrArr(ii)(1)
                oProfileDataArr(ins)(2) = rptStrArr(ii)(2)
                oProfileDataArr(ins)(3) = rptStrArr(ii)(3)

                ins = ins + 1
                movedCount = movedCount + 1
                i = stationCount
            Else
                oProfileDataArr(ins) = trArr
                oProfileDataArr(ins)(0) = rptStrArr(i)(0)
                oProfileDataArr(ins)(1) = rptStrArr(i)(1)
                oProfileDataArr(ins)(2) = rptStrArr(i)(2)
                oProfileDataArr(ins)(3) = rptStrArr(i)(3)
                movedIncCount = movedIncCount + 1
                ins = ins + 1
            End If
'            If ins = 249 Then
'                MsgBox ""
'            End If
        Next
    Next
    For ii = movedCount To specCount
        oProfileDataArr(ins) = trArr
        oProfileDataArr(ins)(0) = rptStrArr(ii)(0)
        oProfileDataArr(ins)(1) = rptStrArr(ii)(1)
        oProfileDataArr(ins)(2) = rptStrArr(ii)(2)
        oProfileDataArr(ins)(3) = rptStrArr(ii)(3)
        ins = ins + 1
    Next
    For i = specCount + movedIncCount + 1 To stationCount - 1
        oProfileDataArr(ins) = trArr
        oProfileDataArr(ins)(0) = rptStrArr(i)(0)
        oProfileDataArr(ins)(1) = rptStrArr(i)(1)
        oProfileDataArr(ins)(2) = rptStrArr(i)(2)
        oProfileDataArr(ins)(3) = rptStrArr(i)(3)
        ins = ins + 1
    Next
    
    
      ' no grade percent reported for first station
    If oProfileDataArr(0)(0) = oProfile.StartingStation Then
        oProfileDataArr(0)(2) = vbNullString
    End If
    If oProfileDataArr(0)(3) = vbNullString Then
        oProfileDataArr(0)(3) = "&nbsp;"
    End If

    For i = 1 To stationCount - 1

        'compute grade percent
        Dim grad As Double
        If oProfileDataArr(i)(0) = oProfileDataArr(i - 1)(0) Then
            oProfileDataArr(i)(2) = vbNullString
        Else
            oProfileDataArr(i)(2) = (oProfileDataArr(i)(1) - oProfileDataArr(i - 1)(1)) / _
                            (oProfileDataArr(i)(0) - oProfileDataArr(i - 1)(0))
        End If

        ' add nbsp to Location column for clean HTML table grid
        If oProfileDataArr(i)(3) = vbNullString Then
            oProfileDataArr(i)(3) = "&nbsp;"
        End If
    Next
    ReDim g_oProfileDataArr(stationCount)
    Dim tmpArr(3)
    
    For i = 0 To stationCount - 1
        tmpArr(nStationIndex) = oProfile.Alignment.GetStationStringWithEquations(oProfileDataArr(i)(nStationIndex))
        tmpArr(nElevationIndex) = FormatElevSettings((oProfileDataArr(i)(nElevationIndex)))
        If oProfileDataArr(i)(2) = vbNullString Then
            tmpArr(nGradeIndex) = "&nbsp;"
        Else
            tmpArr(nGradeIndex) = FormatGradSettings((oProfileDataArr(i)(nGradeIndex)))
        End If
        tmpArr(nLocIndex) = oProfileDataArr(i)(nLocIndex)
        g_oProfileDataArr(i) = tmpArr
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



