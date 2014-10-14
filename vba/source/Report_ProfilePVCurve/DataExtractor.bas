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


Public Const nCurvTypeIndex = 0
Public Const nPVCStationIndex = 1
Public Const nPVCElevationIndex = 2
Public Const nPVIStationIndex = 3
Public Const nPVIElevationIndex = 4
Public Const nPVTStationIndex = 5
Public Const nPVTElevationIndex = 6
Public Const nHighStationIndex = 7
Public Const nHightElevation = 8
Public Const nLowStationIndex = 9
Public Const nLowElevation = 10
Public Const nGradeInIndex = 11
Public Const nGradeOutIndex = 12
Public Const nGradeChange = 13
Public Const nCurveLenIndex = 14
Public Const nCurveRadiusIndex = 15
Public Const nKIndex = 16
Public Const nHeadlightDisIndex = 17
Public Const nPassingDisIndex = 17
Public Const nStoppingDisIndex = 18
Public Const nLastCurInfoIndex = nStoppingDisIndex


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
    Dim oDataArr(nLastCurInfoIndex) As Variant
    Dim oPVI As AeccProfilePVI
    Dim oPVIC As AeccProfilePVICurve
    Dim nCur As Integer
    nCur = 0
    ReDim g_oProfileDataArr(oProfile.PVIs.Count - 1)
        For Each oPVI In oProfile.PVIs
                ' get curve info
            On Error Resume Next
            Erase oDataArr
            Set oPVIC = Nothing
            Set oPVIC = oPVI
            If Not oPVIC Is Nothing Then
                If oPVIC.BeginStation >= stationStart And oPVIC.EndStation <= stationEnd Then
                    oDataArr(nCurvTypeIndex) = oPVIC.VerticalCurveType
                    oDataArr(nPVCStationIndex) = oProfile.Alignment.GetStationStringWithEquations(oPVIC.BeginStation)
                    oDataArr(nPVCElevationIndex) = FormatElevSettings(oPVIC.BeginElevation)
                    oDataArr(nPVIStationIndex) = oProfile.Alignment.GetStationStringWithEquations(oPVIC.Station)
                    oDataArr(nPVIElevationIndex) = FormatElevSettings(oPVIC.Elevation)
                    oDataArr(nPVTStationIndex) = oProfile.Alignment.GetStationStringWithEquations(oPVIC.EndStation)
                    oDataArr(nPVTElevationIndex) = FormatElevSettings(oPVIC.EndElevation)
                    oDataArr(nHighStationIndex) = oProfile.Alignment.GetStationStringWithEquations(oPVIC.HighPointStation)
                    oDataArr(nHightElevation) = FormatElevSettings(oPVIC.HighPointElevation)
                    oDataArr(nLowStationIndex) = oProfile.Alignment.GetStationStringWithEquations(oPVIC.LowPointStation)
                    oDataArr(nLowElevation) = FormatElevSettings(oPVIC.LowPointElevation)
                    If oPVI.Station = oProfile.StartingStation Then
                        oDataArr(nGradeInIndex) = "&nbsp;" 'first One has no GradeIn Data
                    Else
                        oDataArr(nGradeInIndex) = FormatGradSettings(oPVIC.GradeIn)
                    End If
                    If oPVI.Station = oProfile.EndingStation Then
                        oDataArr(nGradeOutIndex) = "&nbsp;"  'last One has no GradeOut Data
                    Else
                        oDataArr(nGradeOutIndex) = FormatGradSettings(oPVIC.GradeOut)
                    End If
                    oDataArr(nGradeChange) = FormatGradSettings(oPVIC.GradeChange)
                    oDataArr(nCurveLenIndex) = FormatDistSettings(oPVIC.CurveLength)
                    
                    If oPVIC.EntityType = aeccParabola Then
                        Dim oPVIPara As AeccProfilePVICurveParabolic
                        Set oPVIPara = oPVIC
                        'how to deal with k
                        oDataArr(nKIndex) = FormatDistSettings(oPVIPara.k)
                        oDataArr(nCurveRadiusIndex) = FormatDistSettings(oPVIPara.Radius)
                        If oPVIC.VerticalCurveType = aeccSag Then
                            oDataArr(nHeadlightDisIndex) = FormatDistSettings(oPVIPara.HeadlightSightDistance)
                        Else ' = aeccCrest
                            oDataArr(nPassingDisIndex) = FormatDistSettings(oPVIPara.PassingSightDistance)
                            oDataArr(nStoppingDisIndex) = FormatDistSettings(oPVIPara.StoppingSightDistance)
                        End If
                    ElseIf oPVIC.EntityType = aeccProfileArc Then
                        Dim oPVIA As AeccProfilePVICurveCircular
                        Set oPVIA = oPVIC
                         oDataArr(nCurveRadiusIndex) = oPVIA.Radius
                    End If
                    
                    g_oProfileDataArr(nCur) = oDataArr
                    nCur = nCur + 1
                End If
            End If

        Next
    ExtractData = True
End Function



