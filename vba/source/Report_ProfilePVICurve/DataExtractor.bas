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
Public Const nElevationIndex = 1
Public Const nGradeOutIndex = 2
Public Const nCurveInfo = 3
Public Const nLastIndex = nCurveInfo

Public Const nCurvTypeIndex = 0
Public Const nPVCStationIndex = 1
Public Const nPVCElevationIndex = 2
Public Const nPVTStationIndex = 3
Public Const nPVTElevationIndex = 4
Public Const nHighStationIndex = 5
Public Const nHightElevation = 6
Public Const nLowStationIndex = 7
Public Const nLowElevation = 8
Public Const nGradeInIndex = 9
Public Const nGradeChange = 10
Public Const nCurveLenIndex = 11
Public Const nKIndex = 12
Public Const nHeadlightDisIndex = 13
Public Const nPassingDisIndex = 13
Public Const nStoppingDisIndex = 14
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
    Dim oDataArr(nLastIndex)
    Dim oPVI As AeccProfilePVI
    Dim oPVIC As AeccProfilePVICurve
    Dim nCur As Integer
    nCur = 0
    
    ReDim g_oProfileDataArr(oProfile.PVIs.Count - 1)
    ' Existing ground profile
    If oProfile.Type = aeccExistingGround Then
        For Each oPVI In oProfile.PVIs
            If oPVI.station >= stationStart And oPVI.station <= stationEnd Then
                Erase oDataArr
                oDataArr(nStationIndex) = oProfile.Alignment.GetStationStringWithEquations(oPVI.station)
                oDataArr(nElevationIndex) = FormatElevSettings(oPVI.Elevation)
                If oPVI.station = oProfile.EndingStation Then
                    oDataArr(nGradeOutIndex) = "&nbsp;"  'last One has no GradeOut Data
                Else
                    oDataArr(nGradeOutIndex) = FormatGradSettings(oPVI.GradeOut)
                End If
                g_oProfileDataArr(nCur) = oDataArr
                nCur = nCur + 1
            End If
        Next
    ' finished ground profile
    ElseIf oProfile.Type = aeccFinishedGround Then
    
        For Each oPVI In oProfile.PVIs
            On Error Resume Next
            Erase oDataArr
            If oPVI.station >= stationStart And oPVI.station <= stationEnd Then
                With oPVI
                    oDataArr(nStationIndex) = oProfile.Alignment.GetStationStringWithEquations(.station)
                    oDataArr(nElevationIndex) = FormatElevSettings(.Elevation)
                    oDataArr(nGradeOutIndex) = FormatGradSettings(.GradeOut)
                    If oDataArr(nGradeOutIndex) = Empty Then
                        oDataArr(nGradeOutIndex) = "&nbsp;"
                    End If
                End With
                ' get curve info
                
                Set oPVIC = Nothing
                Set oPVIC = oPVI
                If Not oPVIC Is Nothing Then
                    Dim oCurveDataArr(nLastCurInfoIndex)
                    Erase oCurveDataArr
                    oCurveDataArr(nCurvTypeIndex) = oPVIC.VerticalCurveType
                    oCurveDataArr(nPVCStationIndex) = oProfile.Alignment.GetStationStringWithEquations(oPVIC.BeginStation)
                    oCurveDataArr(nPVCElevationIndex) = FormatElevSettings(oPVIC.BeginElevation)
                    oCurveDataArr(nPVTStationIndex) = oProfile.Alignment.GetStationStringWithEquations(oPVIC.EndStation)
                    oCurveDataArr(nPVTElevationIndex) = FormatElevSettings(oPVIC.EndElevation)
                    oCurveDataArr(nHighStationIndex) = oProfile.Alignment.GetStationStringWithEquations(oPVIC.HighPointStation)
                    oCurveDataArr(nHightElevation) = FormatElevSettings(oPVIC.HighPointElevation)
                    oCurveDataArr(nLowStationIndex) = oProfile.Alignment.GetStationStringWithEquations(oPVIC.LowPointStation)
                    oCurveDataArr(nLowElevation) = FormatElevSettings(oPVIC.LowPointElevation)
                    If oPVI.station = oProfile.StartingStation Then
                        oCurveDataArr(nGradeInIndex) = "&nbsp;"
                    Else
                        oCurveDataArr(nGradeInIndex) = FormatGradSettings(oPVIC.GradeIn)
                    End If
                    oCurveDataArr(nGradeChange) = FormatGradSettings(oPVIC.GradeChange)
                    oCurveDataArr(nCurveLenIndex) = FormatDistSettings(oPVIC.CurveLength)
    '
                    If oPVIC.EntityType = aeccParabola Then
                        Dim oPVIPara As AeccProfilePVICurveParabolic
                        Set oPVIPara = oPVIC
                        oCurveDataArr(nKIndex) = oPVIPara.k
                        If oPVIC.VerticalCurveType = aeccSag Then
                            oCurveDataArr(nHeadlightDisIndex) = FormatDistSettings(oPVIPara.HeadlightSightDistance)
                        Else ' = aeccCrest
                            oCurveDataArr(nPassingDisIndex) = FormatDistSettings(oPVIPara.PassingSightDistance)
                            oCurveDataArr(nStoppingDisIndex) = FormatDistSettings(oPVIPara.StoppingSightDistance)
                        End If
                    End If
                    oDataArr(nCurveInfo) = oCurveDataArr
    
                End If
    
                g_oProfileDataArr(nCur) = oDataArr
                nCur = nCur + 1
            End If
        Next
    End If
    
    ExtractData = True
End Function



