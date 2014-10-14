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
'
'
'
Option Explicit On

Imports System
Imports AeccUiLandLib = Autodesk.AECC.Interop.UiLand
Imports AeccLandLib = Autodesk.AECC.Interop.Land

Friend Class ProfPVCurve_ExtractData

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


    Private Shared m_oProfileDataArr()
    Public Shared ReadOnly Property ProfileDataArr()
        Get
            Return m_oProfileDataArr
        End Get
    End Property

    Private Shared Function FormatElevSettings(ByVal dDis As Double) As String
        Dim oElevSettings As AeccLandLib.AeccSettingsElevation
        oElevSettings = ReportApplication.AeccXDatabase.Settings.ProfileSettings.AmbientSettings.ElevationSettings

        FormatElevSettings = ReportFormat.FormatDistance(dDis, oElevSettings.Unit.Value, oElevSettings.Precision.Value, _
                    oElevSettings.Rounding.Value, oElevSettings.Sign.Value)
    End Function
    Private Shared Function FormatDistSettings(ByVal dDis As Double) As String
        Dim oDistSettings As AeccLandLib.AeccSettingsDistance
        oDistSettings = ReportApplication.AeccXDatabase.Settings.ProfileSettings.AmbientSettings.DistanceSettings

        FormatDistSettings = ReportFormat.FormatDistance(dDis, oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                    oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
    End Function

    Public Shared Function ExtractData(ByVal oProfile As AeccLandLib.AeccProfile, ByVal stationStart As Double, ByVal stationEnd As Double) As Boolean

        Dim oPVI As AeccLandLib.AeccProfilePVI
        Dim oPVIC As AeccLandLib.AeccProfilePVICurve
        Dim nCur As Integer
        nCur = 0
        ReDim m_oProfileDataArr(oProfile.PVIs.Count - 1)
        For Each oPVI In oProfile.PVIs
            ' get curve info
            On Error Resume Next

            oPVIC = Nothing
            oPVIC = oPVI
            If Not oPVIC Is Nothing Then

                Dim startStationString, endStationString As String
                Dim startStationRounded, endStationRounded As Double
                startStationString = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, oPVIC.BeginStation)
                endStationString = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, oPVIC.EndStation)
                'Fix defect 1462489, the input will be raw station 
                'startStationRounded = ReportUtilities.GetRawStation(startStationString, oProfile.Alignment.StationIndexIncrement)
                'endStationRounded = ReportUtilities.GetRawStation(endStationString, oProfile.Alignment.StationIndexIncrement)
                startStationRounded = oPVIC.BeginStation
                endStationRounded = oPVIC.EndStation
                If startStationRounded >= stationStart And endStationRounded <= stationEnd Then
                    Dim oDataArr(nLastCurInfoIndex) As Object
                    oDataArr(nCurvTypeIndex) = oPVIC.VerticalCurveType
                    oDataArr(nPVCStationIndex) = startStationString
                    oDataArr(nPVCElevationIndex) = FormatElevSettings(oPVIC.BeginElevation)
                    oDataArr(nPVIStationIndex) = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, oPVIC.Station)
                    oDataArr(nPVIElevationIndex) = FormatElevSettings(oPVIC.Elevation)
                    oDataArr(nPVTStationIndex) = endStationString
                    oDataArr(nPVTElevationIndex) = FormatElevSettings(oPVIC.EndElevation)
                    oDataArr(nHighStationIndex) = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, oPVIC.HighPointStation)
                    oDataArr(nHightElevation) = FormatElevSettings(oPVIC.HighPointElevation)
                    oDataArr(nLowStationIndex) = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, oPVIC.LowPointStation)
                    oDataArr(nLowElevation) = FormatElevSettings(oPVIC.LowPointElevation)
                    If oPVI.Station = oProfile.StartingStation Then
                        oDataArr(nGradeInIndex) = "&nbsp;" 'first One has no GradeIn Data
                    Else
                        oDataArr(nGradeInIndex) = ReportFormat_DotNet.FormatGrade(oPVIC.GradeIn)
                    End If
                    If oPVI.Station = oProfile.EndingStation Then
                        oDataArr(nGradeOutIndex) = "&nbsp;"  'last One has no GradeOut Data
                    Else
                        oDataArr(nGradeOutIndex) = ReportFormat_DotNet.FormatGrade(oPVIC.GradeOut)
                    End If
                    oDataArr(nGradeChange) = ReportFormat_DotNet.FormatGrade(oPVIC.GradeChange)
                    oDataArr(nCurveLenIndex) = FormatDistSettings(oPVIC.CurveLength)

                    If oPVIC.EntityType = AeccLandLib.AeccProfilePVICurveType.aeccParabola Then
                        Dim oPVIPara As AeccLandLib.AeccProfilePVICurveParabolic
                        oPVIPara = oPVIC
                        'how to deal with k
                        oDataArr(nKIndex) = FormatDistSettings(oPVIPara.K)
                        oDataArr(nCurveRadiusIndex) = FormatDistSettings(oPVIPara.Radius)
                        If oPVIC.VerticalCurveType = AeccLandLib.AeccProfileVerticalCurveType.aeccSag Then
                            oDataArr(nHeadlightDisIndex) = FormatDistSettings(oPVIPara.HeadlightSightDistance)
                        Else ' = aeccCrest
                            oDataArr(nPassingDisIndex) = FormatDistSettings(oPVIPara.PassingSightDistance)
                            oDataArr(nStoppingDisIndex) = FormatDistSettings(oPVIPara.StoppingSightDistance)
                        End If
                    ElseIf oPVIC.EntityType = AeccLandLib.AeccProfilePVICurveType.aeccProfileArc Then
                        Dim oPVIA As AeccLandLib.AeccProfilePVICurveCircular
                        oPVIA = oPVIC
                        oDataArr(nCurveRadiusIndex) = oPVIA.Radius
                    End If

                    m_oProfileDataArr(nCur) = oDataArr
                    nCur = nCur + 1
                End If
            End If

        Next
        ExtractData = True
    End Function

End Class
