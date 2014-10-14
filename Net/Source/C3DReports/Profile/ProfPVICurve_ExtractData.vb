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

Friend Class ProfPVICurve_ExtractData
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
        ' Existing ground profile
        If oProfile.Type = AeccLandLib.AeccProfileType.aeccExistingGround Then
            For Each oPVI In oProfile.PVIs

                Dim stationString As String
                Dim stationRounded As Double
                stationString = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, oPVI.Station)
                'Fix defect 1462489, the input will be raw station 
                'stationRounded = ReportUtilities.GetRawStation(stationString, oProfile.Alignment.StationIndexIncrement)
                stationRounded = oPVI.Station
                If stationRounded >= stationStart And stationRounded <= stationEnd Then
                    Dim oDataArr(nLastIndex)
                    oDataArr(nStationIndex) = stationString
                    oDataArr(nElevationIndex) = FormatElevSettings(oPVI.Elevation)
                    If oPVI.Station = oProfile.EndingStation Then
                        oDataArr(nGradeOutIndex) = "&nbsp;"  'last One has no GradeOut Data
                    Else
                        oDataArr(nGradeOutIndex) = ReportFormat_DotNet.FormatGrade(oPVI.GradeOut)
                    End If
                    m_oProfileDataArr(nCur) = oDataArr.Clone()
                    nCur = nCur + 1
                End If
            Next
            ' finished ground profile
        ElseIf oProfile.Type = AeccLandLib.AeccProfileType.aeccFinishedGround Then
            For Each oPVI In oProfile.PVIs
                On Error Resume Next

                Dim stationString As String
                Dim stationRounded As Double
                stationString = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, oPVI.Station)
                ''Fix defect 1462489, the input will be raw station 
                'stationRounded = ReportUtilities.GetRawStation(stationString, oProfile.Alignment.StationIndexIncrement)
                stationRounded = oPVI.Station

                If stationRounded >= stationStart And stationRounded <= stationEnd Then
                    Dim oDataArr(nLastIndex)
                    With oPVI
                        oDataArr(nStationIndex) = stationString
                        oDataArr(nElevationIndex) = FormatElevSettings(.Elevation)
                        oDataArr(nGradeOutIndex) = ReportFormat_DotNet.FormatGrade(.GradeOut)
                        If oDataArr(nGradeOutIndex) Is Nothing Then
                            oDataArr(nGradeOutIndex) = "&nbsp;"
                        End If
                    End With
                    ' get curve info

                    oPVIC = Nothing
                    oPVIC = oPVI
                    If Not oPVIC Is Nothing Then
                        Dim oCurveDataArr(nLastCurInfoIndex)
                        'Erase oCurveDataArr
                        oCurveDataArr(nCurvTypeIndex) = oPVIC.VerticalCurveType
                        oCurveDataArr(nPVCStationIndex) = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, oPVIC.BeginStation)
                        oCurveDataArr(nPVCElevationIndex) = FormatElevSettings(oPVIC.BeginElevation)
                        oCurveDataArr(nPVTStationIndex) = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, oPVIC.EndStation)
                        oCurveDataArr(nPVTElevationIndex) = FormatElevSettings(oPVIC.EndElevation)
                        oCurveDataArr(nHighStationIndex) = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, oPVIC.HighPointStation)
                        oCurveDataArr(nHightElevation) = FormatElevSettings(oPVIC.HighPointElevation)
                        oCurveDataArr(nLowStationIndex) = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, oPVIC.LowPointStation)
                        oCurveDataArr(nLowElevation) = FormatElevSettings(oPVIC.LowPointElevation)
                        If oPVI.Station = oProfile.StartingStation Then
                            oCurveDataArr(nGradeInIndex) = "&nbsp;"
                        Else
                            oCurveDataArr(nGradeInIndex) = ReportFormat_DotNet.FormatGrade(oPVIC.GradeIn)
                        End If
                        oCurveDataArr(nGradeChange) = ReportFormat_DotNet.FormatGrade(oPVIC.GradeChange)
                        oCurveDataArr(nCurveLenIndex) = FormatDistSettings(oPVIC.CurveLength)
                        '
                        If oPVIC.EntityType = AeccLandLib.AeccProfilePVICurveType.aeccParabola Then
                            Dim oPVIPara As AeccLandLib.AeccProfilePVICurveParabolic
                            oPVIPara = oPVIC
                            oCurveDataArr(nKIndex) = oPVIPara.K
                            If oPVIC.VerticalCurveType = AeccLandLib.AeccProfileVerticalCurveType.aeccSag Then
                                oCurveDataArr(nHeadlightDisIndex) = FormatDistSettings(oPVIPara.HeadlightSightDistance)
                            Else ' = aeccCrest
                                oCurveDataArr(nPassingDisIndex) = FormatDistSettings(oPVIPara.PassingSightDistance)
                                oCurveDataArr(nStoppingDisIndex) = FormatDistSettings(oPVIPara.StoppingSightDistance)
                            End If
                        End If
                        oDataArr(nCurveInfo) = oCurveDataArr

                    End If

                    m_oProfileDataArr(nCur) = oDataArr.Clone()
                    nCur = nCur + 1
                End If
            Next
        End If

        ExtractData = True
    End Function
End Class
