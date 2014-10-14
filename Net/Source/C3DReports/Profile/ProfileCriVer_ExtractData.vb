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
Imports AutoCAD = Autodesk.AutoCAD.Interop
Imports AeccUiLandLib = Autodesk.AECC.Interop.UiLand
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AEC.Interop.UIBase
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Friend Class ProfileCriVer_ExtractData
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
    Public Const nPVIStationIndex = 15
    Public Const nDesignSpeedIndex = 16
    Public Const nMinKForStoppingSightDisIndex = 17
    Public Const nMinKForPassingSightDisIndex = 18
    Public Const nMinKForHeadlightSightDisIndex = 19
    Public Const nMinOverallCurLenIndex = 20
    Public Const nMaxDrainageCurLenIndex = 21
    Public Const nMinRiderComfortCurLenIndex = 22
    Public Const nMinKForStoppingSightDisResultIndex = 23
    Public Const nMinKForPassingSightDisResultIndex = 24
    Public Const nMinKForHeadlightSightDisResultIndex = 25
    Public Const nMinOverallCurLenResultIndex = 26
    Public Const nMaxDrainageCurLenResultIndex = 27
    Public Const nMinRiderComfortCurLenResultIndex = 28
    Public Const nDesignCheckNamesIndex = 29 ' Array of names
    Public Const nDesignCheckResultsIndex = 30 'Array of Results
    Public Const nLastCurInfoIndex = nDesignCheckResultsIndex


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

    Public Shared Function ExtractData(ByVal oProfile As AeccLandLib.AeccProfile, _
                                ByVal stationStart As Double, _
                                ByVal stationEnd As Double) As Boolean
        Dim oDataArr(nLastIndex) As Object
        Dim oPVI As AeccLandLib.AeccProfilePVI
        Dim oPVIC As AeccLandLib.AeccProfilePVICurve
        Dim nCur As Integer
        Dim bUseCriteriaBasedFile As Boolean

        ReDim m_oProfileDataArr(oProfile.PVIs.Count - 1)
        nCur = 0
        For Each oPVI In oProfile.PVIs

            Dim roundedStation As Double
            Dim stringStation As String
            stringStation = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, oPVI.Station)
            'Fix defect 1462489, the input will be raw station 
            'roundedStation = ReportUtilities.GetRawStation(stringStation, oProfile.Alignment.StationIndexIncrement)
            roundedStation = oPVI.Station
            If roundedStation >= stationStart And roundedStation <= stationEnd Then
                With oPVI
                    oDataArr(nStationIndex) = stringStation
                    Try
                        oDataArr(nGradeOutIndex) = ReportFormat_DotNet.FormatGrade(.GradeOut)
                    Catch ex As Exception

                    End Try
                    If oDataArr(nGradeOutIndex) Is Nothing Then
                        oDataArr(nGradeOutIndex) = "&nbsp;"
                    End If
                End With
                ' get curve info
                oPVIC = Nothing
                Try
                    oPVIC = oPVI
                Catch ex As Exception
                End Try
                If Not oPVIC Is Nothing Then
                    Dim oCurveDataArr(nLastCurInfoIndex)
                    oCurveDataArr(nCurvTypeIndex) = oPVIC.VerticalCurveType
                    oCurveDataArr(nPVCStationIndex) = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, oPVIC.BeginStation)
                    oCurveDataArr(nPVTStationIndex) = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, oPVIC.EndStation)

                    bUseCriteriaBasedFile = oProfile.UseDesignCriteriaFile And oProfile.DesignSpeedBased

                    If oPVI.Station = oProfile.StartingStation Then
                        oCurveDataArr(nGradeInIndex) = "&nbsp;"
                    Else
                        oCurveDataArr(nGradeInIndex) = ReportFormat_DotNet.FormatGrade(oPVIC.GradeIn)
                    End If
                    oCurveDataArr(nCurveLenIndex) = FormatDistSettings(oPVIC.CurveLength)
                    If oPVIC.EntityType = AeccLandLib.AeccProfilePVICurveType.aeccParabola Then
                        Dim oPVIPara As AeccLandLib.AeccProfilePVICurveParabolic
                        oPVIPara = oPVIC
                        oCurveDataArr(nKIndex) = oPVIPara.K.ToString("N2")
                        oCurveDataArr(nDesignSpeedIndex) = oPVIPara.HighestDesignSpeed
                        If bUseCriteriaBasedFile Then
                            If oPVIC.VerticalCurveType = AeccLandLib.AeccProfileVerticalCurveType.aeccSag Then
                                oCurveDataArr(nMinKForHeadlightSightDisIndex) = FormatDistSettings(oPVIPara.MinimumKValueHSD)
                                If Not oCurveDataArr(nMinKForHeadlightSightDisIndex) Is Nothing Then
                                    'Dim temp As Double = oPVIPara.MinimumKValueHSD
                                    'oCurveDataArr(nMinKForHeadlightSightDisIndex) = temp.ToString("f2")
                                    If oPVIPara.K >= oPVIPara.MinimumKValueHSD Then
                                        oCurveDataArr(nMinKForHeadlightSightDisResultIndex) = LocalizedRes.ProfileCriVer_Html_Cleared
                                    Else
                                        oCurveDataArr(nMinKForHeadlightSightDisResultIndex) = "<font color=""#FF0000"">" + LocalizedRes.ProfileCriVer_Html_Violated + "</font>"
                                    End If
                                End If
                            Else ' = aeccCrest
                                oCurveDataArr(nMinKForPassingSightDisIndex) = FormatDistSettings(oPVIPara.MinimumKValuePSD)
                                If Not oCurveDataArr(nMinKForPassingSightDisIndex) Is Nothing Then
                                    'Dim temp As Double = oPVIPara.MinimumKValuePSD
                                    'oCurveDataArr(nMinKForPassingSightDisIndex) = oCurveDataArr(nMinKForPassingSightDisIndex).ToString("f2")
                                    If oPVIPara.K >= oPVIPara.MinimumKValuePSD Then
                                        oCurveDataArr(nMinKForPassingSightDisResultIndex) = LocalizedRes.ProfileCriVer_Html_Cleared
                                    Else
                                        oCurveDataArr(nMinKForPassingSightDisResultIndex) = "<font color=""#FF0000"">" + LocalizedRes.ProfileCriVer_Html_Violated + "</font>"
                                    End If
                                End If
                                oCurveDataArr(nMinKForStoppingSightDisIndex) = FormatDistSettings(oPVIPara.MinimumKValueSSD)
                                If Not oCurveDataArr(nMinKForStoppingSightDisIndex) Is Nothing Then
                                    'Dim temp As Double = oPVIPara.MinimumKValueSSD
                                    'oCurveDataArr(nMinKForStoppingSightDisIndex) = temp.ToString("f2")
                                    If oPVIPara.K >= oPVIPara.MinimumKValueSSD Then
                                        oCurveDataArr(nMinKForStoppingSightDisResultIndex) = LocalizedRes.ProfileCriVer_Html_Cleared
                                    Else
                                        oCurveDataArr(nMinKForStoppingSightDisResultIndex) = "<font color=""#FF0000"">" + LocalizedRes.ProfileCriVer_Html_Violated + "</font>"
                                    End If
                                End If
                            End If
                        End If
                    End If
                    oDataArr(nCurveInfo) = oCurveDataArr
                    m_oProfileDataArr(nCur) = oDataArr.Clone()
                    ExtractDesignCheckData(nCur, oPVIC)
                End If

                nCur = nCur + 1
            End If
        Next

        ExtractData = True
    End Function

    Public Shared Function ExtractDesignCheckData(ByVal nCur As Integer, ByVal oPVIC As AeccLandLib.AeccProfilePVICurve) As Boolean
        Dim oCurveDataArr As Object
        oCurveDataArr = m_oProfileDataArr(nCur)(nCurveInfo)
        Dim oAllDesignCheckArr As AeccLandLib.AeccProfileDesignChecks
        Dim oDesignCheckNameArr()
        Dim oDesignCheckResultArr()
        Dim oEachDesignCheck As AeccLandLib.AeccProfileDesignCheck
        oAllDesignCheckArr = oPVIC.DesignChecks
        Dim checkResult As Boolean
        Dim i As Integer
        i = 0
        If oAllDesignCheckArr.Count > 0 Then
            ReDim oDesignCheckNameArr(oAllDesignCheckArr.Count)
            ReDim oDesignCheckResultArr(oAllDesignCheckArr.Count)
            For Each oEachDesignCheck In oAllDesignCheckArr
                Try
                    oPVIC.ValidateDesignCheck(oEachDesignCheck, checkResult)
                    oDesignCheckNameArr(i) = oEachDesignCheck.Name
                    If (checkResult) Then
                        oDesignCheckResultArr(i) = LocalizedRes.ProfileCriVer_Html_Cleared
                    Else
                        oDesignCheckResultArr(i) = "<font color=""#FF0000"">" + LocalizedRes.ProfileCriVer_Html_Violated + "</font>"
                    End If
                Catch
                    ' Don't add to if eNotApplicable
                End Try
                i = i + 1
            Next
            oCurveDataArr(nDesignCheckNamesIndex) = oDesignCheckNameArr
            oCurveDataArr(nDesignCheckResultsIndex) = oDesignCheckResultArr
        End If
        Try
            Dim temp As Integer
            temp = oPVIC.HighestDesignSpeed
            oCurveDataArr(nDesignSpeedIndex) = temp
        Catch ex As Exception
        End Try
        m_oProfileDataArr(nCur)(nCurveInfo) = oCurveDataArr.Clone()
        ExtractDesignCheckData = True
    End Function
End Class
