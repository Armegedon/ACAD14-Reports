' -----------------------------------------------------------------------------
' <copyright file="CrossSection_ExtractData.vb" company="Autodesk">
' Copyright (C) Autodesk, Inc. All rights reserved.
'
' Permission to use, copy, modify, and distribute this software in
' object code form for any purpose and without fee is hereby granted,
' provided that the above copyright notice appears in all copies and
' that both that copyright notice and the limited warranty and
' restricted rights notice below appear in all supporting
' documentation.
'
' AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
' AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
' MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
' DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
' UNINTERRUPTED OR ERROR FREE.
'
' Use, duplication, or disclosure by the U.S. Government is subject to
' restrictions set forth in FAR 52.227-19 (Commercial Computer
' Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
' (Rights in Technical Data and Computer Software), as applicable.
'
' </copyright>
' -----------------------------------------------------------------------------

Option Explicit On
Option Strict On

Imports System
Imports AutoCAD = Autodesk.AutoCAD.Interop
Imports Autodesk.AECC.Interop.Roadway
Imports Autodesk.AECC.Interop.Land
Imports Autodesk.AECC.Interop

Friend Class CrossSection_ExtractData

    Public Class SampleLineData
        Public Name As String
        Public Station As String
        Public ElevationEG As String
        Public ElevationFG As String
        Public East As String
        Public North As String
        Public CrossCrownLeft As String
        Public CrossCrownRight As String
    End Class

    Public Shared m_oSampleLineDataArr As SampleLineData()


    Private Shared Function FormatElevSettings(ByVal dDis As Double, Optional ByVal unitIndicator As Boolean = True) As String
        Dim oElevSettings As Land.AeccSettingsElevation
        oElevSettings = ReportApplication.AeccXDatabase.Settings.ProfileSettings.AmbientSettings.ElevationSettings
        FormatElevSettings = ReportFormat.FormatDistance(dDis, oElevSettings.Unit.Value, oElevSettings.Precision.Value, _
                    oElevSettings.Rounding.Value, oElevSettings.Sign.Value, unitIndicator)
    End Function
    Private Shared Function FormatDistSettings(ByVal dDis As Double, Optional ByVal unitIndicator As Boolean = True) As String
        Dim oDistSettings As Land.IAeccSettingsDistance
        oDistSettings = ReportApplication.AeccXDatabase.Settings.ProfileSettings.AmbientSettings.DistanceSettings

        FormatDistSettings = ReportFormat.FormatDistance(dDis, oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                    oDistSettings.Rounding.Value, oDistSettings.Sign.Value, unitIndicator)
    End Function

    Public Shared Function ExtractData(ByVal oSampleLineGroup As Land.AeccSampleLineGroup, ByVal stationStart As Double, ByVal stationEnd As Double) As Boolean
        Dim oSampleLine As Land.AeccSampleLine
        Dim oCorridor As Roadway.AeccCorridor = Nothing
        Dim oProfilePR As Land.AeccProfile = Nothing
        Dim oProfileTN As Land.AeccProfile = Nothing
        Dim oBaseLine As Roadway.AeccBaseline
        Dim X As Double
        Dim Y As Double
        Dim dRealStation As Double
        Dim nCur As Integer = 0
        Dim dElev1G As Double
        Dim dElev1D As Double
        Dim dOffset1G As Double
        Dim dOffset1D As Double

        oCorridor = Daylight_ExtractData.FindCorridorByBaseline(oSampleLineGroup.Parent, oProfilePR)
        Dim bCorridorFound As Boolean
        If Not oCorridor Is Nothing Then
            bCorridorFound = True
        Else
            bCorridorFound = False
        End If

        oProfileTN = Daylight_ExtractData.FindSurfaceProfile(oSampleLineGroup.Parent)

        'Arrondi des abscisses
        Dim iAbscissePrec As Integer = ReportApplication.AeccXDatabase.Settings.ProfileSettings.AmbientSettings.StationSettings.Precision.Value

        'redimensionne le tableau des tabulations
        ReDim m_oSampleLineDataArr(oSampleLineGroup.SampleLines.Count - 1)

        'Boucle sur chaque tabulation
        For Each oSampleLine In oSampleLineGroup.SampleLines

            On Error Resume Next
            Dim oDataArr As New SampleLineData
            If (oSampleLine.Station >= stationStart And _
                Math.Round(oSampleLine.Station, iAbscissePrec) <= stationEnd) Then

                'infos sur la tabulation
                oDataArr.Name = oSampleLine.Name
                oDataArr.Station = ReportUtilities.GetStationStringWithDerived(oSampleLineGroup.Parent, oSampleLine.Station)
                oSampleLineGroup.Parent.PointLocation(oSampleLine.Station, 0, X, Y)
                oDataArr.East = FormatDistSettings(X, False)
                oDataArr.North = FormatDistSettings(Y, False)
                oDataArr.ElevationEG = FormatElevSettings(oProfileTN.ElevationAt(oSampleLine.Station), False)

                dRealStation = oSampleLine.Station
                '            dRealStation = GetRawStation(oSampleLineGroup.Parent.GetStationStringWithEquations(dRealStation))

                'info sur le profil en long
                If bCorridorFound Then

                    'calcul du Z Projet
                    oDataArr.ElevationFG = FormatElevSettings(oProfilePR.ElevationAt(oSampleLine.Station), False)
                    'calcul du devers
                    Dim dCrownLeft, dCrownRight As Double
                    dCrownLeft = 0.0
                    dCrownRight = 0.0
                    '                Set oCalculatedLinks = Nothing
                    For Each oBaseLine In oCorridor.Baselines
                        '                    oSortedStationArr = oBaseLine.GetSortedStations
                        If oSampleLine.Station >= oBaseLine.StartStation And oSampleLine.Station <= oBaseLine.EndStation Then
                            '                        'recherche de la station la plus proche
                            '                        For i = 0 To UBound(oSortedStationArr)
                            '                            If ((oSampleLine.Station < oSortedStationArr(i)) _
                            '                                Or (Abs(oSampleLine.Station - oSortedStationArr(i)) < 0.01)) Then
                            '                                'dRealStation = oSortedStationArr(i)
                            '                                Exit For
                            '                            End If
                            '                        Next
                            'Extrait l'assemblage à ce PK
                            Dim oCalculatedPoints As Roadway.AeccCalculatedPoints = oBaseLine.AppliedAssembly(oSampleLine.Station).GetPointsByCode("ETW")
                            For Each oPoint As Roadway.AeccCalculatedPoint In oCalculatedPoints
                                Dim vx As Double()
                                Dim off As Double

                                vx = CType(oPoint.GetStationOffsetElevationToBaseline, Double())
                                off = vx(1)
                                If off < 0 Then
                                    dElev1G = vx(2)
                                    dOffset1G = off
                                End If

                                If off > 0 Then
                                    dElev1D = vx(2)
                                    dOffset1D = off
                                End If

                            Next
                            '                        dCrownLeft = GetSlopeData(0, dOffset1G, 0, dElev1G)
                            '                        dCrownRight = GetSlopeData(0, dOffset1D, 0, dElev1D)

                        End If
                    Next
                    oDataArr.CrossCrownLeft = GetSlopeData(0, dOffset1G, 0, dElev1G)
                    oDataArr.CrossCrownRight = GetSlopeData(0, dOffset1D, 0, dElev1D)
                Else
                    'calcul du Z Projet
                    oDataArr.ElevationFG = FormatElevSettings(0.0#, False)
                    'calcul du devers
                    oDataArr.CrossCrownLeft = "---"
                    oDataArr.CrossCrownRight = "---"
                End If

                m_oSampleLineDataArr(nCur) = oDataArr
                nCur = nCur + 1
            End If
        Next

        ExtractData = True
    End Function

    Private Shared Function GetSlopeData(ByVal curOff As Double, ByVal lastOff As Double, ByVal curElev As Double, ByVal lastElev As Double) As String
        Dim deltaX As Double
        Dim deltaY As Double
        Dim slope As Double
        Const percent2Slope As Double = 0.1
        deltaX = Math.Abs(curOff - lastOff)
        deltaY = curElev - lastElev

        If (deltaX <> 0) Then
            slope = deltaY / deltaX
            If (Math.Abs(slope) <= percent2Slope) Then
                GetSlopeData = Format(slope, "0.0%")
            Else
                GetSlopeData = "1:" & Format(slope, "0.00")
            End If
        Else
            GetSlopeData = Format(0, "0.00")
        End If


    End Function
End Class
