' -----------------------------------------------------------------------------
' <copyright file="Daylight_ExtractData.vb" company="Autodesk">
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

Friend Class Daylight_ExtractData

    Public Class SampleLineData
        Public Name As String
        Public Station As String
        Public LengthLeft As String
        Public EastLeft As String
        Public NorthLeft As String
        Public ElevLeft As String
        Public LengthRight As String
        Public EastRight As String
        Public NorthRight As String
        Public ElevRight As String
        Public Last As String
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


    Public Shared Function ExtractData(ByVal oSampleLineGroup As Land.AeccSampleLineGroup, _
                                       ByVal existingSurface As Land.AeccSurface, _
                                       ByVal stationStart As Double, _
                                       ByVal stationEnd As Double) As Boolean
        Dim oSampleLine As Land.AeccSampleLine
        Dim oProfilePR As Land.AeccProfile = Nothing
        Dim oBaseLine As Roadway.AeccBaseline
        Dim X As Double
        Dim Y As Double
        Dim nCur As Integer = 0
        Dim dElev1G As Double
        Dim dElev1D As Double
        Dim dOffset1G As Double
        Dim dOffset1D As Double


        'Recherche du corridor
        'Find corridor related with the sample line group
        Dim profile As Land.AeccProfile = Nothing
        Dim relatedCorridor As Roadway.AeccCorridor = FindCorridorByBaseline(oSampleLineGroup.Parent, profile)

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
                '            dRealStation = oSampleLine.Station

                'info sur le profil en long
                If Not relatedCorridor Is Nothing And Not existingSurface Is Nothing Then

                    'calcul l'entrée en terre
                    For Each oBaseLine In relatedCorridor.Baselines
                        If oSampleLine.Station >= oBaseLine.StartStation And oSampleLine.Station <= oBaseLine.EndStation Then
                            'Extrait l'assemblage à ce PK
                            Dim oCalculatedPoints As AeccCalculatedPoints
                            'Recherche les points d'entrée en terre
                            oCalculatedPoints = oBaseLine.AppliedAssembly(oSampleLine.Station).GetPointsByCode("Daylight")
                            For Each oPoint As AeccCalculatedPoint In oCalculatedPoints
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

                        End If
                    Next
                    oSampleLineGroup.Parent.PointLocation(oSampleLine.Station, dOffset1G, X, Y)
                    oDataArr.EastLeft = FormatDistSettings(X, False)
                    oDataArr.NorthLeft = FormatDistSettings(Y, False)
                    oDataArr.ElevLeft = FormatElevSettings(existingSurface.FindElevationAtXY(X, Y), False)
                    oDataArr.LengthLeft = FormatDistSettings(dOffset1G, False)
                    oSampleLineGroup.Parent.PointLocation(oSampleLine.Station, dOffset1D, X, Y)
                    oDataArr.EastRight = FormatDistSettings(X, False)
                    oDataArr.NorthRight = FormatDistSettings(Y, False)
                    oDataArr.ElevRight = FormatElevSettings(existingSurface.FindElevationAtXY(X, Y), False)
                    oDataArr.LengthRight = FormatDistSettings(dOffset1D, False)
                Else
                    oDataArr.EastLeft = "---"
                    oDataArr.NorthLeft = "---"
                    oDataArr.ElevLeft = "---"
                    oDataArr.LengthLeft = "---"
                    oDataArr.EastRight = "---"
                    oDataArr.NorthRight = "---"
                    oDataArr.ElevRight = "---"
                    oDataArr.LengthRight = "---"
                End If

                m_oSampleLineDataArr(nCur) = oDataArr
                nCur = nCur + 1
            End If
        Next

        ExtractData = True
    End Function

    Public Shared Function FindCorridorByBaseline(ByVal alignment As Land.IAeccAlignment, _
                                                  ByRef profile As Land.AeccProfile) As Roadway.AeccCorridor
        'iterate the roadway document to find the first corridor which contains the alignment as its baseline.
        For Each corridor As Roadway.AeccCorridor In ReportApplication.AeccXRoadwayDocument.Corridors
            For Each baseline As Roadway.AeccBaseline In corridor.Baselines
                If baseline.Alignment.ObjectID = alignment.ObjectID Then
                    profile = baseline.Profile
                    Return corridor
                End If
            Next
        Next
        Return Nothing
    End Function

    Public Shared Function FindSurfaceProfile(ByVal alignment As Land.IAeccAlignment) As Land.AeccProfile
        For Each profile As Land.AeccProfile In alignment.Profiles
            If profile.Type = Land.AeccProfileType.aeccExistingGround Then
                Return profile
            End If
        Next
        Return Nothing
    End Function

    Public Shared Function FindEGSurface(ByVal alignment As Land.IAeccAlignment) As Land.AeccSurface
        Dim profile As Land.AeccProfile
        profile = FindSurfaceProfile(alignment)

        If Not profile Is Nothing Then
            Return profile.Surface
        Else
            Return Nothing
        End If
    End Function
End Class
