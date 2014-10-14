' -----------------------------------------------------------------------------
' <copyright file="ProfStaInc_ExtractData.vb" company="Autodesk">
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
Imports AeccUiLandLib = Autodesk.AECC.Interop.UiLand
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Friend Class ProfStaInc_ExtractData

    Public Class ProfileData
        Public RawStation As Double
        Public RawElevation As Double
        Public Station As String
        Public Elevation As String
        Public Grade As String
        Public Location As String
    End Class

    Private Const HTML_SPACE As String = "&nbsp;" '"&nbsp;" : space in html

    Private Shared m_oProfileDataDictionary As New Collections.Generic.SortedDictionary(Of Double, ProfileData)

    Private Shared m_profile As AeccLandLib.IAeccProfile
    Private Shared m_stationStart As Double
    Private Shared m_stationEnd As Double

    Public Shared ReadOnly Property DataDictionary() As Collections.Generic.SortedDictionary(Of Double, ProfileData)
        Get
            Return m_oProfileDataDictionary
        End Get
    End Property

    Private Shared Function FormatElevSettings(ByVal dDis As Double) As String
        Dim oElevSettings As AeccLandLib.AeccSettingsElevation
        oElevSettings = ReportApplication.AeccXDatabase.Settings.ProfileSettings.AmbientSettings.ElevationSettings

        FormatElevSettings = ReportFormat.FormatDistance(dDis, oElevSettings.Unit.Value, oElevSettings.Precision.Value, _
                    oElevSettings.Rounding.Value, oElevSettings.Sign.Value)
    End Function
    Private Shared Function FormatDistSettings(ByVal dDis As Double) As String
        Dim oDistSettings As AeccLandLib.IAeccSettingsDistance
        oDistSettings = ReportApplication.AeccXDatabase.Settings.ProfileSettings.AmbientSettings.DistanceSettings

        FormatDistSettings = ReportFormat.FormatDistance(dDis, oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                    oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
    End Function
    Private Shared Function FormatGradSettings(ByVal dGrad As Double) As String
        Dim oGradSettings As AeccLandLib.AeccSettingsGradeSlope
        oGradSettings = ReportApplication.AeccXDatabase.Settings.ProfileSettings.AmbientSettings.GradeSlopeSettings

        FormatGradSettings = ReportFormat.FormatGrade(dGrad, oGradSettings.Format.Value, oGradSettings.Precision.Value, _
                    oGradSettings.Rounding.Value, oGradSettings.Sign.Value)
    End Function

    Private Shared Function AddToDictionary(ByVal keyStation As Double, _
                                            ByVal station As Double, ByVal elev As Double, _
                                            ByVal location As String) As Boolean
        If keyStation >= m_stationStart And keyStation <= m_stationEnd Then
            If Not m_oProfileDataDictionary.ContainsKey(keyStation) Then
                Dim data As New ProfileData
                data.RawStation = station
                data.RawElevation = elev
                data.Grade = HTML_SPACE
                data.Location = location
                m_oProfileDataDictionary.Add(keyStation, data)
            End If
        End If
    End Function

    Private Shared Function MakeKey(ByVal oProfile As AeccLandLib.IAeccProfile, ByVal station As Double) As Double

        Diagnostics.Debug.Assert(Not oProfile Is Nothing, "Profile is nothing")

        Dim stationString As String
        Dim stationRounded As Double
        stationString = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, station)
        stationRounded = ReportUtilities.GetRawStation(stationString, oProfile.Alignment.StationIndexIncrement)

        Return stationRounded

    End Function

    Public Shared Function ExtractData(ByVal oProfile As AeccLandLib.AeccProfile, _
                                ByVal stationStart As Double, _
                                ByVal stationEnd As Double, _
                                ByVal stationInc As Double) As Boolean
        If oProfile Is Nothing Then
            Return True
        End If

        'stationStart and stationEnd might be trimmed to outside of oProfile station range
        'here pull them back
        If oProfile.StartingStation > stationStart Then
            stationStart = oProfile.StartingStation
        End If
        If oProfile.EndingStation < stationEnd Then
            stationEnd = oProfile.EndingStation
        End If

        'init member function for AddToDictionary(...) use
        m_oProfileDataDictionary.Clear()
        m_profile = oProfile
        m_stationStart = stationStart
        m_stationEnd = stationEnd

        'add PVIs
        For Each oPVI As AeccLandLib.AeccProfilePVI In oProfile.PVIs
            If Not TypeOf oPVI Is AeccLandLib.AeccProfilePVICurve Then
                'not use oPVI.Elevation for the station elevation

                'Fix defect 1462489, the input will be raw station 
                AddToDictionary(oPVI.Station, oPVI.Station, oProfile.ElevationAt(oPVI.Station), _
                    LocalizedRes.ProfStaInc_Html_PVI) 'PVI
            Else

                Dim oPVIC As AeccLandLib.AeccProfilePVICurve
                oPVIC = CType(oPVI, AeccLandLib.AeccProfilePVICurve)

                'add PVC, not use oPVIC.BeginElevation
                'Fix defect 1462489, the input will be raw station 
                AddToDictionary(oPVIC.BeginStation, oPVIC.BeginStation, oProfile.ElevationAt(oPVIC.BeginStation), _
                    LocalizedRes.ProfStaInc_Html_PVC) 'PVC

                'add PVI
                Dim crestOrSag As String
                If oPVIC.VerticalCurveType = AeccLandLib.AeccProfileVerticalCurveType.aeccCrest Then
                    crestOrSag = LocalizedRes.ProfStaInc_Html_Crest 'Crest
                ElseIf oPVIC.VerticalCurveType = AeccLandLib.AeccProfileVerticalCurveType.aeccSag Then
                    crestOrSag = LocalizedRes.ProfStaInc_Html_Sag '"Sag"
                Else
                    System.Diagnostics.Debug.Assert(False, "unknown type")
                    crestOrSag = HTML_SPACE
                End If
                'not use oPVIC.Elevation
                'Fix defect 1462489, the input will be raw station 
                AddToDictionary(oPVIC.Station, oPVIC.Station, oProfile.ElevationAt(oPVIC.Station), crestOrSag) 'Crest

                'add PVT, not use oPVIC.EndElevation
                'Fix defect 1462489, the input will be raw station 
                AddToDictionary(oPVIC.EndStation, oPVIC.EndStation, oProfile.ElevationAt(oPVIC.EndStation), _
                    LocalizedRes.ProfStaInc_Html_PVT) '"PVT"

            End If
        Next

        ' Add incremental stations
        For station As Double = stationStart To stationEnd Step stationInc
            'Fix defect 1462489, the input will be raw station 
            AddToDictionary(station, station, oProfile.ElevationAt(station), HTML_SPACE) '"&nbsp;" : space in html
        Next

        'DID 1197617, we should add the End station to the report
        If stationEnd Mod stationInc <> 0 Then
            AddToDictionary(stationEnd, stationEnd, oProfile.ElevationAt(stationEnd), HTML_SPACE) '"&nbsp;" : space in html
        End If


        'fill grade, format data
        Dim preData As ProfileData
        preData = Nothing

        For Each data As ProfileData In m_oProfileDataDictionary.Values
            'first data
            If preData Is Nothing Then
                If data.RawStation = oProfile.StartingStation Then
                    data.Grade = HTML_SPACE
                End If
            Else
                Dim deltaStation As Double
                deltaStation = data.RawStation - preData.RawStation
                If deltaStation = 0 Then
                    'error, we've checked this when add to dictionary
                    System.Diagnostics.Debug.Assert(False)

                    data.Grade = HTML_SPACE
                Else
                    Dim gradeVal As Double
                    gradeVal = (data.RawElevation - preData.RawElevation) / deltaStation
                    data.Grade = FormatGradSettings(gradeVal)
                End If
            End If

            preData = data

            data.Station = ReportUtilities.GetStationStringWithDerived(oProfile.Alignment, data.RawStation)
            data.Elevation = FormatElevSettings(data.RawElevation)
        Next

        Return True
    End Function
End Class
