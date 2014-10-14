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

Imports LocalizedRes = Report.My.Resources.LocalizedStrings
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports AeccRoadLib = Autodesk.AECC.Interop.Roadway

Friend Class CorridorCS_SectionData
    Public Datas As New SortedDictionary(Of Double, CorridorCS_PointData)
End Class

Friend Class CorridorCS_PointData

    Public mOffsetString As String = ""
    Public mOffsetValue As Double
    Public mOffsetRounded As Double
    Public mElevationString As String = ""
    Public mElevationValue As Double
    Public mElevationRounded As Double
    Public mEastingString As String = ""
    Public mEastingValue As Double
    Public mEastingRounded As Double
    Public mNorthingString As String = ""
    Public mNorthingValue As Double
    Public mNorthingRounded As Double
    Public mSlope As String = ""

    Public Sub Fill(ByVal Offset As Double, _
                    ByVal Elev As Double, _
                    ByVal East As Double, ByVal North As Double, Optional ByVal Slope As String = "")
        mOffsetValue = Offset
        mElevationValue = Elev
        mEastingValue = East
        mNorthingValue = North
        mSlope = Slope

        mOffsetString = CorridorCrossSection_ExtractData.FormatDistSettings(mOffsetValue, mOffsetRounded)
        mElevationString = CorridorCrossSection_ExtractData.FormatElevSettings(mElevationValue, mElevationRounded)
        mEastingString = CorridorCrossSection_ExtractData.FormatCoordSettings(mEastingValue, mEastingRounded)
        mNorthingString = CorridorCrossSection_ExtractData.FormatCoordSettings(mNorthingValue, mNorthingRounded)
    End Sub
End Class

Friend Class CorridorCrossSection_ExtractData

    Private Shared m_oCorridorCrossSectionData As New Dictionary(Of Double, CorridorCS_SectionData)

    Public Shared ReadOnly Property SlopeStakeData() As Dictionary(Of Double, CorridorCS_SectionData)
        Get
            Return m_oCorridorCrossSectionData
        End Get
    End Property

    Public Shared Function FormatDistSettings(ByVal dDis As Double, ByRef dDisRounded As Double) As String
        Dim oDistSettings As AeccLandLib.IAeccSettingsDistance
        oDistSettings = ReportApplication.AeccXRoadwayDatabase.Settings.SectionSettings.AmbientSettings.DistanceSettings

        'oDistSettings.Precision.Value
        dDisRounded = ReportFormat.RoundDouble(dDis, 2, oDistSettings.Rounding.Value)

        FormatDistSettings = ReportFormat.FormatDistance(dDis, oDistSettings.Unit.Value, _
            oDistSettings.Precision.Value, oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
    End Function

    Public Shared Function FormatElevSettings(ByVal dElev As Double, ByRef dElevRounded As Double) As String
        Dim oElevSettings As AeccLandLib.AeccSettingsElevation
        oElevSettings = ReportApplication.AeccXRoadwayDatabase.Settings.SectionSettings.AmbientSettings.ElevationSettings

        dElevRounded = ReportFormat.RoundDouble(dElev, oElevSettings.Precision.Value, oElevSettings.Rounding.Value)

        FormatElevSettings = ReportFormat.FormatDistance(dElev, oElevSettings.Unit.Value, oElevSettings.Precision.Value, _
                    oElevSettings.Rounding.Value, oElevSettings.Sign.Value)
    End Function

    Public Shared Function FormatCoordSettings(ByVal dCoord As Double, ByRef dCoordRounded As Double) As String
        Dim oCoordSettings As AeccLandLib.AeccSettingsCoordinate
        oCoordSettings = ReportApplication.AeccXRoadwayDatabase.Settings.SectionSettings.AmbientSettings.CoordinateSettings

        dCoordRounded = ReportFormat.RoundDouble(dCoord, oCoordSettings.Precision.Value, oCoordSettings.Rounding.Value)

        FormatCoordSettings = ReportFormat.FormatDistance(dCoord, oCoordSettings.Unit.Value, oCoordSettings.Precision.Value, _
                    oCoordSettings.Rounding.Value, oCoordSettings.Sign.Value, False)
    End Function

    Private Shared Function calcStepsNeeded(ByVal oBaseAlignment As AeccLandLib.AeccAlignment, _
                                ByVal oSampleLineGroups As AeccLandLib.IAeccSampleLineGroup, _
                                ByVal oSurface As AeccLandLib.AeccSurface, _
                                ByVal stationStart As Double, _
                                ByVal stationEnd As Double) As Integer
        Dim steps As Integer
        Dim lastStation As Double
        'set lastStation invalid
        lastStation = stationStart - 100

        For Each oSampleLine As AeccLandLib.AeccSampleLine In oSampleLineGroups.SampleLines
            Dim curStation As Double
            curStation = oSampleLine.Station


            If curStation < stationStart Or curStation > stationEnd Then
                Continue For
            End If

            'just calc steps, assume precision is 0.0001
            If curStation - lastStation < 0.0001 Then
                Continue For
            Else
                lastStation = curStation
            End If

            steps += 1
        Next

        Return steps
    End Function

    Public Shared Function ExtractData(ByVal oBaseAlignment As AeccLandLib.AeccAlignment, _
                                ByVal oSampleLineGroups As AeccLandLib.IAeccSampleLineGroup, _
                                ByVal oSurface As AeccLandLib.AeccSurface, _
                                ByVal stationStart As Double, _
                                ByVal stationEnd As Double, _
                                ByVal ctlProgress As ProgressBar, _
                                ByVal steps As Integer) As Boolean
        Dim PtSteped As Integer
        Dim oneStepPtCount As Integer
        Dim stepNeeded As Integer
        Dim stepsSoFar As Integer

        If oBaseAlignment Is Nothing Then Return False
        If oSurface Is Nothing Then Return False

        stepNeeded = calcStepsNeeded(oBaseAlignment, oSampleLineGroups, oSurface, stationStart, stationEnd)
        If stepNeeded = 0 Then
            oneStepPtCount = -1
        Else
            oneStepPtCount = Convert.ToInt32(Math.Ceiling(steps / stepNeeded))
        End If

        'for offset rounding
        Dim oDistSettings As AeccLandLib.IAeccSettingsDistance _
            = ReportApplication.AeccXRoadwayDatabase.Settings.SectionSettings.AmbientSettings.DistanceSettings

        m_oCorridorCrossSectionData.Clear()
        Dim lastStation As Double
        'set lastStation invalid
        lastStation = stationStart - 100

        For Each oSampleLine As AeccLandLib.AeccSampleLine In oSampleLineGroups.SampleLines
            Dim curStationRounded As Double
            Dim curStation As Double
            Dim curStationString As String
            curStation = oSampleLine.Station
            curStationString = ReportUtilities.GetStationString(oSampleLineGroups.Parent, curStation)
            curStationRounded = ReportUtilities.GetRawStation(curStationString, oSampleLineGroups.Parent.StationIndexIncrement)
            If curStationRounded < stationStart Or curStationRounded > stationEnd Then
                Continue For
            End If
            If curStationRounded = lastStation Then
                Continue For
            Else
                lastStation = curStationRounded
            End If

            Dim sData As New CorridorCS_SectionData

            ' get start and end vertex
            Dim vertices As AeccLandLib.IAeccSampleLineVertices = oSampleLine.Vertices
            Dim vtxstart As AeccLandLib.IAeccSampleLineVertex = Nothing
            Dim vtxend As AeccLandLib.IAeccSampleLineVertex = Nothing

            If vertices Is Nothing Then Continue For

            If vertices.Count() > 1 Then
                vtxstart = vertices.Item(0)
                vtxend = vertices.Item(vertices.Count() - 1)
            End If

            If vtxstart Is Nothing Or vtxend Is Nothing Then Continue For
            ' now get all the vertices in between
            Dim pCoords As Array
            Try
                pCoords = oSurface.SampleElevations(vtxstart.Location(0), vtxstart.Location(1), vtxend.Location(0), vtxend.Location(1))
            Catch
                Continue For
            End Try

            Dim i As Integer = 0

            While i < pCoords.GetUpperBound(0)
                Dim ptData As New CorridorCS_PointData

                Dim ptElev As Double = pCoords(i + 2)
                Dim ptEast As Double = pCoords(i)
                Dim ptNorth As Double = pCoords(i + 1)
                Dim ptStat As Double = 0.0
                Dim ptOffset As Double = 0.0
                oBaseAlignment.StationOffset(ptEast, ptNorth, ptStat, ptOffset)

                ptData.Fill(ptOffset, ptElev, ptEast, ptNorth)
                If Not sData.Datas.ContainsKey(ptData.mOffsetRounded) Then
                    sData.Datas.Add(ptData.mOffsetRounded, ptData)
                End If

                i += 3
            End While

            'step progress
            If PtSteped = oneStepPtCount Then
                ctlProgress.PerformStep()
                PtSteped = 0
                stepsSoFar += 1
            Else
                PtSteped += 1
            End If

            If Not m_oCorridorCrossSectionData.ContainsKey(curStationRounded) Then
                If sData.Datas.Count > 0 Then
                    m_oCorridorCrossSectionData.Add(curStationRounded, sData)
                End If
            End If
        Next
        For i As Integer = stepsSoFar To steps
            ctlProgress.PerformStep()
        Next i

        Return True
    End Function
End Class
