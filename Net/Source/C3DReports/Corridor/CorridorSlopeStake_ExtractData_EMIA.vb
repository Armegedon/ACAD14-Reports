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

Friend Class Slope_SectionData_Emia
    Public LeftDatas As New SortedDictionary(Of Double, Slope_PointData_Emia)
    Public RightDatas As New SortedDictionary(Of Double, Slope_PointData_Emia)
End Class

Friend Class Slope_PointData_Emia

    Public mCodes As String = ""
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

    Public Sub Fill(ByVal Codes As String, ByVal Offset As Double, _
                    ByVal Elev As Double, _
                    ByVal East As Double, ByVal North As Double, Optional ByVal Slope As String = "")
        mCodes = Codes
        mOffsetValue = Offset
        mElevationValue = Elev
        mEastingValue = East
        mNorthingValue = North
        mSlope = Slope

        mOffsetString = CorridorSlopeStake_ExtractData_Emia.FormatDistSettings(mOffsetValue, mOffsetRounded)
        mElevationString = CorridorSlopeStake_ExtractData_Emia.FormatElevSettings(mElevationValue, mElevationRounded)
        mEastingString = CorridorSlopeStake_ExtractData_Emia.FormatCoordSettings(mEastingValue, mEastingRounded)
        mNorthingString = CorridorSlopeStake_ExtractData_Emia.FormatCoordSettings(mNorthingValue, mNorthingRounded)
    End Sub

    Public Sub New()

    End Sub
End Class

Friend Class CorridorSlopeStake_ExtractData_Emia

    Private Shared m_oSlopeStakeData As New Dictionary(Of Double, Slope_SectionData_Emia)
    Public Shared ReadOnly Property SlopeStakeData() As Dictionary(Of Double, Slope_SectionData_Emia)
        Get
            Return m_oSlopeStakeData
        End Get
    End Property

    Private Shared Function FormatCoordSettings(ByVal dDis As Double) As String
        Dim oCoordSettings As AeccLandLib.AeccSettingsCoordinate
        oCoordSettings = ReportApplication.AeccXRoadwayDatabase.Settings.SectionSettings.AmbientSettings.CoordinateSettings()

        FormatCoordSettings = ReportFormat.FormatDistance(dDis, oCoordSettings.Unit.Value, _
            oCoordSettings.Precision.Value, oCoordSettings.Rounding.Value, oCoordSettings.Sign.Value)
    End Function

    Public Shared Function FormatDistSettings(ByVal dDis As Double, ByRef dDisRounded As Double) As String
        Dim oDistSettings As AeccLandLib.AeccSettingsDistance
        oDistSettings = ReportApplication.AeccXRoadwayDatabase.Settings.SectionSettings.AmbientSettings.DistanceSettings

        'oDistSettings.Precision.Value
        dDisRounded = ReportFormat.RoundDouble(dDis, 2, oDistSettings.Rounding.Value)

        FormatDistSettings = ReportFormat.FormatDistance(dDis, oDistSettings.Unit.Value, _
            oDistSettings.Precision.Value, oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
    End Function

    Private Shared Function FormatDirSettings(ByVal dDirection As Double) As String
        Dim oDirSettings As AeccLandLib.AeccSettingsDirection
        oDirSettings = ReportApplication.AeccXRoadwayDatabase.Settings.SectionSettings.AmbientSettings.DirectionSettings

        FormatDirSettings = ReportFormat.FormatDirection(dDirection, oDirSettings.Unit.Value, _
            oDirSettings.Precision.Value, oDirSettings.Rounding.Value, oDirSettings.Format.Value, _
            oDirSettings.Direction.Value, oDirSettings.Capitalization.Value, _
            oDirSettings.Sign.Value, oDirSettings.MeasurementType.Value, oDirSettings.BearingQuadrant.Value)
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

    Private Shared Function calcStepsNeeded(ByVal oBaseline As AeccRoadLib.IAeccBaseline, _
                                ByVal oSampleLineGroups As AeccLandLib.IAeccSampleLineGroup, _
                                ByVal stationStart As Double, _
                                ByVal stationEnd As Double, _
                                ByVal sCodeName As String)
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

            Dim oLinks As AeccRoadLib.AeccCalculatedLinks
            Try
                oLinks = oBaseline.AppliedAssembly(curStation).GetLinksByCode(sCodeName)
            Catch ex As Exception
                'Diagnostics.Debug.Assert(False, ex.Message)
                'error, try next point
                Continue For
            End Try

            For Each oLink As AeccRoadLib.AeccCalculatedLink In oLinks
                'assume 2 code point in a links
                steps += 2
            Next
        Next

        Return steps
    End Function

    Public Shared Function ExtractData(ByVal oBaseline As AeccRoadLib.IAeccBaseline, _
                                ByVal oSampleLineGroups As AeccLandLib.IAeccSampleLineGroup, _
                                ByVal stationStart As Double, _
                                ByVal stationEnd As Double, _
                                ByVal sCodeName As String, _
                                ByVal ctlProgress As ProgressBar, _
                                ByVal steps As Integer) As Boolean
        Dim PtSteped As Integer
        Dim oneStepPtCount As Integer
        Dim stepNeeded As Integer
        Dim stepsSoFar As Integer
        stepNeeded = calcStepsNeeded(oBaseline, oSampleLineGroups, stationStart, stationEnd, sCodeName)

        If stepNeeded = 0 Then
            oneStepPtCount = -1
        Else
            oneStepPtCount = Math.Ceiling(steps / stepNeeded)
        End If

        'for offset rounding
        Dim oDistSettings As AeccLandLib.AeccSettingsDistance
        oDistSettings = ReportApplication.AeccXRoadwayDatabase.Settings.SectionSettings.AmbientSettings.DistanceSettings

        m_oSlopeStakeData.Clear()
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

            Dim datumElevation As Double
            datumElevation = oBaseline.Profile.ElevationAt(curStation)

            Dim oLinks As AeccRoadLib.AeccCalculatedLinks
            Dim sData As New Slope_SectionData_Emia
            Dim hasCrossLink As Boolean = False

            Try
                oLinks = oBaseline.AppliedAssembly(curStation).GetLinksByCode(sCodeName)
            Catch ex As Exception
                'Diagnostics.Debug.Assert(False, ex.Message)
                'error, try next point
                Continue For
            End Try

            For Each oLink As AeccRoadLib.AeccCalculatedLink In oLinks
                Dim hasLeftData As Boolean = False
                Dim HasRightData As Boolean = False
                Dim oPoint As AeccRoadLib.AeccCalculatedPoint
                For Each oPoint In oLink.CalculatedPoints
                    Dim ptOffsetElev As Object
                    ptOffsetElev = oPoint.GetStationOffsetElevationToBaseline()

                    Dim ptOffset As Double
                    Dim ptElev As Double
                    Dim ptEast As Double
                    Dim ptNorth As Double
                    Dim Codes As String
                    ptOffset = ptOffsetElev(1)
                    ptElev = ptOffsetElev(2) + datumElevation
                    Dim ptXYZ As Object = oBaseline.StationOffsetElevationToXYZ(ptOffsetElev)
                    ptEast = CDbl(ptXYZ(0))
                    ptNorth = CDbl(ptXYZ(1))
                    If oPoint.CorridorCodes.Count = 0 Then
                        Codes = ""
                    Else
                        Codes = oPoint.CorridorCodes.Item(0)
                    End If

                    'fill point data, slope will be calculated later for points not sorted not
                    Dim ptData As New Slope_PointData_Emia
                    ptData.Fill(Codes, ptOffset, ptElev, ptEast, ptNorth)

                    If ptOffset < 0 Then
                        'has data judge should put outside If() statement, for data may added
                        hasLeftData = True
                        If Not sData.LeftDatas.ContainsKey(Math.Abs(ptData.mOffsetRounded)) Then
                            sData.LeftDatas.Add(Math.Abs(ptData.mOffsetRounded), ptData)
                        Else
                            'check point code
                            If ptData.mCodes <> "" Then
                                Dim existPt As Slope_PointData_Emia
                                existPt = sData.LeftDatas.Item(Math.Abs(ptData.mOffsetRounded))
                                If existPt.mCodes = "" Then
                                    existPt.mCodes = ptData.mCodes
                                End If
                            End If
                        End If
                    Else
                        HasRightData = True
                        If Not sData.RightDatas.ContainsKey(ptData.mOffsetRounded) Then
                            sData.RightDatas.Add(ptData.mOffsetRounded, ptData)
                        Else
                            'check point code
                            If ptData.mCodes <> "" Then
                                Dim existPt As Slope_PointData_Emia
                                existPt = sData.RightDatas.Item(ptData.mOffsetRounded)
                                If existPt.mCodes = "" Then
                                    existPt.mCodes = ptData.mCodes
                                End If
                            End If
                        End If
                    End If

                    'step progress
                    If PtSteped = oneStepPtCount Then
                        'ctlProgress.PerformStep()
                        PtSteped = 0
                        stepsSoFar += 1
                    Else
                        PtSteped += 1
                    End If
                Next
                If hasLeftData And HasRightData Then
                    hasCrossLink = True
                End If
            Next

            Dim firstL As Slope_PointData_Emia
            Dim firstR As Slope_PointData_Emia
            firstL = formatPointDatas(sData.LeftDatas)
            firstR = formatPointDatas(sData.RightDatas)

            'calculate first point slope
            If hasCrossLink Then
                firstL.mSlope = GetSlopeData(firstL.mOffsetValue, firstR.mOffsetValue, _
                    firstL.mElevationValue, firstR.mElevationValue)
                firstR.mSlope = firstL.mSlope
            Else
                'the first left and right point not in a link
                If Not firstL Is Nothing Then
                    firstL.mSlope = ""
                End If
                If Not firstR Is Nothing Then
                    firstR.mSlope = ""
                End If
            End If

            If Not m_oSlopeStakeData.ContainsKey(curStationRounded) Then
                If sData.LeftDatas.Count > 0 Or sData.RightDatas.Count > 0 Then
                    m_oSlopeStakeData.Add(curStationRounded, sData)
                End If
            End If
        Next
        'For i As Integer = stepsSoFar To steps
        '    ctlProgress.PerformStep()
        'Next i

        Return True
    End Function

    Private Shared Function formatPointDatas(ByRef datas As SortedDictionary(Of Double, Slope_PointData_Emia)) _
                                            As Slope_PointData_Emia
        If datas.Values.Count = 0 Then
            'no point data
            Return Nothing
        End If

        Dim firstPt As Slope_PointData_Emia = Nothing
        Dim lastElev, curElev, lastOffset, curOffset As Double
        Dim counter As Integer = 0
        For Each ptData As Slope_PointData_Emia In datas.Values
            lastElev = curElev
            lastOffset = curOffset
            curElev = ptData.mElevationValue
            curOffset = ptData.mOffsetValue
            If counter = 0 Then
                'first point slope should be calculated outside this function
                firstPt = ptData
            Else
                ptData.mSlope = GetSlopeData(curOffset, lastOffset, curElev, lastElev)
            End If
            counter += 1
        Next

        Return firstPt
    End Function

    Private Shared Function GetSlopeData(ByVal curOff As Double, ByVal lastOff As Double, _
                                ByVal curElev As Double, ByVal lastElev As Double) As String
        Dim deltaX As Double
        Dim deltaY As Double
        Dim slope As Double
        Const percent2Slope = 0.1
        deltaX = Math.Abs(curOff - lastOff)
        deltaY = curElev - lastElev

        If (deltaX <> 0) Then
            slope = deltaY / deltaX
            If (Math.Abs(slope) <= percent2Slope) Then
                GetSlopeData = slope.ToString("p")
            Else
                GetSlopeData = LocalizedRes.CorridorSlopeStake_Html_One + _
                    slope.ToString("0.00")
            End If
        Else
            ' Vertical link slopes are printed as "vert"
            GetSlopeData = "vert"
        End If

    End Function
End Class
