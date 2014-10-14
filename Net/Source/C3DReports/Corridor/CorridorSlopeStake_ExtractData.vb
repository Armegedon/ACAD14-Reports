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
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.Civil.DatabaseServices

Friend Class Slope_SectionData
    Public LeftDatas As New SortedDictionary(Of Double, Slope_PointData)
    Public LeftEndInfo As Slope_EndInfo
    Public RightDatas As New SortedDictionary(Of Double, Slope_PointData)
    Public RightEndInfo As Slope_EndInfo
    Public CenterLineInfo As New Slope_PointData
    Public LinksOnGraph As New SortedDictionary(Of Integer, List(Of Slope_PointData))
    Public ROWPoints As New List(Of Slope_PointData)
    Public EGDatas As New SortedDictionary(Of Integer, List(Of Slope_PointData))
    Public HaveMaterialInfo As Boolean = False
    Public CutArea As Double = 0.0
    Public FillArea As Double = 0.0
    Public CumulativeNetVolume As Double = 0.0
End Class

Friend Class Slope_EndInfo
    Public EndType As String
    Public DeltaOffset As String
    Public EndSlope As String
End Class

Friend Class Slope_PointData

    Public mCodes As String = ""
    Public mOffsetString As String = ""
    Public mOffsetValue As Double
    Public mOffsetRounded As Double
    Public mElevationString As String = ""
    Public mElevationValue As Double
    Public mElevationRounded As Double
    Public mSlope As String = ""

    Public Sub Fill(ByVal Codes As String, ByVal Offset As Double, _
                    ByVal Elev As Double, Optional ByVal Slope As String = "")
        mCodes = Codes
        mOffsetValue = Offset
        mElevationValue = Elev
        mSlope = Slope

        mOffsetString = CorridorSlopeStake_ExtractData.FormatDistSettings(mOffsetValue, mOffsetRounded)
        mElevationString = CorridorSlopeStake_ExtractData.FormatElevSettings(mElevationValue, mElevationRounded)
    End Sub
End Class

Friend Class Slope_PointData_Sorter
    Implements IComparer(Of Slope_PointData)
    Function Compare(ByVal x As Slope_PointData, ByVal y As Slope_PointData) As Integer _
        Implements IComparer(Of Slope_PointData).Compare
        If x.mOffsetValue < y.mOffsetValue Then
            Return -1
        ElseIf x.mOffsetValue = y.mOffsetValue Then
            Return 0
        Else
            Return 1
        End If
    End Function 'IComparer.Compare
End Class 'Slope_PointData_Sorter

Friend Class CorridorSlopeStake_ExtractData

    Private Shared m_oSlopeStakeData As New Dictionary(Of Double, Slope_SectionData)
    Public Shared ReadOnly Property SlopeStakeData() As Dictionary(Of Double, Slope_SectionData)
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
        Dim oDistSettings As AeccLandLib.IAeccSettingsDistance
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

    Private Shared Function calcStepsNeeded(ByVal oBaseline As AeccRoadLib.IAeccBaseline, _
                                ByVal oSampleLineGroups As AeccLandLib.IAeccSampleLineGroup, _
                                ByVal stationStart As Double, _
                                ByVal stationEnd As Double, _
                                ByVal sCodeName As String) As Integer
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
                oLinks = oBaseline.AppliedAssembly(oSampleLine.Station).GetLinksByCode(sCodeName)
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
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
                                ByVal sLinkCode As String, _
                                ByVal sROWPointCode As String, _
                                ByVal sMaterialListName As String, _
                                ByVal ctlProgress As ProgressBar, _
                                ByVal steps As Integer) As Boolean
        Dim PtSteped As Integer
        Dim oneStepPtCount As Integer
        Dim stepNeeded As Integer
        stepNeeded = calcStepsNeeded(oBaseline, oSampleLineGroups, stationStart, stationEnd, sLinkCode)
        If steps = 0 Then
            oneStepPtCount = -1
        Else
            oneStepPtCount = CInt(Math.Ceiling(stepNeeded / steps))
        End If

        'for offset rounding
        Dim oDistSettings As AeccLandLib.IAeccSettingsDistance
        oDistSettings = ReportApplication.AeccXRoadwayDatabase.Settings.SectionSettings.AmbientSettings.DistanceSettings

        m_oSlopeStakeData.Clear()
        Dim lastStation As Double
        'set lastStation invalid
        lastStation = stationStart - 100

        Dim oLineGroup As AeccLandLib.AeccSampleLineGroup = oSampleLineGroups
        Dim sampleLineGroupId As ObjectId = Autodesk.AutoCAD.DatabaseServices.DBObject.FromAcadObject(oLineGroup)

        Dim bHaveMaterialInfo As Boolean = False
        If sMaterialListName <> LocalizedRes.CorridorSlopeStake_None Then bHaveMaterialInfo = True

        Dim qtoSectionalResult() As QTOSectionalResult = Nothing

        If bHaveMaterialInfo Then
            Using trans As Transaction = sampleLineGroupId.Database.TransactionManager.StartTransaction
                'Open SampleLineGroup
                Dim oSampleLineGroupExt As SampleLineGroup = trans.GetObject(sampleLineGroupId, OpenMode.ForRead)

                'Get volume info
                Dim mappingGuid As System.Guid = oSampleLineGroupExt.GetMappingGuid(sMaterialListName)
                Dim qtoResult As QuantityTakeoffResult = Nothing
                qtoResult = oSampleLineGroupExt.GetTotalVolumeResultDataForMaterialList(mappingGuid)
                qtoSectionalResult = qtoResult.GetResultsAlongSampleLines()
                Debug.Assert(qtoSectionalResult.Length = oSampleLineGroups.SampleLines.Count)
            End Using
        End If

        Dim index As Integer = 0

        For Each oSampleLine As AeccLandLib.AeccSampleLine In oSampleLineGroups.SampleLines
            Dim curStationRounded As Double
            Dim curStation As Double
            Dim curStationString As String
            curStation = oSampleLine.Station
            curStationString = ReportUtilities.GetStationString(oSampleLineGroups.Parent, curStation)
            curStationRounded = ReportUtilities.GetRawStation(curStationString, oSampleLineGroups.Parent.StationIndexIncrement)
            If curStationRounded < stationStart Or curStationRounded > stationEnd Then
                index = index + 1
                Continue For
            End If
            If curStationRounded = lastStation Then
                index = index + 1
                Continue For
            Else
                lastStation = curStationRounded
            End If

            Dim datumElevation As Double
            Dim oLinks As AeccRoadLib.AeccCalculatedLinks
            Dim oLinksOnGraphics As AeccRoadLib.AeccCalculatedLinks
            Dim sData As New Slope_SectionData

            Try
                datumElevation = oBaseline.Profile.ElevationAt(curStation)
                oLinks = oBaseline.AppliedAssembly(curStation).GetLinksByCode(sLinkCode)
                oLinksOnGraphics = oBaseline.AppliedAssembly(curStation).GetLinks()
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
                'error, try next point
                index = index + 1
                Continue For
            End Try

            ' material info
            If bHaveMaterialInfo Then
                sData.HaveMaterialInfo = bHaveMaterialInfo
                sData.CutArea = qtoSectionalResult(index).AreaResult.CutArea
                sData.FillArea = qtoSectionalResult(index).AreaResult.FillArea
                sData.CumulativeNetVolume = qtoSectionalResult(index).VolumeResult.CumulativeCutVolume - qtoSectionalResult(index).VolumeResult.CumulativeFillVolume
            End If

            ' Add center line point
            ' TODO: localization
            sData.CenterLineInfo.Fill("CL", 0.0, datumElevation)

            ' Add link points array to graph as assembly
            For idxOfLinksOnGraph As Integer = 0 To oLinksOnGraphics.Count - 1
                'For Each olinkOnGraph As AeccRoadLib.AeccCalculatedLink In oLinksOnGraphics
                Dim olinkOnGraph As AeccRoadLib.AeccCalculatedLink
                olinkOnGraph = oLinksOnGraphics.Item(idxOfLinksOnGraph)

                Dim oPointOnGraph As AeccRoadLib.AeccCalculatedPoint
                For Each oPointOnGraph In olinkOnGraph.CalculatedPoints
                    Dim ptOffsetElev() As Double
                    ptOffsetElev = CType(oPointOnGraph.GetStationOffsetElevationToBaseline(), Double())

                    Dim ptOffset As Double
                    Dim ptElev As Double
                    Dim Codes As String
                    ptOffset = ptOffsetElev(1)
                    ptElev = ptOffsetElev(2) + datumElevation
                    If oPointOnGraph.CorridorCodes.Count = 0 Then
                        Codes = ""
                    Else
                        Codes = oPointOnGraph.CorridorCodes.Item(0)
                    End If

                    'fill point data, slope will be calculated later for points not sorted not
                    Dim ptData As New Slope_PointData
                    ptData.Fill(Codes, ptOffset, ptElev)

                    If Not sData.LinksOnGraph.ContainsKey(idxOfLinksOnGraph) Then
                        sData.LinksOnGraph.Add(idxOfLinksOnGraph, New List(Of Slope_PointData))
                    End If
                    Dim PtDataList As List(Of Slope_PointData)
                    PtDataList = sData.LinksOnGraph.Item(idxOfLinksOnGraph)
                    PtDataList.Add(ptData)
                Next
            Next

            ' Add ROW points to graph
            Dim oROWPoints As AeccRoadLib.AeccCalculatedPoints
            Try
                oROWPoints = oBaseline.AppliedAssembly(curStation).GetPointsByCode(sROWPointCode)
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
                'error, try next point
                Continue For
            End Try

            For Each oROWPoint As AeccRoadLib.AeccCalculatedPoint In oROWPoints

                Dim ptOffsetElev() As Double
                ptOffsetElev = CType(oROWPoint.GetStationOffsetElevationToBaseline(), Double())

                Dim ptOffset As Double = ptOffsetElev(1)
                Dim ptElev As Double = ptOffsetElev(2) + datumElevation
                Dim code As String = "ROW: " + sROWPointCode

                Dim ptROW As New Slope_PointData
                ptROW.Fill(code, ptOffset, ptElev)

                sData.ROWPoints.Add(ptROW)
            Next

            ' Add EG section point
            Dim nIdxOfEG As Integer = -1 ' Though we expect only 1 eg, but may be multiple surfaces lying there
            For Each oSection As AeccLandLib.AeccSection In oSampleLine.Sections
                If oSection.DataType = AeccLandLib.AeccSectionDataType.aeccSectionDataGridSurface _
                   Or oSection.DataType = AeccLandLib.AeccSectionDataType.aeccSectionDataTIN Then

                    ' in order to know if the surface really intersect with SL,
                    ' try to get section's elevation, which will finally reach AeccImpSection::hasValidElevationData()
                    Dim hasValidElevationData As Boolean = False
                    Try
                        Dim dElevMin As Double
                        dElevMin = oSection.ElevationMin
                        hasValidElevationData = True
                    Catch ex As Exception
                        hasValidElevationData = False
                    End Try
                    If hasValidElevationData = False Then
                        'the links don't pass over a valid portion of the surface, try next
                        Continue For
                    End If

                    nIdxOfEG += 1
                    sData.EGDatas.Add(nIdxOfEG, New List(Of Slope_PointData))
                    Dim ptDataList As List(Of Slope_PointData)
                    ptDataList = sData.EGDatas(nIdxOfEG)

                    Dim oSectionLinks As AeccLandLib.AeccSectionLinks
                    Dim oSectionLink As AeccLandLib.AeccSectionLink
                    oSectionLinks = oSection.Links

                    Dim nLinksCount = oSectionLinks.Count
                    For idxOfSecLink As Integer = 0 To nLinksCount - 1
                        ' For Each oSectionLink In oSectionLinks
                        oSectionLink = oSectionLinks.Item(idxOfSecLink)
                        Dim ptEGDataStart As New Slope_PointData
                        ptEGDataStart.Fill("", oSectionLink.StartPointX, oSectionLink.StartPointY)
                        ptDataList.Add(ptEGDataStart)
                        ' add last end pt
                        If idxOfSecLink = nLinksCount - 1 Then
                            Dim ptEGDataEnd As New Slope_PointData
                            ptEGDataEnd.Fill("", oSectionLink.EndPointX, oSectionLink.EndPointY)
                            ptDataList.Add(ptEGDataEnd)
                        End If
                    Next
                End If
            Next

            For Each oLink As AeccRoadLib.AeccCalculatedLink In oLinks
                Dim oPoint As AeccRoadLib.AeccCalculatedPoint
                For Each oPoint In oLink.CalculatedPoints
                    Dim ptOffsetElev() As Double
                    ptOffsetElev = CType(oPoint.GetStationOffsetElevationToBaseline(), Double())

                    Dim ptOffset As Double
                    Dim ptElev As Double
                    Dim Codes As String
                    ptOffset = ptOffsetElev(1)
                    ptElev = ptOffsetElev(2) + datumElevation
                    If oPoint.CorridorCodes.Count = 0 Then
                        Codes = ""
                    Else
                        Codes = oPoint.CorridorCodes.Item(0)
                    End If

                    'fill point data, slope will be calculated later for points not sorted not
                    Dim ptData As New Slope_PointData
                    ptData.Fill(Codes, ptOffset, ptElev)

                    If ptOffset < 0 Then
                        'has data judge should put outside If() statement, for data may added
                        If Not sData.LeftDatas.ContainsKey(Math.Abs(ptData.mOffsetRounded)) Then
                            sData.LeftDatas.Add(Math.Abs(ptData.mOffsetRounded), ptData)
                        Else
                            'check point code
                            If ptData.mCodes <> "" Then
                                Dim existPt As Slope_PointData
                                existPt = sData.LeftDatas.Item(Math.Abs(ptData.mOffsetRounded))
                                If existPt.mCodes = "" Then
                                    existPt.mCodes = ptData.mCodes
                                End If
                            End If
                        End If
                    Else
                        If Not sData.RightDatas.ContainsKey(ptData.mOffsetRounded) Then
                            sData.RightDatas.Add(ptData.mOffsetRounded, ptData)
                        Else
                            'check point code
                            If ptData.mCodes <> "" Then
                                Dim existPt As Slope_PointData
                                existPt = sData.RightDatas.Item(ptData.mOffsetRounded)
                                If existPt.mCodes = "" Then
                                    existPt.mCodes = ptData.mCodes
                                End If
                            End If
                        End If
                    End If

                    'step progress
                    If PtSteped = oneStepPtCount Then
                        ctlProgress.PerformStep()
                        PtSteped = 0
                    Else
                        PtSteped += 1
                    End If
                Next
            Next

            If (sData.LeftDatas.Count + sData.RightDatas.Count) >= 2 Then
                If sData.LeftDatas.Count < 2 Then ' at leaset 2 last offset points
                    Dim tempDatasForLeftEndInfo As New SortedDictionary(Of Double, Slope_PointData)
                    For Each ptData As Slope_PointData In sData.LeftDatas.Values
                        tempDatasForLeftEndInfo.Add(ptData.mOffsetRounded, ptData)
                    Next
                    Dim i As Integer = 0
                    Dim iMax As Integer = 1 - tempDatasForLeftEndInfo.Count
                    For Each keyOffset As Double In sData.RightDatas.Keys
                        If i <= iMax Then
                            Dim ptData As Slope_PointData = sData.RightDatas(keyOffset)
                            tempDatasForLeftEndInfo.Add(ptData.mOffsetRounded, ptData)
                            i += 1
                        End If
                    Next
                    formatPointDatas(tempDatasForLeftEndInfo, sData.LeftEndInfo)
                    For Each keyOffset As Double In sData.LeftDatas.Keys
                        Dim ptData As Slope_PointData = sData.LeftDatas(keyOffset)
                        ptData = tempDatasForLeftEndInfo(ptData.mOffsetRounded)
                    Next
                Else
                    formatPointDatas(sData.LeftDatas, sData.LeftEndInfo)
                End If

                If sData.RightDatas.Count < 2 Then ' at leaset 2 last offset points
                    Dim tempDatasForRightEndInfo As New SortedDictionary(Of Double, Slope_PointData)
                    For Each ptData As Slope_PointData In sData.RightDatas.Values
                        tempDatasForRightEndInfo.Add(ptData.mOffsetRounded, ptData)
                    Next
                    Dim i As Integer = 0
                    Dim iMax As Integer = 1 - tempDatasForRightEndInfo.Count
                    For Each keyOffset As Double In sData.LeftDatas.Keys
                        If i <= iMax Then
                            Dim ptData As Slope_PointData = sData.LeftDatas(keyOffset)
                            tempDatasForRightEndInfo.Add(ptData.mOffsetRounded, ptData)
                            i += 1
                        End If
                    Next
                    formatPointDatas(tempDatasForRightEndInfo, sData.RightEndInfo)
                    For Each keyOffset As Double In sData.RightDatas.Keys
                        Dim ptData As Slope_PointData = sData.RightDatas(keyOffset)
                        ptData = tempDatasForRightEndInfo(ptData.mOffsetRounded)
                    Next
                Else
                    formatPointDatas(sData.RightDatas, sData.RightEndInfo)
                End If
            End If

            If Not m_oSlopeStakeData.ContainsKey(curStationRounded) Then
                If sData.LeftDatas.Count > 0 Or sData.RightDatas.Count > 0 Then
                    m_oSlopeStakeData.Add(curStationRounded, sData)
                End If
            Else
                Dim existingData As Report.Slope_SectionData = m_oSlopeStakeData.Item(curStationRounded)
                existingData.HaveMaterialInfo = sData.HaveMaterialInfo
                existingData.CutArea = sData.CutArea
                existingData.FillArea = sData.FillArea
                existingData.CumulativeNetVolume = sData.CumulativeNetVolume
            End If

            'next section material info
            index = index + 1
        Next


        Return True
    End Function

    Private Shared Sub formatPointDatas(ByRef datas As SortedDictionary(Of Double, Slope_PointData), _
                                            ByRef endInfo As Slope_EndInfo) '_
        Try
            Dim last2PointData As Slope_PointData = Nothing
            Dim lastPointData As Slope_PointData = Nothing
            For Each pointData As Slope_PointData In datas.Values
                If Not lastPointData Is Nothing Then
                    lastPointData.mSlope = GetSlopeData(pointData.mOffsetValue, lastPointData.mOffsetValue, _
                                                    pointData.mElevationValue, lastPointData.mElevationValue)
                End If
                last2PointData = lastPointData
                lastPointData = pointData
            Next

            lastPointData.mSlope = ""
            endInfo = GetEndColumInfo(last2PointData.mOffsetValue, last2PointData.mElevationValue, _
                                      lastPointData.mOffsetValue, lastPointData.mElevationValue)
        Catch ex As Exception
            Diagnostics.Debug.Assert(False, ex.Message)
            endInfo = New Slope_EndInfo
            endInfo.DeltaOffset = ""
            endInfo.EndType = ""
            endInfo.EndSlope = ""
        End Try
    End Sub

    Private Shared Function GetSlopeData(ByVal curOff As Double, ByVal lastOff As Double, _
                                ByVal curElev As Double, ByVal lastElev As Double) As String
        Dim deltaX As Double
        Dim deltaY As Double
        Dim slope As Double
        Const percent2Slope As Double = 0.5
        deltaX = Math.Abs(curOff - lastOff)
        deltaY = curElev - lastElev

        If (deltaX <> 0) Then
            slope = deltaY / deltaX
            Dim roundedSlope As Double
            roundedSlope = Math.Floor(Math.Abs(slope) * 10000 + 0.5) / 10000
            If (roundedSlope < percent2Slope) Then
                GetSlopeData = slope.ToString("p")
            Else
                GetSlopeData = "1:" + slope.ToString("0.00")
            End If
        Else
            GetSlopeData = "0.00"
        End If

    End Function

    Private Shared Function GetEndColumInfo(ByVal last2Off As Double, ByVal last2Elev As Double, _
                                    ByVal lastOff As Double, ByVal lastElev As Double) As Slope_EndInfo
        Dim endInfo As New Slope_EndInfo

        Dim deltaY As Double
        deltaY = lastElev - last2Elev

        If deltaY < 0 Then
            endInfo.EndType = LocalizedRes.CorridorSlopeStake_Html_F + _
                Math.Abs(deltaY).ToString("0.00")
        Else
            endInfo.EndType = LocalizedRes.CorridorSlopeStake_Html_C + _
                deltaY.ToString("0.00")
        End If
        endInfo.DeltaOffset = LocalizedRes.CorridorSlopeStake_Html_At + _
            Math.Abs(lastOff - last2Off).ToString("0.00")
        endInfo.EndSlope = GetSlopeData(lastOff, last2Off, lastElev, last2Elev)

        Return endInfo
    End Function

End Class
