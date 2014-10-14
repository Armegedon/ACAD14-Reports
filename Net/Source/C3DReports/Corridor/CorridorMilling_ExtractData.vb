' -----------------------------------------------------------------------------
' <copyright file="CorridorMilling_ExtractData.vb" company="Autodesk">
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

Imports LocalizedRes = Report.My.Resources.LocalizedStrings
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports AeccRoadLib = Autodesk.AECC.Interop.Roadway
Imports AcGeom = Autodesk.AutoCAD.Geometry




Friend Class GrindingSectionData
    Public mStartStation As Double
    Public mEndStation As Double
    Public mStartLeftOffset As Double
    Public mStartRightOffset As Double
    Public mEndLeftOffset As Double
    Public mEndRightOffset As Double

End Class

Friend Class CorridorGrindingData
    Public mSectionDataList As List(Of GrindingSectionData)
    Public Function GetTotalArea() As Double

        GetTotalArea = 0.0
    End Function

End Class

Friend Class CorridorMilling_ExtractData
    Public Shared m_dMillWidthTolerance As Double = 0.5
    Public Shared m_dMillHeightTolerance As Double = 0.1
    Private Shared m_oCorridorGrindingData As CorridorGrindingData
    Public Shared m_ctlProgressBar As CtrlProgressBar = Nothing
    Public Shared ReadOnly Property GrindingData() As CorridorGrindingData
        Get
            Return m_oCorridorGrindingData
        End Get
    End Property

    Public Shared Function ExtractData(ByVal oCorridor As AeccRoadLib.IAeccCorridor, ByVal oBaseline As AeccRoadLib.IAeccBaseline, ByVal dInStartStation As Double, ByVal dInEndStation As Double) As Boolean
        If oCorridor Is Nothing Or oBaseline Is Nothing Then
            Return False
        End If

        m_oCorridorGrindingData = New CorridorGrindingData()
        m_oCorridorGrindingData.mSectionDataList = New List(Of GrindingSectionData)

        Dim bInitStat As Boolean = False
        Dim dStartStation As Double
        Dim dEndStation As Double
        Dim dStartLeftOffset As Double
        Dim dStartRightOffset As Double
        Dim dEndLeftOffset As Double
        Dim dEndRightOffset As Double

        Dim dPrevStation As Double = 0.0
        'Dim dLastMaxLeft As Double = 0.0
        'Dim dLastMaxRight As Double = 0.0

        'extract data
        Dim aRegions As AeccRoadLib.IAeccBaselineRegions = oBaseline.BaselineRegions

        Dim nRegionCount As Integer = aRegions.Count
        Dim iRegion As Integer
        For iRegion = 0 To nRegionCount - 1

            Dim oRegion As AeccRoadLib.IAeccBaselineRegion = aRegions.Item(iRegion)

            If Not oRegion.IsProcessed Then
                Continue For
            End If

            Dim adStations() As Double = oRegion.GetSortedStations()

            bInitStat = False

            For Each dStat As Double In adStations

                If dStat < dInStartStation Then
                    Continue For
                End If
                If dStat > dInEndStation Then
                    Exit For
                End If

                ' Step Progress
                If Not m_ctlProgressBar Is Nothing Then
                    m_ctlProgressBar.PerformStep()
                End If

                ' each region
                'For Each region As AeccRoadLib.IAeccBaselineRegion In oBaseline.BaselineRegions
                ' each assembly in region
                'For Each assem As AeccRoadLib.IAeccAppliedAssembly In Region.AppliedAssemblies
                Dim assem As AeccRoadLib.IAeccAppliedAssembly = oBaseline.AppliedAssembly(dStat)
                If assem Is Nothing Then
                    dPrevStation = dStat
                    Continue For
                End If
                Dim shapes As AeccRoadLib.IAeccCalculatedShapes
                ' Get all the "mill"
                shapes = assem.GetShapesByCode("Mill")
                ' ????
                If shapes.Count <= 0 Then
                    dPrevStation = dStat
                    Continue For
                End If

                Dim dMaxLeftOffset As Double
                Dim dMaxRightOffset As Double
                Dim bInitOffset As Boolean
                dMaxLeftOffset = 0.0
                dMaxRightOffset = 0.0
                bInitOffset = False
                For Each shape As AeccRoadLib.IAeccCalculatedShape In shapes
                    Dim dLeftOffset As Double = 0.0
                    Dim dRightOffset As Double = 0.0
                    If GetMaxLeftRightByTol(shape, dLeftOffset, dRightOffset) Then
                        If Not bInitOffset Then
                            dMaxLeftOffset = dLeftOffset
                            dMaxRightOffset = dRightOffset
                            bInitOffset = True
                        Else
                            If dMaxLeftOffset > dLeftOffset Then
                                dMaxLeftOffset = dLeftOffset
                            End If

                            If dMaxRightOffset < dRightOffset Then
                                dMaxRightOffset = dRightOffset
                            End If
                        End If

                    End If

                    'For Each link As AeccRoadLib.IAeccCalculatedLink In shape.CalculatedLinks
                    '    For Each point As AeccRoadLib.IAeccCalculatedPoint In link.CalculatedPoints
                    '        ' Staion-Offset-Elevation -> Soe
                    '        Dim oSoe As Object = point.StationOffsetElevationToSubassembly
                    '        'Dim XYZ As Object = oBaseline.StationOffsetElevationToXYZ(oSoe)
                    '        Dim offset As Double
                    '        offset = oSoe(1)
                    '        If Not bInitOffset Then
                    '            dMaxLeftOffset = offset
                    '            dMaxRightOffset = offset
                    '            bInitOffset = True
                    '        Else
                    '            If dMaxLeftOffset > offset Then
                    '                dMaxLeftOffset = offset
                    '            End If

                    '            If dMaxRightOffset < offset Then
                    '                dMaxRightOffset = offset
                    '            End If
                    '        End If


                    '    Next

                    'Next

                Next


                If Not bInitStat Then
                    dStartStation = dStat
                    dEndStation = dStat
                    dStartLeftOffset = dMaxLeftOffset
                    dStartRightOffset = dMaxRightOffset
                    dEndLeftOffset = dMaxLeftOffset
                    dEndRightOffset = dMaxRightOffset
                    bInitStat = True
                Else
                    ' do more
                    If Math.Abs((dMaxRightOffset - dMaxLeftOffset) - (dStartRightOffset - dStartLeftOffset)) <= m_dMillWidthTolerance Then
                        dEndStation = dStat
                        dEndLeftOffset = dMaxLeftOffset
                        dEndRightOffset = dMaxRightOffset

                        ' handlelast station
                        If dStat = adStations(adStations.Length - 1) And dStartStation <> dEndStation Then
                            Dim sectionData As GrindingSectionData = New GrindingSectionData()
                            sectionData = New GrindingSectionData()
                            sectionData.mStartStation = dStartStation
                            sectionData.mEndStation = dEndStation
                            sectionData.mStartLeftOffset = dStartLeftOffset
                            sectionData.mStartRightOffset = dStartRightOffset
                            sectionData.mEndLeftOffset = dEndLeftOffset
                            sectionData.mEndRightOffset = dEndRightOffset
                            m_oCorridorGrindingData.mSectionDataList.Add(sectionData)
                        End If

                    Else
                        'If m_oCorridorGrindingData.mSectionDataList.Count > 0 Then
                        '    Dim sectionDataCount As Integer = m_oCorridorGrindingData.mSectionDataList.Count
                        '    Dim lastSectionData As GrindingSectionData = m_oCorridorGrindingData.mSectionDataList.Item(sectionDataCount - 1)
                        '    If lastSectionData.mEndStation < dStartStation And lastSectionData.mEndStation = dPrevStation Then ' ???? <> ???
                        '        Dim transitiveSectionData As GrindingSectionData = New GrindingSectionData()
                        '        transitiveSectionData = New GrindingSectionData()
                        '        transitiveSectionData.mStartStation = lastSectionData.mEndStation
                        '        transitiveSectionData.mEndStation = dStartStation
                        '        transitiveSectionData.mStartLeftOffset = lastSectionData.mEndLeftOffset
                        '        transitiveSectionData.mStartRightOffset = lastSectionData.mEndRightOffset
                        '        transitiveSectionData.mEndLeftOffset = dStartLeftOffset
                        '        transitiveSectionData.mEndRightOffset = dStartRightOffset
                        '        m_oCorridorGrindingData.mSectionDataList.Add(transitiveSectionData)
                        '    End If
                        'End If

                        'If dPrevStation = adStations(0) Then
                        '    Dim transitiveSectionData As GrindingSectionData = New GrindingSectionData()
                        '    transitiveSectionData = New GrindingSectionData()
                        '    transitiveSectionData.mStartStation = dPrevStation
                        '    transitiveSectionData.mEndStation = dStat
                        '    transitiveSectionData.mStartLeftOffset = dStartLeftOffset
                        '    transitiveSectionData.mStartRightOffset = dStartRightOffset
                        '    transitiveSectionData.mEndLeftOffset = dMaxLeftOffset
                        '    transitiveSectionData.mEndRightOffset = dMaxRightOffset
                        '    m_oCorridorGrindingData.mSectionDataList.Add(transitiveSectionData)

                        '    'bInitStat = False
                        '    dStartStation = dStat
                        '    dEndStation = dStat
                        '    dStartLeftOffset = dMaxLeftOffset
                        '    dStartRightOffset = dMaxRightOffset
                        '    dEndLeftOffset = dMaxLeftOffset
                        '    dEndRightOffset = dMaxRightOffset
                        '    bInitStat = True
                        'End If

                        If dStartStation < dEndStation Then  ' ???? <> ???
                            Dim grindingSectionData As GrindingSectionData
                            grindingSectionData = New GrindingSectionData()
                            grindingSectionData.mStartStation = dStartStation
                            grindingSectionData.mEndStation = dEndStation
                            grindingSectionData.mStartLeftOffset = dStartLeftOffset
                            grindingSectionData.mStartRightOffset = dStartRightOffset
                            grindingSectionData.mEndLeftOffset = dEndLeftOffset
                            grindingSectionData.mEndRightOffset = dEndRightOffset
                            m_oCorridorGrindingData.mSectionDataList.Add(grindingSectionData)

                            ' Transitive region
                            Dim grindingSectionDataTrans As GrindingSectionData
                            grindingSectionDataTrans = New GrindingSectionData()
                            grindingSectionDataTrans.mStartStation = dEndStation
                            grindingSectionDataTrans.mEndStation = dStat
                            grindingSectionDataTrans.mStartLeftOffset = dEndLeftOffset
                            grindingSectionDataTrans.mStartRightOffset = dEndRightOffset
                            grindingSectionDataTrans.mEndLeftOffset = dMaxLeftOffset
                            grindingSectionDataTrans.mEndRightOffset = dMaxRightOffset
                            m_oCorridorGrindingData.mSectionDataList.Add(grindingSectionDataTrans)

                            'dStartStation = 0.0
                            'dEndStation = 0.0
                            'bInitStat = False
                            dStartStation = dStat
                            dEndStation = dStat
                            dStartLeftOffset = dMaxLeftOffset
                            dStartRightOffset = dMaxRightOffset
                            dEndLeftOffset = dMaxLeftOffset
                            dEndRightOffset = dMaxRightOffset
                            bInitStat = True
                        Else 'If dStartStation = dEndStation Then
                            Dim transitiveSectionData As GrindingSectionData = New GrindingSectionData()
                            transitiveSectionData = New GrindingSectionData()
                            transitiveSectionData.mStartStation = dStartStation
                            transitiveSectionData.mEndStation = dStat
                            transitiveSectionData.mStartLeftOffset = dStartLeftOffset
                            transitiveSectionData.mStartRightOffset = dStartRightOffset
                            transitiveSectionData.mEndLeftOffset = dMaxLeftOffset
                            transitiveSectionData.mEndRightOffset = dMaxRightOffset
                            m_oCorridorGrindingData.mSectionDataList.Add(transitiveSectionData)

                            'bInitStat = False
                            dStartStation = dStat
                            dEndStation = dStat
                            dStartLeftOffset = dMaxLeftOffset
                            dStartRightOffset = dMaxRightOffset
                            dEndLeftOffset = dMaxLeftOffset
                            dEndRightOffset = dMaxRightOffset
                            bInitStat = True
                        End If


                    End If
                End If

                dPrevStation = dStat
                'dLastMaxLeft = dMaxLeftOffset
                'dLastMaxRight = dMaxRightOffset

            Next
        Next
        Return True
    End Function
    Public Shared Function GetMaxLeftRightByTol(ByVal shape As AeccRoadLib.IAeccCalculatedShape, ByRef dMaxLeftOffset As Double, ByRef dMaxRightOffset As Double) As Boolean
        Dim lineEOB As AcGeom.Line3d = Nothing

        ' get the edge of bottom of Mill
        Dim linkEOB As AeccRoadLib.IAeccCalculatedLink = Nothing
        Dim ptSoeBottom1 As Double() = Nothing
        Dim ptSoeBottom2 As Double() = Nothing
        For Each link As AeccRoadLib.IAeccCalculatedLink In shape.CalculatedLinks
            Dim nCode As Integer = link.CorridorCodes.Count
            For index As Integer = 0 To nCode - 1
                Dim code As String = link.CorridorCodes.Item(index)
                If code = "Mill" Then
                    Dim ptSoe1 As Double() = CType(link.CalculatedPoints.Item(0).GetStationOffsetElevationToBaseline(), Double())
                    Dim ptSoe2 As Double() = CType(link.CalculatedPoints.Item(1).GetStationOffsetElevationToBaseline(), Double())
                    Dim dElev1 As Double = ptSoe1(2)
                    Dim dElev2 As Double = ptSoe2(2)
                    If ptSoeBottom1 Is Nothing Then
                        If dElev1 > dElev2 Then
                            ptSoeBottom1 = ptSoe2
                        Else
                            ptSoeBottom1 = ptSoe1
                        End If
                    ElseIf ptSoeBottom2 Is Nothing Then
                        If dElev1 > dElev2 Then
                            ptSoeBottom2 = ptSoe2
                        Else
                            ptSoeBottom2 = ptSoe1
                        End If
                    Else
                        Return False
                    End If
                End If
            Next
            'For Each code As String In link.CorridorCodes
            '    If code = "L2" Then
            '        linkEOB = link
            '    End If

            'Next

        Next

        'linkEOB = link

        'If linkEOB Is Nothing Then
        '    Return False
        'End If

        ' build the lineEOB
        'If linkEOB.CalculatedPoints.Count < 2 Then
        '    Return False
        'End If
        'Dim ptSoeStart As Object = 'linkEOB.CalculatedPoints.Item(0).StationOffsetElevationToSubassembly
        'Dim ptSoeEnd As Object = linkEOB.CalculatedPoints.Item(1).StationOffsetElevationToSubassembly
        If ptSoeBottom1 Is Nothing Or ptSoeBottom2 Is Nothing Then
            Return False
        End If
        Dim ptSoeStart As Double() = ptSoeBottom1 'linkEOB.CalculatedPoints.Item(0).StationOffsetElevationToSubassembly
        Dim ptSoeEnd As Double() = ptSoeBottom2 'linkEOB.CalculatedPoints.Item(1).StationOffsetElevationToSubassembly
        Dim ptEOBStart As AcGeom.Point3d = New AcGeom.Point3d(ptSoeStart(0), ptSoeStart(1), ptSoeStart(2))
        Dim ptEOBEnd As AcGeom.Point3d = New AcGeom.Point3d(ptSoeEnd(0), ptSoeEnd(1), ptSoeEnd(2))
        If ptEOBStart = ptEOBEnd Then
            Return False
        End If

        lineEOB = New AcGeom.Line3d(ptEOBStart, ptEOBEnd)

        Dim bInitProcess As Boolean = False
        dMaxLeftOffset = 0.0
        dMaxRightOffset = 0.0

        For Each link As AeccRoadLib.IAeccCalculatedLink In shape.CalculatedLinks
            Dim bHigher As Boolean = False
            For Each point As AeccRoadLib.IAeccCalculatedPoint In link.CalculatedPoints
                Dim oSoe As Double() = CType(point.GetStationOffsetElevationToBaseline(), Double())
                Dim ptCheck As AcGeom.Point3d = New AcGeom.Point3d(oSoe(0), oSoe(1), oSoe(2))

                Dim dDistance As Double = lineEOB.GetDistanceTo(ptCheck)
                If dDistance >= m_dMillHeightTolerance Then
                    bHigher = True
                    For Each pointDoubleCheck As AeccRoadLib.IAeccCalculatedPoint In link.CalculatedPoints
                        Dim oSoeDoubleCheck As Double() = CType(pointDoubleCheck.GetStationOffsetElevationToBaseline(), Double())
                        If Not bInitProcess Then
                            bInitProcess = True
                            dMaxLeftOffset = oSoeDoubleCheck(1)
                            dMaxRightOffset = oSoeDoubleCheck(1)
                        Else
                            Dim offset As Double = oSoeDoubleCheck(1)
                            If dMaxLeftOffset > offset Then
                                dMaxLeftOffset = offset
                            End If

                            If dMaxRightOffset < offset Then
                                dMaxRightOffset = offset
                            End If
                        End If
                    Next
                End If

                If bHigher Then
                    Exit For
                End If

            Next

        Next

        Return bInitProcess
    End Function

    Public Shared Function FormatDistSettings(ByVal dDis As Double) As String
        Dim oDistSettings As AeccLandLib.IAeccSettingsDistance
        oDistSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.DistanceSettings

        FormatDistSettings = ReportFormat.FormatDistance(dDis, oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                    oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
    End Function

    Public Shared Function FormatAreaSettings(ByVal dArea As Double) As String
        Dim oAreaSettings As AeccLandLib.AeccSettingsArea
        oAreaSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.AreaSettings

        FormatAreaSettings = ReportFormat.FormatArea(dArea, _
                                            oAreaSettings.Unit.Value, _
                                            oAreaSettings.Precision.Value, _
                                            oAreaSettings.Rounding.Value, _
                                            oAreaSettings.Sign.Value)
    End Function

End Class




Friend Class SlopeInfo_SectionData
        Public LeftDatas As New SortedDictionary(Of Double, SlopeInfo_PointData)
        Public LeftEndInfo As SlopeInfo_EndInfo
        Public RightDatas As New SortedDictionary(Of Double, SlopeInfo_PointData)
        Public RightEndInfo As SlopeInfo_EndInfo
        Public CenterLineInfo As New SlopeInfo_PointData
        Public LinksOnGraph As New SortedDictionary(Of Integer, List(Of SlopeInfo_PointData))
        Public EGDatas As New SortedDictionary(Of Integer, List(Of SlopeInfo_PointData))
End Class

Friend Class SlopeInfo_EndInfo
    Public EndType As String
    Public DeltaOffset As String
    Public EndSlope As String
End Class

Friend Class SlopeInfo_PointData

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

        mOffsetString = CorridorSlopeStakeInfo_ExtractData.FormatDistSettings(mOffsetValue, mOffsetRounded)
        mElevationString = CorridorSlopeStakeInfo_ExtractData.FormatElevSettings(mElevationValue, mElevationRounded)
    End Sub
End Class

Friend Class SlopeInfo_PointData_Sorter
    Implements IComparer(Of SlopeInfo_PointData)
    Function Compare(ByVal x As SlopeInfo_PointData, ByVal y As SlopeInfo_PointData) As Integer _
        Implements IComparer(Of SlopeInfo_PointData).Compare
        If x.mOffsetValue < y.mOffsetValue Then
            Return -1
        ElseIf x.mOffsetValue = y.mOffsetValue Then
            Return 0
        Else
            Return 1
        End If
    End Function 'IComparer.Compare
End Class 'SlopeInfo_PointData_Sorter

Friend Class CorridorSlopeStakeInfo_ExtractData

    Private Shared m_oSlopeStakeData As New Dictionary(Of Double, SlopeInfo_SectionData)
    Public Shared ReadOnly Property SlopeStakeData() As Dictionary(Of Double, SlopeInfo_SectionData)
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
                                ByVal sCodeName As String, _
                                ByVal ctlProgress As ProgressBar, _
                                ByVal steps As Integer) As Boolean
        Dim PtSteped As Integer
        Dim oneStepPtCount As Integer
        Dim stepNeeded As Integer
        stepNeeded = calcStepsNeeded(oBaseline, oSampleLineGroups, stationStart, stationEnd, sCodeName)
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
            Dim oLinks As AeccRoadLib.AeccCalculatedLinks
            Dim oLinksOnGraphics As AeccRoadLib.AeccCalculatedLinks
            Dim sData As New SlopeInfo_SectionData

            Try
                datumElevation = oBaseline.Profile.ElevationAt(curStation)
                oLinks = oBaseline.AppliedAssembly(curStation).GetLinksByCode(sCodeName)
                oLinksOnGraphics = oBaseline.AppliedAssembly(curStation).GetLinks()
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
                'error, try next point
                Continue For
            End Try

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
                    Dim ptOffsetElev As Double()
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
                    Dim ptData As New SlopeInfo_PointData
                    ptData.Fill(Codes, ptOffset, ptElev)

                    If Not sData.LinksOnGraph.ContainsKey(idxOfLinksOnGraph) Then
                        sData.LinksOnGraph.Add(idxOfLinksOnGraph, New List(Of SlopeInfo_PointData))
                    End If
                    Dim PtDataList As List(Of SlopeInfo_PointData)
                    PtDataList = sData.LinksOnGraph.Item(idxOfLinksOnGraph)
                    PtDataList.Add(ptData)
                Next
            Next

            ' Add EG section point
            Dim nIdxOfEG As Integer = -1 ' Though we expect only 1 eg, but may be multiple surfaces lying there
            For Each oSection As AeccLandLib.AeccSection In oSampleLine.Sections
                If oSection.DataType = AeccLandLib.AeccSectionDataType.aeccSectionDataGridSurface _
                   Or oSection.DataType = AeccLandLib.AeccSectionDataType.aeccSectionDataTIN Then

                    nIdxOfEG += 1
                    sData.EGDatas.Add(nIdxOfEG, New List(Of SlopeInfo_PointData))
                    Dim ptDataList As List(Of SlopeInfo_PointData)
                    ptDataList = sData.EGDatas(nIdxOfEG)

                    Dim oSectionLinks As AeccLandLib.AeccSectionLinks
                    Dim oSectionLink As AeccLandLib.AeccSectionLink
                    oSectionLinks = oSection.Links

                    Dim nLinksCount As Integer = oSectionLinks.Count
                    For idxOfSecLink As Integer = 0 To nLinksCount - 1
                        ' For Each oSectionLink In oSectionLinks
                        oSectionLink = oSectionLinks.Item(idxOfSecLink)
                        Dim ptEGDataStart As New SlopeInfo_PointData
                        ptEGDataStart.Fill("", oSectionLink.StartPointX, oSectionLink.StartPointY)
                        ptDataList.Add(ptEGDataStart)
                        ' add last end pt
                        If idxOfSecLink = nLinksCount - 1 Then
                            Dim ptEGDataEnd As New SlopeInfo_PointData
                            ptEGDataEnd.Fill("", oSectionLink.EndPointX, oSectionLink.EndPointY)
                            ptDataList.Add(ptEGDataEnd)
                        End If
                    Next
                End If
            Next

            For Each oLink As AeccRoadLib.AeccCalculatedLink In oLinks
                Dim oPoint As AeccRoadLib.AeccCalculatedPoint
                For Each oPoint In oLink.CalculatedPoints
                    Dim ptOffsetElev As Double()
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
                    Dim ptData As New SlopeInfo_PointData
                    ptData.Fill(Codes, ptOffset, ptElev)

                    If ptOffset < 0 Then
                        'has data judge should put outside If() statement, for data may added
                        If Not sData.LeftDatas.ContainsKey(Math.Abs(ptData.mOffsetRounded)) Then
                            sData.LeftDatas.Add(Math.Abs(ptData.mOffsetRounded), ptData)
                        Else
                            'check point code
                            If ptData.mCodes <> "" Then
                                Dim existPt As SlopeInfo_PointData
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
                                Dim existPt As SlopeInfo_PointData
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

            formatPointDatas(sData.LeftDatas, sData.LeftEndInfo)
            formatPointDatas(sData.RightDatas, sData.RightEndInfo)

            If Not m_oSlopeStakeData.ContainsKey(curStationRounded) Then
                If sData.LeftDatas.Count > 0 Or sData.RightDatas.Count > 0 Then
                    m_oSlopeStakeData.Add(curStationRounded, sData)
                End If
            End If
        Next

        Return True
    End Function

    Private Shared Sub formatPointDatas(ByRef datas As SortedDictionary(Of Double, SlopeInfo_PointData), _
                                            ByRef endInfo As SlopeInfo_EndInfo) '_
        Try
            Dim last2PointData As SlopeInfo_PointData = Nothing
            Dim lastPointData As SlopeInfo_PointData = Nothing
            For Each pointData As SlopeInfo_PointData In datas.Values
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
            endInfo = New SlopeInfo_EndInfo
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
                                    ByVal lastOff As Double, ByVal lastElev As Double) As SlopeInfo_EndInfo
        Dim endInfo As New SlopeInfo_EndInfo

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
