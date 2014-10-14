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
Imports Autodesk.AECC.Interop.Roadway
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports LocalizedRes = Report.My.Resources.LocalizedStrings
Imports Autodesk.AECC.Interop

Friend Class ExistingPP_ExtractData

    Public Const nStationIndex As Integer = 0
    Public Const nElevationIndex As Integer = 1
    Public Const nElevationDesignIndex As Integer = 2
    Public Const nElevationDifferenceIndex As Integer = 3
    Public Const nReferenceIndex As Integer = 4
    'Public Const nGradeOutIndex = 4
    'Public Const nVertAlignElevationIndex = 5
    'Public Const nCurveInfo = 6
    Public Const nNorthingIndex As Integer = 5
    Public Const nEastingIndex As Integer = 6
    'Public Const nLastIndex = nVertAlignElevationIndex
    Public Const nLastIndex As Integer = 6

    Public Const nCurvTypeIndex As Integer = 0
    Public Const nPVCStationIndex As Integer = 1
    Public Const nPVCElevationIndex As Integer = 2
    Public Const nPVTStationIndex As Integer = 3
    Public Const nPVTElevationIndex As Integer = 4
    Public Const nHighStationIndex As Integer = 5
    Public Const nHightElevation As Integer = 6
    Public Const nLowStationIndex As Integer = 7
    Public Const nLowElevation As Integer = 8
    Public Const nGradeInIndex As Integer = 9
    Public Const nGradeChange As Integer = 10
    Public Const nCurveLenIndex As Integer = 11
    Public Const nKIndex As Integer = 12
    Public Const nHeadlightDisIndex As Integer = 13
    Public Const nPassingDisIndex As Integer = 13
    Public Const nStoppingDisIndex As Integer = 14
    Public Const nRadius As Integer = 15
    Public Const nLastCurInfoIndex As Integer = nRadius

    Public Shared m_oProfileDataArr()

    Private Shared Function FormatElevSettings(ByVal dDis As Double, Optional ByVal unitIndicator As Boolean = True) As String
        Dim oElevSettings As Land.AeccSettingsElevation
        oElevSettings = ReportApplication.AeccXDatabase.Settings.ProfileSettings.AmbientSettings.ElevationSettings
        FormatElevSettings = ReportFormat.FormatDistance(dDis, oElevSettings.Unit.Value, oElevSettings.Precision.Value, _
                    oElevSettings.Rounding.Value, oElevSettings.Sign.Value, unitIndicator)
    End Function

    Private Shared Function FormatCoordSettings(ByVal dDis As Double) As String
        Dim oCoordSettings As AeccLandLib.AeccSettingsCoordinate
        oCoordSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.CoordinateSettings

        FormatCoordSettings = ReportFormat.FormatSouradnice(dDis, oCoordSettings.Unit.Value, oCoordSettings.Precision.Value, _
                    oCoordSettings.Rounding.Value, oCoordSettings.Sign.Value)
    End Function

    Private Shared Function FormatDistSettings(ByVal dDis As Double, Optional ByVal unitIndicator As Boolean = True) As String
        Dim oDistSettings As Land.AeccSettingsDistance
        oDistSettings = ReportApplication.AeccXDatabase.Settings.ProfileSettings.AmbientSettings.DistanceSettings

        FormatDistSettings = ReportFormat.FormatDistance(dDis, oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                    oDistSettings.Rounding.Value, oDistSettings.Sign.Value, unitIndicator)
    End Function

    Public Shared Function ExtractData(ByVal oAlignment As AeccLandLib.AeccAlignment, ByVal oProfileDesign As AeccLandLib.AeccProfile, ByVal oProfileExisting As AeccLandLib.AeccProfile, ByVal stationStart As Double, ByVal stationEnd As Double, ByVal stationInterval As Double) As Boolean
        Dim oDataArr(nLastIndex)
        Dim iCount As Integer
        Dim oPVI As AeccLandLib.AeccProfilePVI = Nothing
        Dim oPVIC As AeccLandLib.AeccProfilePVICurve
        Dim nCur As Integer
        Dim staniceni As String
        Dim Stan_dbl As Double

        Dim curStation As Double
        Dim curNorthing As Double
        Dim curEasting As Double
        Dim curTangent As Double
        Dim curElevation As Double

        Dim cStations As AeccLandLib.AeccAlignmentStations
        Dim oStation As AeccLandLib.AeccAlignmentStation

        Dim dStation As Double
        Dim StartStation As Double
        Dim Vx As Object

        On Error Resume Next

        nCur = 0

        If InStr(1, oProfileDesign.Alignment.EndingStation, ",", vbTextCompare) > 0 Then
            ReportForm_ExistingProfilePoints.m_bDecimalSeparator = True
        Else
            ReportForm_ExistingProfilePoints.m_bDecimalSeparator = False
        End If

        iCount = 0

        ' ------ REGULAR interval -------
        '    Dim g_staInc As Long
        '    g_staInc = 10
        If ReportForm_ExistingProfilePoints.m_bRegularIntervalCheck = False Then stationInterval = 100000
        cStations = oAlignment.GetStations(AeccLandLib.AeccStationType.aeccMajor, stationInterval, stationInterval / 10)
        Dim cStationCollection As New Collection
        For Each oStation In cStations
            If oStation.Station >= stationStart And oStation.Station <= stationEnd Or Math.Abs(oStation.Station - stationEnd) < 0.011 Then
                Dim tmp1(0 To 1) As Object
                Dim tmp2 As String
                tmp1(0) = oStation.Station
                If oStation.Station = stationStart Then
                    '                    tmp1(1) = "Start"
                    '                    tmp2 = "Start"
                    tmp1(1) = LocalizedRes.ExistingProfilePoints_AnnotationStart
                    tmp2 = LocalizedRes.ExistingProfilePoints_AnnotationStart
                ElseIf oStation.Station = stationEnd Or Math.Abs(oStation.Station - stationEnd) < 0.011 Then
                    '                    tmp1(1) = "End"
                    '                    tmp2 = "End"
                    tmp1(1) = LocalizedRes.ExistingProfilePoints_AnnotationEnd
                    tmp2 = LocalizedRes.ExistingProfilePoints_AnnotationEnd

                Else
                    '                    tmp1(1) = "Regular"
                    '                    tmp2 = "Regular"
                    tmp1(1) = LocalizedRes.ExistingProfilePoints_Regular
                    tmp2 = LocalizedRes.ExistingProfilePoints_Regular
                End If

                cStationCollection.Add(tmp1, CStr(oStation.Station))
                SortStations(oPVI.Station, tmp2, cStationCollection)
                'SortStations(oStation.Station, tmp2, cStationCollection)
            End If
        Next
        '    End If

        ' ------ HTPS points -------
        If ReportForm_ExistingProfilePoints.m_bHTPSCheck = True Then
            cStations = oAlignment.GetStations(AeccLandLib.AeccStationType.aeccGeometryPoint, stationInterval, stationInterval / 10)
            For Each oStation In cStations
                If oStation.Station >= stationStart And oStation.Station <= stationEnd Or Math.Abs(oStation.Station - stationEnd) < 0.011 Then
                    SortStations(oStation.Station, GeometryPointString(oStation.GeometryPointType), cStationCollection)
                End If
            Next
        End If



        ' ------ EXISTING ground profile gradient breaks -------

        If ReportForm_ExistingProfilePoints.m_bExistingCheck = True Then
            If oProfileExisting.Type = AeccLandLib.AeccProfileType.aeccExistingGround Then
                For Each oPVI In oProfileExisting.PVIs
                    If oPVI.Station >= stationStart And oPVI.Station <= stationEnd Then
                        '                    SortStations oPVI.Station, "Existing", cStationCollection
                        SortStations(oPVI.Station, LocalizedRes.ExistingProfilePoints_Existing, cStationCollection)
                    End If
                Next
            End If
        End If
        '        Set cStations = oProfileExisting.Alignment.GetStations(aeccGeometryPoint, g_staInc, g_staInc / 10)
        '        For Each oStation In cStations
        '            If oStation.Station >= stationStart And oStation.Station <= stationEnd Then
        '                SortStations oStation.Station, "Existing", cStationCollection
        '            End If
        '        Next



        ' ------ DESIGN surface profile-------
        Dim cPVIs As AeccLandLib.AeccProfilePVIs
        Dim pvi As AeccLandLib.AeccProfilePVI
        Dim pviCurve As AeccLandLib.AeccProfilePVICurve


        If ReportForm_ExistingProfilePoints.m_bVTPSCheck = True Then
            cPVIs = oProfileDesign.PVIs
            For Each pvi In cPVIs
                If pvi.Station >= stationStart And pvi.Station < stationEnd Then
                    SortStations(pvi.Station, LocalizedRes.ExistingProfilePoints_VPI, cStationCollection)
                    pviCurve = pvi
                    If Not pviCurve Is Nothing Then
                        SortStations(pviCurve.BeginStation, LocalizedRes.ExistingProfilePoints_StartVerticalTP, cStationCollection)
                        SortStations(pviCurve.EndStation, LocalizedRes.ExistingProfilePoints_EndStartVerticalTP, cStationCollection)
                        SortStations(pviCurve.HighPointStation, LocalizedRes.ExistingProfilePoints_HighPoint, cStationCollection)
                        SortStations(pviCurve.LowPointStation, LocalizedRes.ExistingProfilePoints_LowPoint, cStationCollection)
                    End If
                End If
            Next
        End If


        '    End If

        ''     -- WEEDING -----
        '    If g_Weeding = True Then
        '        For Each Vx In cStationCollection
        '            dStation = Vx(0)
        '
        '
        '        Next
        '
        '    End If


        ReDim m_oProfileDataArr(cStationCollection.Count)
        nCur = 0
        Dim Xcoord As Double

        For Each Vx In cStationCollection
            dStation = Vx(0)

            If oAlignment.PointLocationEx(dStation, 0.0#, 0.0#, curEasting, curNorthing, curTangent) = AeccLandLib.AeccAlignmentReturnValue.aeccOK Then
                curElevation = oProfileExisting.ElevationAt(dStation)
                'curEasting = oAlignment.GetStationStringWithEquations(dStation)
                'curNorthing = oAlignment.GetStationStringWithEquations(dStation)

                Erase oDataArr
                oDataArr = Array.CreateInstance(GetType(Object), nLastIndex + 1)
                If ReportForm_ExistingProfilePoints.m_bDecimalSeparator = False Then
                    oDataArr(nStationIndex) = ReportUtilities.GetStationStringWithDerived(oProfileExisting.Alignment, dStation)
                Else
                    staniceni = Replace(ReportUtilities.GetStationStringWithDerived(oProfileExisting.Alignment, dStation), ".", ",", 1, 1)
                    oDataArr(nStationIndex) = staniceni
                End If


                oDataArr(nEastingIndex) = FormatCoordSettings(curEasting)
                oDataArr(nNorthingIndex) = FormatCoordSettings(curNorthing)
                oDataArr(nElevationIndex) = FormatElevSettings(oProfileExisting.ElevationAt(dStation))
                oDataArr(nElevationDesignIndex) = FormatElevSettings(oProfileDesign.ElevationAt(dStation))
                oDataArr(nElevationDifferenceIndex) = FormatElevSettings(oProfileExisting.ElevationAt(dStation) - oProfileDesign.ElevationAt(dStation))
                '                oDataArr(nVertAlignElevationIndex) = FormatElevSettings(ReportForm_ProfPVICurve.CmbVertAlign.Text)
                oDataArr(nReferenceIndex) = Vx(1)
                m_oProfileDataArr(nCur) = oDataArr
                nCur = nCur + 1
            End If



        Next

        ExtractData = True
    End Function

    Private Shared Function GeometryPointString(ByVal geoPt As Integer) As String
        Select Case geoPt
            Case Land.AeccGeometryPointType.aeccBegOfAlign
                GeometryPointString = LocalizedRes.ExistingProfilePoints_BeginOfAlignment
            Case Land.AeccGeometryPointType.aeccEndOfAlign
                GeometryPointString = LocalizedRes.ExistingProfilePoints_EndOfAlignment
            Case Land.AeccGeometryPointType.aeccTanTan
                GeometryPointString = LocalizedRes.ExistingProfilePoints_LineLine
            Case Land.AeccGeometryPointType.aeccCurveTan
                GeometryPointString = LocalizedRes.ExistingProfilePoints_CurveLine
            Case Land.AeccGeometryPointType.aeccTanCurve
                GeometryPointString = LocalizedRes.ExistingProfilePoints_LineCurve
            Case Land.AeccGeometryPointType.aeccCurveCompCurve
                GeometryPointString = LocalizedRes.ExistingProfilePoints_CompCurveCurve
            Case Land.AeccGeometryPointType.aeccCurveRevCurve
                GeometryPointString = LocalizedRes.ExistingProfilePoints_RevCurveCurve
            Case Land.AeccGeometryPointType.aeccLineSpiral
                GeometryPointString = LocalizedRes.ExistingProfilePoints_LineTransition
            Case Land.AeccGeometryPointType.aeccSpiralLine
                GeometryPointString = LocalizedRes.ExistingProfilePoints_TransitionLine
            Case Land.AeccGeometryPointType.aeccCurveSpiral
                GeometryPointString = LocalizedRes.ExistingProfilePoints_CurveTransition
            Case Land.AeccGeometryPointType.aeccSpiralCurve
                GeometryPointString = LocalizedRes.ExistingProfilePoints_TransitionCurve
            Case Land.AeccGeometryPointType.aeccSpiralCompSpiral
                GeometryPointString = LocalizedRes.ExistingProfilePoints_CompTransitionTransition
            Case Land.AeccGeometryPointType.aeccSpiralRevSpiral
                GeometryPointString = LocalizedRes.ExistingProfilePoints_RevTransitionTransition
            Case Land.AeccGeometryPointType.aeccPI
                GeometryPointString = LocalizedRes.ExistingProfilePoints_PI
            Case Land.AeccGeometryPointType.aeccCPI
                GeometryPointString = LocalizedRes.ExistingProfilePoints_CPI
            Case Land.AeccGeometryPointType.aeccSPI
                GeometryPointString = LocalizedRes.ExistingProfilePoints_SPI
            Case Else
                GeometryPointString = " "

                '        Case aeccBegOfAlign
                '            GeometryPointString = "Begin of Alignment"
                '        Case aeccEndOfAlign
                '            GeometryPointString = "End of Alignment"
                '        Case aeccTanTan
                '            GeometryPointString = "Line - Line"
                '        Case aeccCurveTan
                '            GeometryPointString = "Curve - Line"
                '        Case aeccTanCurve
                '            GeometryPointString = "Line - Curve"
                '        Case aeccCurveCompCurve
                '            GeometryPointString = "Curve - Curve"
                '        Case aeccCurveRevCurve
                '            GeometryPointString = "Curve - Curve"
                '        Case aeccLineSpiral
                '            GeometryPointString = "Line - Transition"
                '        Case aeccSpiralLine
                '            GeometryPointString = "Transition - Line"
                '        Case aeccCurveSpiral
                '            GeometryPointString = "Curve - Transition"
                '        Case aeccSpiralCurve
                '            GeometryPointString = "Transition - Curve"
                '        Case aeccSpiralCompSpiral
                '            GeometryPointString = "Transition - Transition"
                '        Case aeccSpiralRevSpiral
                '            GeometryPointString = "Transition - Transition"
                '        Case aeccPI
                '            GeometryPointString = "PI"
                '        Case aeccCPI
                '            GeometryPointString = "CPI"
                '        Case aeccSPI
                '            GeometryPointString = "SPI"
                '        Case Else
                '            GeometryPointString = " "
        End Select
    End Function

    Private Shared Sub SortStations(ByVal dStation As Double, ByVal strType As String, ByRef colStations As Collection)
        Dim oStation As Object
        Dim vStation(0 To 1) As Object
        Dim dCurrentStation As Double
        If colStations.Count = 0 Then
            vStation(0) = dStation
            vStation(1) = strType
            colStations.Add(vStation, CStr(dStation))
            Exit Sub
        End If
        For Each oStation In colStations
            dCurrentStation = oStation(0)
            If dCurrentStation = dStation Then
                Exit For
            ElseIf dCurrentStation > dStation Then
                vStation(0) = dStation
                vStation(1) = strType
                colStations.Add(vStation, CStr(dStation), CStr(oStation(0)))
                Exit For
            End If
        Next
    End Sub
End Class
