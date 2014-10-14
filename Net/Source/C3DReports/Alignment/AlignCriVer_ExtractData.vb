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
Imports System.Math
Imports AutoCAD = Autodesk.AutoCAD.Interop
Imports AeccUiLandLib = Autodesk.AECC.Interop.UiLand
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AEC.Interop.UIBase
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Friend Class AlignCriVer_ExtractData
    Private Shared m_oAlignDataArr()

    Public Const nTypeIndex = 0
    Public Const nLengthIndex = 1
    Public Const nCourseIndex = 2
    Public Const nChordIndex = 3
    Public Const nRadiusIndex = 4
    Public Const nDeltaIndex = 5
    Public Const nThetaIndex = 6
    Public Const nDOCIndex = 7
    Public Const nArcTypeIndex = 8
    Public Const nTangentIndex = 9
    Public Const nMidOrdIndex = 10
    Public Const nExternalIndex = 11
    Public Const nLTanIndex = 12
    Public Const nSTanIndex = 13
    Public Const nPValueIndex = 14
    Public Const nKValueIndex = 15
    Public Const nAValueIndex = 16
    Public Const nXValueIndex = 17
    Public Const nYValueIndex = 18
    Public Const nSpiralTypeIndex = 19
    Public Const nDesignSpeedIndex = 20
    Public Const nMinRadiusIndex = 21
    Public Const nPrevElementIndex = 22
    Public Const nPrevStationIndex = 23
    Public Const nPrevNorthingIndex = 24
    Public Const nPrevEastingIndex = 25
    Public Const nCurrElementIndex = 26
    Public Const nCurrStationIndex = 27
    Public Const nCurrNorthingIndex = 28
    Public Const nCurrEastingIndex = 29
    Public Const nNextElementIndex = 30
    Public Const nNextStationIndex = 31
    Public Const nNextNorthingIndex = 32
    Public Const nNextEastingIndex = 33
    Public Const nMinTransLenIndex = 34
    Public Const nDesignCheckNamesIndex = 35 ' Array of names
    Public Const nDesignCheckResultsIndex = 36 'Array of Results
    Public Const nMinTransLenResultIndex = 37
    Public Const nMinRadiusResultIndex = 38
    Public Const nEntityPrefixIndex = 39
    Public Const nPrevDoubleStationIndex = 40
    Public Const nGroupIndex = 41
    Public Const nCompoundSpiralIndex = 42
    Public Const nLastIndex = nCompoundSpiralIndex

    Public Shared ReadOnly Property AlignDataArr()
        Get
            Return m_oAlignDataArr
        End Get
    End Property
    Private Shared Function FormatCoordSettings(ByVal dDis As Double) As String
        Dim oCoordSettings As AeccLandLib.AeccSettingsCoordinate
        oCoordSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.CoordinateSettings

        FormatCoordSettings = ReportFormat.FormatDistance(dDis, oCoordSettings.Unit.Value, oCoordSettings.Precision.Value, _
                    oCoordSettings.Rounding.Value, oCoordSettings.Sign.Value)
    End Function

    Private Shared Function FormatDistSettings(ByVal dDis As Double) As String
        Dim oDistSettings As AeccLandLib.AeccSettingsDistance
        oDistSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.DistanceSettings

        FormatDistSettings = ReportFormat.FormatDistance(dDis, oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                    oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
    End Function

    Private Shared Function FormatDirSettings(ByVal dDirection As Double) As String
        Dim oDirSettings As AeccLandLib.AeccSettingsDirection
        oDirSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.DirectionSettings
        FormatDirSettings = ReportFormat.FormatDirection(dDirection, oDirSettings.Unit.Value, oDirSettings.Precision.Value, _
                                            oDirSettings.Rounding.Value, oDirSettings.Format.Value, _
                                            oDirSettings.Direction.Value, oDirSettings.Capitalization.Value, _
                                            oDirSettings.Sign.Value, oDirSettings.MeasurementType.Value, oDirSettings.BearingQuadrant.Value)
    End Function
    Private Shared Function FormatAngleSettings(ByVal dAngle As Double) As String
        Dim oAngleSettings As AeccLandLib.AeccSettingsAngle
        oAngleSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.AngleSettings

        FormatAngleSettings = ReportFormat.FormatAngle(dAngle, oAngleSettings.Unit.Value, oAngleSettings.Precision.Value, _
                                         oAngleSettings.Rounding.Value, oAngleSettings.Format.Value, oAngleSettings.Sign.Value)
    End Function

    Public Shared Function ExtractData(ByVal oAlignment As AeccLandLib.AeccAlignment, _
                                ByVal stationStart As Double, _
                                ByVal stationEnd As Double) As Boolean
        Dim oDataArr(nLastIndex) As Object
        Dim oAlignEnt As AeccLandLib.AeccAlignmentEntity
        Dim oAlignLine As AeccLandLib.AeccAlignmentTangent
        Dim oAlignArc As AeccLandLib.AeccAlignmentArc
        Dim oAlignSpiral As AeccLandLib.AeccAlignmentSpiral
        Dim oAlignCS As AeccLandLib.AeccAlignmentCSGroup
        Dim oAlignSC As AeccLandLib.AeccAlignmentSCGroup
        Dim oAlignSS As AeccLandLib.AeccAlignmentSSGroup
        Dim oAlignSCS As AeccLandLib.AeccAlignmentSCSGroup
        Dim oAlignSTS As AeccLandLib.AeccAlignmentSTSGroup
        Dim oAlignST As AeccLandLib.AeccAlignmentSTGroup
        Dim oAlignTS As AeccLandLib.AeccAlignmentTSGroup
        Dim oAlignSSC As AeccLandLib.AeccAlignmentSSCGroup
        Dim oAlignCSS As AeccLandLib.AeccAlignmentCSSGroup
        Dim oAlignSCSCS As AeccLandLib.AeccAlignmentSCSCSGroup
        Dim oAlignSCSSCS As AeccLandLib.AeccAlignmentSCSSCSGroup
        Dim oAlignEnts As AeccLandLib.AeccAlignmentEntities
        oAlignEnts = oAlignment.Entities
        Dim nCur As Integer
        Dim nGroupId As Integer
        nCur = 0
        nGroupId = 0
        ' clac the count of basic element
        Dim nCount As Integer
        nCount = 0
        For Each oAlignEnt In oAlignEnts
            On Error Resume Next
            If oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccTangent Then
                nCount = nCount + 1
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccArc Then
                nCount = nCount + 1
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiral Then
                nCount = nCount + 1
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccCurveSpiralGroup Then
                nCount = nCount + 2
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralCurveGroup Then
                nCount = nCount + 2
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralSpiralGroup Then
                nCount = nCount + 2
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralCurveSpiralGroup Then
                nCount = nCount + 3
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralTangentSpiralGroup Then
                nCount = nCount + 3
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralTangentGroup Then
                nCount = nCount + 2
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccTangentSpiralGroup Then
                nCount = nCount + 2
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralSpiralCurveGroup Then
                nCount = nCount + 3
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccCurveSpiralSpiralGroup Then
                nCount = nCount + 3
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralCurveSpiralCurveSpiralGroup Then
                nCount = nCount + 5
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralCurveSpiralSpiralCurveSpiralGroup Then
                nCount = nCount + 6
            End If
        Next
        ReDim m_oAlignDataArr(nCount)
        For Each oAlignEnt In oAlignEnts
            'On Error Resume Next
            Erase oDataArr
            ReDim oDataArr(nLastIndex)
            oAlignLine = Nothing
            oAlignArc = Nothing
            oAlignSpiral = Nothing
            oAlignCS = Nothing
            oAlignSC = Nothing
            oAlignSCS = Nothing
            oAlignSS = Nothing
            oAlignSTS = Nothing
            oAlignST = Nothing
            oAlignTS = Nothing
            oAlignSSC = Nothing
            oAlignCSS = Nothing
            oAlignSCSCS = Nothing
            oAlignSCSSCS = Nothing
            If oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccTangent Then
                oAlignLine = oAlignEnt
                If Not oAlignLine Is Nothing Then
                    oDataArr(nTypeIndex) = oAlignLine.Type
                    m_oAlignDataArr(nCur) = oDataArr
                    Call ExtractTangentData(oAlignLine, nCur)
                    Call ExtractTangentStationData(oAlignment, oAlignLine, nCur)
                    Call ExtractTangentDesignCheckData(oAlignLine, oAlignLine.DesignChecks, nCur)
                    oDataArr(nEntityPrefixIndex) = 0
                    oDataArr(nGroupIndex) = nGroupId
                    m_oAlignDataArr(nCur) = oDataArr
                End If
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccArc Then
                oAlignArc = oAlignEnt
                If Not oAlignArc Is Nothing Then
                    oDataArr(nTypeIndex) = oAlignArc.Type
                    m_oAlignDataArr(nCur) = oDataArr
                    Call ExtractArcData(oAlignArc, nCur)
                    Call ExtractArcStationData(oAlignment, oAlignArc, nCur)
                    Call ExtractArcDesignCheckData(oAlignArc, oAlignArc.DesignChecks, nCur)
                    oDataArr(nGroupIndex) = nGroupId
                    m_oAlignDataArr(nCur)(nEntityPrefixIndex) = 0
                End If
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiral Then
                oAlignSpiral = oAlignEnt
                If Not oAlignSpiral Is Nothing Then
                    oDataArr(nTypeIndex) = oAlignSpiral.Type
                    m_oAlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralData(oAlignSpiral, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSpiral, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSpiral, oAlignSpiral.DesignChecks, nCur)
                    oDataArr(nEntityPrefixIndex) = 0
                    oDataArr(nGroupIndex) = nGroupId
                    m_oAlignDataArr(nCur) = oDataArr
                End If
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccCurveSpiralGroup Then
                oAlignCS = oAlignEnt
                If Not oAlignCS Is Nothing Then
                    oDataArr(nTypeIndex) = oAlignCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractArcDataEx(oAlignCS.ArcIn, nCur)
                    Call ExtractArcStationData(oAlignment, oAlignCS.ArcIn, nCur)
                    Call ExtractGroupDesignCheckData(oAlignCS.ArcIn, oAlignCS, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 1
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignCS.SpiralOut, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignCS.SpiralOut, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignCS.SpiralOut, oAlignCS.SpiralOut.DesignChecks, nCur)
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 2
                End If
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralCurveGroup Then
                oAlignSC = oAlignEnt
                If Not oAlignSC Is Nothing Then
                    oDataArr(nTypeIndex) = oAlignSC.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSC.SpiralIn, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSC.SpiralIn, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSC.SpiralIn, oAlignSC.SpiralIn.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 1
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSC.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractArcDataEx(oAlignSC.ArcOut, nCur)
                    Call ExtractArcStationData(oAlignment, oAlignSC.ArcOut, nCur)
                    Call ExtractGroupDesignCheckData(oAlignSC.ArcOut, oAlignSC, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 2
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                End If
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralSpiralGroup Then
                oAlignSC = oAlignEnt
                If Not oAlignSC Is Nothing Then
                    oDataArr(nTypeIndex) = oAlignSC.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSS.Spiral1, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSS.Spiral1, nCur)
                    Call ExtractGroupDesignCheckData(oAlignSS.Spiral1, oAlignSS, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 1
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSS.Spiral2, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSS.Spiral2, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSS.Spiral2, oAlignSS.Spiral2.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 2
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                End If
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralCurveSpiralGroup Then
                oAlignSCS = oAlignEnt
                If Not oAlignSCS Is Nothing Then
                    oDataArr(nTypeIndex) = oAlignSCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSCS.SpiralIn, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSCS.SpiralIn, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSCS.SpiralIn, oAlignSCS.SpiralIn.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 1
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractArcDataEx(oAlignSCS.Arc, nCur)
                    Call ExtractGroupDesignCheckData(oAlignSCS.Arc, oAlignSCS, nCur)
                    Call ExtractArcStationData(oAlignment, oAlignSCS.Arc, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 2
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Dim temp = oAlignSCS.SpiralOut
                    Call ExtractSpiralDataEx(oAlignSCS.SpiralOut, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSCS.SpiralOut, oAlignSCS.SpiralOut.DesignChecks, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSCS.SpiralOut, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 3
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                End If
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralTangentSpiralGroup Then
                oAlignSTS = oAlignEnt
                If Not oAlignSTS Is Nothing Then
                    oDataArr(nTypeIndex) = oAlignSTS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSTS.SpiralIn, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSTS.SpiralIn, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSTS.SpiralIn, oAlignSTS.SpiralIn.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 1
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSTS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractTangentData(oAlignSTS.Tangent, nCur)
                    Call ExtractTangentStationData(oAlignment, oAlignSTS.Tangent, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 2
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSTS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSTS.SpiralOut, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSTS.SpiralOut, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSTS.SpiralOut, oAlignSTS.SpiralOut.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 3
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                End If
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralTangentGroup Then
                oAlignST = oAlignEnt
                If Not oAlignST Is Nothing Then
                    oDataArr(nTypeIndex) = oAlignST.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignST.SpiralIn, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignST.SpiralIn, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignST.SpiralIn, oAlignST.SpiralIn.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 1
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignST.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractTangentData(oAlignST.TangentOut, nCur)
                    Call ExtractTangentStationData(oAlignment, oAlignST.TangentOut, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 2
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                End If
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccTangentSpiralGroup Then
                oAlignTS = oAlignEnt
                If Not oAlignTS Is Nothing Then
                    oDataArr(nTypeIndex) = oAlignTS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractTangentData(oAlignTS.TangentIn, nCur)
                    Call ExtractTangentStationData(oAlignment, oAlignTS.TangentIn, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 1
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignTS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignTS.SpiralOut, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignTS.SpiralOut, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignTS.SpiralOut, oAlignTS.SpiralOut.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 2
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                End If
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralSpiralCurveGroup Then
                oAlignSSC = oAlignEnt
                If Not oAlignSSC Is Nothing Then
                    oDataArr(nTypeIndex) = oAlignSSC.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSSC.Spiral1, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSSC.Spiral1, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSSC.Spiral1, oAlignSSC.Spiral1.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 1
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSSC.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSSC.Spiral2, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSSC.Spiral2, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSSC.Spiral2, oAlignSSC.Spiral2.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 2
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSSC.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractArcDataEx(oAlignSSC.Arc, nCur)
                    Call ExtractArcStationData(oAlignment, oAlignSSC.Arc, nCur)
                    Call ExtractArcDesignCheckData(oAlignSSC.Arc, oAlignSSC.Arc.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 3
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                End If
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccCurveSpiralSpiralGroup Then
                oAlignCSS = oAlignEnt
                If Not oAlignCSS Is Nothing Then
                    oDataArr(nTypeIndex) = oAlignCSS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractArcDataEx(oAlignCSS.Arc, nCur)
                    Call ExtractArcStationData(oAlignment, oAlignCSS.Arc, nCur)
                    Call ExtractArcDesignCheckData(oAlignCSS.Arc, oAlignCSS.Arc.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 1
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignCSS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignCSS.Spiral1, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignCSS.Spiral1, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignCSS.Spiral1, oAlignCSS.Spiral1.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 2
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignCSS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignCSS.Spiral2, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignCSS.Spiral2, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignCSS.Spiral2, oAlignCSS.Spiral2.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 3
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                End If
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralCurveSpiralCurveSpiralGroup Then
                oAlignSCSCS = oAlignEnt
                If Not oAlignSCSCS Is Nothing Then
                    oDataArr(nTypeIndex) = oAlignSCSCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSCSCS.Spiral1, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSCSCS.Spiral1, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSCSCS.Spiral1, oAlignSCSCS.Spiral1.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 1
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSCSCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractArcDataEx(oAlignSCSCS.Arc1, nCur)
                    Call ExtractArcStationData(oAlignment, oAlignSCSCS.Arc1, nCur)
                    Call ExtractArcDesignCheckData(oAlignSCSCS.Arc1, oAlignSCSCS.Arc1.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 2
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSCSCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSCSCS.Spiral2, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSCSCS.Spiral2, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSCSCS.Spiral2, oAlignSCSCS.Spiral2.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 3
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSCSCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractArcDataEx(oAlignSCSCS.Arc2, nCur)
                    Call ExtractArcStationData(oAlignment, oAlignSCSCS.Arc2, nCur)
                    Call ExtractArcDesignCheckData(oAlignSCSCS.Arc2, oAlignSCSCS.Arc2.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 4
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSCSCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSCSCS.Spiral3, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSCSCS.Spiral3, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSCSCS.Spiral3, oAlignSCSCS.Spiral3.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 5
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                End If
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralCurveSpiralSpiralCurveSpiralGroup Then
                oAlignSCSSCS = oAlignEnt
                If Not oAlignSCSSCS Is Nothing Then
                    oDataArr(nTypeIndex) = oAlignSCSSCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSCSSCS.Spiral1, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSCSSCS.Spiral1, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSCSSCS.Spiral1, oAlignSCSSCS.Spiral1.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 1
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSCSSCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractArcDataEx(oAlignSCSSCS.Arc1, nCur)
                    Call ExtractArcStationData(oAlignment, oAlignSCSSCS.Arc1, nCur)
                    Call ExtractArcDesignCheckData(oAlignSCSSCS.Arc1, oAlignSCSSCS.Arc1.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 2
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSCSSCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSCSSCS.Spiral2, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSCSSCS.Spiral2, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSCSSCS.Spiral2, oAlignSCSSCS.Spiral2.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 3
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSCSSCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSCSSCS.Spiral3, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSCSSCS.Spiral3, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSCSSCS.Spiral3, oAlignSCSSCS.Spiral3.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 4
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSCSSCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractArcDataEx(oAlignSCSSCS.Arc2, nCur)
                    Call ExtractArcStationData(oAlignment, oAlignSCSSCS.Arc2, nCur)
                    Call ExtractArcDesignCheckData(oAlignSCSSCS.Arc2, oAlignSCSSCS.Arc2.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 5
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                    nCur = nCur + 1
                    ReDim oDataArr(nLastIndex)
                    oDataArr(nTypeIndex) = oAlignSCSSCS.Type
                    AlignDataArr(nCur) = oDataArr
                    Call ExtractSpiralDataEx(oAlignSCSSCS.Spiral4, nCur)
                    Call ExtractSpiralStationData(oAlignment, oAlignSCSSCS.Spiral4, nCur)
                    Call ExtractSpiralDesignCheckData(oAlignSCSSCS.Spiral4, oAlignSCSSCS.Spiral4.DesignChecks, nCur)
                    AlignDataArr(nCur)(nEntityPrefixIndex) = 6
                    AlignDataArr(nCur)(nGroupIndex) = nGroupId
                End If
            End If
            nCur = nCur + 1
            nGroupId = nGroupId + 1
        Next
        Call SortData()
        ExtractData = True
    End Function

    Public Shared Function SortData() As Boolean
        Dim count As Integer
        Dim i As Integer
        Dim j As Integer
        Dim tempDataArr()
        count = 0
        j = 0
        For i = 0 To UBound(AlignDataArr) - 1
            If AlignDataArr(i)(nEntityPrefixIndex) <= 1 Then
                count = count + 1
            End If
        Next
        ReDim tempDataArr(count)
        For i = 0 To UBound(AlignDataArr) - 1
            If AlignDataArr(i)(nEntityPrefixIndex) <= 1 Then
                tempDataArr(j) = AlignDataArr(i)
                j = j + 1
            End If
        Next

        'sort
        Dim k As Integer
        Dim Temp
        Dim Flag As Boolean
        Dim n As Integer = UBound(tempDataArr) - 1
        For k = 1 To n
            Flag = False
            For j = 0 To n - k
                If tempDataArr(j)(nPrevDoubleStationIndex) > tempDataArr(j + 1)(nPrevDoubleStationIndex) Then
                    Temp = tempDataArr(j)
                    tempDataArr(j) = tempDataArr(j + 1)
                    tempDataArr(j + 1) = Temp
                    Flag = True
                End If
            Next
            If Flag = False Then
                Exit For
            End If
        Next
        i = 0
        ReDim Temp(UBound(AlignDataArr))
        For k = 0 To UBound(tempDataArr) - 1
            If tempDataArr(k)(nEntityPrefixIndex) = 0 Then
                Temp(i) = tempDataArr(k)
                i = i + 1
            Else
                For j = 0 To UBound(AlignDataArr) - 1
                    If AlignDataArr(j)(nGroupIndex) = tempDataArr(k)(nGroupIndex) Then
                        Temp(i) = AlignDataArr(j)
                        i = i + 1
                    End If
                Next
            End If
        Next
        For i = 0 To UBound(Temp) - 1
            AlignDataArr(i) = Temp(i)
        Next
        SortData = True
    End Function


    Private Shared Function ExtractTangentData(ByVal oAlignLine As AeccLandLib.AeccAlignmentTangent, ByVal nIndexDataArray As Integer) As Boolean

        Dim oDataArr As Object
        oDataArr = m_oAlignDataArr(nIndexDataArray)
        If Not oAlignLine Is Nothing Then
            oDataArr(nTypeIndex) = AeccLandLib.AeccAlignmentEntityType.aeccTangent
            oDataArr(nLengthIndex) = FormatDistSettings(oAlignLine.Length)
            Try
                Dim temp As Integer
                temp = oAlignLine.HighestDesignSpeed
                oDataArr(nDesignSpeedIndex) = temp
            Catch
            End Try
            m_oAlignDataArr(nIndexDataArray) = oDataArr
        End If
        ExtractTangentData = True
    End Function

    Private Shared Function ExtractTangentStationData(ByVal oAlignment As AeccLandLib.AeccAlignment, ByVal oAlignLine As AeccLandLib.AeccAlignmentTangent, ByVal nIndexDataArray As Integer) As Boolean

        Dim oDataArr As Object
        oDataArr = m_oAlignDataArr(nIndexDataArray)
        If Not oAlignLine Is Nothing Then
            oDataArr(nPrevElementIndex) = LocalizedRes.AlignCriVer_Html_Start
            oDataArr(nPrevDoubleStationIndex) = oAlignLine.StartingStation
            oDataArr(nPrevStationIndex) = ReportUtilities.GetStationStringWithDerived(oAlignment, oAlignLine.StartingStation)
            oDataArr(nNextElementIndex) = LocalizedRes.AlignCriVer_Html_End
            oDataArr(nNextStationIndex) = ReportUtilities.GetStationStringWithDerived(oAlignment, oAlignLine.EndingStation)
            m_oAlignDataArr(nIndexDataArray) = oDataArr
        End If

        ExtractTangentStationData = True
    End Function

    Private Shared Function ExtractArcData(ByVal oAlignArc As AeccLandLib.AeccAlignmentArc, ByVal nIndexDataArray As Integer) As Boolean

        Dim oDataArr As Object

        On Error Resume Next
        oDataArr = m_oAlignDataArr(nIndexDataArray)
        If Not oAlignArc Is Nothing Then
            oDataArr(nTypeIndex) = AeccLandLib.AeccAlignmentEntityType.aeccArc
            If oAlignArc.Clockwise = True Then
                oDataArr(nArcTypeIndex) = LocalizedRes.AlignCriVer_Html_Right
            Else
                oDataArr(nArcTypeIndex) = LocalizedRes.AlignCriVer_Html_Left
            End If
            oDataArr(nRadiusIndex) = FormatDistSettings(oAlignArc.Radius)
            oDataArr(nLengthIndex) = FormatDistSettings(oAlignArc.Length)
            oDataArr(nMinRadiusIndex) = oAlignArc.MinimumRadius
            If Not oDataArr(nMinRadiusIndex) Is Nothing Then
                'compare radius with min radius
                Dim temp As Double = oAlignArc.MinimumRadius
                oDataArr(nMinRadiusIndex) = temp.ToString("N2")
                If oAlignArc.Radius >= oAlignArc.MinimumRadius Then
                    oDataArr(nMinRadiusResultIndex) = LocalizedRes.AlignCriVer_Html_Align_Cleared
                Else
                    oDataArr(nMinRadiusResultIndex) = "<font color=""#FF0000"">" + LocalizedRes.AlignCriVer_Html_Align_Violated + "</font>"
                End If
            End If
            m_oAlignDataArr(nIndexDataArray) = oDataArr
        End If

        ExtractArcData = True
    End Function

    Public Shared Function ExtractArcDataEx(ByVal oAlignArc As AeccLandLib.AeccAlignmentArc, ByVal nIndexDataArray As Integer) As Boolean

        Dim oDataArr As Object

        On Error Resume Next
        oDataArr = AlignDataArr(nIndexDataArray)
        If Not oAlignArc Is Nothing Then
            oDataArr(nTypeIndex) = AeccLandLib.AeccAlignmentEntityType.aeccArc
            If oAlignArc.Clockwise = True Then
                oDataArr(nArcTypeIndex) = LocalizedRes.AlignCriVer_Html_Right
            Else
                oDataArr(nArcTypeIndex) = LocalizedRes.AlignCriVer_Html_Left
            End If
            oDataArr(nRadiusIndex) = FormatDistSettings(oAlignArc.Radius)
            oDataArr(nLengthIndex) = FormatDistSettings(oAlignArc.Length)
            oDataArr(nMinRadiusIndex) = FormatDistSettings(oAlignArc.MinimumRadius)
            'compare radius with min radius
            If Not oDataArr(nMinRadiusIndex) Is Nothing Then
                Dim temp As Double = oAlignArc.MinimumRadius
                oDataArr(nMinRadiusIndex) = temp.ToString("N2")
                If oAlignArc.Radius >= oAlignArc.MinimumRadius Then
                    oDataArr(nMinRadiusResultIndex) = LocalizedRes.AlignCriVer_Html_Align_Cleared
                Else
                    oDataArr(nMinRadiusResultIndex) = "<font color=""#FF0000"">" + LocalizedRes.AlignCriVer_Html_Align_Violated + "</font>"
                End If
            End If
            AlignDataArr(nIndexDataArray) = oDataArr
        End If

        ExtractArcDataEx = True
    End Function


    Private Shared Function ExtractArcStationData(ByVal oAlignment As AeccLandLib.AeccAlignment, ByVal oAlignArc As AeccLandLib.AeccAlignmentArc, ByVal nIndexDataArray As Integer) As Boolean

        Dim oDataArr As Object

        On Error Resume Next
        oDataArr = m_oAlignDataArr(nIndexDataArray)
        If Not oAlignArc Is Nothing Then
            Dim oAlignEnt As AeccLandLib.AeccAlignmentEntity
            oAlignEnt = oAlignment.Entities.EntityAtId(oAlignArc.EntityBefore)
            If oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiral Then
                oDataArr(nPrevElementIndex) = LocalizedRes.AlignCriVer_Html_SC
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccArc Then
                Dim oAlignArcTemp As AeccLandLib.AeccAlignmentArc
                oAlignArcTemp = oAlignEnt
                If Not oAlignArcTemp Is Nothing Then
                    If oAlignArcTemp.ReverseCurve = True Then
                        oDataArr(nPrevElementIndex) = LocalizedRes.AlignCriVer_Html_PRC
                    Else
                        oDataArr(nPrevElementIndex) = LocalizedRes.AlignCriVer_Html_PCC
                    End If
                Else
                    oDataArr(nPrevElementIndex) = LocalizedRes.AlignCriVer_Html_PC
                End If
            Else
                oDataArr(nPrevElementIndex) = LocalizedRes.AlignCriVer_Html_PC
            End If
            oDataArr(nPrevStationIndex) = ReportUtilities.GetStationStringWithDerived(oAlignment, oAlignArc.StartingStation)
            oDataArr(nPrevDoubleStationIndex) = oAlignArc.StartingStation
            oDataArr(nCurrElementIndex) = LocalizedRes.AlignCriVer_Html_RP

            oAlignEnt = oAlignment.Entities.EntityAtId(oAlignArc.EntityAfter)
            If oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiral Then
                oDataArr(nNextElementIndex) = LocalizedRes.AlignCriVer_Html_CS
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccArc Then
                Dim oAlignArcTemp As AeccLandLib.AeccAlignmentArc
                oAlignArcTemp = oAlignEnt
                If Not oAlignArcTemp Is Nothing Then
                    If oAlignArcTemp.ReverseCurve = True Then
                        oDataArr(nNextElementIndex) = LocalizedRes.AlignCriVer_Html_PRC
                    Else
                        oDataArr(nNextElementIndex) = LocalizedRes.AlignCriVer_Html_PCC
                    End If
                Else
                    oDataArr(nNextElementIndex) = LocalizedRes.AlignCriVer_Html_PT
                End If
            Else
                oDataArr(nNextElementIndex) = LocalizedRes.AlignCriVer_Html_PT
            End If
            oDataArr(nNextStationIndex) = ReportUtilities.GetStationStringWithDerived(oAlignment, oAlignArc.EndingStation)
            m_oAlignDataArr(nIndexDataArray) = oDataArr
        End If

        ExtractArcStationData = True
    End Function
    Private Shared Function CalculateDirection(ByVal startNorthing As Double, ByVal startEasting As Double, ByVal endNorthing As Double, ByVal endEasting As Double) As Double
        Dim retAngle As Double
        Dim nTwo As Double
        Dim eTwo As Double
        Dim nOne As Double
        Dim eOne As Double
        Dim xDiff As Double
        Dim yDiff As Double
        Dim tanA As Double
        Dim angle As Double
        retAngle = 0

        nOne = startNorthing
        eOne = startEasting

        nTwo = endNorthing
        eTwo = endEasting

        xDiff = eOne - eTwo
        yDiff = nOne - nTwo

        tanA = yDiff / xDiff
        angle = (Math.Atan(tanA) * 180) / 3.1415926 'Math.PI

        If angle > 0 Then
            If nOne > nTwo Then  ' SW
                retAngle = 180.0# + angle
            Else
                retAngle = angle
            End If
        ElseIf angle < 0 Then
            If nOne > nTwo Then  ' SE
                retAngle = 360.0# + angle
            Else    '// NW
                retAngle = 180.0# + angle
            End If
        Else
            If eOne > eTwo Then
                retAngle = 180.0#
            Else
                retAngle = 0.0#
            End If
        End If
        CalculateDirection = retAngle
    End Function
    Private Shared Function GetSpiralChordLength(ByVal oAlignSpiral As AeccLandLib.AeccAlignmentSpiral) As Double
        Dim xDiff As Double
        Dim yDiff As Double
        xDiff = oAlignSpiral.StartEasting - oAlignSpiral.EndEasting
        yDiff = oAlignSpiral.StartNorthing - oAlignSpiral.EndNorthing
        GetSpiralChordLength = Math.Sqrt(xDiff * xDiff + yDiff * yDiff)

    End Function


    Private Shared Function ExtractSpiralData(ByVal oAlignSpiral As AeccLandLib.AeccAlignmentSpiral, ByVal nIndexDataArray As Integer) As Boolean

        Dim oDataArr As Object

        On Error Resume Next
        oDataArr = m_oAlignDataArr(nIndexDataArray)
        Dim spiralType As String
        spiralType = LocalizedRes.AlignCriVer_Html_Clothoid
        If Not oAlignSpiral Is Nothing Then
            oDataArr(nTypeIndex) = AeccLandLib.AeccAlignmentEntityType.aeccSpiral
            oDataArr(nLengthIndex) = FormatDistSettings(oAlignSpiral.Length)
            oDataArr(nAValueIndex) = FormatDistSettings(oAlignSpiral.A)
            oDataArr(nMinTransLenIndex) = FormatDistSettings(oAlignSpiral.MinimumTransitionLength)
            oDataArr(nCompoundSpiralIndex) = oAlignSpiral.Compound
            If Not oDataArr(nMinTransLenIndex) Is Nothing Then
                Dim temp As Double = oAlignSpiral.MinimumTransitionLength
                oDataArr(nMinTransLenIndex) = temp.ToString("N2")
                'compare length with min transition length
                If oAlignSpiral.Length >= oAlignSpiral.MinimumTransitionLength Then
                    oDataArr(nMinTransLenResultIndex) = LocalizedRes.AlignCriVer_Html_Align_Cleared
                Else
                    oDataArr(nMinTransLenResultIndex) = "<font color=""#FF0000"">" + LocalizedRes.AlignCriVer_Html_Align_Violated + "</font>"
                End If
            End If
            If oAlignSpiral.SpiralType = AeccLandLib.AeccAlignmentSpiralType.aeccAlignmentSpiralBiQuadratic Then
                spiralType = LocalizedRes.AlignCriVer_Html_BiQuadratic
            ElseIf oAlignSpiral.SpiralType = AeccLandLib.AeccAlignmentSpiralType.aeccAlignmentSpiralJapCubic Then
                spiralType = LocalizedRes.AlignCriVer_Html_JapCubic
            ElseIf oAlignSpiral.SpiralType = AeccLandLib.AeccAlignmentSpiralType.aeccAlignmentSpiralSineHalfWave Then
                spiralType = LocalizedRes.AlignCriVer_Html_SineHalfWave
            ElseIf oAlignSpiral.SpiralType = AeccLandLib.AeccAlignmentSpiralType.aeccAlignmentSpiralBloss Then
                spiralType = LocalizedRes.AlignCriVer_Html_Bloss
            ElseIf oAlignSpiral.SpiralType = AeccLandLib.AeccAlignmentSpiralType.aeccAlignmentSpiralCubicParabola Then
                spiralType = LocalizedRes.AlignCriVer_Html_CubicParabola
            ElseIf oAlignSpiral.SpiralType = AeccLandLib.AeccAlignmentSpiralType.aeccAlignmentSpiralSinusoidal Then
                spiralType = LocalizedRes.AlignCriVer_Html_Sinusoidal
            Else
                spiralType = LocalizedRes.AlignCriVer_Html_Clothoid
            End If
            oDataArr(nSpiralTypeIndex) = spiralType
            m_oAlignDataArr(nIndexDataArray) = oDataArr
        End If

        ExtractSpiralData = True
    End Function

    Private Shared Function ExtractSpiralStationData(ByVal oAlignment As AeccLandLib.AeccAlignment, ByVal oAlignSpiral As AeccLandLib.AeccAlignmentSpiral, ByVal nIndexDataArray As Integer) As Boolean

        Dim oDataArr As Object

        On Error Resume Next
        oDataArr = m_oAlignDataArr(nIndexDataArray)
        If Not oAlignSpiral Is Nothing Then
            Dim oAlignEnt As AeccLandLib.AeccAlignmentEntity
            oAlignEnt = oAlignment.Entities.EntityAtId(oAlignSpiral.EntityBefore)
            If oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiral Then
                oDataArr(nPrevElementIndex) = LocalizedRes.AlignCriVer_Html_SS
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccArc Then
                oDataArr(nPrevElementIndex) = LocalizedRes.AlignCriVer_Html_CS
            Else
                oDataArr(nPrevElementIndex) = LocalizedRes.AlignCriVer_Html_TS
            End If
            oDataArr(nPrevStationIndex) = ReportUtilities.GetStationStringWithDerived(oAlignment, oAlignSpiral.StartingStation)
            oDataArr(nPrevDoubleStationIndex) = oAlignSpiral.StartingStation

            oAlignEnt = oAlignment.Entities.EntityAtId(oAlignSpiral.EntityAfter)
            If oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiral Then
                oDataArr(nNextElementIndex) = LocalizedRes.AlignCriVer_Html_SS
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccArc Then
                oDataArr(nNextElementIndex) = LocalizedRes.AlignCriVer_Html_SC
            Else
                oDataArr(nNextElementIndex) = LocalizedRes.AlignCriVer_Html_ST
            End If
            oDataArr(nNextStationIndex) = ReportUtilities.GetStationStringWithDerived(oAlignment, oAlignSpiral.EndingStation)

            m_oAlignDataArr(nIndexDataArray) = oDataArr
        End If

        ExtractSpiralStationData = True
    End Function
    Public Shared Function ExtractSpiralDataEx(ByVal oAlignSpiral As AeccLandLib.AeccAlignmentSpiral, ByVal nIndexDataArray As Integer) As Boolean

        Dim oDataArr As Object
        oDataArr = AlignDataArr(nIndexDataArray)
        Dim spiralType As String
        spiralType = "Clothoid"
        On Error Resume Next
        If Not oAlignSpiral Is Nothing Then
            oDataArr(nTypeIndex) = AeccLandLib.AeccAlignmentEntityType.aeccSpiral
            oDataArr(nLengthIndex) = FormatDistSettings(oAlignSpiral.Length)
            oDataArr(nAValueIndex) = FormatDistSettings(oAlignSpiral.A)
            oDataArr(nMinTransLenIndex) = FormatDistSettings(oAlignSpiral.MinimumTransitionLength)
            oDataArr(nCompoundSpiralIndex) = oAlignSpiral.Compound
            If Not oDataArr(nMinTransLenIndex) Is Nothing Then
                Dim temp As Double = oAlignSpiral.MinimumTransitionLength
                oDataArr(nMinTransLenIndex) = temp.ToString("N2")
                'compare length with min transition length
                If oAlignSpiral.Length >= oAlignSpiral.MinimumTransitionLength Then
                    oDataArr(nMinTransLenResultIndex) = LocalizedRes.AlignCriVer_Html_Align_Cleared
                Else
                    oDataArr(nMinTransLenResultIndex) = "<font color=""#FF0000"">" + LocalizedRes.AlignCriVer_Html_Align_Violated + "</font>"
                End If
            End If

            If oAlignSpiral.SpiralType = AeccLandLib.AeccAlignmentSpiralType.aeccAlignmentSpiralBiQuadratic Then
                spiralType = LocalizedRes.AlignCriVer_Html_BiQuadratic
            ElseIf oAlignSpiral.SpiralType = AeccLandLib.AeccAlignmentSpiralType.aeccAlignmentSpiralJapCubic Then
                spiralType = LocalizedRes.AlignCriVer_Html_JapCubic
            ElseIf oAlignSpiral.SpiralType = AeccLandLib.AeccAlignmentSpiralType.aeccAlignmentSpiralSineHalfWave Then
                spiralType = LocalizedRes.AlignCriVer_Html_SineHalfWave
            ElseIf oAlignSpiral.SpiralType = AeccLandLib.AeccAlignmentSpiralType.aeccAlignmentSpiralBloss Then
                spiralType = LocalizedRes.AlignCriVer_Html_Bloss
            ElseIf oAlignSpiral.SpiralType = AeccLandLib.AeccAlignmentSpiralType.aeccAlignmentSpiralCubicParabola Then
                spiralType = LocalizedRes.AlignCriVer_Html_CubicParabola
            ElseIf oAlignSpiral.SpiralType = AeccLandLib.AeccAlignmentSpiralType.aeccAlignmentSpiralSinusoidal Then
                spiralType = LocalizedRes.AlignCriVer_Html_Sinusoidal
            Else
                spiralType = LocalizedRes.AlignCriVer_Html_Clothoid
            End If
            oDataArr(nSpiralTypeIndex) = spiralType
            AlignDataArr(nIndexDataArray) = oDataArr
        End If

        ExtractSpiralDataEx = True
    End Function

    Public Shared Function ExtractSpiralDesignCheckData(ByVal oAlignSpiral As AeccLandLib.AeccAlignmentSpiral, ByVal oDesignChecks As AeccLandLib.AeccAlignmentDesignChecks, ByVal nIndexDataArray As Integer) As Boolean
        Dim oDataArr As Object
        oDataArr = AlignDataArr(nIndexDataArray)
        If Not oAlignSpiral Is Nothing Then
            'need to filter design check
            Dim oDesignCheckNameArr() ' As Object
            Dim oAllDesignCheckNameArr As AeccLandLib.AeccAlignmentDesignChecks
            Dim oDesignCheckResultArr() ' As Object
            Dim oEachDesignCheckNameArr As AeccLandLib.AeccAlignmentDesignCheck
            oAllDesignCheckNameArr = oDesignChecks
            If oAllDesignCheckNameArr.Count Then
                ReDim oDesignCheckNameArr(oAllDesignCheckNameArr.Count)
                ReDim oDesignCheckResultArr(oAllDesignCheckNameArr.Count)
                Dim checkResult As Boolean
                Dim i As Integer
                i = 0
                For Each oEachDesignCheckNameArr In oAllDesignCheckNameArr
                    Try
                        oAlignSpiral.ValidateDesignCheck(oEachDesignCheckNameArr, checkResult)
                        oDesignCheckNameArr(i) = oEachDesignCheckNameArr.Name
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
                oDataArr(nDesignCheckNamesIndex) = oDesignCheckNameArr
                oDataArr(nDesignCheckResultsIndex) = oDesignCheckResultArr
            End If

            Try
                Dim temp As Integer
                temp = oAlignSpiral.HighestDesignSpeed
                oDataArr(nDesignSpeedIndex) = temp
            Catch ex As Exception

            End Try
            AlignDataArr(nIndexDataArray) = oDataArr
        End If
        ExtractSpiralDesignCheckData = True
    End Function

    Public Shared Function ExtractGroupDesignCheckData(ByVal oAlignElem As AeccLandLib.AeccAlignmentEntity, ByVal oAlignEnt As AeccLandLib.AeccAlignmentEntity, ByVal nIndexDataArray As Integer) As Boolean
        Dim oDataArr As Object
        oDataArr = AlignDataArr(nIndexDataArray)
        'Dim oAlignLine As AeccLandLib.AeccAlignmentTangent
        Dim oAlignArc As AeccLandLib.AeccAlignmentArc
        Dim oAlignSpiral As AeccLandLib.AeccAlignmentSpiral
        Dim oAlignSCS As AeccLandLib.AeccAlignmentSCSGroup
        Dim oAlignSC As AeccLandLib.AeccAlignmentSCGroup
        Dim oAlignCS As AeccLandLib.AeccAlignmentCSGroup
        Dim oAlignSS As AeccLandLib.AeccAlignmentSSGroup

        If Not oAlignElem Is Nothing And Not oAlignEnt Is Nothing Then
            'oAlignLine = Nothing
            oAlignArc = Nothing
            oAlignSpiral = Nothing
            oAlignSCS = Nothing
            oAlignSC = Nothing
            oAlignCS = Nothing
            oAlignSS = Nothing

            Dim oDesignCheckNameArr() ' As Object
            Dim oAllElemDesignCheckNameArr As AeccLandLib.AeccAlignmentDesignChecks
            Dim oAllEntDesignCheckNameArr As AeccLandLib.AeccAlignmentDesignChecks
            Dim oDesignCheckResultArr() ' As Object
            Dim oEachDesignCheckNameArr As AeccLandLib.AeccAlignmentDesignCheck
            Dim oDesignCheckCount As Integer

            If oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccCurveSpiralGroup Then
                oAlignCS = oAlignEnt
                oAllEntDesignCheckNameArr = oAlignCS.DesignChecks
                oDesignCheckCount = oAlignCS.DesignChecks.Count
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralCurveGroup Then
                oAlignSC = oAlignEnt
                oAllEntDesignCheckNameArr = oAlignSC.DesignChecks
                oDesignCheckCount = oAlignSC.DesignChecks.Count
            ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralSpiralGroup Then
                oAlignSS = oAlignEnt
                oAllEntDesignCheckNameArr = oAlignSS.DesignChecks
                oDesignCheckCount = oAlignSS.DesignChecks.Count
                'ElseIf oAlignEnt.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiralCurveSpiralGroup Then
            Else
                oAlignSCS = oAlignEnt
                oAllEntDesignCheckNameArr = oAlignSCS.DesignChecks
                oDesignCheckCount = oAlignSCS.DesignChecks.Count
            End If

            Dim checkResult As Boolean
            Dim i As Integer
            i = 0

            If oAlignElem.Type = AeccLandLib.AeccAlignmentEntityType.aeccSpiral Then
                oAlignSpiral = oAlignElem
                If Not oAlignSpiral Is Nothing Then
                    oAllElemDesignCheckNameArr = oAlignSpiral.DesignChecks
                    oDesignCheckCount = oDesignCheckCount + oAllElemDesignCheckNameArr.Count
                    If oDesignCheckCount Then
                        ReDim oDesignCheckNameArr(oDesignCheckCount)
                        ReDim oDesignCheckResultArr(oDesignCheckCount)

                        For Each oEachDesignCheckNameArr In oAllElemDesignCheckNameArr
                            Try
                                oAlignSpiral.ValidateDesignCheck(oEachDesignCheckNameArr, checkResult)
                                oDesignCheckNameArr(i) = oEachDesignCheckNameArr.Name
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

                        oEachDesignCheckNameArr = Nothing
                        For Each oEachDesignCheckNameArr In oAllEntDesignCheckNameArr
                            Try
                                If Not oAlignSCS Is Nothing Then
                                    oAlignSCS.ValidateDesignCheck(oEachDesignCheckNameArr, checkResult)
                                ElseIf Not oAlignSC Is Nothing Then
                                    oAlignSC.ValidateDesignCheck(oEachDesignCheckNameArr, checkResult)
                                ElseIf Not oAlignCS Is Nothing Then
                                    oAlignCS.ValidateDesignCheck(oEachDesignCheckNameArr, checkResult)
                                ElseIf Not oAlignSS Is Nothing Then
                                    oAlignSS.ValidateDesignCheck(oEachDesignCheckNameArr, checkResult)
                                End If
                                oDesignCheckNameArr(i) = oEachDesignCheckNameArr.Name
                                If (checkResult) Then
                                    oDesignCheckResultArr(i) = LocalizedRes.ProfileCriVer_Html_Cleared
                                Else
                                    oDesignCheckResultArr(i) = "<font color=""#FF0000"">" + LocalizedRes.ProfileCriVer_Html_Violated + "</font>"
                                End If
                            Catch ex As Exception

                            End Try
                            i = i + 1
                        Next
                        oDataArr(nDesignCheckNamesIndex) = oDesignCheckNameArr
                        oDataArr(nDesignCheckResultsIndex) = oDesignCheckResultArr
                    End If 'End of "If oDesignCheckCount"
                    Try
                        Dim temp As Integer
                        temp = oAlignSpiral.HighestDesignSpeed
                        oDataArr(nDesignSpeedIndex) = temp
                    Catch ex As Exception

                    End Try
                End If 'End of  "If Not oAlignSpiral Is Nothing "
            ElseIf oAlignElem.Type = AeccLandLib.AeccAlignmentEntityType.aeccArc Then
                oAlignArc = oAlignElem
                oAllElemDesignCheckNameArr = oAlignArc.DesignChecks
                oDesignCheckCount = oDesignCheckCount + oAllElemDesignCheckNameArr.Count
                If oDesignCheckCount Then
                    ReDim oDesignCheckNameArr(oDesignCheckCount)
                    ReDim oDesignCheckResultArr(oDesignCheckCount)

                    For Each oEachDesignCheckNameArr In oAllElemDesignCheckNameArr
                        Try
                            oAlignArc.ValidateDesignCheck(oEachDesignCheckNameArr, checkResult)
                            oDesignCheckNameArr(i) = oEachDesignCheckNameArr.Name
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

                    oEachDesignCheckNameArr = Nothing
                    For Each oEachDesignCheckNameArr In oAllEntDesignCheckNameArr
                        Try
                            If Not oAlignSCS Is Nothing Then
                                oAlignSCS.ValidateDesignCheck(oEachDesignCheckNameArr, checkResult)
                            ElseIf Not oAlignSC Is Nothing Then
                                oAlignSC.ValidateDesignCheck(oEachDesignCheckNameArr, checkResult)
                            ElseIf Not oAlignCS Is Nothing Then
                                oAlignCS.ValidateDesignCheck(oEachDesignCheckNameArr, checkResult)
                            ElseIf Not oAlignSS Is Nothing Then
                                oAlignSS.ValidateDesignCheck(oEachDesignCheckNameArr, checkResult)
                            End If
                            oDesignCheckNameArr(i) = oEachDesignCheckNameArr.Name
                            If (checkResult) Then
                                oDesignCheckResultArr(i) = LocalizedRes.ProfileCriVer_Html_Cleared
                            Else
                                oDesignCheckResultArr(i) = "<font color=""#FF0000"">" + LocalizedRes.ProfileCriVer_Html_Violated + "</font>"
                            End If
                        Catch ex As Exception

                        End Try
                        i = i + 1
                    Next
                    oDataArr(nDesignCheckNamesIndex) = oDesignCheckNameArr
                    oDataArr(nDesignCheckResultsIndex) = oDesignCheckResultArr

                    Try
                        Dim temp As Integer
                        temp = oAlignArc.HighestDesignSpeed
                        oDataArr(nDesignSpeedIndex) = temp
                    Catch ex As Exception

                    End Try
                End If
            End If
            AlignDataArr(nIndexDataArray) = oDataArr
        End If
        ExtractGroupDesignCheckData = True
    End Function

    Public Shared Function ExtractArcDesignCheckData(ByVal oAlignArc As AeccLandLib.AeccAlignmentArc, ByVal oDesignChecks As AeccLandLib.AeccAlignmentDesignChecks, ByVal nIndexDataArray As Integer) As Boolean
        Dim oDataArr As Object
        oDataArr = AlignDataArr(nIndexDataArray)
        If Not oAlignArc Is Nothing Then
            'need to filter design check
            Dim oDesignCheckNameArr() ' As Object
            Dim oAllDesignCheckNameArr As AeccLandLib.AeccAlignmentDesignChecks
            Dim oDesignCheckResultArr() ' As Object
            Dim oEachDesignCheckNameArr As AeccLandLib.AeccAlignmentDesignCheck
            oAllDesignCheckNameArr = oDesignChecks
            If oAllDesignCheckNameArr.Count Then
                ReDim oDesignCheckNameArr(oAllDesignCheckNameArr.Count)
                ReDim oDesignCheckResultArr(oAllDesignCheckNameArr.Count)
                Dim checkResult As Boolean
                Dim i As Integer
                i = 0
                For Each oEachDesignCheckNameArr In oAllDesignCheckNameArr
                    Try
                        oAlignArc.ValidateDesignCheck(oEachDesignCheckNameArr, checkResult)
                        oDesignCheckNameArr(i) = oEachDesignCheckNameArr.Name
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
                oDataArr(nDesignCheckNamesIndex) = oDesignCheckNameArr
                oDataArr(nDesignCheckResultsIndex) = oDesignCheckResultArr
            End If

            Try
                Dim temp As Integer
                temp = oAlignArc.HighestDesignSpeed
                oDataArr(nDesignSpeedIndex) = temp
            Catch ex As Exception

            End Try
            AlignDataArr(nIndexDataArray) = oDataArr
        End If
        ExtractArcDesignCheckData = True
    End Function

    Public Shared Function ExtractTangentDesignCheckData(ByVal oAlignLine As AeccLandLib.AeccAlignmentTangent, ByVal oDesignChecks As AeccLandLib.AeccAlignmentDesignChecks, ByVal nIndexDataArray As Integer) As Boolean
        Dim oDataArr As Object
        oDataArr = AlignDataArr(nIndexDataArray)
        If Not oAlignLine Is Nothing Then
            'need to filter design check
            Dim oDesignCheckNameArr() ' As Object
            Dim oAllDesignCheckNameArr As AeccLandLib.AeccAlignmentDesignChecks
            Dim oDesignCheckResultArr() ' As Object
            Dim oEachDesignCheckNameArr As AeccLandLib.AeccAlignmentDesignCheck
            oAllDesignCheckNameArr = oDesignChecks
            If oAllDesignCheckNameArr.Count Then
                ReDim oDesignCheckNameArr(oAllDesignCheckNameArr.Count)
                ReDim oDesignCheckResultArr(oAllDesignCheckNameArr.Count)
                Dim checkResult As Boolean
                Dim i As Integer
                i = 0
                For Each oEachDesignCheckNameArr In oAllDesignCheckNameArr
                    Try
                        oAlignLine.ValidateDesignCheck(oEachDesignCheckNameArr, checkResult)
                        oDesignCheckNameArr(i) = oEachDesignCheckNameArr.Name
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
                oDataArr(nDesignCheckNamesIndex) = oDesignCheckNameArr
                oDataArr(nDesignCheckResultsIndex) = oDesignCheckResultArr
            End If
            AlignDataArr(nIndexDataArray) = oDataArr
        End If
        ExtractTangentDesignCheckData = True
    End Function


End Class
