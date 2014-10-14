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

Imports System
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports LocalizedRes = Report.My.Resources.LocalizedStrings


Friend Class ReportFormat
    '-----------------------------------------------------------------------
    '   set rawAngle(radian) into [0,2pi].
    '-----------------------------------------------------------------------
    Public Shared Function AdjustAngle2PI(ByVal rawAngle As Double) As Double
        Do While rawAngle >= Math.PI * 2 Or Math.PI * 2 - rawAngle <= 0.00000001
            rawAngle = rawAngle - Math.PI * 2
        Loop
        Do While rawAngle < 0.0 And 0.0 - rawAngle > 0.00000001
            rawAngle = rawAngle + Math.PI * 2
        Loop

        If Math.Abs(0.0 - rawAngle) <= 0.00000001 Then
            rawAngle = 0
        End If
        AdjustAngle2PI = rawAngle
    End Function

    '-----------------------------------------------------------------------
    '   rounds off values and returns a Double.
    '-----------------------------------------------------------------------
    Public Shared Function RoundDouble(ByVal value As Double, ByVal prec As Integer, _
        Optional ByVal roundType As AeccLandLib.AeccRoundingType = AeccLandLib.AeccRoundingType.aeccRoundingNormal) As Double

        System.Diagnostics.Debug.Assert(prec >= 0 And prec <= 8)

        Dim factor As Double
        factor = 1.0 * (10 ^ prec)

        Dim tmpValue As Double
        tmpValue = value * factor
        Dim absVal As Double = 0.0
        Dim sgnVal As Double = 0.0

        absVal = Math.Abs(tmpValue)
        sgnVal = Math.Sign(tmpValue)

        Select Case roundType
            Case Autodesk.AECC.Interop.Land.AeccRoundingType.aeccRoundingNormal
                tmpValue = sgnVal * Math.Floor(absVal + 0.5)
            Case Autodesk.AECC.Interop.Land.AeccRoundingType.aeccRoundingUp
                tmpValue = sgnVal * Math.Ceiling(absVal)
            Case Autodesk.AECC.Interop.Land.AeccRoundingType.aeccRoundingTruncate
                tmpValue = sgnVal * Math.Floor(absVal)
        End Select

        Return tmpValue / factor
    End Function

    '-----------------------------------------------------------------------
    '   rounds off values and returns a string.
    '-----------------------------------------------------------------------
    Public Shared Function RoundVal(ByVal value As Double, ByVal prec As Integer, _
        Optional ByVal roundType As AeccLandLib.AeccRoundingType = AeccLandLib.AeccRoundingType.aeccRoundingNormal) As String

        Dim roundedValue As Double = RoundDouble(value, prec, roundType)

        Dim sFormatString As String
        sFormatString = "N" + prec.ToString()
        RoundVal = roundedValue.ToString(sFormatString)

    End Function

    Public Shared Function RoundDMS_2(ByVal valueInDegrees As Double, _
                                    ByVal prec As Integer, _
                                    ByVal mode As AeccLandLib.AeccRoundingType) As Double

        Dim roundedVal As Double = 0.0
        Dim remainingPrec As Integer = 0
        Dim factor As Integer
        factor = calcDMSRoundingFactor(prec, remainingPrec)
        roundedVal = ReportFormat.RoundDouble(valueInDegrees * factor, remainingPrec, mode)
        roundedVal /= factor
        RoundDMS_2 = roundedVal
    End Function

    Private Shared Function calcDMSRoundingFactor(ByVal prec As Integer, ByRef remainingPrec As Integer) As Integer

        System.Diagnostics.Debug.Assert(prec >= 0 And prec <= 8)

        Dim factors As Integer() = New Integer() {1, 6, 60, 360, 3600}

        Dim result As Integer = 1

        If (prec <= 4) Then
            remainingPrec = 0
            result = factors(prec)
        Else
            remainingPrec = prec - 4
            result = factors(4)
        End If

        calcDMSRoundingFactor = result
    End Function

    '-----------------------------------------------------------------------
    '   format a grade
    '   e.g.:  1:3.232
    '   Note: given the originate gradeFormatType aeccGradeSlopeFormatDecimal
    '         this function don't take care aeccGradeSlopeFormatRiseRun & aeccGradeSlopeFormatRunRise,which will be in FormatSlope
    '-----------------------------------------------------------------------
    Public Shared Function FormatGrade(ByVal dGrade As Double, _
    Optional ByVal grdFormat As AeccLandLib.AeccGradeSlopeFormatType = AeccLandLib.AeccGradeSlopeFormatType.aeccGradeSlopeFormatPercent, _
    Optional ByVal grdPrec As Integer = 3, _
    Optional ByVal grdRounding As AeccLandLib.AeccRoundingType = AeccLandLib.AeccRoundingType.aeccRoundingNormal, _
    Optional ByVal grdSign As AeccLandLib.AeccSignType = AeccLandLib.AeccSignType.aeccSignNegative) As String
        Dim sPreSign As String = "", sPostSign As String = ""
        Select Case grdSign
            Case AeccLandLib.AeccSignType.aeccSignAlways
                If dGrade < 0.0# Then
                    sPreSign = "-"
                Else
                    sPreSign = "+"
                End If
            Case AeccLandLib.AeccSignType.aeccSignBracketNegative
                If dGrade < 0.0# Then
                    sPreSign = "("
                    sPostSign = ")"
                End If
            Case AeccLandLib.AeccSignType.aeccSignNegative
                If dGrade < 0.0# Then
                    sPreSign = "-"
                End If
        End Select
        dGrade = Math.Abs(dGrade)

        Dim sGrade As String
        sGrade = ""
        Select Case grdFormat
            Case AeccLandLib.AeccGradeSlopeFormatType.aeccGradeSlopeFormatDecimal
                sGrade = RoundVal(dGrade, grdPrec, grdRounding)
            Case AeccLandLib.AeccGradeSlopeFormatType.aeccGradeSlopeFormatPercent
                dGrade = dGrade * 100.0#
                sGrade = RoundVal(dGrade, grdPrec, grdRounding)
                sGrade = sGrade + "%"
        End Select

        FormatGrade = sPreSign + sGrade + sPostSign
    End Function

    '-----------------------------------------------------------------------
    '   format a volume
    '   e.g.:  +35.34Cu.Yd. +35.34Cu.M. +35.34Cu.Ft.
    '   Note: the originate unitType depends on drawing's AeccSettingsUnitZone
    '-----------------------------------------------------------------------
    Public Shared Function FormatArea(ByVal dArea As Double, _
    Optional ByVal areaUnit As AeccLandLib.AeccAreaUnitType = AeccLandLib.AeccAreaUnitType.aeccAreaUnitSquareFoot, _
    Optional ByVal volPrec As Integer = 4, _
    Optional ByVal volRounding As AeccLandLib.AeccRoundingType = AeccLandLib.AeccRoundingType.aeccRoundingNormal, _
    Optional ByVal volSign As AeccLandLib.AeccSignType = AeccLandLib.AeccSignType.aeccSignNegative, _
    Optional ByVal showUnitString As Boolean = True) As String

        Dim sPreSign As String, sPostSign As String
        sPreSign = ""
        sPostSign = ""
        Select Case volSign 'grdSign
            Case AeccLandLib.AeccSignType.aeccSignAlways
                If dArea < 0.0# Then
                    sPreSign = "-"
                Else
                    sPreSign = "+"
                End If
            Case AeccLandLib.AeccSignType.aeccSignBracketNegative
                If dArea < 0.0# Then
                    sPreSign = "("
                    sPostSign = ")"
                End If
            Case AeccLandLib.AeccSignType.aeccSignNegative
                If dArea < 0.0# Then
                    sPreSign = "-"
                End If
        End Select
        dArea = Math.Abs(dArea)
        'trans dDis to aeccCoordinateUnitMeter
        Dim disOriUnit As AeccLandLib.AeccDrawingUnitType
        disOriUnit = ReportApplication.AeccXDatabase.Settings.DrawingSettings.UnitZoneSettings.DrawingUnits
        If disOriUnit = AeccLandLib.AeccDrawingUnitType.aeccDrawingUnitFeet Then
            dArea = dArea * (0.3048006096012 ^ 2)
        End If

        Dim sArea As String
        sArea = ""
        Select Case areaUnit
            Case AeccLandLib.AeccAreaUnitType.aeccAreaUnitSquareMeter  'aeccVolumeUnitCubicMeter
                sArea = RoundVal(dArea, volPrec, volRounding)
                'sArea = sArea + LocalizedRes.ReportFormat_SqM '"Sq.M."
            Case AeccLandLib.AeccAreaUnitType.aeccAreaUnitSquareFoot
                dArea = dArea / (0.3048006096012 ^ 2)
                sArea = RoundVal(dArea, volPrec, volRounding)
                'sArea = sArea + LocalizedRes.ReportFormat_SqFt '"Sq.Ft."
            Case AeccLandLib.AeccAreaUnitType.aeccAreaUnitAcre
                dArea = dArea * 0.0002471
                sArea = RoundVal(dArea, volPrec, volRounding)
                'sArea = sArea + LocalizedRes.ReportFormat_Acre '"Acre"
            Case AeccLandLib.AeccAreaUnitType.aeccAreaUnitHectare
                dArea = dArea * 0.0001
                sArea = RoundVal(dArea, volPrec, volRounding)
                'sArea = sArea + LocalizedRes.ReportFormat_Hectare '"Hectare"
            Case AeccLandLib.AeccAreaUnitType.aeccAreaUnitSquareKilometer
                dArea = dArea * 0.000001
                sArea = RoundVal(dArea, volPrec, volRounding)
                'sArea = sArea + LocalizedRes.ReportFormat_SqKm '"Sq.Km."
            Case AeccLandLib.AeccAreaUnitType.aeccAreaUnitSquareMile
                dArea = dArea * 0.000000386
                sArea = RoundVal(dArea, volPrec, volRounding)
                'sArea = sArea + LocalizedRes.ReportFormat_SqMile '"Sq.Mile"
            Case AeccLandLib.AeccAreaUnitType.aeccAreaUnitSquareYard
                dArea = dArea * 1.196
                sArea = RoundVal(dArea, volPrec, volRounding)
                'sArea = sArea + LocalizedRes.ReportFormat_SqYd '"Sq.Yd."
        End Select

        If showUnitString Then
            sArea = sArea + AreaUnitString(areaUnit)
        End If

        FormatArea = sPreSign + sArea + sPostSign
    End Function

    Public Shared Function AreaUnitString(Optional ByVal areaUnit As AeccLandLib.AeccAreaUnitType = AeccLandLib.AeccAreaUnitType.aeccAreaUnitSquareFoot)
        Dim sArea As String
        sArea = ""
        Select Case areaUnit
            Case AeccLandLib.AeccAreaUnitType.aeccAreaUnitSquareMeter  'aeccVolumeUnitCubicMeter
                sArea = LocalizedRes.ReportFormat_SqM '"Sq.M."
            Case AeccLandLib.AeccAreaUnitType.aeccAreaUnitSquareFoot
                sArea = LocalizedRes.ReportFormat_SqFt '"Sq.Ft."
            Case AeccLandLib.AeccAreaUnitType.aeccAreaUnitAcre
                sArea = LocalizedRes.ReportFormat_Acre '"Acre"
            Case AeccLandLib.AeccAreaUnitType.aeccAreaUnitHectare
                sArea = LocalizedRes.ReportFormat_Hectare '"Hectare"
            Case AeccLandLib.AeccAreaUnitType.aeccAreaUnitSquareKilometer
                sArea = LocalizedRes.ReportFormat_SqKm '"Sq.Km."
            Case AeccLandLib.AeccAreaUnitType.aeccAreaUnitSquareMile
                sArea = LocalizedRes.ReportFormat_SqMile '"Sq.Mile"
            Case AeccLandLib.AeccAreaUnitType.aeccAreaUnitSquareYard
                sArea = LocalizedRes.ReportFormat_SqYd '"Sq.Yd."
        End Select
        AreaUnitString = sArea
    End Function

    '-----------------------------------------------------------------------
    '   format a volume
    '   e.g.:  +35.34Cu.Yd. +35.34Cu.M. +35.34Cu.Ft.
    '   Note: the originate unitType depends on drawing's AeccSettingsUnitZone
    '-----------------------------------------------------------------------
    Public Shared Function FormatVolume(ByVal dVol As Double, _
    Optional ByVal volUnit As AeccLandLib.AeccVolumeUnitType = AeccLandLib.AeccVolumeUnitType.aeccVolumeUnitCubicFoot, _
    Optional ByVal volPrec As Integer = 4, _
    Optional ByVal volRounding As AeccLandLib.AeccRoundingType = AeccLandLib.AeccRoundingType.aeccRoundingNormal, _
    Optional ByVal volSign As AeccLandLib.AeccSignType = AeccLandLib.AeccSignType.aeccSignNegative, _
    Optional ByVal showUnitString As Boolean = True) As String
        Dim sPreSign As String, sPostSign As String
        sPreSign = ""
        sPostSign = ""
        Select Case volSign 'grdSign
            Case AeccLandLib.AeccSignType.aeccSignAlways
                If dVol < 0.0# Then
                    sPreSign = "-"
                Else
                    sPreSign = "+"
                End If
            Case AeccLandLib.AeccSignType.aeccSignBracketNegative
                If dVol < 0.0# Then
                    sPreSign = "("
                    sPostSign = ")"
                End If
            Case AeccLandLib.AeccSignType.aeccSignNegative
                If dVol < 0.0# Then
                    sPreSign = "-"
                End If
        End Select
        dVol = Math.Abs(dVol)

        'trans dDis to aeccCoordinateUnitMeter
        Dim disOriUnit As AeccLandLib.AeccDrawingUnitType
        disOriUnit = ReportApplication.AeccXDatabase.Settings.DrawingSettings.UnitZoneSettings.DrawingUnits
        If disOriUnit = AeccLandLib.AeccDrawingUnitType.aeccDrawingUnitFeet Then
            dVol = dVol * (0.3048006096012 ^ 3)
        End If

        Dim sVol As String
        sVol = ""
        Select Case volUnit
            Case AeccLandLib.AeccVolumeUnitType.aeccVolumeUnitCubicMeter
                sVol = RoundVal(dVol, volPrec, volRounding)
                'sVol = sVol + LocalizedRes.ReportFormat_CuM '"Cu.M."
            Case AeccLandLib.AeccVolumeUnitType.aeccVolumeUnitCubicFoot
                dVol = dVol / (0.3048006096012 ^ 3)
                sVol = RoundVal(dVol, volPrec, volRounding)
                'sVol = sVol + LocalizedRes.ReportFormat_CuFt '"Cu.Ft"
            Case AeccLandLib.AeccVolumeUnitType.aeccVolumeUnitCubicYard
                dVol = dVol / ((3.0# * 0.3048006096012) ^ 3)
                sVol = RoundVal(dVol, volPrec, volRounding)
                'sVol = sVol + LocalizedRes.ReportFormat_CuYd '"Cu.Yd"
        End Select

        If showUnitString Then
            sVol = sVol + VolumeUnitString(volUnit)
        End If
        FormatVolume = sPreSign + sVol + sPostSign
    End Function

    Public Shared Function VolumeUnitString(Optional ByVal volUnit As AeccLandLib.AeccVolumeUnitType = AeccLandLib.AeccVolumeUnitType.aeccVolumeUnitCubicFoot) As String
        Dim sVol As String
        sVol = ""
        Select Case volUnit
            Case AeccLandLib.AeccVolumeUnitType.aeccVolumeUnitCubicMeter
                sVol = LocalizedRes.ReportFormat_CuM '"Cu.M."
            Case AeccLandLib.AeccVolumeUnitType.aeccVolumeUnitCubicFoot
                sVol = LocalizedRes.ReportFormat_CuFt '"Cu.Ft"
            Case AeccLandLib.AeccVolumeUnitType.aeccVolumeUnitCubicYard
                sVol = LocalizedRes.ReportFormat_CuYd '"Cu.Yd"
        End Select
        VolumeUnitString = sVol
    End Function

    '-----------------------------------------------------------------------
    '   format a distance
    '   e.g.:  +35.34(f)
    '   Note: the originate unitType depends on drawing's AeccSettingsUnitZone
    '-----------------------------------------------------------------------
    Public Shared Function FormatDistance(ByVal dDis As Double, Optional ByVal disUnit As AeccLandLib.AeccCoordinateUnitType = AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitFoot, _
    Optional ByVal disPrec As Integer = 6, _
    Optional ByVal disRounding As AeccLandLib.AeccRoundingType = AeccLandLib.AeccRoundingType.aeccRoundingNormal, _
    Optional ByVal disSign As AeccLandLib.AeccSignType = AeccLandLib.AeccSignType.aeccSignNegative, _
    Optional ByVal unitIndicator As Boolean = True) As String
        Dim sPreSign As String = "", sPostSign As String = ""
        Select Case disSign
            Case AeccLandLib.AeccSignType.aeccSignAlways
                If dDis < 0.0# Then
                    sPreSign = "-"
                Else
                    sPreSign = "+"
                End If
            Case AeccLandLib.AeccSignType.aeccSignBracketNegative
                If dDis < 0.0# Then
                    sPreSign = "("
                    sPostSign = ")"
                End If
            Case AeccLandLib.AeccSignType.aeccSignNegative
                If dDis < 0.0# Then
                    sPreSign = "-"
                End If
        End Select
        dDis = Math.Abs(dDis)

        'trans dDis to aeccCoordinateUnitMeter
        Dim disOriUnit As AeccLandLib.AeccDrawingUnitType
        disOriUnit = ReportApplication.AeccXDatabase.Settings.DrawingSettings.UnitZoneSettings.DrawingUnits
        If disOriUnit = AeccLandLib.AeccDrawingUnitType.aeccDrawingUnitFeet Then
            dDis = dDis * 0.3048006096012
        End If

        Dim sDis As String
        sDis = ""
        Select Case disUnit
            Case AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitMeter
                sDis = RoundVal(dDis, disPrec, disRounding)
                If unitIndicator = True Then
                    sDis = sDis + LocalizedRes.ReportFormat_M '"m"
                End If

            Case AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitKilometer
                dDis = dDis / 1000.0#
                sDis = RoundVal(dDis, disPrec, disRounding)
                If unitIndicator = True Then
                    sDis = sDis + LocalizedRes.ReportFormat_Km '"km"
                End If

            Case AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitFoot
                dDis = dDis / 0.3048006096012
                sDis = RoundVal(dDis, disPrec, disRounding)
                If unitIndicator = True Then
                    sDis = sDis + LocalizedRes.ReportFormat_Foot '"'"
                End If

            Case AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitInch
                dDis = dDis * 12.0# / 0.3048006096012
                sDis = RoundVal(dDis, disPrec, disRounding)
                If unitIndicator = True Then
                    sDis = sDis + LocalizedRes.ReportFormat_Inch '""""
                End If

            Case AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitYard
                dDis = dDis / 3.0# / 0.3048006096012
                sDis = RoundVal(dDis, disPrec, disRounding)
                If unitIndicator = True Then
                    sDis = sDis + LocalizedRes.ReportFormat_yd '"yd"
                End If

            Case AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitMile
                dDis = dDis / 1760.0# / 3.0# / 0.3048006096012
                sDis = RoundVal(dDis, disPrec, disRounding)
                If unitIndicator = True Then
                    sDis = sDis + LocalizedRes.ReportFormat_Miles '"miles"
                End If

        End Select
        FormatDistance = sPreSign + sDis + sPostSign
    End Function

    '-----------------------------------------------------------------------
    '   format a radian angle
    '   e.g.:  65^34'32.32"
    '   Note:  do not take care dRadAng's scope
    '-----------------------------------------------------------------------
    Public Shared Function FormatAngle(ByVal dRadAng As Double, _
    Optional ByVal angUnit As AeccLandLib.AeccAngleUnitType = AeccLandLib.AeccAngleUnitType.aeccAngleUnitDegree, _
    Optional ByVal angPrec As Integer = 6, _
    Optional ByVal angRounding As AeccLandLib.AeccRoundingType = AeccLandLib.AeccRoundingType.aeccRoundingNormal, _
    Optional ByVal angFormat As AeccLandLib.AeccFormatType = AeccLandLib.AeccFormatType.aeccFormatDegreeMinuteSecond, _
    Optional ByVal angSign As AeccLandLib.AeccSignType = AeccLandLib.AeccSignType.aeccSignNegative) As String
        Dim sPreSign As String, sPostSign As String
        sPreSign = ""
        sPostSign = ""
        Select Case angSign
            Case AeccLandLib.AeccSignType.aeccSignAlways
                If dRadAng < 0.0# Then
                    sPreSign = "-"
                Else
                    sPreSign = "+"
                End If
            Case AeccLandLib.AeccSignType.aeccSignBracketNegative
                If dRadAng < 0.0# Then
                    sPreSign = "("
                    sPostSign = ")"
                End If
            Case AeccLandLib.AeccSignType.aeccSignNegative
                If dRadAng < 0.0# Then
                    sPreSign = "-"
                End If
        End Select

        dRadAng = Math.Abs(dRadAng)

        Dim sAngle As String
        sAngle = ""
        Select Case angUnit
            Case AeccLandLib.AeccAngleUnitType.aeccAngleUnitDegree
                Dim dAng As Double
                dAng = dRadAng * 180.0# / Math.PI
                
                If angFormat = AeccLandLib.AeccFormatType.aeccFormatDecimal Then
                    sAngle = RoundVal(dAng, angPrec, angRounding)
                    sAngle = sAngle + " " + LocalizedRes.ReportFormat_Decimal '" (d)"
                Else
                    sAngle = formatDMS(dAng, angPrec, angRounding, angFormat)
                End If
            Case AeccLandLib.AeccAngleUnitType.aeccAngleUnitGrad
                dRadAng = dRadAng * 200.0# / Math.PI
                sAngle = RoundVal(dRadAng, angPrec, angRounding)
                sAngle = sAngle + " " + LocalizedRes.ReportFormat_G '" (g)"
            Case AeccLandLib.AeccAngleUnitType.aeccAngleUnitRadian
                dRadAng = dRadAng
                sAngle = RoundVal(dRadAng, angPrec, angRounding)
                sAngle = sAngle + " " + LocalizedRes.ReportFormat_R '" (r)"
        End Select 'angUnit
        FormatAngle = sPreSign + sAngle + sPostSign
    End Function

    '-----------------------------------------------------------------------
    '   format a radian angle to direction.
    '   e.g.:  N65^34'32.32"E
    '   Note:  dRadAng is start from x-axis and rotate counterclockwise
    '-----------------------------------------------------------------------
    Public Shared Function FormatDirection(ByVal dRadAng As Double, _
    Optional ByVal angUnit As AeccLandLib.AeccAngleUnitType = AeccLandLib.AeccAngleUnitType.aeccAngleUnitDegree, _
    Optional ByVal angPrec As Integer = 6, _
    Optional ByVal angRounding As AeccLandLib.AeccRoundingType = AeccLandLib.AeccRoundingType.aeccRoundingNormal, _
    Optional ByVal angFormat As AeccLandLib.AeccFormatType = AeccLandLib.AeccFormatType.aeccFormatDegreeMinuteSecond, _
    Optional ByVal angDirection As AeccLandLib.AeccDirectionType = AeccLandLib.AeccDirectionType.aeccDirectionShortName, _
    Optional ByVal angCap As AeccLandLib.AeccCapitalizationType = AeccLandLib.AeccCapitalizationType.aeccCapitalizationUpperCase, _
    Optional ByVal angSign As AeccLandLib.AeccSignType = AeccLandLib.AeccSignType.aeccSignNegative, _
    Optional ByVal angMeasType As AeccLandLib.AeccMeasurementType = AeccLandLib.AeccMeasurementType.aeccMeasurementBearings, _
    Optional ByVal angBearQuand As AeccLandLib.AeccBearingQuadrantType = AeccLandLib.AeccBearingQuadrantType.aeccBearingQuadrantNorthEast) As String
        dRadAng = AdjustAngle2PI(dRadAng)
        dRadAng = dRadAng * 180.0 / Math.PI

        Dim sPrefix As String, sPostfix As String
        sPrefix = ""
        sPostfix = ""

        Dim sNorth As String = ""
        Dim sSouth As String = ""
        Dim sEast As String = ""
        Dim sWest As String = ""
        getDirectionName(angDirection, angCap, sNorth, sSouth, sEast, sWest)

        Select Case angMeasType
            Case AeccLandLib.AeccMeasurementType.aeccMeasurementBearings
                If dRadAng >= 0 And dRadAng <= 90 Then
                    dRadAng = 90.0 - dRadAng
                    sPrefix = sNorth
                    sPostfix = sEast
                ElseIf dRadAng > 90 And dRadAng <= 180 Then
                    dRadAng = dRadAng - 90.0
                    sPrefix = sNorth
                    sPostfix = sWest
                ElseIf dRadAng > 180 And dRadAng < 270 Then
                    dRadAng = 270.0 - dRadAng
                    sPrefix = sSouth
                    sPostfix = sWest
                ElseIf dRadAng >= 270 And dRadAng <= 360 Then
                    dRadAng = dRadAng - 270.0
                    sPrefix = sSouth
                    sPostfix = sEast
                Else
                    System.Diagnostics.Debug.Assert(False, "dRadAng should be in range 0~360")
                End If
            Case AeccLandLib.AeccMeasurementType.aeccMeasurementNorthAzimuth
                dRadAng = 450.0 - dRadAng
                If dRadAng >= 360 Then
                    dRadAng = dRadAng - 360
                End If
                sPrefix = LocalizedRes.ReportFormat_NUppercase + " " '"N "
            Case AeccLandLib.AeccMeasurementType.aeccMeasurementSouthAzimuth
                dRadAng = 270 - dRadAng
                If dRadAng <= 0 Then
                    dRadAng = dRadAng + 360
                End If
                sPrefix = LocalizedRes.ReportFormat_SUppercase '"S "
        End Select 'angMeasType

        'get dRadAng's string
        Dim sAngle As String
        sAngle = ""
        Select Case angUnit
            Case AeccLandLib.AeccAngleUnitType.aeccAngleUnitDegree
                If angFormat <> AeccLandLib.AeccFormatType.aeccFormatDecimal Then
                    sAngle = formatDMS(dRadAng, angPrec, angRounding, angFormat)
                Else
                    sAngle = RoundVal(dRadAng, angPrec, angRounding)
                End If

                Select Case angFormat
                    Case AeccLandLib.AeccFormatType.aeccFormatDecimal
                        sPostfix = sPostfix + " " + _
                            LocalizedRes.ReportFormat_Decimal '" (d)"
                    Case AeccLandLib.AeccFormatType.aeccFormatDecimalDegreeMinuteSecond
                        sPostfix = sPostfix + " " + _
                            LocalizedRes.ReportFormat_Dms '" (dms)"
                        'Case AeccLandLib.AeccFormatType.aeccFormatDegreeMinuteSecond
                        'Case AeccLandLib.AeccFormatType.aeccFormatDegreeMinuteSecondSpaced
                End Select
            Case AeccLandLib.AeccAngleUnitType.aeccAngleUnitGrad
                dRadAng = dRadAng * 400 / 360
                sAngle = RoundVal(dRadAng, angPrec, angRounding)
                sPostfix = sPostfix + " " + LocalizedRes.ReportFormat_G '" (g)"
            Case AeccLandLib.AeccAngleUnitType.aeccAngleUnitRadian
                dRadAng = dRadAng * Math.PI / 360
                sAngle = RoundVal(dRadAng, angPrec, angRounding)
                sPostfix = sPostfix + " " + LocalizedRes.ReportFormat_R '" (r)"
        End Select 'angUnit
        FormatDirection = sPrefix + sAngle + sPostfix
    End Function

    Private Shared Sub getDirectionName(ByVal angDirection As AeccLandLib.AeccDirectionType, _
                                        ByVal angCap As AeccLandLib.AeccCapitalizationType, _
                                        ByRef sNorth As String, _
                                        ByRef sSouth As String, _
                                        ByRef sEast As String, _
                                        ByRef sWest As String)

        Dim north_l As String = LocalizedRes.ReportFormat_NorthLowercase
        Dim NORTH_u As String = LocalizedRes.ReportFormat_NorthUppercase
        Dim North_c As String = LocalizedRes.ReportFormat_NorthTitleCaps

        Dim west_l As String = LocalizedRes.ReportFormat_WestLowercase
        Dim WEST_u As String = LocalizedRes.ReportFormat_WestUppercase
        Dim West_c As String = LocalizedRes.ReportFormat_WestTitleCaps

        Dim south_l As String = LocalizedRes.ReportFormat_SouthLowercase
        Dim SOUTH_u As String = LocalizedRes.ReportFormat_SouthUppercase
        Dim South_c As String = LocalizedRes.ReportFormat_SouthTitleCaps

        Dim east_l As String = LocalizedRes.ReportFormat_EastLowercase
        Dim EAST_u As String = LocalizedRes.ReportFormat_EastUppercase
        Dim East_c As String = LocalizedRes.ReportFormat_EastTitleCaps

        Dim n_l As String = LocalizedRes.ReportFormat_NLowercase
        Dim N_u As String = LocalizedRes.ReportFormat_NUppercase

        Dim w_l As String = LocalizedRes.ReportFormat_WLowercase
        Dim W_u As String = LocalizedRes.ReportFormat_WUppercase

        Dim s_l As String = LocalizedRes.ReportFormat_SLowercase
        Dim S_u As String = LocalizedRes.ReportFormat_SUppercase

        Dim e_l As String = LocalizedRes.ReportFormat_ELowercase
        Dim E_u As String = LocalizedRes.ReportFormat_EUppercase

        Select Case angDirection
            Case AeccLandLib.AeccDirectionType.aeccDirectionLongName
                Select Case angCap
                    Case AeccLandLib.AeccCapitalizationType.aeccCapitalizationLowerCase
                        sNorth = north_l 'sPrefix = north_l '"north"
                        sWest = west_l 'sPostfix = west_l '"west"
                        sSouth = south_l
                        sEast = east_l
                    Case AeccLandLib.AeccCapitalizationType.aeccCapitalizationTitleCaps
                        sNorth = North_c 'sPrefix = North_c '"North"
                        sWest = West_c 'sPostfix = West_c '"West"
                        sSouth = South_c
                        sEast = East_c
                    Case Else 'AeccLandLib.AeccCapitalizationType.aeccCapitalizationPreserveCase, _
                        'AeccLandLib.AeccCapitalizationType.aeccCapitalizationUpperCase()
                        sNorth = NORTH_u 'sPrefix = NORTH_u '"NORTH"
                        sWest = WEST_u 'sPostfix = WEST_u '"WEST"
                        sSouth = SOUTH_u
                        sEast = EAST_u
                End Select
            Case AeccLandLib.AeccDirectionType.aeccDirectionShortName
                Select Case angCap
                    Case AeccLandLib.AeccCapitalizationType.aeccCapitalizationLowerCase
                        sNorth = n_l 'sPrefix = n_l '"n"
                        sWest = w_l 'sPostfix = w_l '"w"
                        sSouth = s_l
                        sEast = e_l
                    Case Else 'AeccLandLib.AeccCapitalizationType.aeccCapitalizationPreserveCase, _
                        'AeccLandLib.AeccCapitalizationType.aeccCapitalizationTitleCaps, _
                        'AeccLandLib.AeccCapitalizationType.aeccCapitalizationUpperCase()
                        sNorth = N_u 'sPrefix = N_u '"N"
                        sWest = W_u 'sPostfix = W_u '"W"
                        sSouth = S_u
                        sEast = E_u
                End Select
            Case AeccLandLib.AeccDirectionType.aeccDirectionLongNameSpaced
                Select Case angCap
                    Case AeccLandLib.AeccCapitalizationType.aeccCapitalizationLowerCase
                        sNorth = north_l + " " 'sPrefix = north_l + " " '"north "
                        sWest = " " + west_l 'sPostfix = " " + west_l '" west"
                        sSouth = south_l + " "
                        sEast = " " + east_l
                    Case AeccLandLib.AeccCapitalizationType.aeccCapitalizationTitleCaps
                        sNorth = North_c + " " 'sPrefix = North_c + " " '"North "
                        sWest = " " + West_c 'sPostfix = " " + West_c '" West"
                        sSouth = South_c + " "
                        sEast = " " + East_c
                    Case Else 'AeccLandLib.AeccCapitalizationType.aeccCapitalizationPreserveCase, _
                        'AeccLandLib.AeccCapitalizationType.aeccCapitalizationUpperCase()
                        sNorth = NORTH_u + " " 'sPrefix = NORTH_u + " " '"NORTH "
                        sWest = " " + WEST_u 'sPostfix = " " + WEST_u '" WEST"
                        sSouth = SOUTH_u + " "
                        sEast = " " + EAST_u
                End Select


            Case Else 'AeccLandLib.AeccDirectionType.aeccDirectionShortNameSpaced
                Select Case angCap
                    Case AeccLandLib.AeccCapitalizationType.aeccCapitalizationLowerCase
                        sNorth = n_l + " " 'sPrefix = n_l + " " '"n "
                        sWest = " " + w_l 'sPostfix = " " + w_l '" w"
                        sSouth = s_l + " "
                        sEast = " " + e_l
                    Case Else 'AeccLandLib.AeccCapitalizationType.aeccCapitalizationPreserveCase, _
                        'AeccLandLib.AeccCapitalizationType.aeccCapitalizationTitleCaps, _
                        'AeccLandLib.AeccCapitalizationType.aeccCapitalizationUpperCase()
                        sNorth = N_u + " " 'sPrefix = N_u + " " '"N "
                        sWest = " " + W_u 'sPostfix = " " + W_u '" W"
                        sSouth = S_u + " "
                        sEast = " " + E_u
                End Select
        End Select 'angDirection

    End Sub
    Private Shared Function formatDMS(ByVal value As Double, _
                                    ByVal prec As Integer, _
                                    ByVal rounding As AeccLandLib.AeccRoundingType, _
                                    ByVal formatType As AeccLandLib.AeccFormatType) As String
        Dim resultDMS As String = ""

        Dim tagDegree As String = ""
        Dim tagMin As String = ""
        Dim tagSec As String = ""

        Select Case formatType
            Case AeccLandLib.AeccFormatType.aeccFormatDecimalDegreeMinuteSecond
                tagDegree = "."
            Case AeccLandLib.AeccFormatType.aeccFormatDegreeMinuteSecond
                tagDegree = Chr(176)
                tagMin = "'"
                tagSec = LocalizedRes.ReportFormat_Inch
            Case AeccLandLib.AeccFormatType.aeccFormatDegreeMinuteSecondSpaced
                tagDegree = Chr(176) + " "
                tagMin = "' "
                tagSec = LocalizedRes.ReportFormat_Inch
            Case Else 'AeccLandLib.AeccFormatType.aeccFormatDecimal

        End Select

        Dim degreeString As String = ""
        Dim minuteString As String = ""
        Dim secondString As String = ""


        formatDMSStrings(value, prec, rounding, degreeString, minuteString, secondString)

        If degreeString <> "" Then
            resultDMS += degreeString + tagDegree
        End If

        If minuteString <> "" Then
            resultDMS += minuteString + tagMin
        End If

        If secondString <> "" Then
            resultDMS += secondString + tagSec
        End If

        Return resultDMS

    End Function

    Private Shared Sub formatDMSStrings(ByVal angelInDegree As Double, ByVal prec As Integer, _
                                        ByVal rounding As AeccLandLib.AeccRoundingType, _
                                        ByRef degString As String, _
                                        ByRef minString As String, _
                                        ByRef secString As String)
        Dim remainingPrec As Integer = 0
        Dim factor As Integer = calcDMSRoundingFactor(prec, remainingPrec)

        Dim roundedVal As Double = ReportFormat.RoundDouble(angelInDegree * factor, remainingPrec, rounding)

        Dim centSecVal = roundedVal - Math.Floor(roundedVal) ' the fraction of the second

        prec -= remainingPrec
        Debug.Assert(prec <= 4)

        If prec = 1 Or prec = 3 Then
            prec += 1 ' make counting easier for only counting prec 0, 2, 4, namely deg, min, sec
            roundedVal *= 10
        End If

        ' Store Degree Minute Second
        Dim angleVec() As Double = New Double() {Double.NaN, Double.NaN, Double.NaN}

        Dim remainder As Double = roundedVal
        Dim loopCount As Integer = CInt(prec / 2)
        For index As Integer = 0 To loopCount
            Dim fac As Integer = CInt(60 ^ (loopCount - index)) ' can be 3600, 60 or 1
            angleVec(index) = Math.Floor(remainder / fac)
            remainder -= angleVec(index) * fac
        Next

        If Double.IsNaN(angleVec(0)) <> True Then
            degString = angleVec(0).ToString("0")
        End If

        If Double.IsNaN(angleVec(1)) <> True Then
            minString = angleVec(1).ToString("00")
        End If

        If Double.IsNaN(angleVec(2)) <> True Then
            Dim customFormator As String = ""
            For count As Integer = 1 To remainingPrec
                customFormator &= "0"
            Next
            Dim secVal As Double = angleVec(2) + centSecVal
            secString = secVal.ToString("00." + customFormator)
        End If

    End Sub
    Private Shared Sub convertFromRadians(ByRef value As Double, ByVal toType As AeccLandLib.AeccAngleUnitType)
        Select Case toType
            Case Autodesk.AECC.Interop.Land.AeccAngleUnitType.aeccAngleUnitDegree
                value *= 180 / Math.PI
            Case Autodesk.AECC.Interop.Land.AeccAngleUnitType.aeccAngleUnitGrad
                value *= 200 / Math.PI
            Case Else
                System.Diagnostics.Debug.Assert(toType = Autodesk.AECC.Interop.Land.AeccAngleUnitType.aeccAngleUnitRadian)
        End Select
    End Sub

    Private Shared Sub convertFromDegrees(ByRef value As Double, ByVal toType As AeccLandLib.AeccAngleUnitType)

        Select Case toType
            Case Autodesk.AECC.Interop.Land.AeccAngleUnitType.aeccAngleUnitRadian
                value *= Math.PI / 180
            Case Autodesk.AECC.Interop.Land.AeccAngleUnitType.aeccAngleUnitGrad
                value *= 200 / 180
            Case Else
                Diagnostics.Debug.Assert(toType = AeccLandLib.AeccAngleUnitType.aeccAngleUnitDegree)
        End Select
    End Sub

    Private Shared Sub convertFromGrads(ByRef value As Double, ByVal toType As AeccLandLib.AeccAngleUnitType)

        Select Case toType
            Case Autodesk.AECC.Interop.Land.AeccAngleUnitType.aeccAngleUnitRadian
                value *= Math.PI / 200
            Case Autodesk.AECC.Interop.Land.AeccAngleUnitType.aeccAngleUnitDegree
                value *= 180 / 200
            Case Else
                Diagnostics.Debug.Assert(toType = AeccLandLib.AeccAngleUnitType.aeccAngleUnitGrad)
        End Select
    End Sub

    ''' <summary>
    ''' convertUnits is used to convert a double value
    ''' from one type of angular units to another (e.g., radians to degrees,
    ''' grads to degrees, degrees to grads, and so on.)  The converted value
    ''' is returned by the method.
    ''' </summary>
    ''' <param name="value">numeric angular value to be converted </param>
    ''' <param name="fromType">current unit type of angular value</param>
    ''' <param name="toType">desired unit type of angular value</param>
    ''' <returns>The converted value</returns>
    ''' <remarks></remarks>
    Public Shared Function convertUnits( _
        ByVal value As Double, _
        ByVal fromType As AeccLandLib.AeccAngleUnitType, _
        ByVal toType As AeccLandLib.AeccAngleUnitType) As Double

        If (fromType <> toType) Then
            Select Case fromType
                Case Autodesk.AECC.Interop.Land.AeccAngleUnitType.aeccAngleUnitRadian
                    convertFromRadians(value, toType)
                Case Autodesk.AECC.Interop.Land.AeccAngleUnitType.aeccAngleUnitDegree
                    convertFromDegrees(value, toType)
                Case Autodesk.AECC.Interop.Land.AeccAngleUnitType.aeccAngleUnitGrad
                    convertFromGrads(value, toType)
                Case Else
                    Diagnostics.Debug.Assert(False)
            End Select
        End If

        Return value
    End Function

    Public Shared Function FormatSouradnice(ByVal dDis As Double, Optional ByVal disUnit As AeccLandLib.AeccCoordinateUnitType = AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitFoot, _
                                        Optional ByVal disPrec As Integer = 6, _
                                        Optional ByVal disRounding As AeccLandLib.AeccRoundingType = AeccLandLib.AeccRoundingType.aeccRoundingNormal, _
                                        Optional ByVal disSign As AeccLandLib.AeccSignType = AeccLandLib.AeccSignType.aeccSignNegative, _
                                        Optional ByVal unitIndicator As Boolean = True) As String
        Dim sPreSign As String = "", _
            sPostSign As String = ""
        Select Case disSign
            Case AeccLandLib.AeccSignType.aeccSignAlways
                If dDis < 0.0# Then
                    sPreSign = "-"
                Else
                    sPreSign = "+"
                End If
            Case AeccLandLib.AeccSignType.aeccSignBracketNegative
                If dDis < 0.0# Then
                    sPreSign = "("
                    sPostSign = ")"
                End If
            Case AeccLandLib.AeccSignType.aeccSignNegative
                If dDis < 0.0# Then
                    sPreSign = "-"
                End If
        End Select
        dDis = Math.Abs(dDis)

        'trans dDis to aeccCoordinateUnitMeter
        Dim disOriUnit As AeccLandLib.AeccDrawingUnitType
        disOriUnit = ReportApplication.AeccXDatabase.Settings.DrawingSettings.UnitZoneSettings.DrawingUnits
        If disOriUnit = AeccLandLib.AeccDrawingUnitType.aeccDrawingUnitFeet Then
            dDis = dDis * 0.3048006096012
        End If

        Dim sDis As String = ""
        Select Case disUnit
            Case AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitMeter
                dDis = RoundDouble(dDis, disPrec, disRounding)
                sDis = CStr(dDis)
                If unitIndicator = True Then
                    sDis = sDis '+ "m"
                End If

            Case AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitKilometer
                dDis = dDis / 1000.0#
                dDis = RoundDouble(dDis, disPrec, disRounding)
                sDis = CStr(dDis)
                If unitIndicator = True Then
                    sDis = sDis '+ "km"
                End If

            Case AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitFoot
                dDis = dDis / 0.3048006096012
                dDis = RoundDouble(dDis, disPrec, disRounding)
                sDis = CStr(dDis)
                If unitIndicator = True Then
                    sDis = sDis + "'"
                End If

            Case AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitInch
                dDis = dDis * 12.0# / 0.3048006096012
                dDis = RoundDouble(dDis, disPrec, disRounding)
                sDis = CStr(dDis)
                If unitIndicator = True Then
                    sDis = sDis + """"
                End If

            Case AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitYard
                dDis = dDis / 3.0# / 0.3048006096012
                dDis = RoundDouble(dDis, disPrec, disRounding)
                sDis = CStr(dDis)
                If unitIndicator = True Then
                    sDis = sDis + "yd"
                End If

            Case AeccLandLib.AeccCoordinateUnitType.aeccCoordinateUnitMile
                dDis = dDis / 1760.0# / 3.0# / 0.3048006096012
                dDis = RoundDouble(dDis, disPrec, disRounding)
                sDis = CStr(dDis)
                If unitIndicator = True Then
                    sDis = sDis + "miles"
                End If

        End Select
        FormatSouradnice = sPreSign + sDis + sPostSign
    End Function
End Class
