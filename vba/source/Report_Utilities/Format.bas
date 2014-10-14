Attribute VB_Name = "Format"
'-----------------------------------------------------------------------
'   set rawAngle(radian) into [0,2pi].
'-----------------------------------------------------------------------
Function AdjustAngle2PI(ByVal rawAngle As Double) As Double
    pi = 4 * Atn(1)
    Do While rawAngle >= pi * 2
        rawAngle = rawAngle - pi * 2
    Loop
    Do While rawAngle < 0#
        rawAngle = rawAngle + pi * 2
    Loop
    AdjustAngle2PI = rawAngle
End Function

'-----------------------------------------------------------------------
'   rounds off values and returns a string.
'-----------------------------------------------------------------------
Function RoundVal(value As Double, prec As Integer, Optional roundType As AeccRoundingType = aeccRoundingNormal) As String
    If Abs(value) < 10000000000# Then
        Dim dNewValue As Double
        dNewValue = value
        Dim sNewValue As String
        Select Case prec
            Case 0
                sNewValue = VBA.Format(dNewValue, "#0")
            Case 1
                sNewValue = VBA.Format(dNewValue, "#0.0")
            Case 2
                sNewValue = VBA.Format(dNewValue, "#0.00")
            Case 3
                sNewValue = VBA.Format(dNewValue, "#0.000")
            Case 4
                sNewValue = VBA.Format(dNewValue, "#0.0000")
            Case 5
                sNewValue = VBA.Format(dNewValue, "#0.00000")
            Case 6
                sNewValue = VBA.Format(dNewValue, "#0.000000")
            Case 7
                sNewValue = VBA.Format(dNewValue, "#0.0000000")
            Case 8
                sNewValue = VBA.Format(dNewValue, "#0.00000000")
        End Select
        
        If roundType <> aeccRoundingNormal Then
            dNewValue = CDbl(sNewValue)
            Dim dLastBit As Double
            dLastBit = 10 ^ (-1# * prec)
            If roundType = aeccRoundingTruncate Then
                If Abs(dNewValue) > Abs(value) Then
                    dNewValue = Abs(dNewValue) - dLastBit
                    If value < 0 Then
                        dNewValue = -1# * dNewValue
                    End If
                End If
            ElseIf roundType = aeccRoundingUp Then
                If Abs(dNewValue) < Abs(value) Then
                    dNewValue = Abs(dNewValue) + dLastBit
                    If value < 0 Then
                        dNewValue = -1# * dNewValue
                    End If
                End If
            End If
            sNewValue = CStr(dNewValue)
        End If
        RoundVal = sNewValue
    Else
        RoundVal = VBA.Format(value, "0.00000000000000E+00")
    End If
End Function
'-----------------------------------------------------------------------
'   format a grade
'   e.g.:  1:3.232
'   Note: given the originate gradeFormatType aeccGradeSlopeFormatDecimal
'         this function don't take care aeccGradeSlopeFormatRiseRun & aeccGradeSlopeFormatRunRise,which will be in FormatSlope
'-----------------------------------------------------------------------
Function FormatGrade(dGrade As Double, Optional grdFormat As AeccGradeSlopeFormatType = aeccGradeSlopeFormatPercent, _
                                       Optional grdPrec As Integer = 3, _
                                       Optional grdRounding As AeccRoundingType = aeccRoundingNormal, _
                                       Optional grdSign As AeccSignType = aeccSignNegative) As String
    Dim sPreSign As String, sPostSign As String
    Select Case grdSign
        Case aeccSignAlways
            If dGrade < 0# Then
                sPreSign = "-"
            Else
                sPreSign = "+"
            End If
        Case aeccSignBracketNegative
            If dGrade < 0# Then
                sPreSign = "("
                sPostSign = ")"
            End If
        Case aeccSignNegative
            If dGrade < 0# Then
                sPreSign = "-"
            End If
    End Select
    dGrade = Abs(dGrade)
    
    Dim sGrade As String
    Select Case grdFormat
        Case aeccGradeSlopeFormatDecimal
            dGrade = RoundVal(dGrade, grdPrec, grdRounding)
            sGrade = CStr(dGrade)
        Case aeccGradeSlopeFormatPercent
            dGrade = dGrade * 100#
            dGrade = RoundVal(dGrade, grdPrec, grdRounding)
            sGrade = CStr(dGrade) + "%"
    End Select
    
    FormatGrade = sPreSign + sGrade + sPostSign
End Function

'-----------------------------------------------------------------------
'   format a volume
'   e.g.:  +35.34Cu.Yd. +35.34Cu.M. +35.34Cu.Ft.
'   Note: the originate unitType depends on drawing's AeccSettingsUnitZone
'-----------------------------------------------------------------------
Function FormatArea(dArea As Double, Optional areaUnit As AeccAreaUnitType = aeccAreaUnitSquareFoot, _
                                        Optional volPrec As Integer = 4, _
                                        Optional volRounding As AeccRoundingType = aeccRoundingNormal, _
                                        Optional volSign As AeccSignType = aeccSignNegative) As String
    
    Dim sPreSign As String, sPostSign As String
    Select Case grdSign
        Case aeccSignAlways
            If dArea < 0# Then
                sPreSign = "-"
            Else
                sPreSign = "+"
            End If
        Case aeccSignBracketNegative
            If dArea < 0# Then
                sPreSign = "("
                sPostSign = ")"
            End If
        Case aeccSignNegative
            If dArea < 0# Then
                sPreSign = "-"
            End If
    End Select
    dArea = Abs(dArea)
    'trans dDis to aeccCoordinateUnitMeter
    Dim disOriUnit As AeccDrawingUnitType
    disOriUnit = getAeccDb.settings.DrawingSettings.UnitZoneSettings.DrawingUnits
    If disOriUnit = aeccDrawingUnitFeet Then
        dArea = dArea * (0.3048006096012 ^ 2)
    End If
    
    Dim sArea As String
    Select Case areaUnit
        Case aeccAreaUnitSquareMeter  'aeccVolumeUnitCubicMeter
            dArea = RoundVal(dArea, volPrec, volRounding)
            sArea = CStr(dArea) + "Sq.M."
        Case aeccAreaUnitSquareFoot
            dArea = dArea / (0.3048006096012 ^ 2)
            dArea = RoundVal(dArea, volPrec, volRounding)
            sArea = CStr(dArea) + "Sq.Ft."
        Case aeccAreaUnitAcre
            dArea = dArea * 0.0002471
            dArea = RoundVal(dArea, volPrec, volRounding)
            sArea = CStr(dArea) + "Acre"
        Case aeccAreaUnitHectare
            dArea = dArea * 0.0001
            dArea = RoundVal(dArea, volPrec, volRounding)
            sArea = CStr(dArea) + "Hectare"
        Case aeccAreaUnitSquareKilometer
            dArea = dArea * 0.000001
            dArea = RoundVal(dArea, volPrec, volRounding)
            sArea = CStr(dArea) + "Sq.Km."
        Case aeccAreaUnitSquareMile
            dArea = dArea * 0.000000386
            dArea = RoundVal(dArea, volPrec, volRounding)
            sArea = CStr(dArea) + "Sq.Mile"
        Case aeccAreaUnitSquareYard
            dArea = dArea * 1.196
            dArea = RoundVal(dArea, volPrec, volRounding)
            sArea = CStr(dArea) + "Sq.Yd."
    End Select
    FormatArea = sPreSign + sArea + sPostSign
End Function


'-----------------------------------------------------------------------
'   format a volume
'   e.g.:  +35.34Cu.Yd. +35.34Cu.M. +35.34Cu.Ft.
'   Note: the originate unitType depends on drawing's AeccSettingsUnitZone
'-----------------------------------------------------------------------
Function FormatVolume(dVol As Double, Optional volUnit As AeccVolumeUnitType = aeccVolumeUnitCubicFoot, _
                                        Optional volPrec As Integer = 4, _
                                        Optional volRounding As AeccRoundingType = aeccRoundingNormal, _
                                        Optional volSign As AeccSignType = aeccSignNegative) As String
    Dim sPreSign As String, sPostSign As String
    Select Case grdSign
        Case aeccSignAlways
            If dVol < 0# Then
                sPreSign = "-"
            Else
                sPreSign = "+"
            End If
        Case aeccSignBracketNegative
            If dVol < 0# Then
                sPreSign = "("
                sPostSign = ")"
            End If
        Case aeccSignNegative
            If dVol < 0# Then
                sPreSign = "-"
            End If
    End Select
    dVol = Abs(dVol)
    
    'trans dDis to aeccCoordinateUnitMeter
    Dim disOriUnit As AeccDrawingUnitType
    disOriUnit = getAeccDb.settings.DrawingSettings.UnitZoneSettings.DrawingUnits
    If disOriUnit = aeccDrawingUnitFeet Then
        dVol = dVol * (0.3048006096012 ^ 3)
    End If
    
    Dim sVol As String
    Select Case volUnit
        Case aeccVolumeUnitCubicMeter
            dVol = RoundVal(dVol, volPrec, volRounding)
            sVol = CStr(dVol) + "Cu.M."
        Case aeccVolumeUnitCubicFoot
            dVol = dVol / (0.3048006096012 ^ 3)
            dVol = RoundVal(dVol, volPrec, volRounding)
            sVol = CStr(dVol) + "Cu.Ft"
        Case aeccVolumeUnitCubicInch
            dVol = dVol * ((12# / 0.3048006096012) ^ 3)
            dVol = RoundVal(dVol, volPrec, volRounding)
            sVol = CStr(dVol) + "Cu.Inch"
        Case aeccVolumeUnitCubicYard
            dVol = dVol / ((3# * 0.3048006096012) ^ 3)
            dVol = RoundVal(dVol, volPrec, volRounding)
            sVol = CStr(dVol) + "Cu.Yd"
    End Select
    FormatVolume = sPreSign + sVol + sPostSign
End Function

'-----------------------------------------------------------------------
'   format a distance
'   e.g.:  +35.34(f)
'   Note: the originate unitType depends on drawing's AeccSettingsUnitZone
'-----------------------------------------------------------------------
Function FormatDistance(dDis As Double, Optional disUnit As AeccCoordinateUnitType = aeccCoordinateUnitFoot, _
                                        Optional disPrec As Integer = 6, _
                                        Optional disRounding As AeccRoundingType = aeccRoundingNormal, _
                                        Optional disSign As AeccSignType = aeccSignNegative, _
                                        Optional unitIndicator As Boolean = True) As String
    Dim sPreSign As String, sPostSign As String
    Select Case disSign
        Case aeccSignAlways
            If dDis < 0# Then
                sPreSign = "-"
            Else
                sPreSign = "+"
            End If
        Case aeccSignBracketNegative
            If dDis < 0# Then
                sPreSign = "("
                sPostSign = ")"
            End If
        Case aeccSignNegative
            If dDis < 0# Then
                sPreSign = "-"
            End If
    End Select
    dDis = Abs(dDis)
    
    'trans dDis to aeccCoordinateUnitMeter
    Dim disOriUnit As AeccDrawingUnitType
    disOriUnit = getAeccDb.settings.DrawingSettings.UnitZoneSettings.DrawingUnits
    If disOriUnit = aeccDrawingUnitFeet Then
        dDis = dDis * 0.3048006096012
    End If
    
    Dim sDis As String
    Select Case disUnit
        Case aeccCoordinateUnitMeter
            dDis = RoundVal(dDis, disPrec, disRounding)
            sDis = CStr(dDis)
            If unitIndicator = True Then
                sDis = sDis + "m"
            End If
            
        Case aeccCoordinateUnitKilometer
            dDis = dDis / 1000#
            dDis = RoundVal(dDis, disPrec, disRounding)
            sDis = CStr(dDis)
            If unitIndicator = True Then
                sDis = sDis + "km"
            End If
            
        Case aeccCoordinateUnitFoot
            dDis = dDis / 0.3048006096012
            dDis = RoundVal(dDis, disPrec, disRounding)
            sDis = CStr(dDis)
            If unitIndicator = True Then
                sDis = sDis + "'"
            End If
            
        Case aeccCoordinateUnitInch
            dDis = dDis * 12# / 0.3048006096012
            dDis = RoundVal(dDis, disPrec, disRounding)
            sDis = CStr(dDis)
            If unitIndicator = True Then
                sDis = sDis + """"
            End If
            
        Case aeccCoordinateUnitYard
            dDis = dDis / 3# / 0.3048006096012
            dDis = RoundVal(dDis, disPrec, disRounding)
            sDis = CStr(dDis)
            If unitIndicator = True Then
                sDis = sDis + "yd"
            End If
            
        Case aeccCoordinateUnitMile
            dDis = dDis / 1760# / 3# / 0.3048006096012
            dDis = RoundVal(dDis, disPrec, disRounding)
            sDis = CStr(dDis)
            If unitIndicator = True Then
                sDis = sDis + "miles"
            End If
            
    End Select
    FormatDistance = sPreSign + sDis + sPostSign
End Function

'-----------------------------------------------------------------------
'   format a radian angle
'   e.g.:  65^34'32.32"
'   Note:  do not take care dRadAng's scope
'-----------------------------------------------------------------------
Function FormatAngle(dRadAng As Double, Optional angUnit As AeccAngleUnitType = aeccAngleUnitDegree, Optional angPrec As Integer = 6, _
                                        Optional angRounding As AeccRoundingType = aeccRoundingNormal, Optional angFormat As AeccFormatType = aeccFormatDegreeMinuteSecond, _
                                        Optional angSign As AeccSignType = aeccSignNegative) As String
    Dim sPreSign As String, sPostSign As String
    Dim pi As Double
    pi = 4 * Atn(1)
    Select Case angSign
        Case aeccSignAlways
            If dRadAng < 0# Then
                sPreSign = "-"
            Else
                sPreSign = "+"
            End If
        Case aeccSignBracketNegative
            If dRadAng < 0# Then
                sPreSign = "("
                sPostSign = ")"
            End If
        Case aeccSignNegative
            If dRadAng < 0# Then
                sPreSign = "-"
            End If
    End Select
    dRadAng = Abs(dRadAng)
    
    Dim sAngle As String
    Select Case angUnit
        Case aeccAngleUnitDegree
            Dim dAng As Double
            dAng = dRadAng * 180# / pi
            Dim iDeg As Integer, iMin As Integer, iSec As Integer
            Dim iCentSec As Integer

            If angFormat <> aeccFormatDecimal Then
                iDeg = Int(dAng)
                dAng = (dAng - iDeg) * 60
                iMin = Int(dAng)
                dAng = (dAng - iMin) * 60
                iSec = Int(dAng)
                dAng = dAng - iSec
                Dim iPrec As Integer
                iPrec = angPrec - 4
                Do While iPrec > 0
                        dAng = dAng * 10
                        iPrec = iPrec - 1
                Loop
                iCentSec = Int(dAng)
                
                Dim sDeg As String, sMin As String, sSec As String, sCentSec As String
                sDeg = CStr(iDeg)
                sMin = CStr(iMin)
                If Len(sMin) = 1 Then
                    sMin = "0" & sMin
                End If
                sSec = CStr(iSec)
                If Len(sSec) = 1 Then
                    sSec = "0" & sSec
                End If
                sCentSec = CStr(iCentSec)
            Else
                dAng = RoundVal(dAng, angPrec, angRounding)
            End If
                            
            Select Case angFormat
                Case aeccFormatDecimal
                    sAngle = CStr(dAng) + " (d)"
                Case aeccFormatDecimalDegreeMinuteSecond
                    sAngle = sDeg & "." & sMin & sSec & sCnetSec & " (dms)"
                Case aeccFormatDegreeMinuteSecond
                    sAngle = sDeg & Chr(176) & sMin & "'" & sSec & "." & sCentSec & """"
                Case aeccFormatDegreeMinuteSecondSpaced
                    sAngle = sDeg & Chr(176) & " " & sMin & "' " & sSec & "." & sCentSec & """"
            End Select
        Case aeccAngleUnitGrad
            dRadAng = dRadAng * 200# / pi
            dRadAng = RoundVal(dRadAng, angPrec, angRounding)
            sAngle = CStr(dRadAng) + " (g)"
        Case aeccAngleUnitRadian
            dRadAng = dRadAng
            dRadAng = RoundVal(dRadAng, angPrec, angRounding)
            sAngle = CStr(dRadAng) + " (r)"
    End Select 'angUnit
    FormatAngle = sPreSign + sAngle + sPostSign
End Function

'-----------------------------------------------------------------------
'   format a radian angle to direction.
'   e.g.:  N65^34'32.32"E
'   Note:  dRadAng is start from x-axis and rotate counterclockwise
'-----------------------------------------------------------------------
Function FormatDirection(dRadAng As Double, Optional angUnit As AeccAngleUnitType = aeccAngleUnitDegree, Optional angPrec As Integer = 6, _
                                            Optional angRounding As AeccRoundingType = aeccRoundingNormal, Optional angFormat As AeccFormatType = aeccFormatDegreeMinuteSecond, _
                                            Optional angDirection As AeccDirectionType = aeccDirectionShortName, Optional angCap As AeccCapitalizationType = aeccCapitalizationUpperCase, _
                                            Optional angSign As AeccSignType = aeccSignNegative, Optional angMeasType As AeccMeasurementType = aeccMeasurementBearings, _
                                            Optional angBearQuand As AeccBearingQuadrantType = aeccBearingQuadrantNorthEast) As String
    'step 1: get proper dRadAng considering angMeasType and preFixString/postFixString
    Dim pi As Double
    pi = 4 * Atn(1)
    dRadAng = AdjustAngle2PI(dRadAng)
    dRadAng = dRadAng * 180# / pi
    
    Dim sPrefix As String, sPostfix As String
    Select Case angMeasType
        Case aeccMeasurementBearings
            If dRadAng >= 90 And dRadAng < 180 Then
                dRadAng = dRadAng - 90#
                Select Case angDirection
                    Case aeccDirectionLongName
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "north": sPostfix = "west"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "NORTH": sPostfix = "WEST"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "North": sPostfix = "West"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "NORTH": sPostfix = "WEST"
                        End Select
                    Case aeccDirectionShortName
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "n": sPostfix = "w"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "N": sPostfix = "W"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "N": sPostfix = "W"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "N": sPostfix = "W"
                        End Select
                    Case aeccDirectionLongNameSpaced
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "north ": sPostfix = " west"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "NORTH ": sPostfix = " WEST"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "North ": sPostfix = " West"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "NORTH ": sPostfix = " WEST"
                        End Select
                    Case aeccDirectionShortNameSpaced
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "n ": sPostfix = " w"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "N ": sPostfix = " W"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "N ": sPostfix = " W"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "N ": sPostfix = " W"
                        End Select
                End Select 'angDirection
                
            ElseIf dRadAng >= 180 And dRadAng < 270 Then
                dRadAng = 270# - dRadAng
                Select Case angDirection
                    Case aeccDirectionLongName
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "south": sPostfix = "west"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "SOUTH": sPostfix = "WEST"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "South": sPostfix = "West"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "SOUTH": sPostfix = "WEST"
                        End Select
                    Case aeccDirectionShortName
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "s": sPostfix = "w"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "S": sPostfix = "W"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "S": sPostfix = "W"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "S": sPostfix = "W"
                        End Select
                    Case aeccDirectionLongNameSpaced
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "south ": sPostfix = " west"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "SOUTH ": sPostfix = " WEST"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "South ": sPostfix = " West"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "SOUTH ": sPostfix = " WEST"
                        End Select
                    Case aeccDirectionShortNameSpaced
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "s ": sPostfix = " w"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "S ": sPostfix = " W"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "S ": sPostfix = " W"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "S ": sPostfix = " W"
                        End Select
                End Select 'angDirection
                
            ElseIf dRadAng >= 270 And dRadAng <= 360 Then
                dRadAng = dRadAng - 270#
                Select Case angDirection
                    Case aeccDirectionLongName
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "south": sPostfix = "east"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "SOUTH": sPostfix = "EAST"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "South": sPostfix = "East"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "SOUTH": sPostfix = "EAST"
                        End Select
                    Case aeccDirectionShortName
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "s": sPostfix = "e"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "S": sPostfix = "E"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "S": sPostfix = "E"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "S": sPostfix = "E"
                        End Select
                    Case aeccDirectionLongNameSpaced
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "south ": sPostfix = " east"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "SOUTH ": sPostfix = " EAST"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "South ": sPostfix = " East"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "SOUTH ": sPostfix = " EAST"
                        End Select
                    Case aeccDirectionShortNameSpaced
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "s ": sPostfix = " e"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "S ": sPostfix = " E"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "S ": sPostfix = " E"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "S ": sPostfix = " E"
                        End Select
                End Select 'angDirection
                
            Else
                dRadAng = 90# - dRadAng
                Select Case angDirection
                    Case aeccDirectionLongName
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "north": sPostfix = "east"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "NORTH": sPostfix = "EAST"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "North": sPostfix = "East"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "NORTH": sPostfix = "EAST"
                        End Select
                    Case aeccDirectionShortName
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "n": sPostfix = "e"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "N": sPostfix = "E"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "N": sPostfix = "E"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "N": sPostfix = "E"
                        End Select
                    Case aeccDirectionLongNameSpaced
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "north ": sPostfix = " east"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "NORTH ": sPostfix = " EAST"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "North ": sPostfix = " East"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "NORTH ": sPostfix = " EAST"
                        End Select
                    Case aeccDirectionShortNameSpaced
                        Select Case angCap
                            Case aeccCapitalizationLowerCase
                                sPrefix = "n ": sPostfix = " e"
                            Case aeccCapitalizationPreserveCase
                                sPrefix = "N ": sPostfix = " E"
                            Case aeccCapitalizationTitleCaps
                                sPrefix = "N ": sPostfix = " E"
                            Case aeccCapitalizationUpperCase
                                sPrefix = "N ": sPostfix = " E"
                        End Select
                End Select 'angDirection
            End If
        'Case aeccMeasurementBearings End
        Case aeccMeasurementNorthAzimuth
            dRadAng = 450 - dRadAng
            If dRadAng >= 360 Then
                dRadAng = dRadAng - 360
            End If
            sPrefix = "N "
        Case aeccMeasurementSouthAzimuth
            dRadAng = 270 - dRadAng
            If dRadAng <= 0 Then
                dRadAng = dRadAng + 360
            End If
            sPrefix = "S "
    End Select 'angMeasType
    
    'step 2: get dRadAng's string
    Dim sAngle As String
    Select Case angUnit
        Case aeccAngleUnitDegree
            Dim dAng As Double
            dAng = dRadAng
            Dim iDeg As Integer, iMin As Integer, iSec As Integer
            Dim iCentSec As Integer

            If angFormat <> aeccFormatDecimal Then
                iDeg = Int(dAng)
                dAng = (dAng - iDeg) * 60
                iMin = Int(dAng)
                dAng = (dAng - iMin) * 60
                iSec = Int(dAng)
                dAng = dAng - iSec
                Dim iPrec As Integer
                iPrec = angPrec - 4
                Do While iPrec > 0
                        dAng = dAng * 10
                        iPrec = iPrec - 1
                Loop
                iCentSec = Int(dAng)
                
                Dim sDeg As String, sMin As String, sSec As String, sCentSec As String
                sDeg = CStr(iDeg)
                sMin = CStr(iMin)
                If Len(sMin) = 1 Then
                    sMin = "0" & sMin
                End If
                sSec = CStr(iSec)
                If Len(sSec) = 1 Then
                    sSec = "0" & sSec
                End If
                sCentSec = CStr(iCentSec)
            Else
                dAng = RoundVal(dAng, angPrec, angRounding)
            End If
                            
            Select Case angFormat
                Case aeccFormatDecimal
                    sAngle = CStr(dAng)
                    sPostfix = sPostfix + " (d)"
                Case aeccFormatDecimalDegreeMinuteSecond
                    sAngle = sDeg & "." & sMin & sSec & sCnetSec
                    sPostfix = sPostfix + " (dms)"
                Case aeccFormatDegreeMinuteSecond
                    sAngle = sDeg & Chr(176) & sMin & "'" & sSec & "." & sCentSec & """"
                Case aeccFormatDegreeMinuteSecondSpaced
                    sAngle = sDeg & Chr(176) & " " & sMin & "' " & sSec & "." & sCentSec & """"
            End Select
        Case aeccAngleUnitGrad
            dRadAng = dRadAng * 400 / 360
            dRadAng = RoundVal(dRadAng, angPrec, angRounding)
            sAngle = CStr(dRadAng)
            sPostfix = sPostfix + " (g)"
        Case aeccAngleUnitRadian
            dRadAng = dRadAng * pi / 360
            dRadAng = RoundVal(dRadAng, angPrec, angRounding)
            sAngle = CStr(dRadAng)
            sPostfix = sPostfix + " (r)"
    End Select 'angUnit
    FormatDirection = sPrefix + sAngle + sPostfix
End Function


