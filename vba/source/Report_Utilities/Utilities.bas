Attribute VB_Name = "Utilities"
Enum angleType
    TurnedPlus
    TurnedMinus
    DeflectPlus
    DeflectMinus
    Direction
End Enum
Enum rpAngDispType
    cDecDeg
    cDMS
    cGrad
    cRad
End Enum

Enum rpAngleType
    Bearing
    nAzimuth
    sAzimuth
End Enum

'Function AdjustAngle2PI(ByVal rawAngle As Double) As Double
'    pi = 4 * Atn(1)
'    Do While rawAngle >= pi * 2
'        rawAngle = rawAngle - pi * 2
'    Loop
'    Do While rawAngle < 0#
'        rawAngle = rawAngle + pi * 2
'    Loop
'    AdjustAngle2PI = rawAngle
'End Function
'
''-----------------------------------------------------------------------
''   rounds off values and returns a string.
''-----------------------------------------------------------------------
'Function RoundVal(value As Double, prec As Integer) As String
'    If Abs(value) < 10000000000# Then
'        Select Case prec
'        Case 0
'            RoundVal = Format(value, "#0")
'        Case 1
'            RoundVal = Format(value, "#0.0")
'        Case 2
'            RoundVal = Format(value, "#0.00")
'        Case 3
'            RoundVal = Format(value, "#0.000")
'        Case 4
'            RoundVal = Format(value, "#0.0000")
'        Case 5
'            RoundVal = Format(value, "#0.00000")
'        Case 6
'            RoundVal = Format(value, "#0.000000")
'        Case 7
'            RoundVal = Format(value, "#0.0000000")
'        Case 8
'            RoundVal = Format(value, "#0.00000000")
'
'        End Select
'    Else
'        RoundVal = Format(value, "0.00000000000000E+00")
'    End If
'End Function

'Function FormatAngle(dRadian As Double, AngType As rpAngDispType, Optional lRnd As Variant) As String
'
'    Dim dAng As Double
'    Dim bNeg As Boolean
'    Dim iDeg As Integer, iMin As Integer
'    Dim dSec As Double
'    Dim sDeg As String, sMin As String, sSec As String
'    Dim pi As Double
'    pi = 4 * Atn(1)
'
'    ' work with positive values, then add the neg sign after formating
'    If dRadian < 0 Then
'        bNeg = True
'    Else
'        bNeg = False
'    End If
'    dRadian = Abs(dRadian)
'
'    'calc the angle
'    Select Case AngType
'        Case cDecDeg, cDMS
'            dAng = dRadian * 180 / pi
'        Case cGrad
'            dAng = dRadian * 200 / pi
'        Case cRad
'            dAng = dRadian
'    End Select
'
'
'    ' format the return string
'    Select Case AngType
'        ' format DDD.MMSS
'        Case cDMS
'            iDeg = Int(dAng)
'            dAng = (dAng - iDeg) * 60
'
'            iMin = Int(dAng)
'            dAng = (dAng - iMin) * 60
'
'            dSec = dAng
'            ' round the seconds if necessary
'            If TypeName(lRnd) = "Integer" Then
'                Dim lRndSec As Long
'                lRndSec = lRnd - 4 ' for MMSS
'                ' don't allow precision less then a full second
'                If lRndSec < 0 Then lRndSec = 0
'                dSec = Round(dSec, lRndSec)
'            End If
'
'            ' check for values rounded up
'
'            Do While iDeg >= 360
'                iDeg = iDeg - 360
'            Loop
'
'            'get the string parts
'            sDeg = CStr(iDeg)
'            sMin = CStr(iMin)
'            sSec = CStr(dSec)
'            ' pad minutes with a leading 0 if less then 10
'            If Len(sMin) = 1 Then
'                sMin = "0" & sMin
'            End If
'            ' pad seconds with a leading 0 if less then 10
'            If dSec < 10 Then
'                sSec = "0" & sSec
'            End If
'
'            ' set the return value
'            FormatAngle = sDeg & Chr(176) & sMin & "' " & sSec & """ "
'
'        Case Else
'        ' format all others
'            If TypeName(lRnd) = "Integer" Then
'                FormatAngle = Round(dAng, CLng(lRnd))
'            Else
'                FormatAngle = CStr(dAng)
'            End If
'    End Select
'
'    If bNeg = True Then
'        FormatAngle = "-" & FormatAngle
'    End If
'
'End Function

'' convert the direction of tangent to bearing angle mode
'' Parameter: radAngle is the North-Azimuth radian format
'Public Function FormatBearingAngle(radAngle As Double, dispType As rpAngDispType, Optional lRnd As Variant) As String
'    Dim North, East, South, West
'    Dim sPrefix, sPostfix
'    Dim pi As Double
'    pi = 4 * Atn(1)
'    North = "N"
'    East = "E"
'    South = "S"
'    West = "W"
'
'    ' make sure the radAngle is in [0,2pi)
'    radAngle = AdjustAngle2PI(radAngle)
'
'
'    'Calculate
'    If radAngle <= pi / 2 Then
'        sPrefix = North
'        sPostfix = East
'    ElseIf radAngle <= pi Then
'        radAngle = pi - radAngle
'        sPrefix = South
'        sPostfix = East
'    ElseIf radAngle <= 1.5 * pi Then
'        radAngle = radAngle - pi
'        sPrefix = South
'        sPostfix = West
'    Else
'        radAngle = 2 * pi - radAngle
'        sPrefix = North
'        sPostfix = West
'    End If
'
'
'    FormatBearingAngle = sPrefix & " " & FormatAngle(radAngle, dispType, lRnd) & "" & sPostfix
'
'
''End Function
'' convert the direction of tangent to  Azimuth angle mode
'' Parameter: radAngle is the North-Azimuth radian format
'Public Function FormatAzimuthAngle(radAngle As Double, bNorth As Boolean, dispType As rpAngDispType, Optional lRnd As Variant) As String
'    Dim pi As Double
'    pi = 4 * Atn(1)
'
'    If bNorth = False Then 'South Azimuth
'        radAngle = radAngle - pi
'    End If
'
'    ' make sure the radAngle is in [0,pi)
'    radAngle = AdjustAngle2PI(radAngle)
'
'
'    FormatAzimuthAngle = FormatAngle(radAngle, dispType, lRnd)
'
'
'End Function


'-----------------------------------------------------------------------
'   calc the distance between two northing/easting points
'-----------------------------------------------------------------------
Function CalcDist(sNorth As Double, sEast As Double, eNorth As Double, eEast As Double) As Double
    
    CalcDist = (Abs(sEast - eEast) ^ 2 + Abs(sNorth - eNorth) ^ 2) ^ 0.5
    
End Function
'-----------------------------------------------------------------------
'   calc the direction in radians between two points
'-----------------------------------------------------------------------
' returned radAngle is the North-Azimuth radian format
Function CalcDirRad(sNorth As Double, sEast As Double, eNorth As Double, eEast As Double) As Double

    Dim nDiff As Double
    Dim eDiff As Double
    
    Dim pi As Double
    pi = 4 * Atn(1)
    
    nDiff = eNorth - sNorth
    eDiff = eEast - sEast
    
    If eDiff = 0# And nDiff = 0# Then
        ' the points are at the same location. This should never happend, if it does return a 0.00 direction
        CalcDirRad = 0#
        Exit Function
        
    ' return the direction for horizontal/vertical directions
    ElseIf nDiff = 0# Then
        If eDiff > 0 Then
            CalcDirRad = 0.5 * pi
        Else
            CalcDirRad = 1.5 * pi
        End If
        Exit Function
        
    ElseIf eDiff = 0# Then
        If nDiff > 0 Then
            CalcDirRad = 0#
        Else
            CalcDirRad = pi
        End If
        Exit Function
        
    End If
        
    'calc the angle from the x axis
    CalcDirRad = Atn(Abs(nDiff / eDiff))

    ' calc the direction
    If nDiff > 0 And eDiff < 0 Then
        ' 270 to 360
        CalcDirRad = 1.5 * pi + CalcDirRad
    ElseIf nDiff < 0 And eDiff < 0 Then
        ' 180 to 270
        CalcDirRad = 1.5 * pi - CalcDirRad
    ElseIf nDiff < 0 And eDiff > 0 Then
        '90 to 180
        CalcDirRad = 0.5 * pi + CalcDirRad
    Else
        '0-90
        CalcDirRad = 0.5 * pi - CalcDirRad
    End If

    
End Function

'-----------------------------------------------------------------------
'   calc the 3 point angle, return a string formated based on the current drawing setup angular settings
'-----------------------------------------------------------------------
' returned radAngle is the North-Based radian format
' Parameter: oNorth the Northing value of occupied point
' Parameter: oEast the Easting value of occupied point
' Parameter: sNorth the Northing value of Station point
' Parameter: sNorth the Easting value of Station point
' Parameter: bNorth the Northing value of backsight point
' Parameter: bNorth the Easting value of backsight point
' Return value: in [0,2PI)
Function Calc3PointAngle(oNorth As Double, oEast As Double, _
                         sNorth As Double, sEast As Double, _
                         bNorth As Double, bEast As Double, _
                         anType As angleType) As Double
    Dim dirOtoS As Double
    Dim dirOtoB As Double
    Dim pi As Double
    pi = 4 * Atn(1)
    
    If oNorth = bNorth And oEast = bEast Then
        Calc3PointAngle = 0#
        Exit Function
    End If
        
    dirOtoS = CalcDirRad(oNorth, oEast, sNorth, sEast)
    dirOtoB = CalcDirRad(oNorth, oEast, bNorth, bEast)
    
    Calc3PointAngle = dirOtoS - dirOtoB ' for TurnedPlus
    If anType = TurnedPlus Then
        Calc3PointAngle = dirOtoS - dirOtoB
    ElseIf anType = TurnedMinus Then
        Calc3PointAngle = dirOtoB - dirOtoS
    ElseIf anType = DeflectPlus Then
        Calc3PointAngle = dirOtoS + dirOtoB
    ElseIf anType = DeflectMinus Then
        Calc3PointAngle = -dirOtoS - dirOtoB
    End If
    Calc3PointAngle = AdjustAngle2PI(Calc3PointAngle)
End Function

''-----------------------------------------------------------------------
''   calc the direction between two points, return a string formated based on the current drawing setup angular settings
''-----------------------------------------------------------------------
'Function CalcDir_(sNorth As Double, sEast As Double, eNorth As Double, eEast As Double, angPrec As Integer, angDisp As rpAngDispType, angleType As rpAngleType) As String
'
'    Dim quad As Integer
'    Dim nDiff As Double
'    Dim eDiff As Double
'
'    Dim ang As Double
'
'    Dim dir As Double
'    Dim deg As Integer
'    Dim min As Integer
'    Dim sec As Integer
'    Dim sDeg As String
'    Dim sMin As String
'    Dim sSec As String
'
'    Dim sN As String
'    Dim sS As String
'    Dim sE As String
'    Dim sW As String
'
'    sN = "N" 'GetResStr(1470)
'    sS = "S" 'GetResStr(1471)
'    sE = "E" 'GetResStr(1472)
'    sW = "W" 'GetResStr(1473)
'
'    ' define pi
'    Dim pi As Double
'    pi = 4 * Atn(1)
'
'    nDiff = eNorth - sNorth
'    eDiff = eEast - sEast
'
'  '  angPrec = ImportOpts.angPrec
'    If angDisp = cDMS Then
'        angPrec = angPrec - 4
'        If angPrec < 0 Then angPrec = 0
'    End If
'
'    ' format directions exactly straight east or west
'    If nDiff = 0# Then
'        If angleType = Bearing Then
'            ' dms or grads
'            If angDisp = cDMS Then
'                sDeg = "90-00-00"
'            Else ' grads
'                sDeg = RoundVal(100, angPrec)
'            End If
'            ' format the string
'            If eDiff > 0 Then
'                CalcDir_ = sN & " " & sDeg & " " & sE
'            Else
'                CalcDir_ = sN & " " & sDeg & " " & sW
'            End If
'        ElseIf angleType = nAzimuth Then
'            ' dms or grads
'            If angDisp = cDMS Then
'                ' DMS
'                If eDiff > 0 Then
'                    CalcDir_ = "90-00-00"
'                Else
'                    CalcDir_ = "270-00-00"
'                End If
'            Else ' grads
'                If eDiff > 0 Then
'                    CalcDir_ = RoundVal(100#, angPrec)
'                Else
'                    CalcDir_ = RoundVal(300#, angPrec)
'                End If
'            End If
'        Else ' south azimuth
'            ' dms or grads
'            If angDisp = cDMS Then
'                ' DMS
'                If eDiff > 0 Then
'                    CalcDir_ = "270-00-00"
'                Else
'                    CalcDir_ = "90-00-00"
'                End If
'            Else ' grads
'                If eDiff > 0 Then
'                    CalcDir_ = RoundVal(300#, angPrec)
'                Else
'                    CalcDir_ = RoundVal(100#, angPrec)
'                End If
'            End If
'        End If
'
'    ' format directions exactly north or south
'    ElseIf eDiff = 0# Then
'        If angleType = Bearing Then
'            ' dms or grads
'            If angDisp = cDMS Then
'                sDeg = "00-00-00"
'            Else ' grads
'                sDeg = RoundVal(0#, angPrec)
'            End If
'            ' format the string
'            If nDiff > 0 Then
'                CalcDir_ = sN & " " & sDeg & " " & sE
'            Else
'                CalcDir_ = sS & " " & sDeg & " " & sE
'            End If
'        ElseIf angleType = nAzimuth Then
'            ' dms or grads
'            If angDisp = cDMS Then
'                ' DMS
'                If nDiff > 0 Then
'                    CalcDir_ = "00-00-00"
'                Else
'                    CalcDir_ = "180-00-00"
'                End If
'            Else ' grads
'                If nDiff > 0 Then
'                    CalcDir_ = RoundVal(0#, angPrec)
'                Else
'                    CalcDir_ = RoundVal(200#, angPrec)
'                End If
'            End If
'        Else ' south azimuth
'            ' dms or grads
'            If angDisp = cDMS Then
'                ' DMS
'                If nDiff > 0 Then
'                    CalcDir_ = "180-00-00"
'                Else
'                    CalcDir_ = "00-00-00"
'                End If
'            Else ' grads
'                If nDiff > 0 Then
'                    CalcDir_ = RoundVal(200#, angPrec)
'                Else
'                    CalcDir_ = RoundVal(0#, angPrec)
'                End If
'            End If
'        End If
'
'    ' format all other values
'    Else
'        If nDiff > 0 Then
'            ' it's a positive north
'            If eDiff > 0 Then
'                quad = 1
'            Else
'                quad = 4
'            End If
'        Else
'            ' it's a negative north
'            If eDiff > 0 Then
'                quad = 2
'            Else
'                quad = 3
'            End If
'        End If
'
'        'calc the angle from the x axis
'        If angleType = Bearing Then
'
'            If angDisp = cDMS Then
'                ' DMS
'                ang = Atn(Abs(nDiff / eDiff)) * 180 / pi
'
'                dir = 90 - ang
'
'                deg = Int(dir)
'
'                dir = (dir - Int(dir)) * 60
'                min = Int(dir)
'
'                dir = (dir - Int(dir)) * 60
'                sec = CInt(RoundVal(dir, 0))
'
'                ' check for values rounded to 60
'                If sec >= 60 Then
'                    sec = sec - 60
'                    min = min + 1
'                End If
'                If min >= 60 Then
'                    min = min - 60
'                    deg = deg + 1
'                End If
'                If deg >= 90 Then
'                    deg = deg - 90
'                End If
'
'                ' format the strings
'                sDeg = CStr(deg)
'                If sDeg = "0" Then sDeg = "00" & sDeg
'                sMin = CStr(min)
'                If Len(sMin) = 1 Then sMin = "0" & sMin
'                sSec = CStr(sec)
'                If Len(sSec) = 1 Then sSec = "0" & sSec
'
'                CalcDir_ = sDeg & "-" & sMin & "-" & sSec
'
'            Else
'                ' Grad
'                ang = Atn(Abs(nDiff / eDiff)) * 200 / pi
'                dir = 100 - ang
'                CalcDir_ = RoundVal(dir, angPrec)
'            End If
'
'            ' format the bearing string
'            If quad = 1 Then
'                CalcDir_ = sN & " " & CalcDir_ & " " & sE
'            ElseIf quad = 2 Then
'                 CalcDir_ = sS & " " & CalcDir_ & " " & sE
'            ElseIf quad = 3 Then
'                 CalcDir_ = sS & " " & CalcDir_ & " " & sW
'            Else
'                 CalcDir_ = sN & " " & CalcDir_ & " " & sW
'            End If
'
'        ElseIf angleType = nAzimuth Then
'
'            If ImportOpts.angDisp = cDMS Then
'                ' DMS
'                ang = Atn(Abs(nDiff / eDiff)) * 180 / pi
'
'                If quad = 1 Then
'                    dir = 90 - ang
'                ElseIf quad = 2 Then
'                    dir = 90 + ang
'                ElseIf quad = 3 Then
'                    dir = 270 - ang
'                Else
'                    dir = 270 + ang
'                End If
'
'                deg = Int(dir)
'
'                dir = (dir - Int(dir)) * 60
'                min = Int(dir)
'
'                dir = (dir - Int(dir)) * 60
'                sec = CInt(RoundVal(dir, 0))
'
'                ' check for values rounded to 60
'                If sec >= 60 Then
'                    sec = sec - 60
'                    min = min + 1
'                End If
'                If min >= 60 Then
'                    min = min - 60
'                    deg = deg + 1
'                End If
'                If deg >= 360 Then
'                    deg = deg - 90
'                End If
'
'                ' format the strings
'                sDeg = CStr(deg)
'                If sDeg = "0" Then sDeg = "00" & sDeg
'                sMin = CStr(min)
'                If Len(sMin) = 1 Then sMin = "0" & sMin
'                sSec = CStr(sec)
'                If Len(sSec) = 1 Then sSec = "0" & sSec
'
'                CalcDir_ = sDeg & "-" & sMin & "-" & sSec
'
'            Else
'                ' Grad
'                ang = Atn(Abs(nDiff / eDiff)) * 200 / pi
'                If quad = 1 Then
'                    dir = 100 - ang
'                ElseIf quad = 2 Then
'                    dir = 100 + ang
'                ElseIf quad = 3 Then
'                    dir = 300 - ang
'                Else
'                    dir = 300 + ang
'                End If
'                CalcDir_ = RoundVal(dir, angPrec)
'            End If
'
'        Else ' south azimuth
'
'            If angDisp = cDMS Then
'                ' DMS
'                ang = Atn(Abs(nDiff / eDiff)) * 180 / pi
'
'                If quad = 3 Then
'                    dir = 90 - ang
'                ElseIf quad = 4 Then
'                    dir = 90 + ang
'                ElseIf quad = 1 Then
'                    dir = 270 - ang
'                Else
'                    dir = 270 + ang
'                End If
'
'                deg = Int(dir)
'
'                dir = (dir - Int(dir)) * 60
'                min = Int(dir)
'
'                dir = (dir - Int(dir)) * 60
'                sec = CInt(RoundVal(dir, 0))
'
'                ' check for values rounded to 60
'                If sec >= 60 Then
'                    sec = sec - 60
'                    min = min + 1
'                End If
'                If min >= 60 Then
'                    min = min - 60
'                    deg = deg + 1
'                End If
'                If deg >= 360 Then
'                    deg = deg - 90
'                End If
'
'                ' format the strings
'                sDeg = CStr(deg)
'                If sDeg = "0" Then sDeg = "00" & sDeg
'                sMin = CStr(min)
'                If Len(sMin) = 1 Then sMin = "0" & sMin
'                sSec = CStr(sec)
'                If Len(sSec) = 1 Then sSec = "0" & sSec
'
'                CalcDir_ = sDeg & "-" & sMin & "-" & sSec
'
'            Else
'                ' Grad
'                ang = Atn(Abs(nDiff / eDiff)) * 200 / pi
'                If quad = 3 Then
'                    dir = 100 - ang
'                ElseIf quad = 4 Then
'                    dir = 100 + ang
'                ElseIf quad = 1 Then
'                    dir = 300 - ang
'                Else
'                    dir = 300 + ang
'                End If
'                CalcDir_ = RoundVal(dir, angPrec)
'            End If
'        End If
'    End If
'
'End Function
'Function Get2PointDirData(sNorth As Double, sEast As Double, eNorth As Double, eEast As Double, angPrec As Integer, angDisp As rpAngDispType, angleType As rpAngleType) As String
'    Dim radAngle As Double
'    Dim sResult As String
'    ' get the radangle direction first
'    radAngle = CalcDirRad(sNorth, sEast, eNorth, eEast)
'
'    Select Case angleType
'        Case Bearing
'            sResult = FormatBearingAngle(radAngle, angDisp, angPrec)
'        Case nAzimuth
'            sResult = FormatAzimuthAngle(radAngle, True, angDisp, angPrec)
'        Case Else 'sAzimuth
'            sResult = FormatAzimuthAngle(radAngle, False, angDisp, angPrec)
'    End Select
'    Get2PointDirData = sResult
'End Function



