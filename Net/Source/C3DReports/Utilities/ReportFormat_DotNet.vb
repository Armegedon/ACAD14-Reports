Imports System
Imports Autodesk.Civil
Imports Autodesk.Civil.Settings
Imports Autodesk.Civil.ApplicationServices


Friend Class ReportFormat_DotNet
    '-----------------------------------------------------------------------
    '   rounds off values and returns a Double.
    '-----------------------------------------------------------------------
    Public Shared Function RoundDouble(ByVal value As Double, ByVal prec As Integer, _
        Optional ByVal roundType As RoundingType = RoundingType.Normal) As Double

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
            Case RoundingType.Normal
                tmpValue = sgnVal * Math.Floor(absVal + 0.5)
            Case RoundingType.Up
                tmpValue = sgnVal * Math.Ceiling(absVal)
            Case RoundingType.Truncate
                tmpValue = sgnVal * Math.Floor(absVal)
        End Select

        Return tmpValue / factor
    End Function

    '-----------------------------------------------------------------------
    '   rounds off values and returns a string.
    '-----------------------------------------------------------------------
    Public Shared Function RoundVal(ByVal value As Double, ByVal prec As Integer, _
        Optional ByVal roundType As RoundingType = RoundingType.Normal) As String

        Dim roundedValue As Double = RoundDouble(value, prec, roundType)

        Dim sFormatString As String
        sFormatString = "N" + prec.ToString()
        RoundVal = roundedValue.ToString(sFormatString)

    End Function
    Public Shared Function FormatGrade(ByVal dGrade As Double) As String
        Dim ambientSetting As SettingsAmbient = CivilApplication.ActiveDocument.Settings.DrawingSettings.AmbientSettings
        Dim grdFormat As GradeSlopeFormatType = ambientSetting.GradeSlope.Format.Value
        Dim grdPrec As Integer = ambientSetting.GradeSlope.Precision.Value
        Dim grdRounding As RoundingType = ambientSetting.GradeSlope.Rounding.Value
        Dim grdSign As SignType = ambientSetting.GradeSlope.Sign.Value


        Dim sPreSign As String = "", sPostSign As String = ""
        Select Case grdSign
            Case SignType.Always
                If dGrade < 0.0# Then
                    sPreSign = "-"
                Else
                    sPreSign = "+"
                End If
            Case SignType.BracketNegative
                If dGrade < 0.0# Then
                    sPreSign = "("
                    sPostSign = ")"
                End If
            Case SignType.Negative
                If dGrade < 0.0# Then
                    sPreSign = "-"
                End If
        End Select

        dGrade = Math.Abs(dGrade)

        Dim sGrade As String
        sGrade = ""

        Select Case grdFormat
            Case GradeSlopeFormatType.Decimal
                sGrade = RoundVal(dGrade, grdPrec, grdRounding)
            Case GradeSlopeFormatType.Percent
                dGrade = dGrade * 100.0#
                sGrade = RoundVal(dGrade, grdPrec, grdRounding)
                sGrade = sGrade + "%"
            Case GradeSlopeFormatType.RunRise
                dGrade = 1 / dGrade
                sGrade = RoundVal(dGrade, grdPrec, grdRounding)
                sGrade = sGrade + ":1"
            Case GradeSlopeFormatType.RiseRun
                dGrade = 1 / dGrade
                sGrade = RoundVal(dGrade, grdPrec, grdRounding)
                sGrade = "1:" + sGrade
            Case GradeSlopeFormatType.PerMille
                dGrade = dGrade * 1000.0#
                sGrade = RoundVal(dGrade, grdPrec, grdRounding)
                sGrade = sGrade + "‰"
        End Select

        FormatGrade = sPreSign + sGrade + sPostSign
    End Function
End Class
