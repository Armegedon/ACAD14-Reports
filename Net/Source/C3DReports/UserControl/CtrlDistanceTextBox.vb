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

Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Friend Class CtrlDistanceTextBox
    'member variable
    Private WithEvents textBox_ As System.Windows.Forms.TextBox
    Private m_dDistance As Double

    Event StationChanging(ByVal rawStation As Double, ByVal equationStation As String, ByRef okToChange As Boolean)

    Public Sub Initialize(ByVal textBox As System.Windows.Forms.TextBox)
        textBox_ = textBox
        m_dDistance = 0.0

        Reset()
    End Sub

    Public ReadOnly Property TextBox() As System.Windows.Forms.TextBox
        Get
            Return textBox_
        End Get
    End Property

    Private Sub Reset()
        textBox_.Text = m_dDistance.ToString()
    End Sub

    Public Function setValue(ByVal dVal) As Boolean
        'ReportUtilities.ACADMsgBox(LocalizedRes.Msg_Input_Right_Format_Station)
        Dim oDistSettings As AeccLandLib.AeccSettingsDistance
        oDistSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.DistanceSettings
        textBox_.Text = ReportFormat.FormatDistance(dVal, oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                    oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
        m_dDistance = dVal
        Return True
    End Function
    Public Function isValidDistance(ByVal strText As String) As Boolean
        strText = strText.Trim()
        If strText.Length <= 0 Then
            Return False
        End If
        Dim ch As Char = strText.ToLower().Chars(strText.Length - 1)
        If (ch < ChrW(Keys.D0) Or ch > ChrW(Keys.D9)) And ch <> "m" And ch <> "'" Then
            Return False
        End If

        If ch = "m" Or ch = "'" Then
            strText = strText.Substring(0, strText.Length - 1)
            strText = strText.Trim()
        End If

        Try
            Dim dVal As Double = System.Convert.ToDouble(strText)
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    Public Function getDoubleValue(ByRef dVal As Double) As Boolean
        Dim strText As String = textBox_.Text.Trim()
        Dim ch As Char = strText.ToLower().Chars(strText.Length - 1)
        If (ch < ChrW(Keys.D0) Or ch > ChrW(Keys.D9)) And ch <> "m" And ch <> "'" Then
            Return False
        End If

        If ch = "m" Or ch = "'" Then
            strText = strText.Substring(0, strText.Length - 1)
            strText = strText.Trim()
        End If

        Try
            Dim dValue As Double = System.Convert.ToDouble(strText)
            dVal = dValue
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function


    Private Sub textBox__Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles textBox_.Leave
        Dim strText As String = textBox_.Text.Trim()
        Dim dValue As Double = 0.0
        If (isValidDistance(strText)) Then
            getDoubleValue(dValue)
            m_dDistance = dValue
        Else
            'ReportUtilities.ACADMsgBox(LocalizedRes.Msg_Input_Right_Format_Station)
        End If

        Dim oDistSettings As AeccLandLib.AeccSettingsDistance
        oDistSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.DistanceSettings
        textBox_.Text = ReportFormat.FormatDistance(m_dDistance, oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                    oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
    End Sub
End Class
