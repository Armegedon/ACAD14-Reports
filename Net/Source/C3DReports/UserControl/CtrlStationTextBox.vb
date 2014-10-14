' -----------------------------------------------------------------------------
' <copyright file="CtrlStationTextBox.vb" company="Autodesk">
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

Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Friend Class CtrlStationTextBox

    Event StationChanging(ByVal rawStation As Double, ByVal equationStation As String, ByRef okToChange As Boolean)

    Public Sub Initialize(ByVal textBox As System.Windows.Forms.TextBox)
        textBox_ = textBox
        validateRange_ = True

        Reset()
    End Sub

    Public Property Alignment() As AeccLandLib.IAeccAlignment
        Get
            Return alignment_
        End Get
        Set(ByVal value As AeccLandLib.IAeccAlignment)
            alignment_ = value
        End Set
    End Property

    Public Property RawStation() As Double
        Get
            Return rawStation_
        End Get
        Set(ByVal value As Double)
            updateStationByRawValue(value)
        End Set
    End Property

    Public Property EquationStation() As String
        Get
            Return equationStation_
        End Get
        Set(ByVal value As String)
            Dim rawStationValue As Double
            If parseEquationStation(value, rawStationValue) Then
                updateStationByRawValue(rawStationValue)
            Else
                Reset()
            End If
        End Set
    End Property

    Public ReadOnly Property TextBox() As System.Windows.Forms.TextBox
        Get
            Return textBox_
        End Get
    End Property

    Public Property ValidateRange() As Boolean
        Get
            Return validateRange_
        End Get
        Set(ByVal value As Boolean)
            validateRange_ = value
        End Set
    End Property

    Public Property UseAlignmentRange() As Boolean
        Get
            Return useAlignmentRange_
        End Get
        Set(ByVal value As Boolean)
            useAlignmentRange_ = value
        End Set
    End Property

    Public Property StationMin() As Double
        Get
            If Me.UseAlignmentRange Then
                Return roundStation(alignment_.StartingStation)
            Else
                Return stationMin_
            End If
        End Get
        Set(ByVal value As Double)
            stationMin_ = roundStation(value)
        End Set
    End Property

    Public Property StationMax() As Double
        Get
            If Me.UseAlignmentRange Then
                Return roundStation(alignment_.EndingStation)
            Else
                Return stationMax_
            End If
        End Get
        Set(ByVal value As Double)
            stationMax_ = roundStation(value)
        End Set
    End Property

    Private Sub textBox_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles textBox_.Leave
        Dim rawStationValue As Double
        If parseEquationStation(textBox_.Text.Trim(), rawStationValue) Then
            updateStationByRawValue(rawStationValue)
        Else
            ReportUtilities.ACADMsgBox(LocalizedRes.Msg_Input_Right_Format_Station)
            textBox_.Text = equationStation_
        End If
    End Sub

    Private Sub onKeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textBox_.KeyPress
        If IsInputValidKey(e.KeyChar) = False Then
            Beep()
            e.Handled = True
        End If
    End Sub

    Private Function IsInputValidKey(ByVal KeyChar As Char) As Boolean
        If KeyChar <> ChrW(Keys.Back) And _
                (KeyChar < ChrW(Keys.D0) Or KeyChar > ChrW(Keys.D9)) And _
                KeyChar <> ChrW(Keys.Delete) And _
                KeyChar <> ChrW(Keys.Execute) And _
                KeyChar <> ChrW(Keys.Insert) Then
            IsInputValidKey = False
        Else
            IsInputValidKey = True
        End If
    End Function

    Private Function validateStation(ByVal rawStationValue As Double) As Boolean
        Dim valid As Boolean = True
        If ValidateRange Then
            Try
                If rawStationValue < Me.StationMin Or rawStationValue > Me.StationMax Then
                    valid = False
                End If
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
                valid = False
            End Try
        End If
        Return valid
    End Function

    Private Function roundStation(ByVal rawStationValue As Double) As Double
        Return ReportUtilities.GetRawStation(ReportUtilities.GetStationString(alignment_, rawStationValue), alignment_.StationIndexIncrement)
    End Function

    Private Function parseEquationStation(ByVal equationStationValue As String, ByRef rawStationValue As Double) As Boolean
        If equationStationValue = "" Then
            rawStationValue = 0.0
            Return False
        Else
            Try
                rawStationValue = CDbl(equationStationValue)
            Catch ex As Exception
                Try
                    rawStationValue = ReportUtilities.GetRawStation(equationStationValue, alignment_.StationIndexIncrement)
                Catch eex As Exception
                    Return False
                End Try
            End Try
        End If
        Return True
    End Function

    Private Sub updateStationByRawValue(ByVal rawStationValue As Double)
        If validateStation(rawStationValue) = False Then
            ReportUtilities.ACADMsgBox(LocalizedRes.Msg_Station_Out_Range)
            textBox_.Text = equationStation_
        Else
            Dim equationStationValue As String
            Dim okToChange As Boolean = True
            Try
                equationStationValue = ReportUtilities.GetStationString(alignment_, rawStationValue)
                RaiseEvent StationChanging(rawStationValue, equationStationValue, okToChange)
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
                okToChange = False
                equationStationValue = ""
            End Try
            If okToChange Then
                If validateStation(rawStationValue) Then
                    rawStation_ = rawStationValue
                    equationStation_ = equationStationValue
                    textBox_.Text = equationStationValue
                Else
                    textBox_.Text = equationStation_
                End If
            Else
                'set text box to preview value
                textBox_.Text = equationStation_
            End If
        End If
    End Sub

    Private Sub Reset()
        equationStation_ = ""
        textBox_.Text = ""
        rawStation_ = 0.0
    End Sub

    'member variable
    Private WithEvents textBox_ As System.Windows.Forms.TextBox
    Private alignment_ As AeccLandLib.IAeccAlignment
    Private rawStation_ As Double
    Private equationStation_ As String
    Private validateRange_ As Boolean
    Private useAlignmentRange_ As Boolean = True
    Private stationMin_ As Double = 0.0
    Private stationMax_ As Double = 0.0
End Class
