' -----------------------------------------------------------------------------
' <copyright file="CtrlProfileListView.vb" company="Autodesk">
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

Friend MustInherit Class CtrlListViewWithStation

    Protected MustOverride Sub setupListView()
    Protected MustOverride Property startStationText(ByVal item As ListViewItem) As String
    Protected MustOverride Property endStationText(ByVal item As ListViewItem) As String
    Protected MustOverride Function getAssociateAlignment(ByVal item As ListViewItem) As AeccLandLib.IAeccAlignment

    Protected Overridable Sub getStationRange(ByVal item As ListViewItem, ByRef startStation As Double, ByRef endStation As Double)
        Throw New NotImplementedException("The specified child class should implement this function.")
    End Sub

    Public Function StartRawStation(ByVal item As ListViewItem) As Double
        Return ReportUtilities.GetRawStation(startStationText(item), getAssociateAlignment(item).StationIndexIncrement)
    End Function

    Public Function EndRawStation(ByVal item As ListViewItem) As Double
        Return ReportUtilities.GetRawStation(endStationText(item), getAssociateAlignment(item).StationIndexIncrement)
    End Function

    Public ReadOnly Property WinListView() As System.Windows.Forms.ListView
        Get
            Return listView_
        End Get
    End Property

    Public ReadOnly Property SelectedItem() As System.Windows.Forms.ListViewItem
        Get
            Try
                Return listView_.SelectedItems.Item(0)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    Public Overridable Overloads Sub Initialize(ByVal olistView As ListView, ByVal startStationTextBox As CtrlStationTextBox, _
                          ByVal endStationTextBox As CtrlStationTextBox)

        listView_ = olistView
        Try
            listView_.CheckBoxes = True
            listView_.FullRowSelect = True
            listView_.MultiSelect = False
            listView_.HideSelection = False
            listView_.GridLines = True
            listView_.View = View.Details
            setupListView()
        Catch ex As Exception
            Diagnostics.Debug.Assert(False, ex.Message)
        End Try

        startStationTextBox_ = startStationTextBox
        endStationTextBox_ = endStationTextBox
        startStationTextBox_.TextBox.Enabled = False
        endStationTextBox_.TextBox.Enabled = False
    End Sub

    Private Sub onItemSelectionChanged(ByVal sender As Object, ByVal e As ListViewItemSelectionChangedEventArgs) Handles listView_.ItemSelectionChanged
        Dim hasHandled As Boolean = False
        If listView_.SelectedItems.Count = 1 Then
            If e.IsSelected() = True Then
                startStationTextBox_.Alignment = getAssociateAlignment(e.Item)
                endStationTextBox_.Alignment = getAssociateAlignment(e.Item)

                If Not startStationTextBox_.UseAlignmentRange Then
                    'Start and end station should be the same to use alignment range
                    Debug.Assert(Not endStationTextBox_.UseAlignmentRange)

                    Dim startStation As Double
                    Dim endStation As Double
                    getStationRange(e.Item, startStation, endStation)
                    startStationTextBox_.StationMin = startStation
                    startStationTextBox_.StationMax = endStation

                    endStationTextBox_.StationMin = startStation
                    endStationTextBox_.StationMax = endStation
                End If

                endStationTextBox_.EquationStation = endStationText(e.Item)
                startStationTextBox_.EquationStation = startStationText(e.Item)
                startStationTextBox_.TextBox.Enabled = True
                endStationTextBox_.TextBox.Enabled = True
                hasHandled = True
            End If
        End If

        If hasHandled = False Then
            startStationTextBox_.TextBox.Enabled = False
            endStationTextBox_.TextBox.Enabled = False
            startStationTextBox_.EquationStation = ""
            endStationTextBox_.EquationStation = ""
            startStationTextBox_.Alignment = Nothing
            endStationTextBox_.Alignment = Nothing
        End If
    End Sub

    Private Sub onStartStationChanging(ByVal rawStationValue As Double, ByVal equationStationValue As String, ByRef okToChange As Boolean) Handles startStationTextBox_.StationChanging
        If rawStationValue > endStationTextBox_.RawStation Then
            ReportUtilities.ACADMsgBox(LocalizedRes.Msg_StartStation_Less_EndStation)
            okToChange = False
        Else
            Try
                startStationText(SelectedItem) = equationStationValue
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
            End Try

        End If
    End Sub

    Private Sub onEndStationChanging(ByVal rawStationValue As Double, ByVal equationStationValue As String, ByRef okToChange As Boolean) Handles endStationTextBox_.StationChanging
        If rawStationValue < startStationTextBox_.RawStation Then
            ReportUtilities.ACADMsgBox(LocalizedRes.Msg_EndStation_Greater_StartStation)
            okToChange = False
        Else
            Try
                endStationText(SelectedItem) = equationStationValue
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
            End Try
        End If
    End Sub

    'member variable
    Private WithEvents startStationTextBox_ As CtrlStationTextBox
    Private WithEvents endStationTextBox_ As CtrlStationTextBox
    Private WithEvents listView_ As System.Windows.Forms.ListView
End Class
