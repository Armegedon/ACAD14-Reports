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

Friend Class CtrlAlignmentListView
    Inherits CtrlListViewWithStation

    Public Enum ListViewColumn
        kInclude = 0
        kName
        kDescription
        kStartStation
        kEndStation
    End Enum

    Protected Overrides Sub setupListView()
        ' setup columns
        MyBase.WinListView.CheckBoxes = True
        MyBase.WinListView.FullRowSelect = True
        MyBase.WinListView.MultiSelect = False
        MyBase.WinListView.HideSelection = False
        MyBase.WinListView.GridLines = True
        MyBase.WinListView.View = View.Details
        MyBase.WinListView.Columns.Add(LocalizedRes.ListView_Column_Include, 50, HorizontalAlignment.Left)
        MyBase.WinListView.Columns.Add(LocalizedRes.ListView_Column_Name, 100, HorizontalAlignment.Left)
        MyBase.WinListView.Columns.Add(LocalizedRes.ListView_Column_Description, 100, HorizontalAlignment.Left)
        MyBase.WinListView.Columns.Add(LocalizedRes.ListView_Column_StationStart, 80, HorizontalAlignment.Left)
        MyBase.WinListView.Columns.Add(LocalizedRes.ListView_Column_StationEnd, 80, HorizontalAlignment.Left)

        'add item to ListView
        For Each iAlignment As AeccLandLib.IAeccAlignment In alignments_.Values
            Dim li As New ListViewItem
            li.SubItems.Add(iAlignment.Name)
            li.SubItems.Add(iAlignment.Description)
            li.SubItems.Add(ReportUtilities.GetStationString(iAlignment, iAlignment.StartingStation))
            li.SubItems.Add(ReportUtilities.GetStationString(iAlignment, iAlignment.EndingStation))
            li.Tag = iAlignment
            li.Checked = True
            MyBase.WinListView.Items.Add(li)
        Next
    End Sub

    Public Overrides Sub Initialize(ByVal oListView As ListView, ByVal startStationTextBox As CtrlStationTextBox, _
                          ByVal endStationTextBox As CtrlStationTextBox)
        'get alignments from Document
        getAlignments()
        MyBase.Initialize(oListView, startStationTextBox, endStationTextBox)
    End Sub

    Public ReadOnly Property Alignments() As Collections.Generic.IList(Of AeccLandLib.IAeccAlignment)
        Get
            Return alignments_.Values
        End Get
    End Property

    Public Shared Function GetAlignmentCount(Optional ByVal showMsg As Boolean = True) As Integer
        Dim count As Integer
        count = ReportApplication.AeccXDocument.AlignmentsSiteless.Count

        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            count += oSite.Alignments.Count
        Next

        If showMsg = True And count = 0 Then
            Dim sMsg As String = String.Format(LocalizedRes.Msg_No_Alignments_In_Drawing, ReportApplication.AeccXDocument.Name)
            ReportUtilities.ACADMsgBox(sMsg)
        End If

        Return count
    End Function

    Private Sub getAlignments()
        alignments_ = New Collections.Generic.SortedList(Of String, AeccLandLib.IAeccAlignment)
        For Each oAlignment As AeccLandLib.AeccAlignment In ReportApplication.AeccXDocument.AlignmentsSiteless
            alignments_.Add(oAlignment.Name, oAlignment)
        Next

        'get alignments from Site
        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            For Each oAlignment As AeccLandLib.AeccAlignment In oSite.Alignments
                alignments_.Add(oAlignment.Name, oAlignment)
            Next
        Next
    End Sub

    Protected Overrides Property startStationText(ByVal item As ListViewItem) As String
        Get
            Try
                Return item.SubItems.Item(ListViewColumn.kStartStation).Text
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
                Return ""
            End Try
        End Get
        Set(ByVal value As String)
            Try
                item.SubItems.Item(ListViewColumn.kStartStation).Text = value
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
            End Try
        End Set
    End Property

    Protected Overrides Property endStationText(ByVal item As ListViewItem) As String
        Get
            Try
                Return item.SubItems.Item(ListViewColumn.kEndStation).Text
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
                Return ""
            End Try
        End Get
        Set(ByVal value As String)
            Try
                item.SubItems.Item(ListViewColumn.kEndStation).Text = value
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
            End Try
        End Set
    End Property

    Protected Overrides Function getAssociateAlignment(ByVal item As ListViewItem) As AeccLandLib.IAeccAlignment
        Try
            Return item.Tag
        Catch ex As Exception
            Diagnostics.Debug.Assert(False, ex.Message)
            Return Nothing
        End Try
    End Function

    'member variable
    Private alignments_ As Collections.Generic.SortedList(Of String, AeccLandLib.IAeccAlignment)
End Class
