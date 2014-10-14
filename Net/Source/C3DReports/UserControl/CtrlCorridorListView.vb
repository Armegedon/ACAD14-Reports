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

Friend Class CtrlCorridorListView
    Inherits CtrlListViewWithStation
    Public haveMaterialList As Boolean = False

    Public Enum ListViewColumn
        kCorridorName = 0
        kAlignmentName
        kSampleLineGroupName
        kLinkCode
        kStartStation
        kEndStation
        kMaterialList
    End Enum

    Public Function CorridorName(ByVal item As ListViewItem) As String
        Try
            Return item.Text
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function AlignmentName(ByVal item As ListViewItem) As String
        Try
            Return item.SubItems.Item(ListViewColumn.kAlignmentName).Text
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function ListCode(ByVal item As ListViewItem) As String
        Try
            Return item.SubItems.Item(ListViewColumn.kLinkCode).Text
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function SampleLineGroupName(ByVal item As ListViewItem) As String
        Try
            Return item.SubItems.Item(ListViewColumn.kSampleLineGroupName).Text
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Sub AddListItem(ByVal sampleLineGroup As AeccLandLib.IAeccSampleLineGroup, ByVal corridorName As String, _
                           ByVal alignmentName As String, ByVal linkCode As String)

        Try
            Dim startStation As New String("")
            Dim endStation As New String("")
            stationsOfSampleLineGroup(sampleLineGroup, startStation, endStation)
            Dim li As New ListViewItem
            li.Text = corridorName
            li.SubItems.Add(alignmentName)
            li.SubItems.Add(sampleLineGroup.Name)
            li.SubItems.Add(linkCode)
            li.SubItems.Add(startStation)
            li.SubItems.Add(endStation)
            li.Tag = sampleLineGroup.Parent
            WinListView.Items.Add(li)
        Catch ex As Exception
            Diagnostics.Debug.Assert(False, ex.Message)
        End Try
    End Sub

    Public Sub AddListItem(ByVal sampleLineGroup As AeccLandLib.IAeccSampleLineGroup, ByVal corridorName As String, _
                       ByVal alignmentName As String, ByVal linkCode As String, ByVal materialListName As String)

        Try
            Debug.Assert(haveMaterialList)
            Dim startStation As New String("")
            Dim endStation As New String("")
            stationsOfSampleLineGroup(sampleLineGroup, startStation, endStation)
            Dim li As New ListViewItem
            li.Text = corridorName
            li.SubItems.Add(alignmentName)
            li.SubItems.Add(sampleLineGroup.Name)
            li.SubItems.Add(linkCode)
            li.SubItems.Add(startStation)
            li.SubItems.Add(endStation)
            li.SubItems.Add(materialListName)
            li.Tag = sampleLineGroup.Parent
            WinListView.Items.Add(li)
        Catch ex As Exception
            Diagnostics.Debug.Assert(False, ex.Message)
        End Try
    End Sub

    Private Sub stationsOfSampleLineGroup(ByVal oSampleLineGroup As AeccLandLib.IAeccSampleLineGroup, _
                                          ByRef startStation As String, ByRef endStation As String)
        Try
            Dim dStartStation, dEndStation As Double
            dStartStation = oSampleLineGroup.SampleLines.Item(0).Station
            dEndStation = dStartStation
            For Each oSampleLine As AeccLandLib.IAeccSampleLine In oSampleLineGroup.SampleLines
                If oSampleLine.Station > dEndStation Then
                    dEndStation = oSampleLine.Station
                End If
                If oSampleLine.Station < dStartStation Then
                    dStartStation = oSampleLine.Station
                End If
            Next
            startStation = ReportUtilities.GetStationString(oSampleLineGroup.Parent, dStartStation)
            endStation = ReportUtilities.GetStationString(oSampleLineGroup.Parent, dEndStation)
        Catch ex As Exception
            Diagnostics.Debug.Assert(False, ex.Message)
        End Try
    End Sub

    Protected Overrides Sub setupListView()
        ' setup columns
        MyBase.WinListView.CheckBoxes = False
        MyBase.WinListView.MultiSelect = True
        MyBase.WinListView.Columns.Add(LocalizedRes.ListView_Column_Name, 100, HorizontalAlignment.Left)
        MyBase.WinListView.Columns.Add(LocalizedRes.ListView_Column_Alignment, 80, HorizontalAlignment.Left)
        MyBase.WinListView.Columns.Add(LocalizedRes.ListView_Column_Corridor_SampleLineGroup, 100, HorizontalAlignment.Left)
        MyBase.WinListView.Columns.Add(LocalizedRes.ListView_Column_Corridor_LinkCode, 80, HorizontalAlignment.Left)
        MyBase.WinListView.Columns.Add(LocalizedRes.ListView_Column_StationStart, 80, HorizontalAlignment.Left)
        MyBase.WinListView.Columns.Add(LocalizedRes.ListView_Column_StationEnd, 80, HorizontalAlignment.Left)
        If haveMaterialList Then
            MyBase.WinListView.Columns.Add(LocalizedRes.ListView_Column_Material_List, 100, HorizontalAlignment.Left)
        End If
    End Sub

    Protected Overrides Property startStationText(ByVal oListItem As ListViewItem) As String
        Get
            Try
                Return oListItem.SubItems.Item(ListViewColumn.kStartStation).Text
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
                Return ""
            End Try
        End Get
        Set(ByVal value As String)
            Try
                oListItem.SubItems.Item(ListViewColumn.kStartStation).Text = value
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
            End Try
        End Set
    End Property
    Protected Overrides Property endStationText(ByVal oListItem As ListViewItem) As String
        Get
            Try
                Return oListItem.SubItems.Item(ListViewColumn.kEndStation).Text
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
                Return ""
            End Try
        End Get
        Set(ByVal value As String)
            Try
                oListItem.SubItems.Item(ListViewColumn.kEndStation).Text = value
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
End Class
