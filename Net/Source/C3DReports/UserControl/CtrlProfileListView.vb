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

Friend Class CtrlProfileListView
    Inherits CtrlListViewWithStation

    Public Enum ListViewColumn
        kInclude = 0
        kName
        kDescription
        kStartStation
        kEndStation
        kAlignmentName
    End Enum

    Public Shared Function GetProfileCount(Optional ByVal type As AeccLandLib.AeccProfileType = Nothing, _
                                            Optional ByVal openMsg As Boolean = True) As Integer
        Dim nCount As Integer
        Dim nOtherProfileCount As Integer
        nCount = 0
        For Each oAlignment As AeccLandLib.AeccAlignment In ReportApplication.AeccXDocument.AlignmentsSiteless
            If type = Nothing Then
                nCount += oAlignment.Profiles.Count
            Else
                For Each oProfile As AeccLandLib.AeccProfile In oAlignment.Profiles
                    If oProfile.Type = type Then
                        nCount += 1
                    Else
                        nOtherProfileCount += 1
                    End If
                Next
            End If
        Next

        'get alignments from Site
        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            For Each oAlignment As AeccLandLib.AeccAlignment In oSite.Alignments
                If type = Nothing Then
                    nCount += oAlignment.Profiles.Count
                Else
                    For Each oProfile As AeccLandLib.AeccProfile In oAlignment.Profiles
                        If oProfile.Type = type Then
                            nCount += 1
                        Else
                            nOtherProfileCount += 1
                        End If
                    Next
                End If
            Next
        Next

        GetProfileCount = nCount

        If openMsg = True And GetProfileCount = 0 Then
            Dim sNoProfile As String
            If nOtherProfileCount = 0 Then
                'all kind of profile count 0
                sNoProfile = LocalizedRes.Msg_No_Profile_In_Drawing
            Else
                'has other profile, but no specified profile
                If type = AeccLandLib.AeccProfileType.aeccFinishedGround Then
                    'no FG profile
                    sNoProfile = LocalizedRes.Msg_No_FG_Profile_In_Drawing
                Else
                    'no profile
                    sNoProfile = LocalizedRes.Msg_No_Profile_In_Drawing
                End If
            End If

            Dim sMsg As String = String.Format(sNoProfile, ReportApplication.AeccXDocument.Name)
            ReportUtilities.ACADMsgBox(sMsg)
        End If
    End Function

    ''' <summary>
    ''' Indicates to check all items or only the first item initially.
    ''' </summary>
    ''' <value>
    ''' Set True if to check all item of the profile list initially. 
    ''' Set False to check the first item only.
    ''' </value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AllItemChecked() As Boolean
        Get
            Return allItemChecked_
        End Get
        Set(ByVal value As Boolean)
            allItemChecked_ = value
        End Set
    End Property

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
        MyBase.WinListView.Columns.Add(LocalizedRes.ListView_Column_Alignment, 80, HorizontalAlignment.Left)

        Dim itemCheck As Boolean = True

        'add item to ListView
        For Each iProfile As AeccLandLib.IAeccProfile In profiles_.Values
            Dim li As New ListViewItem
            li.SubItems.Add(iProfile.Name)
            li.SubItems.Add(iProfile.Description)
            li.SubItems.Add(ReportUtilities.GetStationString(iProfile.Alignment, iProfile.StartingStation))
            li.SubItems.Add(ReportUtilities.GetStationString(iProfile.Alignment, iProfile.EndingStation))
            li.SubItems.Add(iProfile.Alignment.Name)
            li.Checked = itemCheck
            li.Tag = iProfile
            MyBase.WinListView.Items.Add(li)

            If Not Me.AllItemChecked And itemCheck Then
                itemCheck = False
            End If
        Next
    End Sub

    Public Overrides Sub Initialize(ByVal oListView As ListView, ByVal startStationTextBox As CtrlStationTextBox, _
                          ByVal endStationTextBox As CtrlStationTextBox)
        'get alignments from Document
        getProfiles()

        startStationTextBox.UseAlignmentRange = False
        endStationTextBox.UseAlignmentRange = False

        MyBase.Initialize(oListView, startStationTextBox, endStationTextBox)
    End Sub

    Protected Overrides Sub getStationRange(ByVal item As ListViewItem, ByRef startStation As Double, ByRef endStation As Double)
        Dim profile As AeccLandLib.IAeccProfile = CType(item.Tag, AeccLandLib.IAeccProfile)
        startStation = profile.StartingStation
        endStation = profile.EndingStation
    End Sub

    Public ReadOnly Property Profiles() As Collections.Generic.IList(Of AeccLandLib.IAeccProfile)
        Get
            Return profiles_.Values
        End Get
    End Property

    Private Sub getProfiles()
        profiles_ = New Collections.Generic.SortedList(Of String, AeccLandLib.IAeccProfile)
        For Each oAlignment As AeccLandLib.AeccAlignment In ReportApplication.AeccXDocument.AlignmentsSiteless
            For Each oProfile As AeccLandLib.AeccProfile In oAlignment.Profiles
                If oProfile.Type = AeccLandLib.AeccProfileType.aeccFinishedGround Then
                    profiles_.Add(oAlignment.Name + "-" + oProfile.Name, oProfile)
                End If
            Next
        Next

        'get alignments from Site
        For Each oSite As AeccLandLib.AeccSite In ReportApplication.AeccXDatabase.Sites
            For Each oAlignment As AeccLandLib.AeccAlignment In oSite.Alignments
                For Each oProfile As AeccLandLib.AeccProfile In oAlignment.Profiles
                    If oProfile.Type = AeccLandLib.AeccProfileType.aeccFinishedGround Then
                        profiles_.Add(oAlignment.Name + "-" + oProfile.Name, oProfile)
                    End If
                Next
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
        Dim iProfile As AeccLandLib.IAeccProfile
        Try
            iProfile = CType(item.Tag, AeccLandLib.IAeccProfile)
            Return iProfile.Alignment
        Catch ex As Exception
            Diagnostics.Debug.Assert(False, ex.Message)
            Return Nothing
        End Try
    End Function

    'member variable
    Private profiles_ As Collections.Generic.SortedList(Of String, AeccLandLib.IAeccProfile)
    Private allItemChecked_ As Boolean = True
End Class
