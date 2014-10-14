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

Imports System
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports LocalizedRes = Report.My.Resources.LocalizedStrings
Imports Autodesk.AECC.Interop

Friend Class CtrlProfileListExistingView

    Private ListView_Profile As ListView

    Public Sub Initialize(ByVal oListView As ListView)

        ListView_Profile = oListView

        '2. set list view header
        setupListView()
    End Sub

    Public ReadOnly Property CheckedCount() As Integer
        Get
            Return ListView_Profile.CheckedItems.Count
        End Get
    End Property

    Public ReadOnly Property CheckedItems() As System.Windows.Forms.ListView.CheckedListViewItemCollection
        Get
            Return ListView_Profile.CheckedItems
        End Get
    End Property

    Public ReadOnly Property HasSelectedItem() As Boolean
        Get
            If ListView_Profile.SelectedItems.Count = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    Protected Sub setupListView()
        ' list view grid setup
        ' setup columns
        ListView_Profile.View = View.Details
        ListView_Profile.Columns.Add(LocalizedRes.ProfileListView_ColumnTitle_Include, _
            50, HorizontalAlignment.Left)
        ListView_Profile.Columns.Add(LocalizedRes.ProfileListView_ColumnTitle_Name, _
            120, HorizontalAlignment.Left)
        ListView_Profile.Columns.Add(LocalizedRes.ProfileListView_ColumnTitle_Description, _
            100, HorizontalAlignment.Left)
        ListView_Profile.Columns.Add(LocalizedRes.ProfileListView_ColumnTitle_StationStart, _
            80, HorizontalAlignment.Left)
        ListView_Profile.Columns.Add(LocalizedRes.ProfileListView_ColumnTitle_StationEnd, _
            80, HorizontalAlignment.Left)
        ListView_Profile.Columns.Add(LocalizedRes.ProfileListView_ColumnTitle_Alignment, _
            120, HorizontalAlignment.Left)
    End Sub

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
                sNoProfile = LocalizedRes.ProfileListView_NoProfileInDrawing
            Else
                'has other profile, but no specified profile
                If type = AeccLandLib.AeccProfileType.aeccFinishedGround Then
                    'no FG profile
                    sNoProfile = LocalizedRes.ProfileListView_NoProfile
                Else
                    'no profile
                    sNoProfile = LocalizedRes.ProfileListView_NoProfileInDrawing
                End If
            End If

            Dim sMsg As String = sNoProfile.Replace("&DrawingName&", ReportApplication.AeccXDocument.Name)
            ReportUtilities.ACADMsgBox(sMsg)
        End If
    End Function

    Public Sub FillListView_Profile(ByVal profileDesignView As CtrlProfileListView)

        Dim t As Integer
        Dim oProfile As Land.AeccProfile = Nothing
        Dim curListItem As ListViewItem = Nothing
        Dim AlignmentName As String = ""
        Dim oProfileExisting As Land.AeccProfile
        Dim oProfileDesign As Land.AeccProfile = Nothing

        Dim curListItemDesign As ListViewItem = Nothing

        ListView_Profile.Items.Clear()

        For t = 0 To profileDesignView.WinListView.Items.Count() - 1
            curListItem = profileDesignView.WinListView.Items(t)
            If curListItem.Checked = True Then
                oProfile = CType(curListItem.Tag, Land.AeccProfile)
                If Not oProfile Is Nothing Then
                    AlignmentName = oProfile.Alignment.DisplayName
                End If
            End If
        Next

        Dim nCount As Long
        For t = 0 To profileDesignView.WinListView.Items.Count() - 1
            curListItemDesign = profileDesignView.WinListView.Items(t)
            If curListItemDesign.Checked = True Then
                oProfileDesign = CType(curListItemDesign.Tag, Land.AeccProfile)
                For Each oProfileExisting In oProfileDesign.Alignment.Profiles
                    If oProfileDesign.Alignment.Profiles.Count = 0 Then Exit For
                    If oProfileExisting.Type = Land.AeccProfileType.aeccExistingGround And oProfileExisting.Alignment.DisplayName = AlignmentName Then
                        Dim li As New ListViewItem
                        li.SubItems.Add(oProfileExisting.Name)
                        li.SubItems.Add(oProfileExisting.Description)
                        li.SubItems.Add(ReportUtilities.GetStationString(oProfileDesign.Alignment, oProfileDesign.StartingStation))
                        li.SubItems.Add(ReportUtilities.GetStationString(oProfileDesign.Alignment, oProfileDesign.EndingStation))
                        li.SubItems.Add(oProfileExisting.Alignment.Name)
                        li.Checked = True
                        li.Tag = oProfileExisting
                        nCount += 1
                        ListView_Profile.Items.Add(li)
                    End If
                Next
            End If
        Next
    End Sub
End Class
