' -----------------------------------------------------------------------------
' <copyright file="CtrlCrossSectionListView.vb" company="Autodesk">
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
Imports Autodesk.AECC.Interop.Land
Imports Autodesk.AECC.Interop.Roadway
Imports AutoCAD = Autodesk.AutoCAD.Interop
Imports LocalizedRes = Report.My.Resources.LocalizedStrings
Imports Autodesk.AECC.Interop

Friend Class CtrlCrossSectionListView
    Inherits CtrlListViewWithStation
    'Column index of ListView
    Public Const INDEX_INCLUDE As Integer = 0
    Public Const INDEX_NAME As Integer = 1
    Public Const INDEX_DESC As Integer = 2
    Public Const INDEX_STASTART As Integer = 3
    Public Const INDEX_STAEND As Integer = 4
    Public Const INDEX_ALIGNMENT As Integer = 5
    Public Const INDEX_PROFILE As Integer = 6
    Public Const INDEX_CORRIDOR As Integer = 7
    Public Const INDEX_COL_SUM As Integer = INDEX_CORRIDOR

    Private m_oSampleLineGroupeArr As Land.AeccSampleLineGroup()
    Private needSurfaceProfile_ As Boolean = False

    Public ReadOnly Property SampleLineGroupeArr() As Land.AeccSampleLineGroup()
        Get
            Return m_oSampleLineGroupeArr
        End Get
    End Property

    Public Property NeedSurfaceProfile() As Boolean
        Get
            Return Me.needSurfaceProfile_
        End Get
        Set(ByVal value As Boolean)
            Me.needSurfaceProfile_ = value
        End Set
    End Property

    Public Overrides Sub Initialize(ByVal oListView As ListView, _
                                    ByVal startStationTextBox As CtrlStationTextBox, _
                                    ByVal endStationTextBox As CtrlStationTextBox)

        MyBase.Initialize(oListView, startStationTextBox, endStationTextBox)

        '3. fill list view
        FillListView_CrossSect()
    End Sub


    Public ReadOnly Property CheckedCount() As Integer
        Get
            Return WinListView.CheckedItems.Count
        End Get
    End Property

    Public ReadOnly Property CheckedItems() As System.Windows.Forms.ListView.CheckedListViewItemCollection
        Get
            Return WinListView.CheckedItems
        End Get
    End Property

    Public ReadOnly Property HasSelectedItem() As Boolean
        Get
            If WinListView.SelectedItems.Count = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    Private ReadOnly Property SelectedLineGroupe() As Land.AeccSampleLineGroup
        Get
            If HasSelectedItem = True Then
                Return m_oSampleLineGroupeArr(WinListView.SelectedItems.Item(0).Index)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Protected Overrides Sub setupListView()
        ' list view grid setup
        ' setup columns
        WinListView.View = View.Details
        WinListView.Columns.Add(LocalizedRes.CrossSection_ListView_ColumnTitle_Include, _
            50, HorizontalAlignment.Left)
        WinListView.Columns.Add(LocalizedRes.CrossSection_ListView_ColumnTitle_Name, _
            100, HorizontalAlignment.Left)
        WinListView.Columns.Add(LocalizedRes.CrossSection_ListView_ColumnTitle_Description, _
            100, HorizontalAlignment.Left)
        WinListView.Columns.Add(LocalizedRes.CrossSection_ListView_ColumnTitle_StationStart, _
            80, HorizontalAlignment.Left)
        WinListView.Columns.Add(LocalizedRes.CrossSection_ListView_ColumnTitle_StationEnd, _
            80, HorizontalAlignment.Left)
        WinListView.Columns.Add(LocalizedRes.CrossSection_ListView_ColumnTitle_Alignment, _
            80, HorizontalAlignment.Left)
        WinListView.Columns.Add(LocalizedRes.CrossSection_ListView_ColumnTitle_Profile, _
            80, HorizontalAlignment.Left)
        WinListView.Columns.Add(LocalizedRes.CrossSection_ListView_ColumnTitle_Corridor, _
            80, HorizontalAlignment.Left)
    End Sub

    Public Shared Function CheckConditions(Optional ByVal openMsg As Boolean = True, _
                                           Optional ByVal needStaticSurface As Boolean = False) As Boolean
        ' Check if such an alignment exists, that has a profile and also that a corridor and a Sample Line Group exists in the drawing


        ' not using following resume can cause this function to terminate unexpectedly :
        ' in case of using for each statment on an object array that has count = 0 the execution
        ' would continue on the next statment in the function, from where this function was called
        ' and that was last using 'On Error Resume Next' statment
        On Error Resume Next

        Dim bExistInDoc As Boolean = False

        For Each oAlignment As Land.AeccAlignment In ReportApplication.AeccXDocument.AlignmentsSiteless
            If Not oAlignment.Profiles Is Nothing Then
                bExistInDoc = True
                Exit For
            End If
        Next

        If Not bExistInDoc Then
            'get alignments from Site
            For Each oSite As Land.AeccSite In ReportApplication.AeccXDatabase.Sites
                For Each oAlignment As Land.AeccAlignment In oSite.Alignments
                    If Not oAlignment.Profiles Is Nothing Then
                        bExistInDoc = True
                        Exit For
                    End If
                Next
                If bExistInDoc Then
                    Exit For
                End If
            Next
        End If

        If Not bExistInDoc Then
            If openMsg Then
                Dim sNoProfile As String = LocalizedRes.CrossSection_ListView_NoProfileInDrawing
                ReportUtilities.ACADMsgBox(sNoProfile.Replace("&DrawingName&", ReportApplication.AeccXDocument.Name))
            End If
            CheckConditions = False
            Exit Function
        End If

        ' check also for a corridor
        bExistInDoc = True
        If ReportApplication.AeccXRoadwayDocument.Corridors.Count < 1 Then
            bExistInDoc = False
        End If

        If Not bExistInDoc Then
            If openMsg Then
                Dim sNoCorridor As String = LocalizedRes.CrossSection_ListView_NoCorridorInDrawing
                ReportUtilities.ACADMsgBox(sNoCorridor.Replace("&DrawingName&", ReportApplication.AeccXDocument.Name))
            End If
            CheckConditions = False
            Exit Function
        End If

        ' and for a Sample Line Group :
        ' for the last and the most thorough check before FillListView_CrossSect() :
        ' check for corridor and profile on an alignment, that has line groups - if this succeeds 
        ' there is at least one row in the report list view, for which to run the report

        Dim oCorridor As Roadway.AeccCorridor = Nothing
        Dim oProfile As Land.AeccProfile = Nothing
        bExistInDoc = False

        Dim noSurfaceProfile As Boolean = False

        For Each oAlignment As Land.AeccAlignment In ReportApplication.AeccXDocument.AlignmentsSiteless
            If FindCorridor(oAlignment, oCorridor, oProfile) Then

                For Each oSampleLineGroup As Land.AeccSampleLineGroup In oAlignment.SampleLineGroups
                    If Not oSampleLineGroup Is Nothing Then
                        bExistInDoc = True
                        Exit For
                    End If
                Next

                If bExistInDoc Then
                    If needStaticSurface Then
                        If Daylight_ExtractData.FindEGSurface(oAlignment) Is Nothing Then
                            noSurfaceProfile = True
                        Else
                            noSurfaceProfile = False
                        End If
                    End If
                End If

            End If
            If bExistInDoc Then
                Exit For
            End If
        Next

        If Not bExistInDoc Then
            'get alignments from Site
            For Each oSite As Land.AeccSite In ReportApplication.AeccXDatabase.Sites
                For Each oAlignment As Land.AeccAlignment In oSite.Alignments
                    If FindCorridor(oAlignment, oCorridor, oProfile) Then
                        For Each oSampleLineGroup As Land.AeccSampleLineGroup In oAlignment.SampleLineGroups
                            If Not oSampleLineGroup Is Nothing Then
                                bExistInDoc = True
                                Exit For
                            End If
                        Next

                        If bExistInDoc Then
                            If needStaticSurface Then
                                If Daylight_ExtractData.FindEGSurface(oAlignment) Is Nothing Then
                                    noSurfaceProfile = True
                                Else
                                    noSurfaceProfile = False
                                End If
                            End If
                        End If

                    End If
                    If bExistInDoc Then
                        Exit For
                    End If
                Next
                If bExistInDoc Then
                    Exit For
                End If
            Next 'site
        End If

        If Not bExistInDoc Then
            If openMsg Then
                Dim sNoSLG As String = LocalizedRes.CrossSection_ListView_NoSampleLineGroupsInDrawing
                ReportUtilities.ACADMsgBox(sNoSLG.Replace("&DrawingName&", ReportApplication.AeccXDocument.Name))
            End If
            CheckConditions = False
            Exit Function
        End If

        If bExistInDoc Then
            If needStaticSurface Then
                If noSurfaceProfile Then
                    If openMsg Then
                        Dim errorMsg As String = LocalizedRes.CrossSectionListView_NoSurfaceProfile
                        ReportUtilities.ACADMsgBox(errorMsg)
                    End If
                    CheckConditions = False
                    Exit Function
                End If
            End If
        End If

        CheckConditions = bExistInDoc
    End Function

    Private Sub FillListView_CrossSect()
        On Error Resume Next

        Dim oCorridor As Roadway.AeccCorridor = Nothing
        Dim oProfile As Land.AeccProfile = Nothing
        Dim nSampleLineGroup As Integer = 0
        Dim nCurIndex As Integer = 0

        For Each oAlignment As Land.AeccAlignment In ReportApplication.AeccXDocument.AlignmentsSiteless
            If ReportApplication.AeccXDocument.AlignmentsSiteless.Count = 0 Then Exit For
            If FindCorridor(oAlignment, oCorridor, oProfile) Then

                If Me.NeedSurfaceProfile Then
                    If Daylight_ExtractData.FindEGSurface(oAlignment) Is Nothing Then
                        Continue For
                    End If
                End If

                nSampleLineGroup = nSampleLineGroup + oAlignment.SampleLineGroups.Count
                ReDim Preserve m_oSampleLineGroupeArr(nSampleLineGroup)
                For Each oSampleLineGroup As Land.AeccSampleLineGroup In oAlignment.SampleLineGroups
                    If oAlignment.SampleLineGroups.Count = 0 Then
                        Exit For ' an exception occured because of empty oAlignment.SampleLineGroups
                    End If
                    Dim li As New ListViewItem

                    li.SubItems.Add(oSampleLineGroup.Name)
                    li.SubItems.Add(oSampleLineGroup.Description)
                    li.SubItems.Add(ReportUtilities.GetStationString(oAlignment, oAlignment.StartingStation))
                    li.SubItems.Add(ReportUtilities.GetStationString(oAlignment, oAlignment.EndingStation))
                    li.SubItems.Add(oAlignment.Name)
                    li.SubItems.Add(oProfile.Name)
                    li.SubItems.Add(oCorridor.Name)
                    li.Checked = True
                    m_oSampleLineGroupeArr(nCurIndex) = oSampleLineGroup
                    nCurIndex = nCurIndex + 1
                    nSampleLineGroup = nSampleLineGroup + 1
                    WinListView.Items.Add(li)
                Next
            End If
        Next

        For Each oSite As Land.AeccSite In ReportApplication.AeccXDatabase.Sites
            If ReportApplication.AeccXDatabase.Sites.Count = 0 Then Exit For
            For Each oAlignment As Land.AeccAlignment In oSite.Alignments
                If oSite.Alignments.Count = 0 Then Exit For
                If FindCorridor(oAlignment, oCorridor, oProfile) Then

                    If Me.NeedSurfaceProfile Then
                        If Daylight_ExtractData.FindEGSurface(oAlignment) Is Nothing Then
                            Continue For
                        End If
                    End If

                    nSampleLineGroup = nSampleLineGroup + oAlignment.SampleLineGroups.Count
                    ReDim Preserve m_oSampleLineGroupeArr(nSampleLineGroup)
                    For Each oSampleLineGroup As Land.AeccSampleLineGroup In oAlignment.SampleLineGroups
                        If oAlignment.SampleLineGroups.Count = 0 Then
                            Exit For ' an exception occured because of empty oAlignment.SampleLineGroups
                        End If

                        Dim li As New ListViewItem

                        li.SubItems.Add(oSampleLineGroup.Name)
                        li.SubItems.Add(oSampleLineGroup.Description)
                        li.SubItems.Add(ReportUtilities.GetStationString(oAlignment, oAlignment.StartingStation))
                        li.SubItems.Add(ReportUtilities.GetStationString(oAlignment, oAlignment.EndingStation))
                        li.SubItems.Add(oAlignment.Name)
                        li.SubItems.Add(oProfile.Name)
                        li.SubItems.Add(oCorridor.Name)
                        li.Checked = True
                        m_oSampleLineGroupeArr(nCurIndex) = oSampleLineGroup
                        nCurIndex = nCurIndex + 1
                        nSampleLineGroup = nSampleLineGroup + 1
                        WinListView.Items.Add(li)
                    Next
                End If
            Next
        Next 'site
        If (nSampleLineGroup = 0) Then
            Dim sMsg As String = LocalizedRes.CrossSection_ListView_NoSampleLineGroupsInDrawing.Replace("&DrawingName&", ReportApplication.AeccXDocument.Name)
            'MsgBox("Pas de groupe de tabulations présent dans ce dessin : " & ThisDrawing.Name)
            ReportUtilities.ACADMsgBox(sMsg)
        End If
    End Sub

    Private Shared Function FindCorridor(ByVal oAlignment As Land.AeccAlignment, _
                                         ByRef oCorridor As Roadway.AeccCorridor, _
                                         ByRef oProfile As Land.AeccProfile) As Boolean

        Dim bCorridorFound As Boolean = False

        'sélection de tous les corridors
        Dim gpCode(0) As Short
        Dim dataValue(0) As Object
        gpCode(0) = 0
        dataValue(0) = "AECC_CORRIDOR"

        Dim groupCode As Object = gpCode, dataCode As Object = dataValue

        ' Create the selection set

        Dim ssetObj As AutoCAD.AcadSelectionSet = Nothing
        Dim bSelectionSetExist As Boolean = False
        For Each ssetObjTemp As AutoCAD.AcadSelectionSet In ReportApplication.AeccXDocument.SelectionSets
            If ssetObjTemp.Name = "SS_AECC_CORRIDOR" Then
                bSelectionSetExist = True
                ssetObjTemp.Clear()
                ssetObj = ssetObjTemp
            End If
        Next
        If Not bSelectionSetExist Then
            ssetObj = ReportApplication.AeccXDocument.SelectionSets.Add("SS_AECC_CORRIDOR")
        End If
        'sélection de tous les corridors
        ssetObj.Select(AutoCAD.Common.AcSelect.acSelectionSetAll, , , groupCode, dataCode)

        'boucle sur tous les corridors
        For Each oCorridorTemp As Roadway.AeccCorridor In ssetObj
            'boucle sur les ligne de base
            For Each oBaseLine As Roadway.AeccBaseline In oCorridorTemp.Baselines
                If oBaseLine.Alignment.Name = oAlignment.Name Then
                    oProfile = oBaseLine.Profile
                    oCorridor = oCorridorTemp
                    bCorridorFound = True
                End If
            Next
        Next

        FindCorridor = bCorridorFound

    End Function

    Protected Overrides Function getAssociateAlignment(ByVal item As ListViewItem) As Land.IAeccAlignment
        Try
            Return SelectedLineGroupe.Parent
        Catch
            Return Nothing
        End Try
    End Function

    Protected Overrides Property startStationText(ByVal item As ListViewItem) As String
        Get
            Try
                Return item.SubItems.Item(INDEX_STASTART).Text
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
                Return ""
            End Try
        End Get
        Set(ByVal value As String)
            Try
                item.SubItems.Item(INDEX_STASTART).Text = value
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
            End Try
        End Set
    End Property

    Protected Overrides Property endStationText(ByVal item As ListViewItem) As String
        Get
            Try
                Return item.SubItems.Item(INDEX_STAEND).Text
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
                Return ""
            End Try
        End Get
        Set(ByVal value As String)
            Try
                item.SubItems.Item(INDEX_STAEND).Text = value
            Catch ex As Exception
                Diagnostics.Debug.Assert(False, ex.Message)
            End Try
        End Set
    End Property
End Class
