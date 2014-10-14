' -----------------------------------------------------------------------------
' <copyright file="ReportApplication.vb" company="Autodesk">
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
Imports Autodesk.AutoCAD.Runtime
Imports AcMgApp = Autodesk.AutoCAD.ApplicationServices
Imports AeccXLand = Autodesk.AECC.Interop.Land
Imports AeccXUiLand = Autodesk.AECC.Interop.UiLand
Imports AeccXUiRoadway = Autodesk.AECC.Interop.UiRoadway


''' <summary>
''' Implement AutoCAD .NET Runtime extension application.
''' </summary>
''' <remarks></remarks>
Public Class ReportApplication
    Implements IExtensionApplication

    Public Sub Initialize() Implements IExtensionApplication.Initialize

        _aeccXApplication = Nothing
        _aeccXRoadwayApplication = Nothing

        'Dim acadApp As Autodesk.AutoCAD.Interop.IAcadApplication
        'acadApp = CType(AcMgApp.Application.AcadApplication, Autodesk.AutoCAD.Interop.IAcadApplication)

        '_aeccXApplication = New AeccXUiLand.AeccApplication
        '_aeccXApplication.Init(CType(acadApp, Autodesk.AutoCAD.Interop.AcadApplication))

        '_aeccXRoadwayApplication = New AeccXUiRoadway.AeccRoadwayApplication
        '_aeccXRoadwayApplication.Init(CType(acadApp, Autodesk.AutoCAD.Interop.AcadApplication))
    End Sub

    Public Sub Terminate() Implements IExtensionApplication.Terminate

    End Sub

    Public Shared ReadOnly Property ActiveDocument() As AcMgApp.Document
        Get
            Return AcMgApp.Application.DocumentManager.MdiActiveDocument
        End Get
    End Property

    Public Shared ReadOnly Property AeccXApplication() As AeccXUiLand.IAeccApplication
        Get
            If (_aeccXApplication Is Nothing) Then
                Dim acadApp As Autodesk.AutoCAD.Interop.IAcadApplication
                acadApp = CType(AcMgApp.Application.AcadApplication, Autodesk.AutoCAD.Interop.IAcadApplication)

                _aeccXApplication = New AeccXUiLand.AeccApplication
                _aeccXApplication.Init(CType(acadApp, Autodesk.AutoCAD.Interop.AcadApplication))
            End If
            Return _aeccXApplication
        End Get
    End Property

    Public Shared ReadOnly Property AeccXRoadwayApplication() As AeccXUiRoadway.IAeccRoadwayApplication
        Get
            If (_aeccXRoadwayApplication Is Nothing) Then
                Dim acadApp As Autodesk.AutoCAD.Interop.IAcadApplication
                acadApp = CType(AcMgApp.Application.AcadApplication, Autodesk.AutoCAD.Interop.IAcadApplication)

                _aeccXRoadwayApplication = New AeccXUiRoadway.AeccRoadwayApplication
                _aeccXRoadwayApplication.Init(CType(acadApp, Autodesk.AutoCAD.Interop.AcadApplication))
            End If
            Return _aeccXRoadwayApplication
        End Get
    End Property

    Public Shared ReadOnly Property AeccXDocument() As AeccXUiLand.IAeccDocument
        Get
            Return CType(AeccXApplication.ActiveDocument, AeccXUiLand.IAeccDocument)
        End Get
    End Property

    Public Shared ReadOnly Property AeccXRoadwayDocument() As AeccXUiRoadway.IAeccRoadwayDocument
        Get
            Return CType(AeccXRoadwayApplication.ActiveDocument, AeccXUiRoadway.IAeccRoadwayDocument)
        End Get
    End Property

    Public Shared ReadOnly Property AeccXDatabase() As AeccXLand.IAeccDatabase
        Get
            Return CType(AeccXApplication.ActiveDocument.Database, AeccXLand.IAeccDatabase)
        End Get
    End Property

    Public Shared ReadOnly Property AeccXRoadwayDatabase() As AeccXLand.IAeccDatabase
        Get
            Return CType(AeccXRoadwayApplication.ActiveDocument.Database, AeccXLand.IAeccDatabase)
        End Get
    End Property

    ''' <summary>
    ''' Invoke Civil3D help file for specified topic.
    ''' </summary>
    ''' <param name="topic"></param>
    ''' <remarks></remarks>
    Public Shared Sub InvokeHelp(ByVal topic As String)
        Try
            Dim helpPath As String
            helpPath = _aeccXApplication.Preferences.Files.HelpFilePath
            AcMgApp.Application.InvokeHelp(helpPath, topic)
        Catch ex As System.Exception
            Diagnostics.Debug.Assert(False, "Open help file failed : " + ex.Message)
        End Try
    End Sub

    Private Shared _aeccXApplication As AeccXUiLand.AeccApplication
    Private Shared _aeccXRoadwayApplication As AeccXUiRoadway.AeccRoadwayApplication

End Class

