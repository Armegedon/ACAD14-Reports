Attribute VB_Name = "Application"
''//
''// (C) Copyright 2005 by Autodesk, Inc.
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
'
'
'
Option Explicit
Public g_oCivilApp As AeccApplication
Public g_oAeccDb As AeccDatabase
Public g_oCivilRoadwayApp As AeccRoadwayApplication
Public g_oRoadwayDocument As AeccRoadwayDocument


Private Const sAppName = "AeccXUiLand.AeccApplication." + MODVER
Private Const sRoadwayAppName = "AeccXUiRoadway.AeccRoadwayApplication." + MODVER

'
' Start Civil 3D, create document and database objects.
'


Public Function InitCooridorApp() As Boolean
    Dim oApp As AcadApplication
    Set oApp = ThisDrawing.Application
    
    Set g_oCivilRoadwayApp = oApp.GetInterfaceObject(sRoadwayAppName)
    If g_oCivilRoadwayApp Is Nothing Then
        MsgBox "Error creating " & sRoadwayAppName & ", exit."
        InitCooridorApp = False
        Exit Function
    End If
    Set g_oRoadwayDocument = g_oCivilRoadwayApp.ActiveDocument
    Set g_oAeccDb = g_oCivilRoadwayApp.ActiveDocument.Database
    InitCooridorApp = True
End Function
