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
Public g_oAeccDoc As AeccDocument
Public g_oAeccDb As AeccDatabase

Public g_oSurveyApplication As AeccSurveyApplication
Public g_oSurveyDocument As AeccSurveyDocument
'
' Start Civil 3D, create document and database objects.
'
Public Function InitCivil3D() As Boolean
    Dim oApp As AcadApplication
    Set oApp = ThisDrawing.Application
    Const sAppName = "AeccXUiLand.AeccApplication." + MODVER
    Set g_oCivilApp = oApp.GetInterfaceObject(sAppName)
    If g_oCivilApp Is Nothing Then
        MsgBox "Error creating " & sAppName & ", exit."
        InitCivil3D = False
        Exit Function
    End If
    Set g_oAeccDoc = g_oCivilApp.ActiveDocument
    Set g_oAeccDb = g_oAeccDoc.Database
   
    InitCivil3D = True
End Function

'
'
' Function to set up the Civil 3D Survey application, document
' and database objects.
'
Function GetBaseSurveyObjects() As Boolean
    Dim oApp As AcadApplication
    Set oApp = ThisDrawing.Application
    Dim sAppName As String
    sAppName = "AeccXUiSurvey.AeccSurveyApplication"
    On Error Resume Next
    Set g_oSurveyApplication = oApp.GetInterfaceObject(sAppName)
    On Error GoTo 0
    If (g_oSurveyApplication Is Nothing) Then
        MsgBox "Error creating " & sAppName & ", exit."
        GetBaseSurveyObjects = False
        Exit Function
    End If
    Set g_oSurveyDocument = g_oSurveyApplication.ActiveDocument
    'Set g_oSurveyDatabase = g_oSurveyDocument.Database ' Do not need this currently
    GetBaseSurveyObjects = True
End Function

