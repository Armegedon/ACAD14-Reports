VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
'
'Main entry point for this reports application
'
Sub AlignStakeout_RunReport()
    On Error Resume Next
    
    Dim rptDlg As ReportForm_AlignStakeout
    
    ' initialize C3D
    If InitCivil3D = False Then
        Call MsgBox("Error: Unable to initialize Civil 3D Object Model", vbOKOnly, "Civil 3D Reports")
        Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    ' Create report dialog and display it
    Set rptDlg = New ReportForm_AlignStakeout
    If Err.Number <> 0 Then
        Exit Sub
    End If
    rptDlg.Show
    
    Set rptDlg = Nothing
End Sub


