' -----------------------------------------------------------------------------
' <copyright file="CtrlProgressBar.vb" company="Autodesk">
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

Public Class CtrlProgressBar

    Public Sub Initialize(ByVal progressBar As ProgressBar, ByVal groupBox As GroupBox)
        progressBar_ = progressBar
        groupBox_ = groupBox
        groupBox_.Visible = False
    End Sub

    Public ReadOnly Property ProgressBar() As System.Windows.Forms.ProgressBar
        Get
            Return progressBar_
        End Get
    End Property

    Public Sub ProgressBegin(ByVal stepCount As Integer)
        groupBox_.Visible = True
        groupBox_.Refresh()

        If stepCount <= 0 Then
            stepCount = 1
        End If

        'update progress bar
        progressBar_.Minimum = 0
        progressBar_.Maximum = stepCount
        progressBar_.Step = 1

        progressBar_.Value = 0
    End Sub

    Public Sub PerformStep()
        ProgressBar.PerformStep()
    End Sub

    Public Sub ProgressEnd()
        progressBar_.Value = ProgressBar.Maximum
        groupBox_.Visible = False
    End Sub

    Private WithEvents progressBar_ As ProgressBar
    Private WithEvents groupBox_ As GroupBox
End Class
