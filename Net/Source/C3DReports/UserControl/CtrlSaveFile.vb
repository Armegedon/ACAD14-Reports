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

Imports LocalizedRes = Report.My.Resources.LocalizedStrings
Imports Microsoft.Office.Interop

Public Class CtrlSaveFile

    Public Sub New(ByVal textBoxFilePath As TextBox, ByVal buttonOpen As Button)
        textBoxFilePath_ = textBoxFilePath
        textBoxFilePath_.ReadOnly = True
        buttonSelectFile_ = buttonOpen
        Try
            Dim bp As Drawing.Bitmap = My.Resources.Resources.FileSelectionBmp
            bp.MakeTransparent(Color.White)
            buttonSelectFile_.Image = bp
            buttonSelectFile_.ImageAlign = ContentAlignment.MiddleCenter
        Catch ex As Exception
            Diagnostics.Debug.Assert(False, ex.Message)
        End Try

    End Sub


    Public Property SavePath() As String
        Get
            Return textBoxFilePath_.Text
        End Get
        Set(ByVal value As String)
            textBoxFilePath_.Text = value
        End Set
    End Property


    Public Property SaveFileIndex() As Integer
        Get
            Return saveDialogIndex_
        End Get
        Protected Set(ByVal value As Integer)
            saveDialogIndex_ = value
        End Set
    End Property

    Property SaveDialogDefaultExt() As String
        Get
            Return saveDialogDefaultExt_
        End Get
        Set(ByVal value As String)
            saveDialogDefaultExt_ = value
        End Set
    End Property

    Property SaveDialogTitle() As String
        Get
            Return saveDialogTitle_
        End Get
        Set(ByVal value As String)
            saveDialogTitle_ = value
        End Set
    End Property

    Property SaveDialogFilter() As String
        Get
            Return saveDialogFilter_
        End Get
        Set(ByVal value As String)
            saveDialogFilter_ = value
        End Set
    End Property

    Protected Overridable Sub OnSelectSaveFile(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonSelectFile_.Click
        Dim fileDialog As New SaveFileDialog
        With fileDialog
            .DefaultExt = saveDialogDefaultExt_
            .Title = saveDialogTitle_
            .Filter = saveDialogFilter_
            .ValidateNames = True
            .FileName = "CivilReport"
            .FilterIndex = saveDialogIndex_
            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                SavePath = .FileName
                saveDialogIndex_ = .FilterIndex
            Else
                Exit Sub
            End If
        End With
    End Sub


    Private WithEvents textBoxFilePath_ As TextBox
    Private WithEvents buttonSelectFile_ As Button
    Private saveDialogTitle_ As String
    Private saveDialogFilter_ As String
    Private saveDialogDefaultExt_ As String
    Private saveDialogIndex_ As Integer = 0

End Class
