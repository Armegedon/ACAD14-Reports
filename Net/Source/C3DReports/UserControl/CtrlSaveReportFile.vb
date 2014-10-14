' -----------------------------------------------------------------------------
' <copyright file="CtrlSaveReportFile.vb" company="Autodesk">
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


Imports System.IO
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Public Class CtrlSaveReportFile
    Inherits CtrlSaveFile

    Public Sub New(ByVal textBoxFilePath As TextBox, ByVal buttonOpen As Button)
        MyBase.New(textBoxFilePath, buttonOpen)

        ' add default html filter
        AddFilter(LocalizedRes.FileDialog_Filter_Html, _
                  ReportConverter.FileType.Html)
        'add word doc filter
        If ReportConverter.IsWordInstalled Then
            AddFilter(LocalizedRes.FileDialog_Filter_Doc, _
                      ReportConverter.FileType.Doc)
        End If
        'add excel and text filter
        If ReportConverter.IsExcelInstalled Then
            AddFilter(LocalizedRes.FileDialog_Filter_Xls, _
                      ReportConverter.FileType.Excel)
            AddFilter(LocalizedRes.FileDialog_Filter_Txt, _
                      ReportConverter.FileType.Text)
        End If

        'add pdf filter
        AddFilter(LocalizedRes.FileDialog_Filter_Pdf, ReportConverter.FileType.Pdf)

        SaveDialogDefaultExt = ReportConverter.ExtHtml.Substring(1) ' ExtHtml includes ".". Drop it and get "html".
        SaveDialogFilter = CreateFilter()

        SavePath = Environ("temp") + "\civilreport" + ReportConverter.ExtHtml

    End Sub


    Public ReadOnly Property ReportFileType() As ReportConverter.FileType
        Get
            Return reportFileType_
        End Get
    End Property

    Private Function ExtractFileType(ByVal filePath As String) As ReportConverter.FileType
        Dim type As ReportConverter.FileType

        Dim ext As String = Path.GetExtension(filePath).ToLower
        type = ReportConverter.GetFileType(ext)

        If Not ReportConverter.IsSupported(type) Then
            type = ReportConverter.FileType.None
        End If

        Return type
    End Function

    Protected Overrides Sub OnSelectSaveFile(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Do
            Dim tempPath As String = SavePath
            MyBase.OnSelectSaveFile(sender, e)
            reportFileType_ = ExtractFileType(SavePath)
            If reportFileType_ = ReportConverter.FileType.None Then
                ReportUtilities.ACADMsgBox(LocalizedRes.FileDialog_Error_Ext, _
                                           LocalizedRes.Report_InitError)
                SavePath = tempPath 'restore the path
            Else
                Exit Do
            End If
        Loop While True

    End Sub


    Private Sub AddFilter(ByVal filterStr As String, _
                          ByVal filterType As ReportConverter.FileType)
        filterArray_.Add(New Filter(filterStr, filterType))
    End Sub

    Private Function CreateFilter() As String
        If filterArray_.Count > 0 Then
            Dim filterStr As String = filterArray_.Item(0).filterString

            Dim i As Integer = 1
            Do While (i < filterArray_.Count)
                filterStr += "|" & filterArray_.Item(i).filterString
                i = i + 1
            Loop

            Return filterStr
        Else
            Return LocalizedRes.FileDialog_Filter_All
        End If
    End Function

    Public Class Filter
        Public Sub New(ByVal filtr As String, ByVal type As ReportConverter.FileType)
            filterString = filtr
            filterType = type
        End Sub

        Public filterString As String
        Public filterType As ReportConverter.FileType
    End Class

    Private reportFileType_ As ReportConverter.FileType = ReportConverter.FileType.Html
    Private filterArray_ As Generic.List(Of Filter) = New List(Of Filter)

End Class
