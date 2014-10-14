''//
''// (C) Copyright 2009 by Autodesk, Inc.
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
Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Diagnostics.Debug

Public Class ReportConverter
    Public Enum FileType
        Html
        Doc
        Text
        Excel
        Pdf
        None
    End Enum

    Public Const ExtHtml As String = ".html"
    Public Const ExtDoc As String = ".doc"
    Public Const ExtText As String = ".txt"
    Public Const ExtExcel As String = ".xls"
    Public Const ExtPdf As String = ".pdf"

    Public Shared Function IsWordInstalled() As Boolean
        Return CheckProgID("Word.Application")
    End Function

    Public Shared Function IsExcelInstalled() As Boolean
        Return CheckProgID("Excel.Application")
    End Function

    Public Shared Function IsSupported(ByVal type As FileType) As Boolean
        Select Case type
            Case FileType.Html, FileType.Pdf
                Return True
            Case FileType.Doc
                Return ReportConverter.IsWordInstalled
            Case FileType.Excel, FileType.Text
                Return ReportConverter.IsExcelInstalled
            Case Else
                Return False
        End Select
    End Function

    Public Shared Function GetFileType(ByVal ext As String) As FileType
        Dim type As FileType

        If String.Compare(ext, ExtHtml, StringComparison.InvariantCultureIgnoreCase) = 0 Then
            type = FileType.Html
        ElseIf String.Compare(ext, ExtDoc, StringComparison.InvariantCultureIgnoreCase) = 0 Then
            type = FileType.Doc
        ElseIf String.Compare(ext, ExtExcel, StringComparison.InvariantCultureIgnoreCase) = 0 Then
            type = FileType.Excel
        ElseIf String.Compare(ext, ExtText, StringComparison.InvariantCultureIgnoreCase) = 0 Then
            type = FileType.Text
        ElseIf String.Compare(ext, ExtPdf, StringComparison.InvariantCultureIgnoreCase) = 0 Then
            type = FileType.Pdf
        Else
            type = FileType.None
        End If

        Return type
    End Function

    Public Shared Sub GenerateReportFileFromHTML(ByVal htmlFileName As String, _
                                                 ByVal reportFileName As String, _
                                                 ByVal type As FileType)
        If (type = FileType.Html) Then
            ConvertToHtml(htmlFileName, reportFileName)
        ElseIf type = FileType.Doc Then
            ConvertToDoc(htmlFileName, reportFileName)
        ElseIf type = FileType.Excel Then
            ConvertToXsl(htmlFileName, reportFileName)
        ElseIf type = FileType.Text Then
            ConvertToTxt(htmlFileName, reportFileName)
        ElseIf type = FileType.Pdf Then
            ConvertToPdf(htmlFileName, reportFileName)
        Else
            Assert(False, "Cannot be here")
        End If
    End Sub

    Public Shared Sub ConvertToHtml(ByVal fileFrom As String, ByVal fileTo As String)
        If File.Exists(fileTo) Then
            File.Delete(fileTo) 'If file existed, firstly we need delete it.
        End If
        File.Move(fileFrom, fileTo)

    End Sub

    Public Shared Sub ConvertToPdf(ByVal fileFrom As String, ByVal fileTo As String)
        Try
            If File.Exists(fileTo) Then
                File.Delete(fileTo) 'If file existed, firstly we need delete it.
            End If

            AeccConvertHtmlToPdf(fileFrom, fileTo)
        Catch ex As Exception
            Assert(False, ex.Message)
        End Try
    End Sub

    Public Shared Sub ConvertToDoc(ByVal fileFrom As String, ByVal fileTo As String)
        Dim wordApp As Word.Application = Nothing
        Dim doc As Word.Document = Nothing
        Try

            wordApp = New Word.Application
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
            doc = wordApp.Documents.Open(FileName:=CType(fileFrom, Object), _
                                         ReadOnly:=True)

            If File.Exists(fileTo) Then
                File.Delete(fileTo) 'If file existed, firstly we need delete it.
            End If

            doc.SaveAs(FileName:=CType(fileTo, Object), _
                       FileFormat:=Word.WdSaveFormat.wdFormatDocument)

        Finally
            If doc IsNot Nothing Then
                doc.Close()
                doc = Nothing
            End If
            If wordApp IsNot Nothing Then
                wordApp.Quit()
                wordApp = Nothing
            End If
        End Try

    End Sub

    Public Shared Sub ConvertToXsl(ByVal fileFrom As String, ByVal fileTo As String)

        Dim excelApp As Excel.Application = Nothing
        Dim xls As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application
            excelApp.DisplayAlerts = False
            xls = excelApp.Workbooks.Open(Filename:=fileFrom, _
                                          ReadOnly:=True)

            Dim sheet As Excel.Worksheet = CType(xls.Worksheets().Item(1), Excel.Worksheet)
            sheet.Name = "CivilReport"
            If File.Exists(fileTo) Then
                File.Delete(fileTo) 'If file existed, firstly we need delete it.
            End If

            xls.SaveAs(Filename:=fileTo, _
                       FileFormat:=Excel.XlFileFormat.xlWorkbookNormal, _
                       ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)

        Finally
            If xls IsNot Nothing Then
                xls.Close()
                xls = Nothing
            End If
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                excelApp = Nothing
            End If
        End Try
    End Sub

    Public Shared Sub ConvertToTxt(ByVal fileFrom As String, ByVal fileTo As String)
        Dim excelApp As Excel.Application = Nothing
        Dim xls As Excel.Workbook = Nothing
        Try
            excelApp = New Excel.Application
            excelApp.DisplayAlerts = False
            xls = excelApp.Workbooks.Open(Filename:=fileFrom, ReadOnly:=True)

            If File.Exists(fileTo) Then
                File.Delete(fileTo) 'If file existed, firstly we need delete it.
            End If

            xls.SaveAs(Filename:=fileTo, _
                       FileFormat:=Excel.XlFileFormat.xlUnicodeText, _
                       ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)

        Finally
            If xls IsNot Nothing Then
                xls.Close()
                xls = Nothing
            End If
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                excelApp = Nothing
            End If
        End Try
    End Sub

    Private Shared Function CheckProgID(ByVal progId As String) As Boolean
        Try
            Dim appType As Type = Type.GetTypeFromProgID(progId)
            If appType Is Nothing Then
                Return False
            End If
            appType = Nothing
            Return True
        Catch ex As Exception
            System.Diagnostics.Debug.Assert(False, ex.Message)
            Return False
        End Try
    End Function

    Public Shared Function GetRandomTempHtmlPath() As String
        Dim tempPath As String = System.IO.Path.GetTempPath()
        Dim fileName As String = System.IO.Path.GetRandomFileName()
        Return tempPath & fileName & ReportConverter.ExtHtml
    End Function


    ''' <summary>
    ''' Convert a Html file to Pdf format file.
    ''' </summary>
    ''' <param name="fileFrom">Html file name.</param>
    ''' <param name="fileTo">Pdf file name.</param>
    ''' <remarks>This is implemented in AeccDotNetUtils.</remarks>
    <DllImportAttribute(AeccDotNetUtilsDllName, EntryPoint:="AeccConvertHtmlToPdf")> _
    Private Shared Sub AeccConvertHtmlToPdf(<InAttribute(), MarshalAsAttribute(UnmanagedType.LPWStr)> ByVal fileFrom As String, _
                                           <InAttribute(), MarshalAsAttribute(UnmanagedType.LPWStr)> ByVal fileTo As String)
    End Sub

#If DEBUG Then
    Private Const AeccDotNetUtilsDllName As String = "AeccDotNetUtilsD.dll"
#Else
    Private Const AeccDotNetUtilsDllName As String = "AeccDotNetUtils.dll"
#End If

End Class
