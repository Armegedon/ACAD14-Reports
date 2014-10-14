' -----------------------------------------------------------------------------
' <copyright file="ReportWriter.vb" company="Autodesk">
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
Imports System.Web.UI
Imports Microsoft
Imports LocalizedRes = Report.My.Resources.LocalizedStrings


Public Class ReportWriter

    Private sw As IO.StreamWriter
    Private writer As HtmlTextWriter

    Public Sub New(ByVal sFilePath As String)
        Try
            sw = New IO.StreamWriter(sFilePath)
            writer = New HtmlTextWriter(sw)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RenderEndTag()
        If writer Is Nothing Then
            Exit Sub
        End If
        '<html>
        writer.RenderEndTag()
    End Sub

    Public Sub HtmlBegin()
        If writer Is Nothing Then
            Exit Sub
        End If
        '<html>
        writer.RenderBeginTag(HtmlTextWriterTag.Html)
    End Sub

    Public Sub HtmlEnd()
        RenderEndTag()

        writer.Flush()
        sw.Flush()

        sw.Close()
        writer.Close()
    End Sub

    Public Sub HeaderBegin()
        If writer Is Nothing Then
            Exit Sub
        End If
        '<head>
        writer.RenderBeginTag(HtmlTextWriterTag.Head)
        Dim sEncoding As String
        sEncoding = "<meta http-equiv=""Content-Type"" content=""text/html; charset=" & sw.Encoding.WebName & """ />"
        writer.WriteLine(sEncoding)
    End Sub

    Public Sub HeaderEnd()
        '</head>
        RenderEndTag()
    End Sub

    Public Sub RenderTitle(ByVal sTitle As String)
        If writer Is Nothing Then
            Exit Sub
        End If

        writer.RenderBeginTag(HtmlTextWriterTag.Title)
        writer.Write(sTitle)
        writer.RenderEndTag()
    End Sub

    Public Sub RenderHeader(ByVal sTitle As String)
        HeaderBegin()
        RenderTitle(sTitle)
        HeaderEnd()
    End Sub

    
    Public Sub BodyBegin()
        If writer Is Nothing Then
            Exit Sub
        End If

        '<body>
        writer.RenderBeginTag(HtmlTextWriterTag.Body)
    End Sub

    Public Sub BodyEnd()
        '</body>
        RenderEndTag()
    End Sub

    'alignValue : left, center, right, justify
    Public Sub RenderH(ByVal sHeader As String, ByVal iLevel As Integer, Optional ByVal AlignValue As String = "center")
        If writer Is Nothing Then
            Exit Sub
        End If

        writer.AddAttribute(HtmlTextWriterAttribute.Align, AlignValue)
        
        Dim tags As HtmlTextWriterTag() = New HtmlTextWriterTag() _
        {HtmlTextWriterTag.H1, HtmlTextWriterTag.H2, HtmlTextWriterTag.H3, _
        HtmlTextWriterTag.H4, HtmlTextWriterTag.H5, HtmlTextWriterTag.H6}

        If iLevel < 1 And iLevel > 6 Then
            iLevel = 0
        End If

        writer.RenderBeginTag(tags(iLevel - 1))
        writer.Write(sHeader)
        writer.RenderEndTag()
    End Sub

    Public Sub ImgBegin(Optional ByVal sBorder As String = "0", Optional ByVal sWidth As String = "100%")
        If writer Is Nothing Then
            Exit Sub
        End If

        '<Img>
        If Not sBorder Is Nothing Then
            writer.AddAttribute(HtmlTextWriterAttribute.Border, sBorder)
        End If
        If Not sWidth Is Nothing Then
            writer.AddAttribute(HtmlTextWriterAttribute.Width, sWidth)
        End If
        writer.RenderBeginTag(HtmlTextWriterTag.Img)
    End Sub

    Public Sub ImgEnd()
        '</Img>
        RenderEndTag()
    End Sub

    Public Sub TableBegin(Optional ByVal sBorder As String = "0", Optional ByVal sWidth As String = "100%")
        If writer Is Nothing Then
            Exit Sub
        End If

        '<body>
        If Not sBorder Is Nothing Then
            writer.AddAttribute(HtmlTextWriterAttribute.Border, sBorder)
        End If
        If Not sWidth Is Nothing Then
            writer.AddAttribute(HtmlTextWriterAttribute.Width, sWidth)
        End If
        writer.RenderBeginTag(HtmlTextWriterTag.Table)
    End Sub

    Public Sub AddAttribute(ByVal name As System.Web.UI.HtmlTextWriterAttribute, ByVal value As String)
        If writer Is Nothing Then
            Exit Sub
        End If

        writer.AddAttribute(name, value)
    End Sub

    Public Sub TableEnd()
        '</body>
        RenderEndTag()
    End Sub

    Public Sub TrBegin()
        If writer Is Nothing Then
            Exit Sub
        End If

        writer.RenderBeginTag(HtmlTextWriterTag.Tr)
    End Sub

    Public Sub TrEnd()
        '</tr>
        RenderEndTag()
    End Sub

    Public Sub TdBegin(Optional ByVal sAlign As String = "left")
        If writer Is Nothing Then
            Exit Sub
        End If

        writer.AddAttribute(HtmlTextWriterAttribute.Align, sAlign)
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
    End Sub

    Public Sub TdEnd()
        '</td>
        RenderEndTag()
    End Sub

    Public Sub TBodyBegin()
        If writer Is Nothing Then
            Exit Sub
        End If

        writer.RenderBeginTag(HtmlTextWriterTag.Tbody)
    End Sub

    Public Sub TBodyEnd()
        '</tbody>
        RenderEndTag()
    End Sub

    Public Sub RenderTd(ByVal str As String, _
                        Optional ByVal bBold As Boolean = False, _
                        Optional ByVal sAlign As String = "left", _
                        Optional ByVal paddingPtLeft As String = Nothing, _
                        Optional ByVal paddingPtTop As String = Nothing, _
                        Optional ByVal paddingPtRight As String = Nothing, _
                        Optional ByVal paddingPtBottom As String = Nothing, _
                        Optional ByVal colSpan As String = Nothing)

        If writer Is Nothing Then
            Exit Sub
        End If

        Dim styleValue As String
        styleValue = Nothing
        If Not paddingPtLeft Is Nothing Then
            styleValue = "PADDING-RIGHT: " + paddingPtLeft + "pt"
        End If
        If Not paddingPtTop Is Nothing Then
            If Not styleValue Is Nothing Then
                styleValue += "; "
            End If
            styleValue += "PADDING-TOP: " + paddingPtTop + "pt"
        End If
        If Not paddingPtRight Is Nothing Then
            If Not styleValue Is Nothing Then
                styleValue += "; "
            End If
            styleValue += "PADDING-RIGHT: " + paddingPtRight + "pt"
        End If
        If Not paddingPtBottom Is Nothing Then
            If Not styleValue Is Nothing Then
                styleValue += "; "
            End If
            styleValue += "PADDING-BOTTOM: " + paddingPtBottom + "pt"
        End If
        If Not styleValue Is Nothing Then
            writer.AddAttribute(HtmlTextWriterAttribute.Style, styleValue)
        End If

        If Not colSpan Is Nothing Then
            writer.AddAttribute(HtmlTextWriterAttribute.Colspan, colSpan)
        End If

        TdBegin(sAlign)
        If bBold = True Then
            str = "<b>" & str & "</b>"
        End If
        writer.Write(str)
        TdEnd()
    End Sub

    Public Sub RenderDiv(ByVal str As String, _
                        Optional ByVal paddingPtLeft As String = Nothing, _
                        Optional ByVal paddingPtTop As String = Nothing, _
                        Optional ByVal paddingPtRight As String = Nothing, _
                        Optional ByVal paddingPtBottom As String = Nothing)

        If writer Is Nothing Then
            Exit Sub
        End If

        Dim styleValue As String
        styleValue = Nothing
        If Not paddingPtLeft Is Nothing Then
            styleValue = "PADDING-RIGHT: " + paddingPtLeft + "pt"
        End If
        If Not paddingPtTop Is Nothing Then
            If Not styleValue Is Nothing Then
                styleValue += "; "
            End If
            styleValue += "PADDING-TOP: " + paddingPtTop + "pt"
        End If
        If Not paddingPtRight Is Nothing Then
            If Not styleValue Is Nothing Then
                styleValue += "; "
            End If
            styleValue += "PADDING-RIGHT: " + paddingPtRight + "pt"
        End If
        If Not paddingPtBottom Is Nothing Then
            If Not styleValue Is Nothing Then
                styleValue += "; "
            End If
            styleValue += "PADDING-BOTTOM: " + paddingPtBottom + "pt"
        End If
        If Not styleValue Is Nothing Then
            writer.AddAttribute(HtmlTextWriterAttribute.Style, styleValue)
        End If

        writer.RenderBeginTag(HtmlTextWriterTag.Div)
        writer.Write(str)
        writer.RenderEndTag()
    End Sub

    Public Sub RenderDateInfo()
        If writer Is Nothing Then
            Exit Sub
        End If

        Dim sDataInfo As String
        Dim dispDt As DateTime = DateTime.Now

        sDataInfo = String.Format(LocalizedRes.Common_Html_Date, dispDt.ToString())

        writer.WriteLine(sDataInfo)
    End Sub

    Public Sub RenderHr()
        If writer Is Nothing Then
            Exit Sub
        End If

        writer.RenderBeginTag(HtmlTextWriterTag.Hr)
        writer.RenderEndTag()
    End Sub

    Public Sub RenderLine(ByVal s As String, Optional ByVal bCenter As Boolean = False)
        If writer Is Nothing Then
            Exit Sub
        End If

        If bCenter = True Then
            writer.RenderBeginTag(HtmlTextWriterTag.Center)
        End If
        writer.WriteLine(s)
        If bCenter = True Then
            writer.RenderEndTag()
        End If
        RenderBr()
    End Sub

    Public Sub RenderBr()
        If writer Is Nothing Then
            Exit Sub
        End If

        ' Following statement can't create self-closing tag <br/>.
        ' so we have to hard code "br"
        '  writer.RenderBeginTag(HtmlTextWriterTag.Br)
        '  writer.RenderEndTag()
        writer.Write(HtmlTextWriter.TagLeftChar & "br" & HtmlTextWriter.SelfClosingTagEnd)
    End Sub

    Public Sub Render(ByVal s As String, Optional ByVal bBold As Boolean = False)
        If writer Is Nothing Then
            Exit Sub
        End If

        If bBold = True Then
            writer.RenderBeginTag(HtmlTextWriterTag.B)
        End If

        writer.Write(s)

        If bBold = True Then
            writer.RenderEndTag()
        End If
    End Sub


    Public Sub RenderUserSetting()
        Try
            Dim sSetting As String
            Dim oUserSetting As New ReportUserSetting

            'add setting
            TableBegin()

            TrBegin()
            RenderTd(LocalizedRes.ReportWriter_Client, True)
            RenderTd(LocalizedRes.ReportWriter_PreparedBy, True)
            TrEnd()

            TrBegin()
            sSetting = oUserSetting.GetSettingValue("Client", "Contact")
            RenderTd(sSetting)
            sSetting = oUserSetting.GetSettingValue("Owner", "Preparer")
            RenderTd(sSetting)
            TrEnd()

            TrBegin()
            sSetting = oUserSetting.GetSettingValue("Client", "Company")
            RenderTd(sSetting)
            sSetting = oUserSetting.GetSettingValue("Owner", "Company")
            RenderTd(sSetting)
            TrEnd()

            TrBegin()
            sSetting = oUserSetting.GetSettingValue("Client", "Address1")
            RenderTd(sSetting)
            sSetting = oUserSetting.GetSettingValue("Owner", "Address1")
            RenderTd(sSetting)
            TrEnd()
            TableEnd()
        Catch ex As Exception
            ReportUtilities.ACADMsgBox(LocalizedRes.AlignSta_Html_ReadConfigError)
        End Try
    End Sub

End Class
