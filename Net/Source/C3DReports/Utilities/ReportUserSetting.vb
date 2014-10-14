''//
''// (C) Copyright 2008 by Autodesk, Inc.
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

Imports System.Globalization
Imports System.Threading.Thread
Imports Microsoft

Public Class ReportUserSetting

    Private Const CapitalizationDefault As String = "default"
    Private Const CapitalizationLowercase As String = "Lowercase"
    Private Const CapitalizationUppercase As String = "Uppercase"
    Private Const CapitalizationSentenceCase As String = "Sentence Case"
    Private Const CapitalizationTitleCase As String = "Title Case"

    Private m_oXmlDoc As Xml.XmlDocument

    Private Function GetXmlDoc() As Xml.XmlDocument
        If m_oXmlDoc Is Nothing Then
            Dim sFullKeyPath As String = "HKEY_CURRENT_USER\Software\Autodesk\LandXML Reporting\8\Settings\"
            Dim sValueName As String = "SettingsFile"

            'get setting file path from reg table
            Dim sSettingFilePath As String = Win32.Registry.GetValue(sFullKeyPath, sValueName, String.Empty).ToString()

            m_oXmlDoc = New Xml.XmlDocument
            m_oXmlDoc.Load(sSettingFilePath)
        End If

        GetXmlDoc = m_oXmlDoc
    End Function

    Public Function GetSettingValue(ByVal sCategoryName As String, ByVal sSettingName As String) As String
        If GetXmlDoc() Is Nothing Then
            GetSettingValue = String.Empty
            Exit Function
        End If

        Dim sValue As String = String.Empty
        Dim sFormater As String = String.Empty

        Dim oSettingNode As Xml.XmlNode
        Dim oSettingElem As Xml.XmlElement
        Dim oSettingPairNode As Xml.XmlNode
        Dim oSettingPairElem As Xml.XmlElement

        Dim sQuery As String
        sQuery = "//SettingCategory[@name='" + sCategoryName + "']/Setting[@name='" + sSettingName + "']"

        oSettingNode = GetXmlDoc().SelectSingleNode(sQuery)

        If oSettingNode IsNot Nothing Then
            If TypeOf oSettingNode Is Xml.XmlElement Then
                oSettingElem = CType(oSettingNode, Xml.XmlElement)

                'get name value
                oSettingPairNode = oSettingElem.SelectSingleNode("SettingPair[@name='name']")
                If TypeOf oSettingPairNode Is Xml.XmlElement Then
                    oSettingPairElem = CType(oSettingPairNode, Xml.XmlElement)
                    sValue = oSettingPairElem.GetAttribute("value")
                End If

                'get format value
                oSettingPairNode = oSettingElem.SelectSingleNode("SettingPair[@name='capitalization']")
                If TypeOf oSettingPairNode Is Xml.XmlElement Then
                    oSettingPairElem = CType(oSettingPairNode, Xml.XmlElement)
                    sFormater = oSettingPairElem.GetAttribute("value")
                End If
            End If
        End If

        GetSettingValue = capitalize(sValue, sFormater)

    End Function

    Private Function capitalize(ByVal content As String, ByVal capitalization As String) As String
        capitalization.Trim()
        content.Trim()

        If content.Length = 0 Then
            Return String.Empty
        End If

        Dim capitalizedContent As String

        If String.Compare(capitalization, CapitalizationLowercase, StringComparison.InvariantCultureIgnoreCase) = 0 Then
            capitalizedContent = content.ToLower()
        ElseIf String.Compare(capitalization, CapitalizationUppercase, StringComparison.InvariantCultureIgnoreCase) = 0 Then
            capitalizedContent = content.ToUpper()
        ElseIf String.Compare(capitalization, CapitalizationSentenceCase, StringComparison.InvariantCultureIgnoreCase) = 0 Then
            capitalizedContent = content.Substring(0, 1).ToUpper() & content.Substring(1).ToLower()
        ElseIf String.Compare(capitalization, CapitalizationTitleCase, StringComparison.InvariantCultureIgnoreCase) = 0 Then
            capitalizedContent = CurrentThread.CurrentCulture.TextInfo.ToTitleCase(content)
        Else  'capitalizationDefault
            capitalizedContent = content
        End If

        Return capitalizedContent
    End Function
End Class
