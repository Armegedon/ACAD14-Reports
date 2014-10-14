Attribute VB_Name = "UserSetting"
Public Const CountOfParams = 24

Public settingsPair(1 To CountOfParams, 1 To 2) As String

Private Function FormatString(str As String, formmater As String) As String

    Dim temp As String
    temp = ""
    formatter = Trim(formatter)
    If StrComp(formmater, "default", vbTextCompare) = 0 Then
        temp = str
    ElseIf StrComp(formmater, "Lowercase", vbTextCompare) = 0 Then
        temp = LCase(str)
    ElseIf StrComp(formmater, "Uppercase", vbTextCompare) = 0 Then
        temp = UCase(str)
    ElseIf StrComp(formmater, "Sentence Case", vbTextCompare) = 0 Then
        Dim str1 As String
        str1 = Trim(str)
        temp = UCase(Left(str1, 1)) + Right(str1, Len(str1) - 1)
    ElseIf StrComp(formmater, "Title Case", vbTextCompare) = 0 Then
        temp = str
    End If
    
    FormatString = temp

End Function

Function GetReportSettingsParam(key As String) As String
    For i = 1 To CountOfParams
        If StrComp(key, settingsPair(i, 1), vbTextCompare) = 0 Then
            GetReportSettingsParam = settingsPair(i, 2)
        End If
    Next
End Function

Private Sub BuildList(objRootElem As MSXML2.IXMLDOMElement, settingCategoryName As String, start As Integer)
    Dim queryString As String
    queryString = "//SettingCategory[@name='" + settingCategoryName + "']/Setting"
    
     Dim objSettingCategoryList As MSXML2.IXMLDOMNodeList
    Set objSettingCategoryList = objRootElem.SelectNodes(queryString)
    
    Dim objSetting As MSXML2.IXMLDOMElement
    
    i = 0
    For Each objSetting In objSettingCategoryList
        settingsPair(start + i, 1) = settingCategoryName + "." + objSetting.getAttribute("name")
        
        Dim objSettingPairList As MSXML2.IXMLDOMNodeList
        Dim querySettingPairString As String
        querySettingPairString = "./SettingPair[@name='capitalization']"
        Set objSettingPairList = objSetting.SelectNodes(querySettingPairString)
        
        Dim strValue As String
        
        Dim valueAttr As MSXML2.IXMLDOMAttribute
        Set valueAttr = objSetting.FirstChild.Attributes.getNamedItem("value")
        
        strValue = valueAttr.NodeValue
            
        If objSettingPairList.Length <> 0 Then
            Set valueAttr = objSettingPairList.Item(0).Attributes.getNamedItem("value")
            strValue = FormatString(strValue, valueAttr.NodeValue)
        End If
                
        settingsPair(start + i, 2) = strValue
        
        i = i + 1
    
    Next

End Sub

Sub LoadReportSettingsParams(strFileName As String)
    
    Dim objDoc As New MSXML2.DOMDocument
    objDoc.async = False
    objDoc.resolveExternals = False
    
    
    objDoc.Load (strFileName)
    
    Dim objRootElem As MSXML2.IXMLDOMElement
    
    Set objRootElem = objDoc.documentElement
    
    BuildList objRootElem, "Client", 1
    BuildList objRootElem, "Owner", 13
    
    Set objDoc = Nothing

End Sub


Public Sub WriteUserSetting(ts As TextStream)
    On Error GoTo ErrHandle
    Dim str As String
    Dim curSetting As Integer
    Dim sKeyPath As String
    Dim sKeyName As String
    Dim sSettingFilePath As String
    Dim pos As Long
       sKeyPath = "Software\Autodesk\LandXML Reporting\7\Settings"
    sKeyName = "SettingsFile"

    sSettingFilePath = GetSettingString(HKEY_CURRENT_USER, sKeyPath, sKeyName)
    LoadReportSettingsParams sSettingFilePath
    str = "<table border=""0"" width=""100%"">"
    ts.WriteLine (str)
    
    'add setting
    ts.WriteLine ("<tr>")
    str = "<td align=""left""><b>" + "Client:" + "</b></td>"
    ts.WriteLine (str)
    str = "<td align=""left""><b>" + "Prepared by:" + "</b></td>"
    ts.WriteLine (str)
    ts.WriteLine ("</tr>")
    
    For curSetting = 1 To 3
        ts.WriteLine ("<tr>")
        str = "<td align=""left"">" + settingsPair(curSetting, 2) + "</td>"
        ts.WriteLine (str)
        str = "<td align=""left"">" + settingsPair(curSetting + 12, 2) + "</td>"
        ts.WriteLine (str)
        ts.WriteLine ("</tr>")
    Next
    ts.WriteLine ("</table>")
    Exit Sub
ErrHandle:
    MsgBox "Read config file error."
End Sub
