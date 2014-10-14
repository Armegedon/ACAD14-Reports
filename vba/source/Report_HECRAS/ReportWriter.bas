Attribute VB_Name = "ReportWriter"
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

Public g_oDataToReport As New DataSteamNetwork

'This commented block was for html before. May be useful if reports need to be port to .net

'Function WriteReportHeader(rptName As String) As Boolean
'    Dim str As String
'    Dim unitsStr As String
'    Dim fs
'    Dim ts As TextStream
'
'    'unitsStr = GetUnitStr
'
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set ts = fs.CreateTextFile(rptName, True)
'
'
'    str = "<HTML title='HEC-RAS Report'><Body>"
'    ts.WriteLine (str)
'    ts.WriteLine ("<Body>")
'    ts.WriteLine ("<center>")
'
'    str = "<H1>HEC-RAS Report</H1>"
'    ts.WriteLine (str)
'    ts.WriteLine ("</center>")
'
'
'
'    ' write owner and client setting
'    WriteUserSetting ts
'
'    ' data info
'    str = "Date: " & DateTime.Month(DateTime.Date) & "-" & DateTime.Day(DateTime.Date) & "-" & DateTime.Year(DateTime.Date) & "<br/>"
'    ts.WriteLine (str)
'
'    ts.WriteLine ("<HR/>")
'    ts.Close
'
'
'    Set fs = Nothing
'    Set ts = Nothing
'
'
'    WriteReportHeader = True
'End Function
'
'' Write the end of html file
'Function WriteReportFooter(rptName As String) As Boolean
'    Dim str As String
'    Dim fs, ts
'
'
'    Set fs = CreateObject("Scripting.FileSystemObject")
'
'    Set ts = fs.OpenTextFile(rptName, ForAppending)
'
'    str = "</Body></HTML>"
'
'    ts.WriteLine (str)
'    ts.Close
'
'    Set ts = Nothing
'    Set fs = Nothing
'
'
'    WriteReportFooter = True
'End Function
''
'
'
'' Append one item's report to html file
'Function AppendReport(oSteam As CSTEAM, rptName As String) As Boolean
'    On Error Resume Next
'    Dim str As String
'    Dim fs As FileSystemObject
'    Dim ts As TextStream
'    Dim sTable As String
'
'    Dim i As Long
'
''
''
'    Call ExtractData(oSteam)
'
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set ts = fs.OpenTextFile(rptName, ForAppending)
'
'
'    'Write HEADER
'    ts.WriteLine ("<b>BEGIN HEADER:</b>")
'    WriteBlankLine 1, ts
'    sTable = "<br><table border='1' width='60%' frame='void' rules='none' id='table1' cellspacing='2'>"
'    ts.WriteLine sTable
'
'    Dim oHeader As DataHeader
'    Set oHeader = g_oDataToReport.header
'    WriteHeaderRecord "UNITS:", oHeader.Units, ts
'    WriteHeaderRecord "DTM TYPE:", oHeader.DTMType, ts
'    WriteHeaderRecord "DTM:", oHeader.DTM, ts
'    WriteHeaderRecord "NUMBER OF REACHES:", oHeader.NumberOfReaches, ts
'    WriteHeaderRecord "NUMBER OF CROSS-SECTIONS:", oHeader.NumberOfCrossSection, ts
'    WriteHeaderRecord "MAP PROJECTION:", oHeader.MapProjection, ts
'    WriteHeaderRecord "DATUM:", oHeader.Datum, ts
'    ts.WriteLine ("</Table>")
'    WriteBlankLine 1, ts
'    ts.WriteLine ("<b>END HEADER:</b><br>")
'    WriteBlankLine 2, ts
'
'    ' Write Stream NetWork
'    ts.WriteLine ("<b>BEGIN STREAM NETWORK:</b>")
'    WriteBlankLine 1, ts
'    ts.WriteLine sTable
'    Dim oEndPts As Collection
'    Set oEndPts = g_oDataToReport.EndPoints
'    For i = 1 To oEndPts.Count Step 1
'        Dim arrEndPt As Variant
'        arrEndPt = oEndPts(i)
'        WriteStreamRecord "ENDPOINT:", "", arrEndPt(nEasting) + ",", arrEndPt(nNorthing) + ",", arrEndPt(nElevation) + ",", arrEndPt(nID), ts
'    Next
'    Dim oReaches As Collection
'    Set oReaches = g_oDataToReport.Reaches
'    For i = 1 To oReaches.Count Step 1
'        Dim oReach As DataReach
'        Set oReach = oReaches.item(i)
'        WriteStreamRecord "", "", "", "", "", "", ts
'        WriteStreamRecord "REACH:", "", "", "", "", "", ts
'        WriteStreamRecord "", "STEAM ID:", oReach.SteamID, "", "", "", ts
'        WriteStreamRecord "", "REACH ID:", oReach.ReachID, "", "", "", ts
'        WriteStreamRecord "", "FROM POINT:", CStr(oReach.FromPoint), "", "", "", ts
'        WriteStreamRecord "", "TO POINT:", CStr(oReach.ToPoint), "", "", "", ts
'        WriteStreamRecord "", "CENTERLINE:", "", "", "", "", ts
'        Dim j As Long
'        For j = 1 To oReach.EndPoints.Count Step 1
'            arrEndPt = oReach.EndPoints(j)
'            WriteStreamRecord "", "", arrEndPt(nEasting) + ",", arrEndPt(nNorthing) + ",", arrEndPt(nElevation) + ",", arrEndPt(nStation), ts
'        Next
'        WriteStreamRecord "", "END:", "", "", "", "", ts
'    Next
'    ts.WriteLine ("</Table>")
'    WriteBlankLine 1, ts
'    ts.WriteLine ("<b>END STREAM NETWORK:</b><br>")
'    WriteBlankLine 2, ts
'
'    'Write Cross Section
'    ts.WriteLine ("<b>BEGIN CROSS SECTIONS:</b>")
'    ts.WriteLine sTable
'    Dim oSections As Collection
'    Dim oSection As DataSection
'    Set oSections = g_oDataToReport.CrossSections
'    For i = 1 To oSections.Count Step 1
'        Set oSection = oSections.item(i)
'        WriteStreamRecord "", "", "", "", "", "", ts
'        WriteStreamRecord "CROSS SECTION:", "", "", "", "", "", ts
'        WriteStreamRecord "", "STREAM ID:", oSection.SteamID, "", "", "", ts
'        WriteStreamRecord "", "REACH ID:", oSection.ReachID, "", "", "", ts
'        WriteStreamRecord "", "STATION:", oSection.Station, "", "", "", ts
'        WriteStreamRecord "", "CUT LINE:", "", "", "", "", ts
'        For j = 1 To oSection.CutLines.Count Step 1
'            Dim arrCutLine As Variant
'            arrCutLine = oSection.CutLines(j)
'            WriteStreamRecord "", "", arrCutLine(nEasting) + ",", arrCutLine(nNorthing), "", "", ts
'        Next j
'    Next i
'    ts.WriteLine ("</Table>")
'    WriteBlankLine 1, ts
'    ts.WriteLine ("<b>END CROSS SECTIONS:</b><br>")
'    WriteBlankLine 1, ts
'    str = DateTime.Month(DateTime.Date) & "-" & DateTime.Day(DateTime.Date) & "-" & DateTime.Year(DateTime.Date)
'    str = str & "," & DateTime.Time$
'
'    ts.WriteLine ("<b>" + "FILE COMPLETE: " + str + "</b><br>")
'    ts.Close
'
'    Set ts = Nothing
'    Set fs = Nothing
'
'
'    AppendReport = True
'
'
'End Function
''
'
'Private Function TableDataStr(ByVal sVal As String, Optional ByVal bBold As Boolean = False) As String
'    If sVal = "" Then
'        sVal = "&nbsp"
'    End If
'    If bBold = True Then
'        TableDataStr = "<td><b>" & sVal & "</b></td>"
'    Else
'        TableDataStr = "<td>" & sVal & "</td>"
'    End If
'End Function
'
'Private Function TableRecordStr(ByVal sVal As String) As String
'    TableRecordStr = "<tr>" & sVal & "</tr>"
'End Function
'
'Private Sub WriteHeaderRecord(ByVal sType As String, ByVal sValue As String, ts As TextStream)
'    Dim str As String
'    str = "<td width=20>&nbsp</td>" + TableDataStr(sType, True) + TableDataStr(sValue)
'    ts.WriteLine (TableRecordStr(str))
'End Sub
'
'Private Sub WriteStreamRecord(ByVal str1 As String, ByVal str2 As String, ByVal str3 As String, ByVal str4 As String, ByVal str5 As String, ByVal str6 As String, ts As TextStream)
'    Dim str As String
'    str = "<td width=20>&nbsp</td>" + TableDataStr(str1, True) + TableDataStr(str2, True) + TableDataStr(str3) + TableDataStr(str4) + TableDataStr(str5) + TableDataStr(str6)
'    ts.WriteLine (TableRecordStr(str))
'End Sub
'
'Private Sub WriteBlankLine(ByVal num As Long, ts As TextStream)
'    Dim i As Long
'    For i = 1 To num Step 1
'        ts.WriteLine "<br/>"
'    Next i
'
'End Sub



Function WriteReportHeader(rptName As String) As Boolean
    Dim str As String
    Dim unitsStr As String
    Dim fs
    Dim ts As TextStream

    'unitsStr = GetUnitStr

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.CreateTextFile(rptName, True)


    str = "HEC-RAS Report"
    ts.WriteLine (str)
  
    ' data info
    str = "Date: " & DateTime.Month(DateTime.Date) & "-" & DateTime.Day(DateTime.Date) & "-" & DateTime.Year(DateTime.Date)
    ts.WriteLine (str)
    
    ts.Close


    Set fs = Nothing
    Set ts = Nothing


    WriteReportHeader = True
End Function

' Write the end of html file
Function WriteReportFooter(rptName As String) As Boolean
    Dim str As String
    Dim fs
    Dim ts As TextStream


    Set fs = CreateObject("Scripting.FileSystemObject")

    Set ts = fs.OpenTextFile(rptName, ForAppending)

    str = DateTime.Month(DateTime.Date) & "-" & DateTime.Day(DateTime.Date) & "-" & DateTime.Year(DateTime.Date)
    str = str & "," & DateTime.Time$
    
    ts.WriteLine ("FILE COMPLETE: " + str)
    
    ts.Close

    Set ts = Nothing
    Set fs = Nothing


    WriteReportFooter = True
End Function
'


' Append one item's report to html file
Function AppendReport(oSteam As CSTEAM, rptName As String) As Boolean
    On Error Resume Next
    Dim str As String
    Dim fs As FileSystemObject
    Dim ts As TextStream
    Dim sTable As String

    Dim i As Long

'
'
    Call ExtractData(oSteam)
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.OpenTextFile(rptName, ForAppending)
    
    
    'Write HEADER
    WriteBlankLine 2, ts
    ts.WriteLine ("BEGIN HEADER:")
    WriteBlankLine 1, ts
    Dim indent1 As Integer, indent2 As Integer, indent3 As Integer
    indent1 = 7
    indent2 = indent1 + 6
    indent3 = indent2 + 6
    Dim oHeader As DataHeader
    Set oHeader = g_oDataToReport.header
    WriteLineWithSpaces indent1, "UNITS: " + oHeader.Units, ts
    WriteLineWithSpaces indent1, "DTM TYPE: " + oHeader.DTMType, ts
    WriteLineWithSpaces indent1, "DTM: " + oHeader.DTM, ts
    WriteLineWithSpaces indent1, "NUMBER OF REACHES: " + CStr(oHeader.NumberOfReaches), ts
    WriteLineWithSpaces indent1, "NUMBER OF CROSS-SECTIONS: " + CStr(oHeader.NumberOfCrossSection), ts
    WriteLineWithSpaces indent1, "MAP PROJECTION: " + oHeader.MapProjection, ts
    WriteLineWithSpaces indent1, "DATUM: " + oHeader.Datum, ts
   
    WriteBlankLine 1, ts
    ts.WriteLine ("END HEADER:")
    WriteBlankLine 2, ts
    
    ' Write Stream NetWork
    ts.WriteLine ("BEGIN STREAM NETWORK:")
    WriteBlankLine 1, ts
    
    Dim oEndPts As Collection
    Set oEndPts = g_oDataToReport.EndPoints
    For i = 1 To oEndPts.Count Step 1
        Dim arrEndPt As Variant
        arrEndPt = oEndPts(i)
        WriteLineWithSpaces indent1, "ENDPOINT: " + arrEndPt(nEasting) + "," + arrEndPt(nNorthing) + "," + arrEndPt(nElevation) + "," + arrEndPt(nID), ts
    Next
    
    Dim oReaches As Collection
    Set oReaches = g_oDataToReport.Reaches
    For i = 1 To oReaches.Count Step 1
        Dim oReach As DataReach
        Set oReach = oReaches.item(i)
        WriteBlankLine 1, ts
        WriteLineWithSpaces indent1, "REACH:", ts
        WriteLineWithSpaces indent2, "STEAM ID: " + oReach.SteamID, ts
        WriteLineWithSpaces indent2, "REACH ID: " + oReach.ReachID, ts
        WriteLineWithSpaces indent2, "FROM POINT: " + CStr(oReach.FromPoint), ts
        WriteLineWithSpaces indent2, "TO POINT: " + CStr(oReach.ToPoint), ts
        WriteLineWithSpaces indent2, "CENTERLINE:", ts
        Dim j As Long
        For j = 1 To oReach.EndPoints.Count Step 1
            arrEndPt = oReach.EndPoints(j)
            WriteLineWithSpaces indent3, arrEndPt(nEasting) + "," + arrEndPt(nNorthing) + "," + arrEndPt(nElevation) + "," + arrEndPt(nStation), ts
        Next
        WriteLineWithSpaces indent2, "End:", ts
    Next
   
    WriteBlankLine 1, ts
    ts.WriteLine ("END STREAM NETWORK:")
    WriteBlankLine 2, ts
    
    'Write Cross Section
    ts.WriteLine ("BEGIN CROSS SECTIONS:")
    Dim oSections As Collection
    Dim oSection As DataSection
    Set oSections = g_oDataToReport.CrossSections
    For i = 1 To oSections.Count Step 1
        Set oSection = oSections.item(i)
        WriteBlankLine 1, ts
        WriteLineWithSpaces indent1, "CROSS SECTION:", ts
        WriteLineWithSpaces indent2, "STREAM ID: " + oSection.SteamID, ts
        WriteLineWithSpaces indent2, "REACH ID: " + oSection.ReachID, ts
        WriteLineWithSpaces indent2, "STATION: " + oSection.Station, ts
        WriteLineWithSpaces indent2, "CUT LINE:", ts
        For j = 1 To oSection.CutLines.Count Step 1
            Dim arrCutLine As Variant
            arrCutLine = oSection.CutLines(j)
            WriteLineWithSpaces indent3, arrCutLine(nEasting) + "," + arrCutLine(nNorthing), ts
        Next j
    Next i
    
    WriteBlankLine 1, ts
    ts.WriteLine ("END CROSS SECTIONS:")
    WriteBlankLine 1, ts

    ts.Close

    Set ts = Nothing
    Set fs = Nothing
 
    
    AppendReport = True
    

End Function
'





Private Sub WriteBlankLine(ByVal num As Long, ts As TextStream)
    Dim i As Long
    For i = 1 To num Step 1
        ts.WriteLine ""
    Next i
    
End Sub
Private Sub WriteLineWithSpaces(ByVal numberOfSpaces As Integer, ByVal str As String, ts As TextStream)
    Dim i As Integer
    For i = 0 To numberOfSpaces Step 1
        str = " " + str
    Next i
    ts.WriteLine str
End Sub

Private Function GetSpaces(num As Integer) As String
    Dim i As Integer
    For i = 1 To num Step 1
        GetSpaces = GetSpaces + " "
    Next i
End Function
