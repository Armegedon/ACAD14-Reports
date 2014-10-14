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

Private g_tlbHeadArr() As String

Option Explicit
Function WriteReportHeader(rptName As String) As Boolean
    Dim str As String
    Dim unitsStr As String
    Dim fs
    Dim ts As TextStream
    
    'unitsStr = GetUnitStr
 
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.CreateTextFile(rptName, ForWriting, True)
    
    
    str = "<HTML title='Alignment PI Station Report'><Body>"
    ts.WriteLine (str)
    ts.WriteLine ("<Body>")
    ts.WriteLine ("<center>")
    
    str = "<H1>Alignment PI Station Report</H1>"
    ts.WriteLine (str)
    ts.WriteLine ("</center>")
    
      
    
    ' write owner and client setting
    WriteUserSetting ts
    
    ' data info
    str = "Date: " & DateTime.Month(DateTime.Date) & "-" & DateTime.Day(DateTime.Date) & "-" & DateTime.Year(DateTime.Date) & "<br/>"
    ts.WriteLine (str)
    
    ts.WriteLine ("<HR/>")
    ts.Close

   
    Set fs = Nothing
    Set ts = Nothing
    
    GetTableHead
    
    WriteReportHeader = True
End Function

' Write the end of html file
Function WriteReportFooter(rptName As String) As Boolean
    Dim str As String
    Dim fs, rf, ts
    
  
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set rf = fs.GetFile(rptName)
    Set ts = rf.OpenAsTextStream(ForAppending, TristateUseDefault)

    str = "</Body></HTML>"
    
    ts.WriteLine (str)
    ts.Close

    Set ts = Nothing
    Set fs = Nothing
    Set rf = Nothing

    WriteReportFooter = True
End Function

' Append one item's report to html file
Function AppendReport(oAlignment As AeccAlignment, stationStart As Double, stationEnd As Double, rptName As String) As Boolean
    Dim str As String
    Dim fs, rf, ts
    Dim i As Integer
   
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set rf = fs.GetFile(rptName)
    Set ts = rf.OpenAsTextStream(ForAppending, TristateUseDefault)
    
    ' alignment info
    str = str & "Alignment Name: " & oAlignment.Name & "<br/>"
    str = str & "Description: " & oAlignment.Description & "<br/>"
    str = str & "Station Range: " & "Start: " & oAlignment.GetStationStringWithEquations(stationStart) _
                & ",  End: " & oAlignment.GetStationStringWithEquations(stationEnd) & "<br/>"
    
    ts.Write (str)
    
    str = "<br><table border='1' width='100%' id='table1'>"
    str = str & "<b><tr><td>" _
          & g_tlbHeadArr(0) & "</td> <td>" _
          & g_tlbHeadArr(1) & "</td> <td>" _
          & g_tlbHeadArr(2) & "</td> <td>" _
          & g_tlbHeadArr(3) & "</td> <td>" _
          & g_tlbHeadArr(4) & "</td> </tr></b>"
          
    ts.Write (str)

   ' extract data
    Call ExtractData(oAlignment, stationStart, stationEnd)

    For i = 0 To UBound(g_oAlignDataArr)

        If Not TypeName(g_oAlignDataArr(i)) = "Variant()" Then
            Exit For
        End If
        'format string
       
        If i > 0 Then
            ts.Write ("<b><tr><td>" _
                        & "&nbsp" & "</td><td>" _
                        & "&nbsp" & "</td><td>" _
                        & "&nbsp" & "</td><td>" _
                        & g_oAlignDataArr(i)(nLengthValueIndex) & "</td><td>" _
                        & g_oAlignDataArr(i)(nDirectionIndex) & "</td><tr></b>")
        End If
        ts.Write ("<b><tr><td>" _
                    & g_oAlignDataArr(i)(nStationIndex) & "</td><td>" _
                    & g_oAlignDataArr(i)(nNorthingIndex) & "</td><td>" _
                    & g_oAlignDataArr(i)(nEastingIndex) & "</td><td>" _
                    & "&nbsp" & "</td><td>" _
                    & "&nbsp" & "</td><tr></b>")

    Next
    ts.Write ("</Table>")
    ts.Close

    Set ts = Nothing
    Set fs = Nothing
    Set rf = Nothing
    
    AppendReport = True
End Function

Private Sub GetTableHead()
    ReDim g_tlbHeadArr(4)
    g_tlbHeadArr(0) = "PI Station"
    g_tlbHeadArr(1) = "Northing"
    g_tlbHeadArr(2) = "Easting"
    g_tlbHeadArr(3) = "Distance"
    g_tlbHeadArr(4) = "Direction"
    
End Sub
