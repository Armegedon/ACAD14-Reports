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
' Write the head of html file
Function WriteReportHeader(rptName As String) As Boolean
    Dim str As String
    Dim unitsStr As String
    Dim fs
    Dim ts As TextStream
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.CreateTextFile(rptName, True)
 
    str = "<HTML title='Vertical Curve Report'><Body>"
    str = str & "<center><H1>Profile Vertical Curve Report</H1></center>"
    ts.WriteLine (str)
    
    'write user's setting
    WriteUserSetting ts
    
    str = "Date: " & DateTime.Month(DateTime.Date) & "-" & DateTime.Day(DateTime.Date) & "-" & DateTime.Year(DateTime.Date) & "<br/>"
        
    ts.WriteLine (str)
    ts.WriteLine ("<HR/>")
    
    ts.Close
   
    Set fs = Nothing
    Set ts = Nothing
    
   ' GetTableHead
    
    WriteReportHeader = True
End Function

' Write the end of html file
Function WriteReportFooter(rptName As String) As Boolean
    Dim str As String
    Dim fs, rf, ts
    
  
    Set fs = CreateObject("Scripting.FileSystemObject")
  '  Set rf = fs.GetFile(rptName)
    Set ts = fs.OpenTextFile(rptName, ForAppending)

    str = "</Body></HTML>"
    
    ts.WriteLine (str)
    ts.Close

    Set ts = Nothing
    Set fs = Nothing
    Set rf = Nothing

    WriteReportFooter = True
End Function

' Append one item's report to html file
Function AppendReport(oProfile As AeccProfile, stationStart As Double, stationEnd As Double, rptName As String) As Boolean
    Dim str As String
    Dim fs
    Dim ts As TextStream
    Dim i As Long


    Set fs = CreateObject("Scripting.FileSystemObject")
 '   Set rf = fs.GetFile(rptName)
    Set ts = fs.OpenTextFile(rptName, ForAppending)

  
    str = "Vertical Alignment: " & oProfile.Name & "<br/>"
    str = str & "Description: " & oProfile.Description & "<br/>"
    str = str & "Station Range: " & "Start: " & oProfile.Alignment.GetStationStringWithEquations(stationStart) _
                & ",  End: " & oProfile.Alignment.GetStationStringWithEquations(stationEnd) & "<br/>"
    ts.WriteLine (str)
'
 

    'extract data
    Call ExtractData(oProfile, stationStart, stationEnd)
'
    str = "<TABLE id=ID_ProfileTable style=""MARGIN-TOP: 4px"" cellSpacing=0 cellPadding=3 border=1>"
    ts.WriteLine (str)
    ts.WriteLine ("<TBODY>")
    For i = 0 To UBound(g_oProfileDataArr)
        If Not TypeName(g_oProfileDataArr(i)) = "Variant()" Then
            Exit For
        End If
        
        WriteCurveInfoStr i, ts
        
    Next i


    ts.WriteLine ("</TBODY></Table><HR/>")

    ts.Close

    Set ts = Nothing
    Set fs = Nothing
   
'
    AppendReport = True
End Function

Private Sub WriteCurveInfoStr(nNumber As Long, ts As TextStream)
    On Error Resume Next
 
    Dim str As String
    Dim oCurveDataArr As Variant
    oCurveDataArr = g_oProfileDataArr(nNumber)
    
    ts.Write ("<TR>")
    ts.Write ("<TD colSpan=4>")
    str = "<DIV style=""PADDING-BOTTOM: 4pt; PADDING-TOP: 10pt"" >Vertical Curve Information:  "
    If oCurveDataArr(nCurvTypeIndex) = aeccSag Then
        str = str & "(sag curve)" & "</DIV>"
    Else
        str = str & "(crest curve)" & "</DIV>"
    End If
    ts.WriteLine (str)
    str = "<TABLE id=ID_CurveTable style=""BORDER-TOP: black 1px dashed"" cellSpacing=0>"
    ts.WriteLine (str)
    ts.WriteLine ("<TBODY>")
        
    'station
    WriteTypeValueRecord2 "PVC Station:", oCurveDataArr(nPVCStationIndex), "Elevation:", oCurveDataArr(nPVCElevationIndex), ts, True
    WriteTypeValueRecord2 "PVI Station:", oCurveDataArr(nPVIStationIndex), "Elevation:", oCurveDataArr(nPVIElevationIndex), ts
    WriteTypeValueRecord2 "PVT Station:", oCurveDataArr(nPVTStationIndex), "Elevation:", oCurveDataArr(nPVTElevationIndex), ts
    
    If TypeName(oCurveDataArr(nHighStationIndex)) = "String" Then
        WriteTypeValueRecord2 "High Point:", oCurveDataArr(nHighStationIndex), "Elevation:", oCurveDataArr(nHightElevation), ts
    End If
    
    If TypeName(oCurveDataArr(nLowStationIndex)) = "String" Then
        WriteTypeValueRecord2 "Low Point:", oCurveDataArr(nLowStationIndex), "Elevation:", oCurveDataArr(nLowElevation), ts
    End If
    
    WriteTypeValueRecord2 "Grade in:", oCurveDataArr(nGradeInIndex), "Grade out:", oCurveDataArr(nGradeOutIndex), ts, True
    WriteTypeValueRecord2 "Change:", oCurveDataArr(nGradeChange), "K:", oCurveDataArr(nKIndex), ts
    
    WriteTypeValueRecord2 "Curve Length:", oCurveDataArr(nCurveLenIndex), "Curve Radius", oCurveDataArr(nCurveRadiusIndex), ts, True
    If oCurveDataArr(nCurvTypeIndex) = aeccSag Then
        WriteTypeValueRecord2 "Headlight Distance:", oCurveDataArr(nHeadlightDisIndex), "", "", ts
    Else
        WriteTypeValueRecord2 "Passing Distance:", oCurveDataArr(nPassingDisIndex), "Stopping Distance:", oCurveDataArr(nStoppingDisIndex), ts, True
    End If
    
    str = "</TBODY></TABLE></TD></TR>"
    ts.WriteLine (str)
   
  
    
End Sub
Sub WriteTypeValueRecord2(sType1 As Variant, sValue1 As Variant, sType2 As Variant, sValue2 As Variant, ts As TextStream, Optional bInterval As Boolean = False)
    Dim sTypePrefix, sValuePrefix
    Dim str
    If bInterval = True Then
        sTypePrefix = "<TD style=""PADDING-TOP: 6pt"" align=left>"
    Else
        sTypePrefix = "<TD align=left>"
    End If
    sValuePrefix = "<TD style=""PADDING-RIGHT: 10pt; PADDING-LEFT: 10pt; PADDING-TOP: 6pt"" align=right>"
    
    ts.WriteLine ("<TR>")
    
    str = sTypePrefix & sType1 & "</TD>"
    ts.WriteLine (str)
    str = sValuePrefix & sValue1 & "</TD>"
    ts.WriteLine (str)
    str = sTypePrefix & sType2 & "</TD>"
    ts.WriteLine (str)
    str = sValuePrefix & sValue2 & "</TD>"
    ts.WriteLine (str)
    
    ts.WriteLine ("</TR>")
End Sub

