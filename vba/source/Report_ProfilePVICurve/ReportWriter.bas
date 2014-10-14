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
 
    str = "<HTML title='PVI Station&Curve Report'><Body>"
    str = str & "<H1>Profile PVI Station & Curve Report</H1>"
   
    ts.WriteLine (str)
    
    'write user's setting
    WriteUserSetting ts
    str = "Date: " & DateTime.Month(DateTime.Date) & "-" & DateTime.Day(DateTime.Date) & "-" & DateTime.Year(DateTime.Date) & "<br/>"
 '   str = str & "Distance Units: " & unitsStr & "<br/><HR/>"
    
    

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
    Dim i As Integer


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

    If oProfile.Type = aeccExistingGround Then
        str = "<br><table border='1' width='100%' id='table1'>"
        str = str & "<b><tr><td>" _
              & g_tlbHeadArr(0) & "</td> <td>" _
              & g_tlbHeadArr(1) & "</td> <td>" _
              & g_tlbHeadArr(2) & "</td> </tr></b>"
    
        str = str & "<HR/>"
        ts.WriteLine (str)
        
        For i = 0 To UBound(g_oProfileDataArr)
            If Not TypeName(g_oProfileDataArr(i)) = "Variant()" Then
                Exit For
            End If
                    
                                  
            ts.WriteLine ("<b><tr> <td>" _
                        & CStr(i) & "</td> <td>" _
                        & g_oProfileDataArr(i)(nStationIndex) & "</td> <td>" _
                        & g_oProfileDataArr(i)(nGradeOutIndex) & "</td></tr></b>")

        Next i
    ' Finished Ground Profile
    ElseIf oProfile.Type = aeccFinishedGround Then
        Dim bCurve As Boolean
        str = "<br><table border='1' width='100%' id='table1'>"
        str = str & "<b><tr><td>" _
              & g_tlbHeadArr(0) & "</td> <td>" _
              & g_tlbHeadArr(1) & "</td> <td>" _
              & g_tlbHeadArr(2) & "</td> <td>" _
              & g_tlbHeadArr(3) & "</td> </tr></b>"
    
        str = str & "<HR/>"
        ts.WriteLine (str)
        
        For i = 0 To UBound(g_oProfileDataArr)
            If Not TypeName(g_oProfileDataArr(i)) = "Variant()" Then
                Exit For
            End If
                    
            If TypeName(g_oProfileDataArr(i)(nCurveInfo)) = "Variant()" Then
                bCurve = True
            Else
                bCurve = False
            End If
                   
                  
            
            Dim sCurveLen As String
            If bCurve Then
                sCurveLen = CStr(g_oProfileDataArr(i)(nCurveInfo)(nCurveLenIndex))
            Else
                sCurveLen = "&nbsp;"
            End If
            ts.WriteLine ("<b><tr> <td>" _
                        & CStr(i) & "</td> <td>" _
                        & g_oProfileDataArr(i)(nStationIndex) & "</td> <td>" _
                        & g_oProfileDataArr(i)(nGradeOutIndex) & "</td> <td>" _
                        & sCurveLen & "</td></tr></b>")

            If bCurve Then
                WriteCurveInfoStr i, ts
            End If
        Next i
    End If
    


    ts.WriteLine ("</Table><HR/>")
'    ts.WriteBlankLines (1)
'    ts.WriteBlankLines (1)
'    ts.WriteBlankLines (1)

    ts.Close

    Set ts = Nothing
    Set fs = Nothing
   
'
    AppendReport = True
End Function

Private Sub GetTableHead()
    ReDim g_tlbHeadArr(4)
    g_tlbHeadArr(0) = "PVI"
    g_tlbHeadArr(1) = "Station"
    g_tlbHeadArr(2) = "Grade Out (%)"
    g_tlbHeadArr(3) = "Curve Length"
End Sub


Private Sub WriteCurveInfoStr(nNumber As Integer, ts As TextStream)
'    Public Const nStationIndex = 0
'    Public Const nElevationIndex = 1
'    Public Const nGradeOutIndex = 2
'    Public Const nCurveInfo = 3
'    Public Const nLastIndex = nCurveInfo
'
'    Public Const nCurvTypeIndex = 0
'    Public Const nPVCStationIndex = 1
'    Public Const nPVCElevationIndex = 2
'    Public Const nPVTStationIndex = 3
'    Public Const nPVTElevationIndex = 4
'    Public Const nHighStationIndex = 5
'    Public Const nHightElevation = 6
'    Public Const nLowStationIndex = 7
'    Public Const nLowElevation = 8
'    Public Const nGradeInIndex = 9
'    Public Const nGradeChange = 10
'    Public Const nCurveLenIndex = 11
'    Public Const nKIndex = 12
'    Public Const nHeadlightDisIndex = 13
'    Public Const nPassingDisIndex = 13
'    Public Const nStoppingDisIndex = 14
'    Public Const nLastCurInfoIndex = nStoppingDisIndex

'<DIV style="PADDING-BOTTOM: 4pt; PADDING-TOP: 10pt">Vertical Curve Information: (sag curve)</DIV>
'<TABLE id=ID_CurveTable style="BORDER-TOP: black 1px dashed" cellSpacing=0>
'<TBODY>
'<TR>
'<TD style="PADDING-TOP: 6pt" align=left>PVC Station:</TD>
'<TD style="PADDING-RIGHT: 10pt; PADDING-LEFT: 10pt; PADDING-TOP: 6pt" align=right>0+020</TD>
'<TD style="PADDING-TOP: 6pt" align=left>Elevation:</TD>
'<TD style="PADDING-RIGHT: 10pt; PADDING-LEFT: 10pt; PADDING-TOP: 6pt" align=right>127.40</TD></TR>
'<TR>
'<TD align=left>PVI Station: </TD>
'<TD style="PADDING-RIGHT: 10pt; PADDING-LEFT: 10pt" align=right>0+040</TD>
'<TD align=left>Elevation: </TD>
'<TD style="PADDING-RIGHT: 10pt; PADDING-LEFT: 10pt" align=right>127.05</TD></TR>
'<TR>
'<TD align=left>PVT Station:</TD>
'<TD style="PADDING-RIGHT: 10pt; PADDING-LEFT: 10pt" align=right>0+060</TD>
'<TD align=left>Elevation:</TD>
'<TD style="PADDING-RIGHT: 10pt; PADDING-LEFT: 10pt" align=right>126.91</TD></TR>
'<TR>
'<TD style="PADDING-TOP: 6pt" align=left>Grade in (%): </TD>
'<TD style="PADDING-RIGHT: 10pt; PADDING-LEFT: 10pt; PADDING-TOP: 6pt" align=right>-1.750</TD>
'<TD style="PADDING-TOP: 6pt" align=left>Grade out (%): </TD>
'<TD style="PADDING-RIGHT: 10pt; PADDING-LEFT: 10pt; PADDING-TOP: 6pt" align=right>-0.698</TD></TR>
'<TR>
'<TD align=left>Change (%): </TD>
'<TD style="PADDING-RIGHT: 10pt; PADDING-LEFT: 10pt" align=right>1.052</TD>
'<TD align=left>K: </TD>
'<TD style="PADDING-RIGHT: 10pt; PADDING-LEFT: 10pt" align=right>38.033</TD></TR>
'<TR>
'<TD style="PADDING-TOP: 6pt" align=left>Curve Length: </TD>
'<TD style="PADDING-RIGHT: 10pt; PADDING-LEFT: 10pt; PADDING-TOP: 6pt" align=right>40.000</TD>
'<TD style="PADDING-TOP: 6pt" align=left colSpan=2>&nbsp;</TD></TR>
'<TR>
'<TD align=left>Headlight Distance: </TD>
'<TD style="PADDING-RIGHT: 10pt; PADDING-LEFT: 10pt" align=right>Infinite</TD>
'<TD align=right colSpan=2>&nbsp;</TD></TR></TBODY></TABLE></TD></TR>
    On Error Resume Next
    Dim sType, sValue
    Dim str As String
    Dim oCurveDataArr As Variant
    oCurveDataArr = g_oProfileDataArr(nNumber)(nCurveInfo)
    sType = "<TD style=""PADDING-TOP: 6pt"" align=left>"
    sValue = "<TD style=""PADDING-RIGHT: 10pt; PADDING-LEFT: 10pt; PADDING-TOP: 6pt"" align=right>"
    
    ts.Write ("<TR>")
    ts.Write ("<TD>&nbsp;</TD>")
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
    WriteTypeValueRecord2 "PVI Station:", g_oProfileDataArr(nNumber)(nStationIndex), "Elevation:", g_oProfileDataArr(nNumber)(nElevationIndex), ts
    WriteTypeValueRecord2 "PVT Station:", oCurveDataArr(nPVTStationIndex), "Elevation:", oCurveDataArr(nPVTElevationIndex), ts
    
    If TypeName(oCurveDataArr(nHighStationIndex)) = "String" Then
        WriteTypeValueRecord2 "High Point:", oCurveDataArr(nHighStationIndex), "Elevation:", oCurveDataArr(nHightElevation), ts
    End If
    
    If TypeName(oCurveDataArr(nLowStationIndex)) = "String" Then
        WriteTypeValueRecord2 "Low Point:", oCurveDataArr(nLowStationIndex), "Elevation:", oCurveDataArr(nLowElevation), ts
    End If
    
    WriteTypeValueRecord2 "Grade in(%):", oCurveDataArr(nGradeInIndex), "Grade out(%):", g_oProfileDataArr(nNumber)(nGradeOutIndex), ts, True
    WriteTypeValueRecord2 "Change(%):", oCurveDataArr(nGradeChange), "K:", oCurveDataArr(nKIndex), ts
    
    WriteTypeValueRecord2 "Curve Length:", oCurveDataArr(nCurveLenIndex), "", "", ts, True
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

