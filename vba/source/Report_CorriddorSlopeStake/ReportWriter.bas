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
Function WriteReportHeader(rptName As String) As Boolean
    Dim str As String
    Dim unitsStr As String
    Dim fs
    Dim ts As TextStream
    
    'unitsStr = GetUnitStr
 
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.CreateTextFile(rptName, True)
    
    
    str = "<HTML title='Corridor Slope Stake Report'><Body>"
    ts.WriteLine (str)
    ts.WriteLine ("<Body>")
    ts.WriteLine ("<center>")
    
    str = "<H1>Corridor Slope Stake Report</H1>"
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
    
    
    WriteReportHeader = True
End Function

' Write the end of html file
Function WriteReportFooter(rptName As String) As Boolean
    Dim str As String
    Dim fs, ts
    
  
    Set fs = CreateObject("Scripting.FileSystemObject")
 
    Set ts = fs.OpenTextFile(rptName, ForAppending)
        
    str = "</Body></HTML>"
    
    ts.WriteLine (str)
    ts.Close

    Set ts = Nothing
    Set fs = Nothing
  

    WriteReportFooter = True
End Function

' Append one item's report to html file
Function AppendReport(oCorridor As AeccCorridor, oBaseline As AeccBaseline, oSampleLineGroup As AeccSampleLineGroup, stationStart As Double, stationEnd As Double, rptName As String, sCodeName As String) As Boolean
    Dim str As String
    Dim fs As FileSystemObject
    Dim ts As TextStream
    Dim oBaseAlignment As AeccAlignment
    Dim i As Integer
    Dim Stations As Variant
    Dim curStation As Double
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Set ts = fs.OpenTextFile(rptName, ForAppending)
    
    Set oBaseAlignment = oCorridor.Baselines(0).Alignment
    ' Corridor info
    str = "Corridor Name: " & oCorridor.Name & "<br/>"
    str = str & "Description: " & oCorridor.Description & "<br/>"
    str = str & "Base Alignment Name: " & oBaseAlignment.Name & "<br/>"
    str = str & "Sample Line Group Name: " & oSampleLineGroup.Name & "<br/>"
    str = str & "Link Code Name: " & sCodeName & "<br/>"
    str = str & "Station Range: " & "Start: " & oBaseAlignment.GetStationStringWithEquations(stationStart) _
                & ",  End: " & oBaseAlignment.GetStationStringWithEquations(stationEnd) & "<br/>"

    ts.WriteLine (str)
    
 
'    extract data
    Call ExtractData(oCorridor, oBaseline, oSampleLineGroup, stationStart, stationEnd, sCodeName)
    
    Stations = g_oSlopeStakeData.Keys
    For i = 0 To g_oSlopeStakeData.count - 1
        curStation = Stations(i)
        AppendSingleTable oCorridor, curStation, ts
    Next i
     
    ts.WriteLine ("<HR/>")
    ts.Close

    Set ts = Nothing
    Set fs = Nothing
    
    AppendReport = True
End Function

Private Sub AppendSingleTable(oCorridor As AeccCorridor, dStation As Double, ts As TextStream)
    Dim str As String
    Dim arrDataBySta As Variant
    Dim cLeftDatas As Dictionary
    Dim cRightDatas As Dictionary
    Dim arrLeftEndInfo As Variant
    Dim arrRightEndInfo As Variant
    
    Dim i As Long
    
    If Not g_oSlopeStakeData.Exists(dStation) Then
        Exit Sub
    End If
    
    arrDataBySta = g_oSlopeStakeData.item(dStation)
    Set cLeftDatas = arrDataBySta(nLeftIndex)
    Set cRightDatas = arrDataBySta(nRightIndex)
    arrLeftEndInfo = arrDataBySta(nLeftEndIndex)
    arrRightEndInfo = arrDataBySta(nRightEndIndex)
    
    ts.WriteLine ("<center>")
    str = oCorridor.Baselines(0).Alignment.Name & "<br/>"
    str = str & "Station: " & oCorridor.Baselines(0).Alignment.GetStationStringWithEquations(dStation)
    ts.WriteLine (str)
    
    ts.WriteLine ("</center>")
    str = "<br><table border='1' width='100%' id='table1'>"
    ts.WriteLine (str)
    
    i = 1
    
    Do
        Dim iRow As Long '1 based
        For iRow = 1 To 4
            Dim arrRowData(0 To 12) As String
            Erase arrRowData
            'leftend and rightend info
            If i = 1 And iRow < 4 Then
                arrRowData(0) = arrLeftEndInfo(iRow - 1)
                arrRowData(12) = arrRightEndInfo(iRow - 1)

            End If
            
            Dim iColum As Integer
            For iColum = 1 To 5
                Dim arrLeftData As Variant
                Dim arrRightData As Variant
                Dim nCurColIndex As Long '1 based
                'left side
                nCurColIndex = (i - 1) * 5 + iColum
                If nCurColIndex <= cLeftDatas.count Then
                
                    arrLeftData = cLeftDatas.Items(nCurColIndex - 1)
                    arrRowData(6 - iColum) = arrLeftData(iRow - 1)

                End If
                'right side
                If nCurColIndex <= cRightDatas.count Then
                    arrRightData = cRightDatas.Items(nCurColIndex - 1)
                    arrRowData(6 + iColum) = arrRightData(iRow - 1)

                End If
                
            Next iColum
            WriteOneTableRow arrRowData, ts
        Next iRow
        'seperate row
        Dim arrBlank(0 To 12) As String
        WriteOneTableRow arrBlank, ts
        i = i + 1
    Loop Until i * 5 >= (cLeftDatas.count + 5) Or i * 5 >= (cRightDatas.count + 5)
    
    ts.WriteLine ("</Table>")
    
End Sub

Private Sub WriteOneTableRow(arrData As Variant, ts As TextStream)
    Dim i As Long
    Dim str As String
    str = "<b><tr>"
    For i = LBound(arrData) To UBound(arrData)
        If arrData(i) = "" Then
            str = str & "<td>" & "&nbsp;" & "</td>"
        Else
            str = str & "<td>" & arrData(i) & "</td>"
        End If
    Next i
    str = str & "</tr></b>"
    ts.WriteLine (str)
    
End Sub
