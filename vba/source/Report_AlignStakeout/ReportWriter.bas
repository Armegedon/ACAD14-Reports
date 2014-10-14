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
    Set ts = fs.CreateTextFile(rptName, ForWriting, True)
    
    str = "<HTML title='Stakeout Alignment Report'><Body>"
    str = str & "<H1>Stakeout Alignment Report</H1>"
    ts.WriteLine (str)
    
    
    ' write owner and client setting
    WriteUserSetting ts
    
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
    'angle type
    str = str & "Stakeout Angle Type: " & GetStakeAngleTypeStr(g_angleType) & "<br/>"
    'Get occupied and backsight pt info
    str = str & "Occupied Pt: " & "<b>Northing </b>" & g_oOccupiedPt.Northing & "," & "<b>Easting </b>" & g_oOccupiedPt.Easting & "<br/>"
    If Not g_angleType = Direction Then
        str = str & "BackSight Pt: " & "<b>Northing </b>" & g_oBacksightPt.Northing & "," & "<b>Easting </b>" & g_oBacksightPt.Easting & "<br/>"
    End If
    'increment
    str = str & "Station interval: " & g_staInc & "<br/>"
    ' offset
    str = str & "OffSet: " & g_offSet & "<br/>"
    ts.Write (str)

    str = "<br><table border='1' width='100%' id='table1'>"
    str = str & "<b><tr><td>" _
          & g_tlbHeadArr(0) & "</td> <td>" _
          & g_tlbHeadArr(1) & "</td> <td>" _
          & g_tlbHeadArr(2) & "</td> <td>" _
          & g_tlbHeadArr(3) & "</td> <td>" _
          & g_tlbHeadArr(4) & "</td> </tr></b>"

    str = str & "<HR/>"

    ts.Write (str)

    'extract data
    Call ExtractData(oAlignment, stationStart, stationEnd)

    For i = 0 To UBound(g_oAlignDataArr)
        'Dim a As String
        'a = TypeName(g_oAlignDataArr(i))
        'Dim sTable As String
        'sTable = str
        If Not TypeName(g_oAlignDataArr(i)) = "Variant()" Then
            Exit For
        End If
        'format string
        Dim j
        Dim arrStr(nLastIndex)
        For j = 0 To nLastIndex
            arrStr(j) = CStr(VBA.Format(g_oAlignDataArr(i)(j), "0.00"))
        Next j

        ts.Write ("<b><tr> <td>" _
                    & arrStr(nStationIndex) & "</td> <td>" _
                    & arrStr(nDirectionIndex) & "</td> <td>" _
                    & arrStr(nDistance) & "</td> <td>" _
                    & arrStr(nNorthingIndex) & "</td> <td>" _
                    & arrStr(nEastingIndex) _
                    & "</td></tr></b>")

    Next
    ts.Write ("</Table><HR/>")

    ts.Close

    Set ts = Nothing
    Set fs = Nothing
    Set rf = Nothing
    
    AppendReport = True
End Function

Private Sub GetTableHead()
    ReDim g_tlbHeadArr(4)
    g_tlbHeadArr(0) = "Station"
    Select Case g_angleType
        Case TurnedPlus
            g_tlbHeadArr(1) = "Turned.Right"
        Case TurnedMinus
            g_tlbHeadArr(1) = "Turned.Left"
        Case DeflectPlus
            g_tlbHeadArr(1) = "Defl.Right"
        Case DeflectMinus
            g_tlbHeadArr(1) = "Defl.Left"
        Case Else     'Direction
            g_tlbHeadArr(1) = "Direction"
        End Select
    g_tlbHeadArr(2) = "Distance"
    g_tlbHeadArr(3) = "Coordinate.N"
    g_tlbHeadArr(4) = "Coordinate.E"
End Sub

Private Function GetStakeAngleTypeStr(aType As StakeAngleType) As String
    If aType = DeflectMinus Then
        GetStakeAngleTypeStr = "DeflectMinus"
    ElseIf aType = DeflectPlus Then
        GetStakeAngleTypeStr = "DeflectPlus"
    ElseIf aType = Direction Then
        GetStakeAngleTypeStr = "Direction"
    ElseIf aType = TurnedMinus Then
        GetStakeAngleTypeStr = "TurnedMinus"
    Else
        GetStakeAngleTypeStr = "TurnedPlus "
    End If
End Function
