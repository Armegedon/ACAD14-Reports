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
' Write the head of html file
Function WriteReportHeader(rptName As String) As Boolean
    Dim str As String
    Dim fs
    Dim ts As TextStream
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.CreateTextFile(rptName, ForWriting, True)
    
    str = "<HTML title='Parcel Map Check Report'><Body>"
    str = str & "<H1>Parcel Map Check Report</H1>"
    ts.WriteLine (str)
    
    
    ' write owner and client setting
    WriteUserSetting ts
    
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
Function AppendReport(oParcel As AeccParcel, bCounterClockWise As Boolean, bAcrossChord As Boolean, rptName As String) As Boolean
    Dim str As String
    Dim fs, rf
    Dim ts As TextStream
    Dim i As Integer
   
    If ExtractData(oParcel, bCounterClockWise, bAcrossChord) = False Then
        AppendReport = False
        Exit Function
    End If
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set rf = fs.GetFile(rptName)
    Set ts = rf.OpenAsTextStream(ForAppending, TristateUseDefault)
    
    ' Parcel info
    str = str & "Parcel Name: " & oParcel.Parent.Name & "-" & oParcel.Name & "<br/>"
    str = str & "Description: " & oParcel.Description & "<br/>"
    ts.Write (str)
    str = "Process segment order counterclockwise: "
    If bCounterClockWise = True Then
        str = str + "True"
    Else
        str = str + "False"
    End If
    str = str + "<br/>"
    ts.Write str
    str = "Enable mapcheck across chord: "
    If bAcrossChord = True Then
        str = str + "True"
    Else
        str = str + "False"
    End If
    str = str + "<br/>"
    ts.Write str
    
    'POB
    str = "<br><table border='0' width='500' id='table1'>"
    ts.Write (str)
    
    Call AppendOneRecordInTable(ts, _
                                "North: " & FormateCoordSettings_Parcel(g_ParcelCheckData.POB_North), _
                                "East: " & FormateCoordSettings_Parcel(g_ParcelCheckData.POB_East))
    BreakLine ts
    'Segments
    Call AppendSegments(ts)
    
    'Check Error
    Call AppendCheckErrorInfo(ts)
    

    
    ts.Write ("</Table>")
    ts.WriteLine ("<br/>" & "<HR/>")
    ts.Close

    Set ts = Nothing
    Set fs = Nothing
    Set rf = Nothing
    
    AppendReport = True
    AppendReport = True
End Function

Private Sub AppendSegments(ts As TextStream)
    Dim i As Long
    Dim segCount As Long
    Dim seg As Object
    Dim segLine As SegmentLine
    Dim segCurve As SegmentCurve
    segCount = g_ParcelCheckData.SegmentCount
    
    For i = 1 To segCount Step 1
        Set seg = g_ParcelCheckData.Item(i)
        If TypeOf seg Is SegmentLine Then
            Set segLine = seg
            Call AppendOneRecordInTable(ts, "Segment#" + CStr(i) + " : Line", "&nbsp")
            Call AppendOneRecordInTable(ts, _
                                        "Course: " + FormatDirSettings_Parcel(segLine.Course), _
                                        "Length: " + FormatDistSettings_Parcel(segLine.Length))
            Call AppendOneRecordInTable(ts, _
                                        "North: " & FormateCoordSettings_Parcel(segLine.End_North), _
                                        "East: " & FormateCoordSettings_Parcel(segLine.End_East))
            
        Else
            Set segCurve = seg
            Call AppendOneRecordInTable(ts, "Segment#" + CStr(i) + " : Curve", "&nbsp")
            Call AppendOneRecordInTable(ts, _
                                        "Length: " & FormatDistSettings_Parcel(segCurve.Length), _
                                        "Radius: " & FormatDistSettings_Parcel(segCurve.Radius))
            Call AppendOneRecordInTable(ts, _
                                        "Delta: " & FormatAngleSettings_Parcel(segCurve.Delta), _
                                        "Tangent: " & FormatDistSettings_Parcel(segCurve.Tangent))
            Call AppendOneRecordInTable(ts, _
                                        "Chord: " & FormatDistSettings_Parcel(segCurve.Chord), _
                                        "Course: " & FormatDirSettings_Parcel(segCurve.Course))
            Call AppendOneRecordInTable(ts, _
                                        "Course In: " & FormatDirSettings_Parcel(segCurve.CourseIn), _
                                        "Course Out: " & FormatDirSettings_Parcel(segCurve.CourseOut))
            Call AppendOneRecordInTable(ts, _
                                        "RP North: " & FormateCoordSettings_Parcel(segCurve.RP_North), _
                                        "East: " & FormateCoordSettings_Parcel(segCurve.RP_East))
            Call AppendOneRecordInTable(ts, _
                                        "End North: " & FormateCoordSettings_Parcel(segCurve.End_North), _
                                        "East: " & FormateCoordSettings_Parcel(segCurve.End_East))
        End If
    
        BreakLine ts
    Next
   
    
End Sub

Private Sub AppendCheckErrorInfo(ts As TextStream)
    Dim ErrorClosure As Double
    Dim ErrorCourse As Double
    Dim ErrorNorth As Double, ErrorEast As Double
    Dim ErrorPrecision As Double
    
    g_ParcelCheckData.GetErrorInfo ErrorClosure, ErrorCourse, ErrorNorth, ErrorEast, ErrorPrecision
    
    
    Call AppendOneRecordInTable(ts, _
                                "Perimeter: " & FormatDistSettings_Parcel(g_ParcelCheckData.Perimeter), _
                                "Area: " & FormatAreaSettings_Parcel(g_ParcelCheckData.Area))
    'get Closure
    Dim coordPrecision As Integer
    Dim North As Double
    Dim East As Double
    coordPrecision = g_oAeccDb.Settings.ParcelSettings.AmbientSettings.CoordinateSettings.precision.Value
    North = RoundVal(ErrorNorth, coordPrecision + 1)
    East = RoundVal(ErrorEast, coordPrecision + 1)
    Call AppendOneRecordInTable(ts, _
                                "Error Closure: " & CStr(ErrorClosure), _
                                "Course: " & FormatDirSettings_Parcel(ErrorCourse))
    Call AppendOneRecordInTable(ts, _
                                "Error North : " & CStr(North), _
                                "East: " & CStr(East))
      
    BreakLine ts
    'Precision
    Dim distPrecision As Integer
    Dim prec As Double
    distPrecision = g_oAeccDb.Settings.ParcelSettings.AmbientSettings.DistanceSettings.precision.Value
    prec = RoundVal(ErrorPrecision, distPrecision)
    Call AppendOneRecordInTable(ts, _
                                "Precision " & "1:" & CStr(prec), _
                                "&nbsp")
End Sub

Private Sub AppendOneRecordInTable(ts As TextStream, str1 As String, str2 As String)
    Dim str As String
    str = "<b><tr><td>" _
          & str1 & "</td> <td>" _
          & str2 & "</td> </tr></b>"
    ts.Write (str)
End Sub

Private Sub BreakLine(ts As TextStream)
    AppendOneRecordInTable ts, "&nbsp", "&nbsp"
End Sub

