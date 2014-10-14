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
    Dim fs
    Dim ts As TextStream
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.CreateTextFile(rptName, ForWriting, True)
    
    str = "<HTML title='Parcel Volume Report'><Body>"
    str = str & "<H1>Parcel Volume Report</H1>"
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
Function AppendReport(oParcel As AeccParcel, oVolumeSurface As AeccSurface, filCor As Double, cutCor As Double, eleTol As Double, rptName As String) As Boolean
    
    Dim str As String
    Dim fs, rf, ts
    Dim i As Integer


    Set fs = CreateObject("Scripting.FileSystemObject")
    Set rf = fs.GetFile(rptName)
    Set ts = rf.OpenAsTextStream(ForAppending, TristateUseDefault)

    ' alignment info
    str = str & "Parcel Name: " & oParcel.Parent.Name + " - " + oParcel.Name & "<br/>"
    str = str & "Description: " & oParcel.Description & "<br/>"
    str = str & "Volume Surface: " & oVolumeSurface.Name & "<br/>"
    str = str & "Fill correction: " & filCor & "<br/>"
    str = str & "Cut correction: " & cutCor & "<br/>"
    str = str & "Elevation Tolerance: " & eleTol & "<br/>"
    ts.Write (str)

    str = "<br><table border='1' width='100%' id='table1'>"
    str = str & "<b><tr><td>" _
          & g_tlbHeadArr(nFillIndex) & "</td> <td>" _
          & g_tlbHeadArr(nCutIndex) & "</td> <td>" _
          & g_tlbHeadArr(nNetIndex) & "</td> </tr></b>"

    str = str & "<HR/>"

    ts.Write (str)

    'extract data
    Call ExtractData(oParcel, oVolumeSurface, filCor, cutCor)

    ts.Write ("<b><tr> <td>" _
                & g_oVolumeDataArr(nFillIndex) & "</td> <td>" _
                & g_oVolumeDataArr(nCutIndex) & "</td> <td>" _
                & g_oVolumeDataArr(nNetIndex) _
                & "</td></tr></b>")

   
    ts.Write ("</Table><HR/>")

    ts.Close

    Set ts = Nothing
    Set fs = Nothing
    Set rf = Nothing
    
    AppendReport = True
End Function

Private Sub GetTableHead()
    ReDim g_tlbHeadArr(2)
    g_tlbHeadArr(nFillIndex) = "Fill Volume"
    g_tlbHeadArr(nCutIndex) = "Cut Volume"
    g_tlbHeadArr(nNetIndex) = "Net Volume"
End Sub


