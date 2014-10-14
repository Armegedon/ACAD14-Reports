Attribute VB_Name = "DataExtractor"
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


Public Const nEasting = 0
Public Const nNorthing = 1
Public Const nElevation = 2
Public Const nID = 3
Public Const nStation = 3

Public g_Alignments As New Dictionary


Function FormatCoordSettings(ByVal dDis As Double) As String
    Dim oCoordSettings As AeccSettingsCoordinate
    Set oCoordSettings = g_oAeccDb.Settings.AlignmentSettings.AmbientSettings.CoordinateSettings

    FormatCoordSettings = FormatDistance(dDis, oCoordSettings.Unit.Value, oCoordSettings.Precision.Value, _
                oCoordSettings.Rounding.Value, oCoordSettings.Sign.Value, False)
End Function
'
'
Public Function ExtractData(oSteam As CSTEAM) As Boolean
    On Error Resume Next
    g_oDataToReport.Clear
    

    
    'Get Stream EndPts
    Dim sAlignmentName As String
    Dim arrEndPt(3) As String
    Dim oAlignmentPIs As CAlignmentPIs
    Dim oAlignment As AeccAlignment
    Dim i As Long
    Dim dStation As Double
    Dim dNorthing As Double
    Dim dEasting As Double
    Dim dElevation As Double
    
    sAlignmentName = oSteam.SampleLineGroup.Parent.Name
    Set oAlignmentPIs = g_Alignments.item(sAlignmentName)
    Set oAlignment = oAlignmentPIs.Alignment
    For i = 1 To oAlignmentPIs.PICounts Step 1
        dStation = oAlignmentPIs.GetStationByNum(i)
        oAlignment.PointLocation dStation, 0#, dEasting, dNorthing
        arrEndPt(nEasting) = FormatCoordSettings(dEasting)
        arrEndPt(nNorthing) = FormatCoordSettings(dNorthing)
        Dim oSurface As AeccSurface
        Set oSurface = GetSampledSurface(oSteam.SampleLineGroup)
        arrEndPt(nElevation) = FormatCoordSettings(oSurface.FindElevationAtXY(dEasting, dNorthing))
        arrEndPt(nID) = CStr(i)
        g_oDataToReport.EndPoints.Add arrEndPt
    Next
    
    'Get Reaches
    GetReachesData oSteam
    
    'Get CrossSections
    GetSectionsData oSteam
    ExtractData = True
    
    ' Get Header Data
    GetHeaderData oSteam

End Function
'
Public Sub GetHeaderData(oSteam As CSTEAM)
    Dim oHeader As DataHeader
    Dim oSampleLineGroup As AeccSampleLineGroup
    Dim oSurface As AeccSurface
    Set oSampleLineGroup = oSteam.SampleLineGroup
    Set oHeader = g_oDataToReport.header
    
    'get Units:
    Dim disOriUnit As AeccDrawingUnitType
    disOriUnit = g_oAeccDb.Settings.DrawingSettings.UnitZoneSettings.DrawingUnits
    If disOriUnit = aeccDrawingUnitFeet Then
        oHeader.Units = "ENGLISH"
    Else
        oHeader.Units = "METRIC"
    End If
    
    'get DTM TYPE:
    Set oSurface = GetSampledSurface(oSteam.SampleLineGroup)
    oHeader.DTM = oSurface.Name
    Select Case oSurface.Type
    Case aecckGridSurface
        oHeader.DTMType = "GRID"
    Case aecckGridVolumeSurface
        oHeader.DTMType = "GRIDVOLUME"
    Case aecckTinSurface
        oHeader.DTMType = "TIN"
    Case Else
        oHeader.DTMType = "TINVOLUME"
    End Select
    
    'get Number of Reaches;
    oHeader.NumberOfReaches = g_oDataToReport.Reaches.Count
    
    'get Number of Sections
    oHeader.NumberOfCrossSection = g_oDataToReport.CrossSections.Count
    
    'get MAP Projections:
    oHeader.MapProjection = g_oAeccDb.Settings.DrawingSettings.UnitZoneSettings.CoordinateSystem.Projection
    
    'get Datum:
    oHeader.Datum = g_oAeccDb.Settings.DrawingSettings.UnitZoneSettings.CoordinateSystem.Datum
    
    
End Sub

Public Sub GetReachesData(oSteam As CSTEAM)
    Dim sAlignmentName As String
    Dim oAlignmentPIs As CAlignmentPIs
    Dim oAlignment As AeccAlignment
    Dim dicReaches As Dictionary
    Dim oReachData As DataReach
    Dim oReachRegion As CReachRegion
    Dim sReachName As String
    Dim i As Long
    
    sAlignmentName = oSteam.SampleLineGroup.Parent.Name
    Set oAlignmentPIs = g_Alignments.item(sAlignmentName)
    Set oAlignment = oAlignmentPIs.Alignment
    
    Set dicReaches = oSteam.mReaches
    For i = 0 To dicReaches.Count - 1 Step 1
        sReachName = dicReaches.Keys(i)
        Set oReachRegion = dicReaches.item(sReachName)
        Set oReachData = New DataReach
        oReachData.SteamID = oSteam.SampleLineGroup.Parent.Name
        oReachData.ReachID = sReachName
        oReachData.FromPoint = CStr(oReachRegion.BeginPt)
        oReachData.ToPoint = CStr(oReachRegion.EndPt)
        
        Dim j As Long
        Dim arrEndPt(3) As String
        Dim dStation As Double, dEasting As Double, dNorthing As Double
        Dim dStep As Long
        If oReachRegion.BeginPt < oReachRegion.EndPt Then
            dStep = 1
        Else
            dStep = -1
        End If
        For j = oReachRegion.BeginPt To oReachRegion.EndPt Step dStep
            dStation = oAlignmentPIs.GetStationByNum(j)
            oAlignment.PointLocation dStation, 0#, dEasting, dNorthing
            arrEndPt(nEasting) = FormatCoordSettings(dEasting)
            arrEndPt(nNorthing) = FormatCoordSettings(dNorthing)
            Dim oSurface As AeccSurface
            Set oSurface = GetSampledSurface(oSteam.SampleLineGroup)
            arrEndPt(nElevation) = FormatCoordSettings(oSurface.FindElevationAtXY(dEasting, dNorthing))
            arrEndPt(nStation) = oAlignment.GetStationStringWithEquations(dStation)
            oReachData.EndPoints.Add arrEndPt
        Next j
        g_oDataToReport.Reaches.Add oReachData
    Next i
End Sub

Public Sub GetSectionsData(oSteam As CSTEAM)
    Dim oSampleLine As AeccSampleLine
    Dim oReachRegion As CReachRegion
    Dim oAlignmentPIs As CAlignmentPIs
    Dim oDataSection As DataSection
    Dim sReachName As String
    Dim dBeginStation As Double
    Dim dEndStation As Double
    Dim i As Long
    Dim tmpCol As New Collection
    
    Set oAlignmentPIs = g_Alignments.item(oSteam.SampleLineGroup.Parent.Name)
    For i = 0 To oSteam.mReaches.Count - 1 Step 1
        sReachName = oSteam.mReaches.Keys(i)
        Set oReachRegion = oSteam.mReaches.item(sReachName)
        dBeginStation = oAlignmentPIs.GetStationByNum(oReachRegion.BeginPt)
        dEndStation = oAlignmentPIs.GetStationByNum(oReachRegion.EndPt)
            
        For Each oSampleLine In oSteam.SampleLineGroup.SampleLines
            Dim dStation As Double
            dStation = oSampleLine.Station
            If (dStation >= dBeginStation And dStation <= dEndStation) Or (dStation <= dBeginStation And dStation >= dEndStation) Then
                Set oDataSection = New DataSection
                oDataSection.SteamID = oSteam.SampleLineGroup.Parent.Name
                oDataSection.ReachID = sReachName
                oDataSection.Station = oSteam.SampleLineGroup.Parent.GetStationStringWithEquations(dStation)
                
                'sort the vertexs
                Dim vertexs As AeccSampleLineVertices
                Dim oVertex As AeccSampleLineVertex
                Set tmpCol = New Collection
                Dim k As Integer
                Dim j As Integer
                Dim tmpVer1 As AeccSampleLineVertex
                Dim tmpVer2 As AeccSampleLineVertex
                Set vertexs = oSampleLine.Vertices
                tmpCol.Add vertexs.item(0)
                For k = 1 To (vertexs.Count - 1)
                    Set tmpVer1 = vertexs.item(k)
                    For j = 1 To tmpCol.Count
                        Set tmpVer2 = tmpCol.item(j)
                        If tmpVer1.StationOffset < tmpVer2.StationOffset Then
                            Exit For
                        End If
                        
                    Next j
                    If j > tmpCol.Count Then
                        tmpCol.Add item:=tmpVer1, after:=tmpCol.Count
                    Else
                        tmpCol.Add item:=tmpVer1, before:=j
                    End If
                    
                Next k
                '
                ' Add to oDataSection.CutLines
                Dim location(1) As String
                
                For k = 0 To vertexs.CountLeft - 1
                    Set oVertex = tmpCol.item(k + 1)
                    location(nEasting) = FormatCoordSettings(oVertex.location(0))
                    location(nNorthing) = FormatCoordSettings(oVertex.location(1))
                    oDataSection.CutLines.Add location
                Next k
                
                Dim dEasting As Double
                Dim dNorthing As Double
                oSteam.SampleLineGroup.Parent.PointLocation dStation, 0#, dEasting, dNorthing
                location(nEasting) = FormatCoordSettings(dEasting)
                location(nNorthing) = FormatCoordSettings(dNorthing)
                oDataSection.CutLines.Add location
                
                For k = 0 To vertexs.CountRight - 1
                    Set oVertex = tmpCol.item(vertexs.CountLeft + k + 1)
                    location(nEasting) = FormatCoordSettings(oVertex.location(0))
                    location(nNorthing) = FormatCoordSettings(oVertex.location(1))
                    oDataSection.CutLines.Add location
                Next k
                
                
                g_oDataToReport.CrossSections.Add oDataSection
            End If
        Next
        
    Next i
End Sub

Public Function GetSampledSurface(oSampleLineGroup As AeccSampleLineGroup) As AeccSurface
    Set GetSampledSurface = Nothing
    Dim oSampledSurface As AeccSampledSurface
    Dim i As Long
    Dim bFound As Boolean
    bFound = False
    i = 0
    Do While (bFound = False And i < oSampleLineGroup.SampledSurfaces.Count)
        Dim oSurface As AeccSurface
        Set oSurface = oSampleLineGroup.SampledSurfaces.item(i).Surface
        If oSurface.Type = aecckTinSurface Then
            Set GetSampledSurface = oSurface
            bFound = True
        End If
        i = i + 1
    Loop
End Function
