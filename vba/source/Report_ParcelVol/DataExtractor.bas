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

Public Const nFillIndex = 0
Public Const nCutIndex = 1
Public Const nNetIndex = 2
Public Const nLastIndex = nNetIndex




Public g_oVolumeDataArr(nLastIndex) As String

Function FormatVolumeSettings(ByVal dVolume As Double) As String
    Dim oVolumeSettings As AeccSettingsVolume
    Set oVolumeSettings = g_oAeccDb.settings.SurfaceSettings.AmbientSettings.VolumeSettings
        
    FormatVolumeSettings = FormatVolume(dVolume, oVolumeSettings.Unit, oVolumeSettings.Precision, oVolumeSettings.Rounding, oVolumeSettings.Sign)
End Function

Public Function ExtractData(oParcel As AeccParcel, oVolumeSurface As AeccSurface, filCor As Double, cutCor As Double) As Boolean
    Dim dCut As Double
    Dim dFill As Double
    Dim dNet As Double
    
    Dim arrBoundary() As Variant
    Dim oVolumeSurfaceStatics As Object
'   Dim oVolumeSurfaceStatics As AeccTinVolumeSurfaceStatistics
    'Dim oSurface As AeccTinVolumeSurface
    Dim oSurface As Object
    
    arrBoundary = GetBoundaryPtArr(oParcel.ParcelLoops(0))
    CleanBoundaryPtArr arrBoundary
    Set oSurface = oVolumeSurface
    Set oVolumeSurfaceStatics = oSurface.Statistics
    Dim newArr() As Double
    newArr = ResetPointArray(arrBoundary)
    oVolumeSurfaceStatics.BoundedVolumes newArr, dCut, dFill, dNet
    g_oVolumeDataArr(nFillIndex) = FormatVolumeSettings(dFill * filCor)
    g_oVolumeDataArr(nCutIndex) = FormatVolumeSettings(dCut * cutCor)
    g_oVolumeDataArr(nNetIndex) = FormatVolumeSettings(dFill * filCor - dCut * cutCor)
    ExtractData = True
End Function

Public Function GetBoundaryPtArr(oLoop As AeccParcelLoop) As Variant
    Dim bounds() As Variant
    Dim point(2) As Double

    Dim oSegment As AeccParcelSegmentElement
    Dim lPtCount As Long
    lPtCount = 0#
    For Each oSegment In oLoop
        If TypeOf oSegment Is AeccParcelSegmentCurve Then
            Dim arcPoints() As Variant
            Dim arcPtCounts As Long
            Dim iIter As Long
            arcPtCounts = GetArcSamplePoint(arcPoints, oSegment)
            lPtCount = lPtCount + arcPtCounts
            ReDim Preserve bounds(lPtCount - 1)
            For iIter = 1 To arcPtCounts Step 1
                bounds(lPtCount - arcPtCounts + iIter - 1) = arcPoints(iIter - 1)
            Next iIter
        Else
            point(0) = oSegment.startX
            point(1) = oSegment.startY
            point(2) = 0#
            lPtCount = lPtCount + 1
            ReDim Preserve bounds(lPtCount - 1)
            bounds(lPtCount - 1) = point
        End If
          
    Next
    point(0) = oLoop.Item(0).startX
    point(1) = oLoop.Item(0).startY
    point(2) = 0#
    lPtCount = lPtCount + 1
    ReDim Preserve bounds(lPtCount - 1)
    bounds(lPtCount - 1) = point
    GetBoundaryPtArr = bounds
    
    
End Function

' clean the duplicate line in the boundary
Private Sub CleanBoundaryPtArr(arrBoundary() As Variant)
    If UBound(arrBoundary) < 2 Then
        Exit Sub
    End If
    
    
    Dim i As Integer
    Dim pt1, pt2 As Variant
    For i = 0 To UBound(arrBoundary) - 2 Step 1
        pt1 = arrBoundary(i)        ' pt2's location is pt1+2
        pt2 = arrBoundary(i + 2)
        If pt1(0) = pt2(0) And pt1(1) = pt2(1) Then '
            GoTo Rebuild1
        End If
    Next
    pt1 = arrBoundary(UBound(arrBoundary) - 1) ' pt2 is the last 2nd, and pt1 is the 2nd in array.
    pt2 = arrBoundary(1)
    If pt1(0) = pt2(0) And pt1(1) = pt2(1) Then
        GoTo Rebuild2
    End If
    Exit Sub
Rebuild1:
    ReDim bounds(UBound(arrBoundary) - 2) As Variant
    Dim k As Integer
    For k = 0 To UBound(bounds) Step 1
        If k <= i Then
            bounds(k) = arrBoundary(k)
        Else
            bounds(k) = arrBoundary(k + 2)
        End If
    Next
    arrBoundary = bounds
    CleanBoundaryPtArr arrBoundary 'recursion
    Exit Sub
    
Rebuild2:
    ReDim bounds(UBound(arrBoundary) - 2) As Variant
    For k = 0 To UBound(bounds) Step 1
        bounds(k) = bounds(k + 1)
    Next
    CleanBoundaryPtArr arrBoundary 'recursion
    Exit Sub
    
End Sub

' return the count of sample Point
Private Function GetArcSamplePoint(oPoints() As Variant, oArc As AeccParcelSegmentCurve) As Long
    Dim ptCenter As New CPoint2D
    Dim ptStart As New CPoint2D
    Dim ptEnd As New CPoint2D
    Dim chordVector As New CVector2D
    Dim interval As Double
    
    ' initial value from curve segment
    ptStart.m_X = oArc.startX
    ptStart.m_Y = oArc.startY
    ptEnd.m_X = oArc.endX
    ptEnd.m_Y = oArc.EndY
    chordVector.m_X = ptEnd.m_X - ptStart.m_X
    chordVector.m_Y = ptEnd.m_Y - ptStart.m_Y
    
    ' get arc angle
    Dim rotAng As Double
    Dim halfAng As Double
    rotAng = oArc.length / oArc.Radius
    If oArc.Bulge > 0 Then
        halfAng = rotAng / 2
        interval = 0.2
    Else
        halfAng = -rotAng / 2
       interval = -0.2
    End If
    
    ' get the center of the arc
    Dim ptTemp As New CPoint2D
    Dim vecTemp As New CVector2D
    ptTemp.m_X = ptStart.m_X
    ptTemp.m_Y = ptStart.m_Y
    vecTemp.m_X = chordVector.m_X
    vecTemp.m_Y = chordVector.m_Y
    If oArc.Bulge > 0 Then
        Call vecTemp.Rotate(2 * Atn(1) - halfAng)
    Else
        Call vecTemp.Rotate(-2 * Atn(1) - halfAng)
    End If
    
    vecTemp.SetLength oArc.Radius
    ptTemp.PlusVector vecTemp
    ptCenter.m_X = ptTemp.m_X
    ptCenter.m_Y = ptTemp.m_Y
    
    vecTemp.ScaleV -1#
        
    ' get sample points
    Dim count As Long
    Dim pt(2) As Double
    Dim tmpAngle As Double
    Dim samplePt As New CPoint2D
    Dim vectSample As New CVector2D
    count = 1
    ReDim oPoints(0)
    pt(0) = oArc.startX
    pt(1) = oArc.startY
    pt(2) = 0#
    oPoints(0) = pt
    For tmpAngle = interval To 2 * halfAng Step interval
        vectSample.m_X = vecTemp.m_X
        vectSample.m_Y = vecTemp.m_Y
        vectSample.Rotate tmpAngle
        samplePt.m_X = ptCenter.m_X
        samplePt.m_Y = ptCenter.m_Y
        samplePt.PlusVector vectSample
        pt(0) = samplePt.m_X
        pt(1) = samplePt.m_Y
        pt(2) = 0#
        count = count + 1
        ReDim Preserve oPoints(count - 1)
        oPoints(count - 1) = pt
    Next tmpAngle
   
    GetArcSamplePoint = count
End Function

Public Function GetUnitStr() As String
    Dim settings
    Dim unitsStr As String
    
    Set settings = g_oAeccDb.settings.DrawingSettings.AmbientSettings.DistanceSettings.Unit
    If settings = AeccCoordinateUnitType.aeccCoordinateUnitFoot Then
        unitsStr = "Feet"
    ElseIf settings = AeccCoordinateUnitType.aeccCoordinateUnitMeter Then
        unitsStr = "Meters"
    End If
    GetUnitStr = unitsStr
End Function

Public Function ResetPointArray(oPoints() As Variant) As Double()
    Dim l As Long, u As Long, count As Long
    Dim newArr() As Double
    l = LBound(oPoints)
    u = UBound(oPoints)
    count = u - l + 1
    ReDim newArr(count * 3 - 1)
    Dim i As Long
    For i = l To u Step 1
        newArr(3 * (i - l)) = oPoints(i)(0)
        newArr(3 * (i - l) + 1) = oPoints(i)(1)
        newArr(3 * (i - l) + 2) = oPoints(i)(2)
    Next
    
    
    ResetPointArray = newArr

End Function
