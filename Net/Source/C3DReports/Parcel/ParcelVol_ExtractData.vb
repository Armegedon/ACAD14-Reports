' -----------------------------------------------------------------------------
' <copyright file="ReportForm_ParcelVol.vb" company="Autodesk">
' Copyright (C) Autodesk, Inc. All rights reserved.
'
' Permission to use, copy, modify, and distribute this software in
' object code form for any purpose and without fee is hereby granted,
' provided that the above copyright notice appears in all copies and
' that both that copyright notice and the limited warranty and
' restricted rights notice below appear in all supporting
' documentation.
'
' AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
' AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
' MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
' DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
' UNINTERRUPTED OR ERROR FREE.
'
' Use, duplication, or disclosure by the U.S. Government is subject to
' restrictions set forth in FAR 52.227-19 (Commercial Computer
' Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
' (Rights in Technical Data and Computer Software), as applicable.
'
' </copyright>
' -----------------------------------------------------------------------------

Option Explicit On
Option Strict On

Imports System
Imports AeccLandLib = Autodesk.AECC.Interop.Land

Friend Class ParcelVol_ExtractData

    Public Class VolumeData
        Public Fill As String
        Public Cut As String
        Public Net As String
    End Class

    Private Shared _volumeData As New VolumeData

    Public Shared ReadOnly Property ParcelVolumeData() As VolumeData
        Get
            Return _volumeData
        End Get
    End Property


    Private Shared Function FormatVolumeSettings(ByVal dVolume As Double) As String
        Dim oVolumeSettings As AeccLandLib.AeccSettingsVolume
        oVolumeSettings = ReportApplication.AeccXDatabase.Settings.SurfaceSettings.AmbientSettings.VolumeSettings

        FormatVolumeSettings = ReportFormat.FormatVolume(dVolume, oVolumeSettings.Unit.Value, _
                    oVolumeSettings.Precision.Value, oVolumeSettings.Rounding.Value, oVolumeSettings.Sign.Value)
    End Function

    Public Shared Function ExtractData(ByVal oParcel As AeccLandLib.AeccParcel, _
                                ByVal oVolumeSurface As AeccLandLib.AeccSurface, _
                                ByVal filCor As Double, ByVal cutCor As Double) As Boolean
        Dim dCut As New Double
        Dim dFill As New Double
        Dim dNet As New Double

        Dim pts As Collections.Generic.List(Of Autodesk.AutoCAD.Geometry.Point3d)

        pts = GetBoundaryPts(oParcel.ParcelLoops(0))
        CleanBoundaryPts(pts)

        Dim varPoints() As Double
        varPoints = ResetPointArray(pts)

        If oVolumeSurface.Type = AeccLandLib.AeccSurfaceType.aecckGridVolumeSurface Then
            Dim gridSurface As AeccLandLib.AeccGridVolumeSurface
            gridSurface = CType(oVolumeSurface, AeccLandLib.AeccGridVolumeSurface)
            gridSurface.Statistics.BoundedVolumes(varPoints, dCut, dFill, dNet)
        ElseIf oVolumeSurface.Type = AeccLandLib.AeccSurfaceType.aecckTinVolumeSurface Then
            Dim TinSurface As AeccLandLib.AeccTinVolumeSurface
            TinSurface = CType(oVolumeSurface, AeccLandLib.AeccTinVolumeSurface)
            TinSurface.Statistics.BoundedVolumes(varPoints, dCut, dFill, dNet)
        Else
            System.Diagnostics.Debug.Assert(False, "Wrong surface type, must be volume surface")
        End If

        _volumeData.Fill = FormatVolumeSettings(dFill * filCor)
        _volumeData.Cut = FormatVolumeSettings(dCut * cutCor)
        _volumeData.Net = FormatVolumeSettings(dFill * filCor - dCut * cutCor)
        ExtractData = True

    End Function

    Private Shared Function GetBoundaryPts(ByVal oLoop As AeccLandLib.AeccParcelLoop) _
        As Collections.Generic.List(Of Autodesk.AutoCAD.Geometry.Point3d)

        Dim pts As New Collections.Generic.List(Of Autodesk.AutoCAD.Geometry.Point3d)

        For Each oSegment As AeccLandLib.AeccParcelSegmentElement In oLoop
            If TypeOf oSegment Is AeccLandLib.AeccParcelSegmentCurve Then
                Dim arcPts As Collections.Generic.List(Of Autodesk.AutoCAD.Geometry.Point3d)
                arcPts = GetArcSamplePoint(CType(oSegment, AeccLandLib.AeccParcelSegmentCurve))
                pts.AddRange(arcPts)
            Else
                pts.Add(New Autodesk.AutoCAD.Geometry.Point3d(oSegment.StartX, oSegment.StartY, 0.0))
            End If

        Next
        pts.Add(New Autodesk.AutoCAD.Geometry.Point3d(oLoop.Item(0).StartX, oLoop.Item(0).StartY, 0.0))

        Return pts

    End Function

    ' clean the duplicate line in the boundary
    Private Shared Sub CleanBoundaryPts(ByRef pts As Collections.Generic.List(Of Autodesk.AutoCAD.Geometry.Point3d))

        'at least have 4 points for a parcel loop
        Diagnostics.Debug.Assert(pts.Count >= 4)

        'if points count less then 5, it's impossible with duplicated lines
        If pts.Count <= 5 Then
            Exit Sub
        End If

        Dim i As Integer = 0
        While i < pts.Count
            Dim pt0, pt2 As Autodesk.AutoCAD.Geometry.Point3d
            Dim pt0Index, pt1Index, pt2Index As Integer
            pt0Index = i
            pt1Index = (i + 1) Mod (pts.Count - 1)
            pt2Index = (i + 2) Mod (pts.Count - 1)
            pt0 = pts.Item(pt0Index)
            pt2 = pts.Item(pt2Index)
            If pt0.X = pt2.X And pt0.Y = pt2.Y Then
                'remove back point first, otherwise the index will be wrong after removed a point
                If pt1Index < pt2Index Then
                    pts.RemoveAt(pt2Index)
                    pts.RemoveAt(pt1Index)
                Else
                    pts.RemoveAt(pt1Index)
                    pts.RemoveAt(pt2Index)
                End If
            End If
            i += 1
        End While
    End Sub

    ' return the count of sample Point
    Private Shared Function GetArcSamplePoint(ByVal oArc As AeccLandLib.AeccParcelSegmentCurve) _
        As Collections.Generic.List(Of Autodesk.AutoCAD.Geometry.Point3d)

        Dim ptCenter As New CPoint2D
        Dim ptStart As New CPoint2D
        Dim ptEnd As New CPoint2D
        Dim chordVector As New CVector2D
        Dim interval As Double

        ' initial value from curve segment
        ptStart.m_X = oArc.StartX
        ptStart.m_Y = oArc.StartY
        ptEnd.m_X = oArc.EndX
        ptEnd.m_Y = oArc.EndY
        chordVector.m_X = ptEnd.m_X - ptStart.m_X
        chordVector.m_Y = ptEnd.m_Y - ptStart.m_Y

        ' get arc angle
        Dim rotAng As Double
        Dim halfAng As Double
        rotAng = oArc.Length / oArc.Radius
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
            Call vecTemp.Rotate(2 * Math.Atan(1) - halfAng)
        Else
            Call vecTemp.Rotate(-2 * Math.Atan(1) - halfAng)
        End If

        vecTemp.SetLength(oArc.Radius)
        ptTemp.PlusVector(vecTemp)
        ptCenter.m_X = ptTemp.m_X
        ptCenter.m_Y = ptTemp.m_Y

        vecTemp.ScaleV(-1.0)

        ' get sample points
        Dim tmpAngle As Double
        Dim samplePt As New CPoint2D
        Dim vectSample As New CVector2D

        Dim pts As New Collections.Generic.List(Of Autodesk.AutoCAD.Geometry.Point3d)
        pts.Add(New Autodesk.AutoCAD.Geometry.Point3d(oArc.StartX, oArc.StartY, 0.0))

        For tmpAngle = interval To 2 * halfAng Step interval
            vectSample.m_X = vecTemp.m_X
            vectSample.m_Y = vecTemp.m_Y
            vectSample.Rotate(tmpAngle)
            samplePt.m_X = ptCenter.m_X
            samplePt.m_Y = ptCenter.m_Y
            samplePt.PlusVector(vectSample)
            pts.Add(New Autodesk.AutoCAD.Geometry.Point3d(samplePt.m_X, samplePt.m_Y, 0.0))
        Next tmpAngle

        Return pts

    End Function

    Private Shared Function ResetPointArray(ByVal pts As Collections.Generic.List(Of Autodesk.AutoCAD.Geometry.Point3d)) As Double()
        Dim varPoints As New Collections.Generic.List(Of Double)

        For Each pt As Autodesk.AutoCAD.Geometry.Point3d In pts
            varPoints.Add(pt.X)
            varPoints.Add(pt.Y)
            varPoints.Add(pt.Z)
        Next

        Return varPoints.ToArray()

    End Function

End Class
