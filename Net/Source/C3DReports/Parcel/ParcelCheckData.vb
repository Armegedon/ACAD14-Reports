''//
''// (C) Copyright 2008 by Autodesk, Inc.
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

Option Explicit On

Imports System
Imports AeccLandLib = Autodesk.AECC.Interop.Land


Friend Class ParcelCheckData
    Private Segments As Collection
    Private mCounterClockWise As Boolean
    Private mAcrossChord As Boolean

    Private mPOB_North As Double
    Private mPOB_East As Double

    Public Area As Double

    Private mPerimeter As Double
    Public Perimeter As Double

    Public ReadOnly Property CounterClockWise() As Boolean
        Get
            Return mCounterClockWise
        End Get
    End Property

    Public ReadOnly Property AcrossChord() As Boolean
        Get
            Return mAcrossChord
        End Get
    End Property

    Public ReadOnly Property POB_North() As Double
        Get
            Return mPOB_North
        End Get
    End Property

    Public ReadOnly Property POB_East() As Double
        Get
            Return mPOB_East
        End Get
    End Property

    Public ReadOnly Property SegmentCount() As Long
        Get
            Return Segments.Count
        End Get
    End Property

    Public Function Item(ByVal index As Long) As Object
        Item = Segments.Item(index)
    End Function

    Public Sub Reset(ByVal couterClockWise As Boolean, ByVal AcrossChord As Boolean)
        Segments = New Collection
        mPOB_North = 0.0
        mPOB_East = 0.0
        mPerimeter = 0.0
        Perimeter = 0.0
        Area = 0.0
        mCounterClockWise = couterClockWise
        mAcrossChord = AcrossChord
    End Sub

    Public Sub AddLine(ByVal oLine As AeccLandLib.AeccParcelSegmentLine)
        Dim segCount As Long
        Dim curNorth As Double, curEast As Double
        segCount = Segments.Count
        If segCount = 0 Then
            If mCounterClockWise = False Then
                mPOB_North = oLine.StartY
                mPOB_East = oLine.StartX
            Else
                mPOB_North = oLine.EndY
                mPOB_East = oLine.EndX
            End If
            curNorth = mPOB_North
            curEast = mPOB_East
        Else
            curNorth = Segments(segCount).End_North
            curEast = Segments(segCount).End_East
        End If

        Dim segLine As New SegmentLine
        segLine.FillData(curNorth, curEast, oLine, mCounterClockWise)
        Segments.Add(segLine)

        Dim dLength As Double
        Dim distSettings As AeccLandLib.AeccSettingsDistance
        distSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.DistanceSettings
        dLength = ReportFormat.RoundDouble(segLine.Length, distSettings.Precision.Value, distSettings.Rounding.Value)
        mPerimeter += dLength

    End Sub

    Public Sub AddCurve(ByVal oCurve As AeccLandLib.AeccParcelSegmentCurve)
        Dim segCount As Long
        Dim curNorth As Double, curEast As Double
        segCount = Segments.Count
        If segCount = 0 Then
            If mCounterClockWise = False Then
                mPOB_North = oCurve.StartY
                mPOB_East = oCurve.StartX
            Else
                mPOB_North = oCurve.EndY
                mPOB_East = oCurve.EndX
            End If
            curNorth = mPOB_North
            curEast = mPOB_East
        Else
            curNorth = Segments(segCount).End_North
            curEast = Segments(segCount).End_East
        End If

        Dim segCurve As New SegmentCurve
        segCurve.FillData(curNorth, curEast, oCurve, mCounterClockWise, mAcrossChord)
        Segments.Add(segCurve)

        Dim offsetLength As Double
        If mAcrossChord = True Then
            offsetLength = segCurve.Chord
        Else
            offsetLength = segCurve.Length
        End If
        Dim distSettings As AeccLandLib.IAeccSettingsDistance
        distSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.DistanceSettings
        offsetLength = ReportFormat.RoundDouble(offsetLength, distSettings.Precision.Value, distSettings.Rounding.Value)
        mPerimeter += offsetLength

    End Sub

    Public Sub New()
        Reset(False, False)
    End Sub

    Public Sub GetErrorInfo(ByRef Closure As Double, ByRef Course As Double, _
                            ByRef North As Double, ByRef East As Double, ByRef precision As Double)

        Dim count As Long
        count = Segments.Count
        If count = 0 Then
            Exit Sub
        End If

        Dim endNorth As Double, endEast As Double
        endNorth = Segments(count).End_North
        endEast = Segments(count).End_East


        'Get Error Closure
        Dim dx As Double
        Dim dy As Double
        Dim distSettings As AeccLandLib.AeccSettingsDistance
        Dim coordSettings As AeccLandLib.AeccSettingsCoordinate
        distSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.DistanceSettings
        coordSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.CoordinateSettings
        dx = endEast - mPOB_East
        dy = endNorth - mPOB_North
        Closure = Math.Sqrt(dx * dx + dy * dy)
        Closure = ReportFormat.RoundDouble(Closure, coordSettings.Precision.Value, distSettings.Rounding.Value)

        'Get Error Course
        Course = ParcelMapCheck_ExtractData.RoundDirection_Parcel(mPOB_North, mPOB_East, endNorth, endEast)

        'Get Error North and East
        North = dy
        East = dx

        'Get Error Precision
        Dim errColsure As Double
        If Closure < 0.000000001 Then
            errColsure = 0.000001
        Else
            errColsure = Closure
        End If
        precision = mPerimeter / errColsure
    End Sub

End Class
