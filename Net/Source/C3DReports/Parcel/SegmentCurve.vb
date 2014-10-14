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


Friend Class SegmentCurve
    Public Length As Double
    Public Radius As Double
    Public Delta As Double
    Public Tangent As Double
    Public Chord As Double
    Public Course As Double
    Public CourseIn As Double
    Public CourseOut As Double
    Public RP_North As Double
    Public RP_East As Double
    Public End_North As Double
    Public End_East As Double

    Public Sub FillData(ByVal dStartNorth As Double, ByVal dStartEast As Double, _
                        ByVal oSegmentCurve As AeccLandLib.AeccParcelSegmentCurve, _
                        ByVal bCounterClockWise As Boolean, ByVal bAcrossChord As Boolean)

        'radius
        Radius = ParcelMapCheck_ExtractData.RoundLength_Parcel(oSegmentCurve.Radius)

        'Length
        Length = Radius * oSegmentCurve.Delta

        'delta
        Delta = oSegmentCurve.Delta

        'Tangent
        Tangent = Math.Abs(oSegmentCurve.Tangent)

        'start/end point
        Dim startNorth As Double, startEast As Double, endNorth As Double, endEast As Double
        If bCounterClockWise = False Then
            startNorth = oSegmentCurve.StartY
            startEast = oSegmentCurve.StartX
            endNorth = oSegmentCurve.EndY
            endEast = oSegmentCurve.EndX
        Else
            startNorth = oSegmentCurve.EndY
            startEast = oSegmentCurve.EndX
            endNorth = oSegmentCurve.StartY
            endEast = oSegmentCurve.StartX
        End If

        'get center point
        Dim centerNorth As Double, centerEast As Double
        Dim factor As Double
        factor = 0.5 * (1.0 / oSegmentCurve.Bulge - oSegmentCurve.Bulge)
        centerEast = 0.5 * (oSegmentCurve.StartX + oSegmentCurve.EndX - factor * (oSegmentCurve.EndY - oSegmentCurve.StartY))
        centerNorth = 0.5 * (oSegmentCurve.StartY + oSegmentCurve.EndY + factor * (oSegmentCurve.EndX - oSegmentCurve.StartX))

        'CousrIn and CourseOut
        CourseIn = ParcelMapCheck_ExtractData.RoundDirection_Parcel(startNorth, startEast, centerNorth, centerEast)
        CourseOut = ParcelMapCheck_ExtractData.RoundDirection_Parcel(centerNorth, centerEast, endNorth, endEast)

        'chord length
        Chord = ParcelMapCheck_ExtractData.RoundLength_Parcel(Math.Sqrt((oSegmentCurve.EndY _
            - oSegmentCurve.StartY) ^ 2 + (oSegmentCurve.EndX - oSegmentCurve.StartX) ^ 2))

        ' calculate the center of the arc for display
        ParcelMapCheck_ExtractData.CalculateNextPoint(dStartNorth, dStartEast, _
            RP_North, RP_East, CourseIn, Radius)

        'chord Course
        Course = ParcelMapCheck_ExtractData.RoundDirection_Parcel(startNorth, _
            startEast, endNorth, endEast)

        ' perform map check across? to calculate End Pt
        If bAcrossChord = True Then
            Dim roundedChord As Double
            Dim distSettings As AeccLandLib.IAeccSettingsDistance
            distSettings = ReportApplication.AeccXDatabase.Settings.ParcelSettings.AmbientSettings.DistanceSettings
            roundedChord = ReportFormat.RoundDouble(Chord, distSettings.Precision.Value, distSettings.Rounding.Value)
            ParcelMapCheck_ExtractData.CalculateNextPoint(dStartNorth, _
                dStartEast, End_North, End_East, Course, roundedChord)
        Else
            ParcelMapCheck_ExtractData.CalculateNextPoint(RP_North, _
                RP_East, End_North, End_East, CourseOut, Radius)
        End If
    End Sub
End Class
