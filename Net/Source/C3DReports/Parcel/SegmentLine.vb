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

Imports System
Imports AeccLandLib = Autodesk.AECC.Interop.Land


Friend Class SegmentLine

    Public Course As Double
    Public Length As Double
    Public End_North As Double
    Public End_East As Double

    Public Sub FillData(ByVal dStartNorth As Double, _
                            ByVal dStartEast As Double, _
                            ByVal oSegmentLine As AeccLandLib.AeccParcelSegmentLine, _
                            ByVal bCounterClockWise As Boolean)
        Dim roundedAngle As Double
        Dim roundedLength As Double

        Dim startNorth As Double, endNorth As Double, startEast As Double, endEast As Double
        If bCounterClockWise = False Then
            startNorth = oSegmentLine.StartY
            startEast = oSegmentLine.StartX
            endNorth = oSegmentLine.EndY
            endEast = oSegmentLine.EndX
        Else
            startNorth = oSegmentLine.EndY
            startEast = oSegmentLine.EndX
            endNorth = oSegmentLine.StartY
            endEast = oSegmentLine.StartX
        End If

        Course = Math.PI / 2 - ReportUtilities.CalcDirRad(startNorth, startEast, endNorth, endEast)
        roundedLength = ParcelMapCheck_ExtractData.RoundLength_Parcel(oSegmentLine.Length)

        roundedAngle = ParcelMapCheck_ExtractData.RoundDirection_Parcel(startNorth, _
                        startEast, endNorth, endEast)
        ParcelMapCheck_ExtractData.CalculateNextPoint(dStartNorth, dStartEast, _
            End_North, End_East, roundedAngle, roundedLength)
        Length = oSegmentLine.Length
    End Sub
End Class
