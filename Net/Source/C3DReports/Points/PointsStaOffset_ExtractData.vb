' -----------------------------------------------------------------------------
' <copyright file="PointsStaOffset_ExtractData.vb" company="Autodesk">
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

Imports AutoCAD = Autodesk.AutoCAD.Interop
Imports AeccLandLib = Autodesk.AECC.Interop.Land
Imports LocalizedRes = Report.My.Resources.LocalizedStrings

Friend Class PointsStaOffset_ExtractData

    Public Const nPointIndex As Integer = 0
    Public Const nStationIndex As Integer = 1
    Public Const nOffsetIndex As Integer = 2
    Public Const nElevation As Integer = 3
    Public Const nDescriptionRaw As Integer = 4
    Public Const nDescriptionFull As Integer = 5
    Public Const nLastIndex As Integer = nDescriptionFull

    Private Shared m_oPointDataArr(nLastIndex) As String

    Public Shared ReadOnly Property PointDataArr() As String()
        Get
            Return m_oPointDataArr
        End Get
    End Property

    Shared Function FormatPtElevationSettings(ByVal dDis As Double) As String
        Dim oElevationSettings As AeccLandLib.AeccSettingsElevation
        oElevationSettings = ReportApplication.AeccXDatabase.Settings.PointSettings.AmbientSettings.ElevationSettings

        FormatPtElevationSettings = ReportFormat.FormatDistance(dDis, _
                    oElevationSettings.Unit.Value, oElevationSettings.Precision.Value, _
                    oElevationSettings.Rounding.Value, oElevationSettings.Sign.Value)
    End Function

    Shared Function FormatPtCoordSettings(ByVal dDis As Double) As String
        Dim oCoordSettings As AeccLandLib.AeccSettingsCoordinate
        oCoordSettings = ReportApplication.AeccXDatabase.Settings.PointSettings.AmbientSettings.CoordinateSettings

        FormatPtCoordSettings = ReportFormat.FormatDistance(dDis, _
                    oCoordSettings.Unit.Value, oCoordSettings.Precision.Value, _
                    oCoordSettings.Rounding.Value, oCoordSettings.Sign.Value)
    End Function

    Shared Function FormatDistSettings(ByVal dDis As Double) As String
        Dim oDistSettings As AeccLandLib.IAeccSettingsDistance
        oDistSettings = ReportApplication.AeccXDatabase.Settings.AlignmentSettings.AmbientSettings.DistanceSettings

        FormatDistSettings = ReportFormat.FormatDistance(dDis, _
                    oDistSettings.Unit.Value, oDistSettings.Precision.Value, _
                    oDistSettings.Rounding.Value, oDistSettings.Sign.Value)
    End Function

    Public Shared Function ExtractData(ByVal oAlignment As AeccLandLib.IAeccAlignment, _
                                ByVal oPoint As AeccLandLib.AeccPoint) As Boolean
        Dim ptNorthing As Double
        Dim ptEasting As Double
        Dim station As Double
        Dim offSet As Double

        ptNorthing = oPoint.Northing
        ptEasting = oPoint.Easting

        m_oPointDataArr(nPointIndex) = oPoint.Number.ToString()
        m_oPointDataArr(nElevation) = FormatPtElevationSettings(oPoint.Elevation)
        m_oPointDataArr(nDescriptionRaw) = oPoint.RawDescription
        m_oPointDataArr(nDescriptionFull) = oPoint.FullDescription

        If oAlignment.StationOffsetEx(ptEasting, ptNorthing, 0.0#, station, offSet) = _
            AeccLandLib.AeccAlignmentReturnValue.aeccStationOutOfRange Then
            m_oPointDataArr(nStationIndex) = LocalizedRes.PointsStaOffset_Form_OutRange '"Out of range"
            m_oPointDataArr(nOffsetIndex) = LocalizedRes.PointsStaOffset_Form_OutRange '"Out of range"
        Else
            m_oPointDataArr(nStationIndex) = ReportUtilities.GetStationStringWithDerived(oAlignment, station)
            m_oPointDataArr(nOffsetIndex) = FormatDistSettings(offSet)
        End If

        ExtractData = True
    End Function

End Class
