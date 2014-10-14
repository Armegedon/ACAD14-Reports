' -----------------------------------------------------------------------------
' <copyright file="ReportEntry.vb" company="Autodesk">
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

''' <summary>
''' Report entries for tool box. The shared method names are used as macro
''' name in the tool box editor.
''' </summary>
''' <remarks>
''' Module <c>ReportEntry</c> holds shared functions for Civil Toolbox.
''' Before Civil 2011, the toolbox can only access shared functions in
''' .NET assembly. We add this to make compatible with legacy toolbox 
''' configuration file. This can be removed after we update the 
''' configuration file to use command names.
''' </remarks>
Public Module ReportEntry
    Public Sub AlignCriVer_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportAlignCriVer)
    End Sub

    Public Sub AlignPISta_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportAlignPISta)
    End Sub

    Public Sub AlignStaInc_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportAlignStaInc)
    End Sub

    Public Sub AlignStakeout_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportAlignStakeout)
    End Sub

    Public Sub CorridorFL_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportCorridorFL)
    End Sub

    Public Sub CorridorMilling_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportCorridorMilling)
    End Sub

    Public Sub CorridorSlopeStake_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportCorridorSlopeStake)
    End Sub

    Public Sub CorridorVolume_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportCorridorVolume)
    End Sub

    Public Sub CorridorSectionPoints_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportCorridorSectionPoints)
    End Sub

    Public Sub ParcelMapCheck_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportParcelMapCheck)
    End Sub

    Public Sub ParcelVol_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportParcelVol)
    End Sub

    Public Sub PointsStaOffset_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportPointsStaOffset)
    End Sub

    Public Sub CrossSection_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportCrossSection)
    End Sub

    Public Sub Daylight_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportDaylight)
    End Sub

    Public Sub ExistingProfilePoints_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportExistingProfilePoints)
    End Sub

    Public Sub ProfileCriVer_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportProfileCriVer)
    End Sub

    Public Sub ProfilePVCurve_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportProfilePVCurve)
    End Sub

    Public Sub ProfilePVICurve_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportProfilePVICurve)
    End Sub

    Public Sub ProfileStaInc_RunReport()
        ReportCommand.Run(ReportCommand.CmdReportProfileStaInc)
    End Sub
End Module

