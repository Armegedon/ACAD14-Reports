' -----------------------------------------------------------------------------
' <copyright file="ReportCommand.vb" company="Autodesk">
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
Imports Autodesk.AutoCAD.Runtime


''' <summary>
''' Implements AutoCAD .NET extension commands for Reports.
''' </summary>
''' <remarks>
''' Most of the reports are implemented with Interop for Civil3D COM API for now.
''' This is because .NET reports are initially ported from VBA reports. That's 
''' why most of the .NET report code is VBA style, but not .NET style. For example,
''' array o string is used to hold the report data rather than a class or struct.
''' 
''' As Civil 3D .NET API becomes mature, new reports should be implemented using 
''' .NET API. DO NOT just copy the existed report code.
''' </remarks>
Public Class ReportCommand

    ' Command names
    Friend Const CmdReportAlignCriVer As String = "ReportAlignCriVer"
    Friend Const CmdReportAlignPISta As String = "ReportAlignPISta"
    Friend Const CmdReportAlignStaInc As String = "ReportAlignStaInc"
    Friend Const CmdReportAlignStakeout As String = "ReportAlignStakeout"

    Friend Const CmdReportCorridorFL As String = "ReportCorridorFL"
    Friend Const CmdReportCorridorMilling As String = "ReportCorridorMilling"
    Friend Const CmdReportCorridorSlopeStake As String = "ReportCorridorSlopeStake"
    Friend Const CmdReportCorridorVolume As String = "ReportCorridorVolume"
    Friend Const CmdReportCorridorSectionPoints As String = "ReportCorridorSectionPoints"

    Friend Const CmdReportParcelMapCheck As String = "ReportParcelMapCheck"
    Friend Const CmdReportParcelVol As String = "ReportParcelVol"

    Friend Const CmdReportPointsStaOffset As String = "ReportPointsStaOffset"

    Friend Const CmdReportCrossSection As String = "ReportCrossSection"
    Friend Const CmdReportDaylight As String = "ReportDaylight"
    Friend Const CmdReportExistingProfilePoints As String = "ReportExistingProfilePoints"
    Friend Const CmdReportProfileCriVer As String = "ReportProfileCriVer"
    Friend Const CmdReportProfilePVCurve As String = "ReportProfilePVCurve"
    Friend Const CmdReportProfilePVICurve As String = "ReportProfilePVICurve"
    Friend Const CmdReportProfileStaInc As String = "ReportProfileStaInc"

    Private Const ReportCommandGroupName As String = "ReportCommand"

    <CommandMethod(ReportCommandGroupName, CmdReportAlignCriVer, CommandFlags.Modal)> _
    Public Shared Sub AlignCriVer()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_AlignCriVer.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportAlignPISta, CommandFlags.Modal)> _
    Public Shared Sub AlignPISta()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_AlignSta.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportAlignStaInc, CommandFlags.Modal)> _
    Public Shared Sub AlignStaInc()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_AlignStaInc.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportAlignStakeout, CommandFlags.Modal)> _
    Public Shared Sub AlignStakeout()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_AlignStakeout.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportCorridorFL, CommandFlags.Modal)> _
    Public Shared Sub CorridorFL()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_CorridorCrossSection.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportCorridorMilling, CommandFlags.Modal)> _
    Public Shared Sub CorridorMilling()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_CorridorMilling.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportCorridorSlopeStake, CommandFlags.Modal)> _
    Public Shared Sub CorridorSlopeStake()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_CorridorSlopeStake.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportCorridorVolume, CommandFlags.Modal)> _
    Public Shared Sub CorridorVolume()
		ReportUtilities.UpdateUserLocalSetting()
		try
	        ReportForm_CorridorVolume.Run()
		Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportCorridorSectionPoints, CommandFlags.Modal)> _
    Public Shared Sub CorridorSectionPoints()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_CorridorSectionPoints.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub
    
    <CommandMethod(ReportCommandGroupName, CmdReportParcelMapCheck, CommandFlags.Modal)> _
    Public Shared Sub ParcelMapCheck()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_ParcelMapCheck.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportParcelVol, CommandFlags.Modal)> _
    Public Shared Sub ParcelVol()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_ParcelVol.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportPointsStaOffset, CommandFlags.Modal)> _
    Public Shared Sub PointsStaOffset()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_PointsStaOffset.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportCrossSection, CommandFlags.Modal)> _
    Public Shared Sub CrossSection()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_CrossSection.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportDaylight, CommandFlags.Modal)> _
    Public Shared Sub Daylight()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_Daylight.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportExistingProfilePoints, CommandFlags.Modal)> _
    Public Shared Sub ExistingProfilePoints()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_ExistingProfilePoints.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportProfileCriVer, CommandFlags.Modal)> _
    Public Shared Sub ProfileCriVer()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_ProfileCriVer.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportProfilePVCurve, CommandFlags.Modal)> _
    Public Shared Sub ProfilePVCurve()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_ProfPVCurve.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportProfilePVICurve, CommandFlags.Modal)> _
    Public Shared Sub ProfilePVICurve()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_ProfPVICurve.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub

    <CommandMethod(ReportCommandGroupName, CmdReportProfileStaInc, CommandFlags.Modal)> _
    Public Shared Sub ProfileStaInc()
        ReportUtilities.UpdateUserLocalSetting()
        Try
            ReportForm_ProfStaInc.Run()
        Finally
            ReportUtilities.RestoreThreadLocalSetting()
        End Try
    End Sub


    ''' <summary>
    ''' Run AutoCAD command.
    ''' </summary>
    ''' <param name="commandName"></param>
    ''' <remarks>
    ''' Since toolbox will clear any command and partially typed command 
    ''' on the command line before calling the shared method for the command,
    ''' <c>Run()</c> sends command name to execute directly.
    ''' </remarks>
    Public Shared Sub Run(ByVal commandName As String)
        ReportApplication.ActiveDocument.SendStringToExecute(commandName & vbLf, True, False, True)
    End Sub
End Class

