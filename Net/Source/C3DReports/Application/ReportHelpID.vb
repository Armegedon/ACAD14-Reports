' -----------------------------------------------------------------------------
' <copyright file="ReportHelpID.vb" company="Autodesk">
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
''' Help ID definition, copy from Source\Locale\Inc\_Idh_Aecc.h
''' </summary>
''' <remarks></remarks>
Friend Module ReportHelpID

    ' Create Reports - Alignment Criteria Verification
    Public Const IDH_AECC_REPORTS_CREATE_ALIGNMENT_DESIGN_CRITERIA_VERIFICATION As String = _
        "IDH_AECC_REPORTS_CREATE_ALIGNMENT_DESIGN_CRITERIA_VERIFICATION"

    '  Create Reports - PI Station Report 
    Public Const IDH_AECC_REPORTS_CREATE_PI_STATION_REPORT As String = _
        "IDH_AECC_REPORTS_CREATE_PI_STATION_REPORT"

    ' Create Reports - Incremental Stationing Report (Alignments)
    Public Const IDH_AECC_REPORTS_CREATE_INCREMENTAL_STATIONING_REPORT_ALIGNMENTS As String = _
        "IDH_AECC_REPORTS_CREATE_INCREMENTAL_STATIONING_REPORT_ALIGNMENTS"

    ' Create Reports - Stakeout Alignment Report 
    Public Const IDH_AECC_REPORTS_CREATE_STAKEOUT_ALIGNMENT_REPORT As String = _
        "IDH_AECC_REPORTS_CREATE_STAKEOUT_ALIGNMENT_REPORT"

    ' Create Reports - Corridor Cross Section Report
    Public Const IDH_AECC_REPORTS_CREATE_CORRIDOR_FEATURE_LINE_CROSS_SECTION_REPORT As String = _
        "IDH_AECC_REPORTS_CREATE_CORRIDOR_FEATURE_LINE_CROSS_SECTION_REPORT"

    ' Create Reports - Corridor Section Points Report
    Public Const IDH_AECC_REPORTS_CORRIDOR_SECTION_POINTS_REPORT As String = _
        "IDH_AECC_REPORTS_CORRIDOR_SECTION_POINTS_REPORT"

    ' Create Reports - Slope Grade Report
    Public Const IDH_AECC_REPORTS_CREATE_SLOPE_GRADE_REPORT As String = _
        "IDH_AECC_REPORTS_CREATE_SLOPE_GRADE_REPORT"

    ' Create Reports - Milling Report
    Public Const IDH_AECC_REPORTS_CREATE_MILLING_REPORT As String = _
        "IDH_AECC_REPORTS_CREATE_MILLING_REPORT"

    '  Create Reports - Parcel Mapcheck Report
    Public Const IDH_AECC_REPORTS_CREATE_PARCEL_MAPCHECK_REPORT As String = _
        "IDH_AECC_REPORTS_CREATE_PARCEL_MAPCHECK_REPORT"

    ' Create Reports - Parcel Volumes Report
    Public Const IDH_AECC_REPORTS_CREATE_PARCEL_VOLUMES_REPORT As String = _
        "IDH_AECC_REPORTS_CREATE_PARCEL_VOLUMES_REPORT"

    '  Create Reports - Station Offset to Points
    Public Const IDH_AECC_REPORTS_CREATE_STATION_OFFSET_TO_POINTS_REPORT As String = _
        "IDH_AECC_REPORTS_CREATE_STATION_OFFSET_TO_POINTS_REPORT"

    ' Create Reports - CrossSection
    Public Const IDH_AECC_REPORTS_CREATE_EG_AND_LAYOUT_PROFILE_GRADES_REPORT As String = _
        "IDH_AECC_REPORTS_CREATE_EG_AND_LAYOUT_PROFILE_GRADES_REPORT"

    ' Create Reports - Daylight
    Public Const IDH_AECC_REPORTS_CREATE_DAYLIGHT_REPORT As String = _
        "IDH_AECC_REPORTS_CREATE_DAYLIGHT_REPORT"

    ' Create Reports - Design to Existing level differences
    Public Const IDH_AECC_REPORTS_CREATE_EXISTING_PROFILE_POINTS_REPORT As String = _
        "IDH_AECC_REPORTS_CREATE_EXISTING_PROFILE_POINTS_REPORT"

    '  Create Reports - Profile Design Criteria Verification
    Public Const IDH_AECC_REPORTS_CREATE_PROFILE_DESIGN_CRITERIA_VERIFICATION As String = _
        "IDH_AECC_REPORTS_CREATE_PROFILE_DESIGN_CRITERIA_VERIFICATION"

    '  Create Reports - Vertical Curve Report
    Public Const IDH_AECC_REPORTS_CREATE_VERTICAL_CURVE_REPORT As String = _
        "IDH_AECC_REPORTS_CREATE_VERTICAL_CURVE_REPORT"

    '  Create Reports - PVI Station and Curve Report
    Public Const IDH_AECC_REPORTS_CREATE_PVI_STATION_AND_CURVE_REPORT As String = _
        "IDH_AECC_REPORTS_CREATE_PVI_STATION_AND_CURVE_REPORT"

    '  Create Reports - Incremental Stationing Report (Profiles)
    Public Const IDH_AECC_REPORTS_CREATE_INCREMENTAL_STATIONING_REPORT_PROFILES As String = _
        "IDH_AECC_REPORTS_CREATE_INCREMENTAL_STATIONING_REPORT_PROFILES"
End Module
