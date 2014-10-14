Attribute VB_Name = "General"
'mod version

Public Const MODVER = "6.0"

' Help file window handle
Public ReportsHH_hWnd As Long
Public Const HH_HELP_CONTEXT = &HF            ' Display mapped numeric value in dwData.

' CommonDialog control flags
Public Const IDH_AECC_REPORTS_CREATE_PI_STATION_REPORT = 5789      '  Create Reports – PI Station Report      18-Oct
Public Const IDH_AECC_REPORTS_CREATE_INCREMENTAL_STATIONING_REPORT_ALIGNMENTS = 5790      '  Create Reports – Incremental Stationing Report (Alignments)     18-Oct
Public Const IDH_AECC_REPORTS_CREATE_STAKEOUT_ALIGNMENT_REPORT = 5791     '  Create Reports – Stakeout Alignment Report      18-Oct
Public Const IDH_AECC_REPORTS_CREATE_PVI_STATION_AND_CURVE_REPORT = 5792      '  Create Reports – PVI Station and Curve Report       18-Oct
Public Const IDH_AECC_REPORTS_CREATE_INCREMENTAL_STATIONING_REPORT_PROFILES = 5793    '  Create Reports – Incremental Stationing Report (Profiles)       18-Oct
Public Const IDH_AECC_REPORTS_CREATE_VERTICAL_CURVE_REPORT = 5794     '  Create Reports – Vertical Curve Report      18-Oct
Public Const IDH_AECC_REPORTS_CREATE_SLOPE_GRADE_REPORT = 5795    '  Create Reports – Slope Grade Report         18-Oct
Public Const IDH_AECC_REPORTS_CREATE_SECTION_DESIGN_REPORT = 5796     '  Create Reports – Section Design Report      18-Oct
Public Const IDH_AECC_REPORTS_CREATE_PARCEL_VOLUMES_REPORT = 5797     '  Create Reports – Parcel Volumes Report      18-Oct
Public Const IDH_AECC_REPORTS_CREATE_PARCEL_MAPCHECK_REPORT = 5798    '  Create Reports – Parcel Mapcheck Report     18-Oct
Public Const IDH_AECC_REPORTS_CREATE_STATION_OFFSET_TO_POINTS_REPORT = 5799   '  Create Reports – Station Offset to Points       18-Oct
Public Const IDH_AECC_REPORTS_CREATE_HEC_RAS_REPORT = 5800   'Create Reports - HEC-RAS
'-----------------------------------------------------------------------
' declare the HtmlHelp function, used by the Help buttons in the dialog boxes
' and IsWindow to test to see if the help file needs to be closed on exit
'-----------------------------------------------------------------------
Public Declare Function ReportsHtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
         (ByVal hwndCaller As Long, ByVal pszFile As String, _
         ByVal uCommand As Long, ByVal dwData As Long) As Long

Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Function getReportsHelpFile() As String
    getReportsHelpFile = AcadApplication.Path & "\Help\" & "Civil3D2008.chm"
End Function


Function getAeccDb() As AeccDatabase
    Dim oApp As AcadApplication
    Set oApp = ThisDrawing.Application
    Const sAppName = "AeccXUiLand.AeccApplication." + MODVER
    Dim oCivilApp As AeccApplication
    Set oCivilApp = oApp.GetInterfaceObject(sAppName)
    Dim oAeccDoc As AeccDocument
    Set oAeccDoc = oCivilApp.ActiveDocument
    Set getAeccDb = oAeccDoc.Database
End Function


Public Function GetRawStation(ByVal equaStation As String) As Double
    Dim tmpStr As String
    Dim station As Double
    tmpStr = Replace(equaStation, "+", "", 1, 1)
    GetRawStation = Val(tmpStr)
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

Public Sub QuickSort(varArray As Variant, Optional iLeft As Integer = -2, Optional iRight As Integer = -2)
    Dim i As Integer
    Dim j As Integer
    Dim iMid As Integer
    
    Dim varTestValue As Variant
    
    If iLeft = -2 Then iLeft = LBound(varArray)
    If iRight = -2 Then iRight = UBound(varArray)
    
    If iLeft < iRight Then
        iMid = (iLeft + iRight) \ 2
        varTestValue = varArray(iMid)
        i = iLeft
        j = iRight
        
        Do
            Do While varArray(i) < varTestValue
                i = i + 1
            Loop
            Do While varArray(j) > varTestValue
                j = j - 1
            Loop
            If i <= j Then
                SwapElements varArray, i, j
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        'To optimize the sort, always sort the smallest segment first
        If j <= iMid Then
            Call QuickSort(varArray, iLeft, j)
            Call QuickSort(varArray, i, iRight)
        Else
            Call QuickSort(varArray, i, iRight)
            Call QuickSort(varArray, iLeft, j)
        End If
    End If
End Sub

Public Sub SwapElements(varItems As Variant, iItem1 As Integer, iItem2 As Integer)
    Dim varTemp As Variant
    varTemp = varItems(iItem2)
    varItems(iItem2) = varItems(iItem1)
    varItems(iItem1) = varTemp
End Sub

Public Sub SwapVars(varItem1 As Variant, varItem2 As Variant)
    Dim varTemp As Variant
    varTemp = varItem2
    varItem2 = varItem1
    varItem1 = varTemp
End Sub

Public Sub OpenFileByDefaultBrowser(fileName As String)
    Call ShellExecute(0, "open", fileName, 0, 0, 1)
End Sub
