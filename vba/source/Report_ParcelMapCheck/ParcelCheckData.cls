VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParcelCheckData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Segments As Collection
Private mCounterClockWise As Boolean
Private mAcrossChord As Boolean

Private mPOB_North As Double
Private mPOB_East As Double

Public Area As Double

Private mPerimeter As Double
Public Perimeter As Double




Public Property Get CounterClockWise() As Boolean
    CounterClockWise = mCounterClockWise
End Property

Public Property Get AcrossChord() As Boolean
    AcrossChord = mAcrossChord
End Property



Public Property Get POB_North() As Double
    POB_North = mPOB_North
End Property

Public Property Get POB_East() As Double
    POB_East = mPOB_East
End Property

Public Property Get SegmentCount() As Long
    SegmentCount = Segments.count
End Property

Public Function Item(index As Long) As Object
    Set Item = Segments.Item(index)
End Function

Public Sub Reset(couterClockWise As Boolean, AcrossChord As Boolean)
    Set Segments = New Collection
    mPOB_North = 0#
    mPOB_East = 0#
    mPerimeter = 0#
    Perimeter = 0#
    Area = 0#
    mCounterClockWise = couterClockWise
    mAcrossChord = AcrossChord
End Sub

Public Sub AddLine(oLine As AeccParcelSegmentLine)
    Dim segCount As Long
    Dim curNorth As Double, curEast As Double
    segCount = Segments.count
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
    segLine.FillData curNorth, curEast, oLine, mCounterClockWise
    Segments.Add segLine
    mPerimeter = mPerimeter + segLine.Length
    
End Sub

Public Sub AddCurve(oCurve As AeccParcelSegmentCurve)
    Dim segCount As Long
    Dim curNorth As Double, curEast As Double
    segCount = Segments.count
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
    segCurve.FillData curNorth, curEast, oCurve, mCounterClockWise, mAcrossChord
    Segments.Add segCurve
    If mAcrossChord = True Then
        mPerimeter = mPerimeter + segCurve.Chord
    Else
        mPerimeter = mPerimeter + segCurve.Length
    End If
End Sub

Private Sub Class_Initialize()
    Reset False, False
End Sub

Public Sub GetErrorInfo(ByRef Closure As Double, ByRef Course As Double, _
                        ByRef North As Double, ByRef East As Double, ByRef precision As Double)
    
    Dim count As Long
    count = Segments.count
    If count = 0 Then
        Exit Sub
    End If
    
    Dim endNorth As Double, endEast As Double
    endNorth = Segments(count).End_North
    endEast = Segments(count).End_East
    
    
    'Get Error Closure
    Dim dx As Double
    Dim dy As Double
    Dim distSettings As AeccSettingsDistance
    Dim coordSettings As AeccSettingsCoordinate
    Set distSettings = g_oAeccDb.Settings.ParcelSettings.AmbientSettings.DistanceSettings
    Set coordSettings = g_oAeccDb.Settings.ParcelSettings.AmbientSettings.CoordinateSettings
    dx = endEast - mPOB_East
    dy = endNorth - mPOB_North
    Closure = Sqr(dx * dx + dy * dy)
    Closure = RoundVal(Closure, coordSettings.precision, distSettings.Rounding)
    
    'Get Error Course
    Course = 2 * Atn(1) - CalcDirRad(mPOB_North, mPOB_East, endNorth, endEast)
   
    'Get Error North and East
    North = dy
    East = dx
    
    'Get Error Precision
    Dim errColsure As Double
    If Closure < 0.000001 Then
        errColsure = 0.000001
    Else
        errColsure = Closure
    End If
    precision = mPerimeter / errColsure
End Sub
                   
