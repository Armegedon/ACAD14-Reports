VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SegmentCurve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Public Sub FillData(dStartNorth As Double, dStartEast As Double, oSegmentCurve As AeccParcelSegmentCurve, bCounterClockWise As Boolean, bAcrossChord As Boolean)
  
    'radius
    Radius = RoundLength_Parcel(oSegmentCurve.Radius)
    
    'Length
    Length = Radius * oSegmentCurve.Delta
    
    'delta
    Delta = oSegmentCurve.Delta
    
    'Tangent
    Tangent = Abs(oSegmentCurve.Tangent)
  
    
    'get real centerPoint
    Dim dirS2E As Double
    Dim dirS2C As Double
    Dim centerNorth As Double, centerEast As Double
    dirS2E = 2 * Atn(1) - CalcDirRad(oSegmentCurve.StartY, oSegmentCurve.StartX, oSegmentCurve.EndY, oSegmentCurve.EndX)
    dirS2C = dirS2E - 2 * Atn(1) + Delta / 2
    CourseIn = dirS2C
    CalculateNextPoint oSegmentCurve.StartY, oSegmentCurve.StartX, centerNorth, centerEast, dirS2C, oSegmentCurve.Radius
    
    'CourseIn and CourseOut
    Dim realSNorth As Double, realSEast As Double, realENorth As Double, realEEast As Double
    If bCounterClockWise = False Then
        realSNorth = oSegmentCurve.StartY
        realSEast = oSegmentCurve.StartX
        realENorth = oSegmentCurve.EndY
        realEEast = oSegmentCurve.EndX
    Else
        realSNorth = oSegmentCurve.EndY
        realSEast = oSegmentCurve.EndX
        realENorth = oSegmentCurve.StartY
        realEEast = oSegmentCurve.StartX
    End If
    CourseIn = RoundDirection_Parcel(realSNorth, realSEast, centerNorth, centerEast)
    CourseOut = RoundDirection_Parcel(centerNorth, centerEast, realENorth, realEEast)
    
    'chord length
    Chord = RoundLength_Parcel(Sqr((oSegmentCurve.EndY - oSegmentCurve.StartY) ^ 2 + (oSegmentCurve.EndX - oSegmentCurve.StartX) ^ 2))
    
    ' calculate the center of the arc for display
    CalculateNextPoint dStartNorth, dStartEast, RP_North, RP_East, CourseIn, Radius
    
    'chord Course
    Course = RoundDirection_Parcel(realSNorth, realSEast, realENorth, realEEast)
    
    ' perform map check across? to calculate End Pt
    If bAcrossChord = True Then
        CalculateNextPoint dStartNorth, dStartEast, End_North, End_East, Course, Chord
    Else
        CalculateNextPoint RP_North, RP_East, End_North, End_East, CourseOut, Radius
    End If
    
End Sub
