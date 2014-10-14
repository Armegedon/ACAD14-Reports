VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SegmentLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Course As Double
Public Length As Double
Public End_North As Double
Public End_East As Double

Public Sub FillData(dStartNorth As Double, dStartEast As Double, oSegmentLine As AeccParcelSegmentLine, bCounterClockWise As Boolean)
    Dim tmpCourse As Double
    Dim roundedAngle As Double
    Dim roundedLength As Double
    
    tmpCourse = 2 * Atn(1) - CalcDirRad(oSegmentLine.StartY, oSegmentLine.StartX, oSegmentLine.EndY, oSegmentLine.EndX)
    roundedLength = RoundLength_Parcel(oSegmentLine.Length)
    If bCounterClockWise = False Then
        Course = tmpCourse
        roundedAngle = RoundDirection_Parcel(oSegmentLine.StartY, oSegmentLine.StartX, oSegmentLine.EndY, oSegmentLine.EndX)
        
    Else
        Course = -tmpCourse
        roundedAngle = RoundDirection_Parcel(oSegmentLine.EndY, oSegmentLine.EndX, oSegmentLine.StartY, oSegmentLine.StartX)
    End If
    CalculateNextPoint dStartNorth, dStartEast, End_North, End_East, roundedAngle, roundedLength
    Length = oSegmentLine.Length
End Sub


