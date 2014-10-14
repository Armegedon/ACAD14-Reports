VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CReachRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mBeginPt As Long
Private mEndPt As Long


Public Property Get BeginPt() As Long
    BeginPt = mBeginPt
End Property

Public Property Get EndPt() As Long
    EndPt = mEndPt
End Property

Public Function SetValue(ByVal dBeginPt As Long, ByVal dEndPt As Long) As Boolean
    If (Not dBeginPt = dEndPt) And dBeginPt >= 1# Then
        mBeginPt = dBeginPt
        mEndPt = dEndPt
        SetValue = True
    Else
        SetValue = False
    End If
End Function
