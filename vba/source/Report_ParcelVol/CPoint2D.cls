VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPoint2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_X As Double
Public m_Y As Double




Private Sub Class_Initialize()
    m_X = 0#
    m_Y = 0#
End Sub

Public Sub PlusVector(oVector As CVector2D)
    m_X = m_X + oVector.m_X
    m_Y = m_Y + oVector.m_Y
End Sub
