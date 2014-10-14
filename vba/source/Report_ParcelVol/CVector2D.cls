VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVector2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public m_X As Double
Public m_Y As Double




Private Sub Class_Initialize()
    m_X = 0#
    m_Y = 1#
End Sub

Public Sub Rotate(rotAngle As Double)
    Dim x As Double
    Dim y As Double
    x = m_X
    y = m_Y
    m_X = x * Cos(rotAngle) - y * Sin(rotAngle)
    m_Y = x * Sin(rotAngle) + y * Cos(rotAngle)
End Sub


Public Sub SetLength(length As Double)
    Dim x As Double
    Dim y As Double
    x = m_X
    y = m_Y

    m_X = m_X * length / GetLength()
    m_Y = m_Y * length / GetLength()
End Sub

Public Sub PlusVector(oVector As CVector2D)
    m_X = m_X + oVector.m_X
    m_Y = m_Y + oVector.m_Y
End Sub

Public Function GetLength() As Double
    GetLength = Sqr(m_X * m_X + m_Y * m_Y)
End Function

Public Sub ScaleV(value As Double)
    m_X = m_X * value
    m_Y = m_Y * value
End Sub

