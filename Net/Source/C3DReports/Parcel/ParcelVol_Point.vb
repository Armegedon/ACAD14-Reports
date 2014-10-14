''//
''// (C) Copyright 2008 by Autodesk, Inc.
''//
''// Permission to use, copy, modify, and distribute this software in
''// object code form for any purpose and without fee is hereby granted,
''// provided that the above copyright notice appears in all copies and
''// that both that copyright notice and the limited warranty and
''// restricted rights notice below appear in all supporting
''// documentation.
''//
''// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
''// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
''// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
''// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
''// UNINTERRUPTED OR ERROR FREE.
''//
''// Use, duplication, or disclosure by the U.S. Government is subject to
''// restrictions set forth in FAR 52.227-19 (Commercial Computer
''// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
''// (Rights in Technical Data and Computer Software), as applicable.

Option Explicit On

Friend Class CPoint2D
    Public m_X As Double
    Public m_Y As Double

    Public Sub New()
        m_X = 0.0#
        m_Y = 0.0#
    End Sub

    Public Sub PlusVector(ByVal oVector As CVector2D)
        m_X = m_X + oVector.m_X
        m_Y = m_Y + oVector.m_Y
    End Sub

End Class

Friend Class CVector2D
    Public m_X As Double
    Public m_Y As Double

    Public Sub New()
        m_X = 0.0#
        m_Y = 1.0#
    End Sub

    Public Sub Rotate(ByVal rotAngle As Double)
        Dim x As Double
        Dim y As Double
        x = m_X
        y = m_Y
        m_X = x * Math.Cos(rotAngle) - y * Math.Sin(rotAngle)
        m_Y = x * Math.Sin(rotAngle) + y * Math.Cos(rotAngle)
    End Sub


    Public Sub SetLength(ByVal length As Double)
        Dim x As Double
        Dim y As Double
        x = m_X
        y = m_Y

        m_X = m_X * length / GetLength()
        m_Y = m_Y * length / GetLength()
    End Sub

    Public Sub PlusVector(ByVal oVector As CVector2D)
        m_X = m_X + oVector.m_X
        m_Y = m_Y + oVector.m_Y
    End Sub

    Public Function GetLength() As Double
        GetLength = Math.Sqrt(m_X * m_X + m_Y * m_Y)
    End Function

    Public Sub ScaleV(ByVal value As Double)
        m_X = m_X * value
        m_Y = m_Y * value
    End Sub

End Class
