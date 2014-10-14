VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSTEAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mSampleLineGroup As AeccSampleLineGroup
Public mReaches As New Dictionary 'key is reaches name, value is the begin and End Position

Public Function AddReach(sReachName As String, lBegin As Long, lEnd As Long) As Boolean
    Dim bRes As Boolean
    bRes = False
    If Not mReaches.Exists(sReachName) Then
        Dim oReach As New CReachRegion
        If oReach.SetValue(lBegin, lEnd) Then
            mReaches.Add sReachName, oReach
            bRes = True
        End If
    End If
    AddReach = bRes
End Function

Public Function RemoveReach(sReachName As String) As Boolean
    Dim bRes As Boolean
    bRes = False
    If mReaches.Exists(sReachName) Then
       mReaches.Remove sReachName
       bRes = True
       
    End If
    RemoveReach = bRes
End Function

Public Property Let SampleLineGroup(oSampleLineGroup As AeccSampleLineGroup)
    Set mSampleLineGroup = oSampleLineGroup
    mReaches.removeAll
End Property


Public Property Get SampleLineGroup() As AeccSampleLineGroup
    Set SampleLineGroup = mSampleLineGroup
End Property
