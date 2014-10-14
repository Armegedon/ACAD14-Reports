VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBaseline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public mObject As AeccBaseline
Public mSampleLineGroups As Dictionary
Public mLinkCodesGroups As Dictionary

Private Sub Class_Initialize()
    Set mSampleLineGroups = New Dictionary
    Set mLinkCodesGroups = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set mObject = Nothing
End Sub
