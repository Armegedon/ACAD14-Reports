VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCorridor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public mObject As AeccCorridor
Public mBaselines As Dictionary

Private Sub Class_Initialize()
    Set mBaselines = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set mObject = Nothing
End Sub
