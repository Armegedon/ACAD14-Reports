VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAlignmentPIs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mObject As AeccAlignment
Private mPIs As New Dictionary



Private Sub Class_Initialize()
    Set mObject = Nothing
End Sub

Public Property Get Alignment() As AeccAlignment
  Set Alignment = mObject
End Property

Public Property Let Alignment(oAlignment As AeccAlignment)
  Set mObject = oAlignment
  If Not oAlignment Is Nothing Then
    mPIs.removeAll
    Dim lCount As Long
    Dim oStation As AeccAlignmentStation
    lCount = 0
    For Each oStation In oAlignment.GetStations(aeccGeometryPoint, 0#, 0#)
        lCount = lCount + 1
        mPIs.Add lCount, oStation.Station
    Next
  End If
End Property

Public Property Get PICounts() As Long
    PICounts = mPIs.Count
End Property

Public Function GetStationByNum(ByVal num As Long) As Double
    Dim lRes As Double
    
    If mPIs.Exists(num) Then
        lRes = mPIs.item(num)
    End If
    GetStationByNum = lRes
End Function
