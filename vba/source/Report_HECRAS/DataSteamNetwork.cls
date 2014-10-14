VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSteamNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public header As New DataHeader
Public EndPoints As New Collection
Public Reaches As New Collection
Public CrossSections As New Collection

Public Sub Clear()
    Set EndPoints = New Collection
    Set Reaches = New Collection
    Set CrossSections = New Collection
End Sub
