VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UNIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written 22APR2021 by Nicholas Stom


Option Explicit
' These are properties
    Public startTime As Date
    Public endTime As Date
    Public piTag As Variant
    Public dataMatrix As Variant
    Public batchId As String
    Public name As Variant


Sub GenerateMyUnit()
    Debug.Print name
    Debug.Print batchId
    Debug.Print startTime
    Debug.Print endTime
End Sub
