VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GOTOPIVISION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written 24AUG2020 by Nicholas Stom
'Open AdHoc Pi Vision

Option Explicit
' These are properties
    Public startTime As String
    Public endTime As String
    Public piTag As Variant
    Public browserPath As String
    Public piVisionUrl As String

' These are methods
Sub OpenPiVision()
    Dim i As Integer
    Dim sCmd As String
    piVisionUrl = piVisionUrl & "/pivision/#/Displays/AdHoc?DataItems="
    'Concat Pi tags to url
    If IsArray(piTag) = True Then
        For i = 0 To UBound(piTag)
            If i = UBound(piTag) Then
                piVisionUrl = piVisionUrl & piTag(i)
            Else
                piVisionUrl = piVisionUrl & piTag(i) & ";"
            End If
        Next i
    Else
        piVisionUrl = piVisionUrl & piTag & ";"
    End If
    'concat time to url
    piVisionUrl = piVisionUrl & "&starttime=" & startTime
    piVisionUrl = piVisionUrl & "&endtime=" & endTime

  
     
    If InStr(browserPath, "MicrosoftEdge") > 0 Then
        sCmd = "start microsoft-edge:" & piVisionUrl
        Shell "cmd /c """ & sCmd & """", vbHide
    Else
        Shell (browserPath & " -url " & piVisionUrl)
    End If

End Sub
