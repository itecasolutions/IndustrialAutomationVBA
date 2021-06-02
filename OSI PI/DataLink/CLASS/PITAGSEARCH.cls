VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PITAGSEARCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Covington OSI PI Library
'Written 24AUG2020 by Nicholas Stom
'Class is used for a finding Pi Tags
'   * Can be used as WildCards
'   Remember to Add your PI Referemces under Tools


Option Explicit
' These are properties
    Public piTag As String
    Public piServer As String
    Public dataMatrix As Variant
    Public printOutput As Boolean
' These are methods
    Public Sub Get_PiTags()
        Dim oServer As PISDK.Server
        Dim ptList As PISDK.PointList
        Dim pt As PISDK.PIPoint
        Dim strQuery As String
        Set oServer = PISDK.Servers(piServer)
        strQuery = "Tag = '" & piTag & "'"
        Set ptList = oServer.GetPoints(strQuery)
        Dim i As Integer
        Dim internalDataMatrix As Variant
        ReDim internalDataMatrix(0)
        i = 0
        For Each pt In ptList
            ReDim Preserve internalDataMatrix(i)
            internalDataMatrix(i) = pt.name
            If printOutput = True Then
                Debug.Print pt.name
            End If
            i = i + 1
        Next
        
        dataMatrix = internalDataMatrix
    End Sub
