VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OPCDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written 08JUN2021 by Nicholas Stom
'Class is used to OPC DA using OPC Foundation DA Libraries

Option Explicit
' These are properties
    Public opcServerName As String
    Public opcItems As Variant
    Public opcItemsValue As Variant


Public Function GetOpcServers()
    Dim serverNames As Variant
    Dim listServers As Variant
    Dim i As Integer
    Dim theStates As Variant
    
    On Error GoTo ErrorHandler
    Dim theServer As New opcServer
    serverNames = theServer.GetOpcServers
    theStates = Array("Disconnected", "Running", "Failed", "No Configuration", "Suspended", "In Test")
    For i = LBound(serverNames) To UBound(serverNames)
        Debug.Print (serverNames(i))
    Next i
    Exit Function
ErrorHandler:
    Debug.Print ("Error Recieved: " & Err.Number)
End Function

Public Function WriteValueToItems()
    On Error GoTo ErrorHandler
        
        Dim serverNames As Variant
        Dim listServers As Variant
        Dim i As Integer
        Dim theStates As Variant

        Dim theServer As New opcServer
        theServer.Connect (opcServerName)
        
        'Groups
        Dim theGroup As OPCGroup
        Dim theGroups As OPCGroups
        Dim txtName As String
        If theGroups Is Nothing Then
                Set theGroups = theServer.OPCGroups
        End If
        If theGroup Is Nothing Then
                Set theGroup = theGroups.Add("generalItems")
                txtName = theGroup.name
        End If

        theGroup.UpdateRate = CLng(1000)
        theGroup.DeadBand = CLng(1)
        theGroup.TimeBias = CLng(0)
        theGroup.IsActive = CBool(1)
        theGroup.IsSubscribed = CBool(1)



        Dim theItem As OPCItem
        Dim myItems As Variant
        Dim myWriteValues As Variant
        Dim handles(1) As Long
        Dim Errors() As Long
        Dim currentHandle As Long
        myItems = opcItems
        myWriteValues = opcItemsValue

        For i = LBound(myItems) To UBound(myItems)
                Set theItem = theGroup.opcItems.AddItem(myItems(i), currentHandle)
                theItem.Write (myWriteValues(i))
                handles(1) = theGroup.opcItems.Item(1).ServerHandle
                theGroup.opcItems.Remove 1, handles, Errors

        Next i
    
        Exit Function
ErrorHandler:
    Debug.Print ("Error Recieved: " & Err.Number)
End Function


