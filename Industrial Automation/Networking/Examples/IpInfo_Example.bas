Attribute VB_Name = "IpInfo_Example"
Sub IpInfo_Example()
    Dim myIpInfo As New IPINFO
    myIpInfo.ipAddress = Array("")
    myIpInfo.timeOut = 50
    myIpInfo.GetIpInfo
    myFunc = NetworkGeneralFunctions.OutputToSheet(myIpInfo.dataMatrix, "Ip Table", , , True)
    Debug.Print 1
End Sub


Sub IpInfo_Example1()
    Dim ipString() As String
    Dim size As Integer
    Dim myIpInfo As New IPINFO
    size = 255
    ReDim ipString(size)
    For i = 0 To size
        ipString(i) = "0.0.0." & i
    Next i
    myIpInfo.ipAddress = ipString
    myIpInfo.timeOut = 50
    myIpInfo.GetIpInfo
    newMtx = NetworkGeneralFunctions.FilterMtxOnWithOutString(myIpInfo.dataMatrix, "Request timed out", 3)
    myFunc = NetworkGeneralFunctions.OutputToSheet(newMtx, "Ip Table", , , True)
    
    'Add Ping Button
        PingButton
End Sub

Sub IpInfo_Example2()
    Dim myIpInfo As New IPINFO
    Debug.Print myIpInfo.Ping("0.0.0.0")
End Sub

Sub IpInfo_Example3()
    Dim myIpInfo As New IPINFO
    myIpInfo.PingMonitorWithGraphic ("Ip Table")
End Sub

Sub PingButton()
    ActiveSheet.Buttons.Add(690, 17.25, 138, 29.25).Select
    Selection.OnAction = "IpInfo_Example3"
    Selection.Characters.text = "Ping"
    With Selection.Characters(Start:=1, length:=4).Font
        .name = "Calibri"
        .FontStyle = "Regular"
        .size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
End Sub


