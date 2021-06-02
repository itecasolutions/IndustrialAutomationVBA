VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written 26MAY2021 by Nicholas Stom
'Class is used to get information of IP Addresses

Option Explicit
' These are properties
    Public ipAddress As Variant
    Public dataMatrix As Variant
    Public timeOut As Integer


Public Function GetIpInfo() As Variant
    Dim Pinger As Object
    Dim PingResult As Variant
    Dim tempMatrix As Variant
    Dim i As Integer
    Dim item As Variant
    ReDim dataMatrix(10, 0)
    If IsArray(ipAddress) = True Then
        For Each item In ipAddress
            Debug.Print item
            Set Pinger = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where TimeOut = " & timeOut & " and address = '" & item & "'")
            For Each PingResult In Pinger
                dataMatrix(0, UBound(dataMatrix, 2)) = item
                On Error Resume Next
                dataMatrix(1, UBound(dataMatrix, 2)) = PingResult.ProtocolAddress
                On Error Resume Next
                dataMatrix(3, UBound(dataMatrix, 2)) = "Script Error"
                dataMatrix(3, UBound(dataMatrix, 2)) = GetPingStatus(PingResult.StatusCode)
                On Error Resume Next
                If dataMatrix(3, UBound(dataMatrix, 2)) = "Connected" Then
                    dataMatrix(2, UBound(dataMatrix, 2)) = GetComputerName(PingResult.ProtocolAddress)
                End If
                On Error Resume Next
               
            Next
            ReDim Preserve dataMatrix(10, (UBound(dataMatrix, 2) + 1))
        Next
        ReDim Preserve dataMatrix(10, (UBound(dataMatrix, 2) - 1))
    Else
        Set Pinger = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where TimeOut = " & timeOut & " and address = '" & ipAddress & "'")
        For Each PingResult In Pinger
            dataMatrix(0, UBound(dataMatrix, 2)) = ipAddress
            On Error Resume Next
            dataMatrix(1, UBound(dataMatrix, 2)) = PingResult.ProtocolAddress
            On Error Resume Next
            dataMatrix(3, UBound(dataMatrix, 2)) = "Script Error"
            dataMatrix(3, UBound(dataMatrix, 2)) = GetPingStatus(PingResult.StatusCode)
            On Error Resume Next
            If dataMatrix(3, UBound(dataMatrix, 2)) = "Connected" Then
                dataMatrix(2, UBound(dataMatrix, 2)) = GetComputerName(PingResult.ProtocolAddress)
            End If
            On Error Resume Next
        Next
    End If
   
   GetIpInfo = dataMatrix
    
End Function

Private Function GetComputerName(currentIp As String) As String
    On Error GoTo ErrorHandler
    Dim colPings, oTempWk As Object, objPing As Object, strQuery, TmpTxt, myComputerName As String
    Dim LString() As String
    Dim vUserProf, vDir, TmpFile, vx As String
    vUserProf = Environ("USERPROFILE")
    vDir = vUserProf & "\Downloads\"
    TmpFile = vDir & "IndustrialAutomationTempFile.txt"
    Dim oFSO As Object
    Dim oShell As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oShell = CreateObject("Wscript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 0

     vx = "cmd.exe /c nslookup " & currentIp & " > " & TmpFile
    oShell.Run "cmd.exe /c nslookup " & currentIp & " > " & TmpFile, windowStyle, waitOnReturn
    TmpTxt = ""
    Set oTempWk = oFSO.OpenTextFile(TmpFile, 1)
    Do While oTempWk.AtEndOfStream <> True
        vx = Trim(oTempWk.ReadLine)
        If InStr(vx, "Name:    ") > 0 Then
             LString = Split(vx, "Name:    ")
             myComputerName = LString(1)
             Exit Do
        End If
    Loop
    oTempWk.Close
    oFSO.DeleteFile (TmpFile)
    GetComputerName = myComputerName
Exit Function
ErrorHandler:
    GetComputerName = "Not Found"
    Exit Function
End Function



Private Function GetPingStatus(StatCd As String)
    Dim strResult As String
    Select Case StatCd
    Case 0: strResult = "Connected"
    Case 11001: strResult = "Buffer too small"
    Case 11002: strResult = "Destination net unreachable"
    Case 11003: strResult = "Destination host unreachable"
    Case 11004: strResult = "Destination protocol unreachable"
    Case 11005: strResult = "Destination port unreachable"
    Case 11006: strResult = "No resources"
    Case 11007: strResult = "Bad option"
    Case 11008: strResult = "Hardware error"
    Case 11009: strResult = "Packet too big"
    Case 11010: strResult = "Request timed out"
    Case 11011: strResult = "Bad request"
    Case 11012: strResult = "Bad route"
    Case 11013: strResult = "Time-To-Live (TTL) expired transit"
    Case 11014: strResult = "Time-To-Live (TTL) expired reassembly"
    Case 11015: strResult = "Parameter problem"
    Case 11016: strResult = "Source quench"
    Case 11017: strResult = "Option too big"
    Case 11018: strResult = "Bad destination"
    Case 11032: strResult = "Negotiating IPSEC"
    Case 11050: strResult = "General failure"
    Case Else: strResult = "Unknown host"
    End Select
    GetPingStatus = strResult
End Function

Function Ping1(ip)
    Dim objShell, boolCode
    Set objShell = CreateObject("Wscript.Shell")
    boolCode = objShell.Run("ping -n 1 -w 1000 " & ip, 0, True)
    If boolCode = 0 Then
        Ping = True
    Else
        Ping = False
    End If
End Function

Function Ping(ip)
    Dim Pinger As Object
    Dim PingResult As Variant
    Set Pinger = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where TimeOut = 500 and address = '" & ip & "'")
    For Each PingResult In Pinger
        If GetPingStatus(PingResult.StatusCode) = "Connected" Then
            Ping = True
            Exit For
        Else
            Ping = False
            Exit For
        End If
    Next
End Function


Sub PingMonitorWithGraphic(wrkSheetName As String)
    Dim ip As String
    Dim i, ipCol, conCol As Integer
    ipCol = 2
    conCol = 4
    For i = 1 To Worksheets(wrkSheetName).Cells(Rows.Count, 1).End(xlUp).Row
        If Worksheets(wrkSheetName).Cells(i, ipCol) <> "" Then
            ip = Worksheets(wrkSheetName).Cells(i, ipCol)
            If Ping(ip) = True Then
                GrapicPingOnline wrkSheetName, i, 4
            Else
                 If Ping(ip) = True Then
                    GrapicPingOnline wrkSheetName, i, 4
                Else
                     If Ping(ip) = True Then
                       GrapicPingOnline wrkSheetName, i, 4
                    Else
                        If Ping(ip) = True Then
                            GrapicPingOnline wrkSheetName, i, 4
                        Else
                           If Ping(ip) = True Then
                                GrapicPingOnline wrkSheetName, i, 4
                            Else
                                GrapicPingOffline wrkSheetName, i, 4
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

Private Function GrapicPingOnline(wrkSheetName, i, j)
    Worksheets(wrkSheetName).Cells(i, j).Interior.ColorIndex = 0
    Worksheets(wrkSheetName).Cells(i, j).Font.Color = RGB(0, 0, 0)
    Worksheets(wrkSheetName).Cells(i, j).Value = "Online"
    Worksheets(wrkSheetName).Cells(i, j).Font.Color = RGB(0, 200, 0)
End Function
Private Function GrapicPingOffline(wrkSheetName, i, j)
    Worksheets(wrkSheetName).Cells(i, j).Interior.ColorIndex = 0
    Worksheets(wrkSheetName).Cells(i, j).Font.Color = RGB(200, 0, 0)
    Worksheets(wrkSheetName).Cells(i, j).Value = "No Response after 5 attempts"
    Worksheets(wrkSheetName).Cells(i, j).Interior.ColorIndex = 6
End Function
