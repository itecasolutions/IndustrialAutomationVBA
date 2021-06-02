Attribute VB_Name = "DeviceData_Example"
Sub DeviceData_Example()

    Dim DEVICEDATA As New DEVICEDATA
    DEVICEDATA.piTag = ""
    DEVICEDATA.startTime = "1/1/2020  1:38:50 PM"
    DEVICEDATA.endTime = "2/2/2020  1:39:02 PM"
    DEVICEDATA.Get_DeviceData

    Debug.Print 1

End Sub


