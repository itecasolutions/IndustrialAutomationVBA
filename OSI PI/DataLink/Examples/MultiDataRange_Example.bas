Attribute VB_Name = "MultiDataRange_Example"
Sub MultiDataRange_Example()

    Dim MULTIRANGE As New MULTIRANGE
    MULTIRANGE.piTag = Array("pi tag 1", "pi tag 2")
    MULTIRANGE.startTime = "1/1/2020  1:38:50 PM"
    MULTIRANGE.endTime = "1/1/2020  1:40:02 PM"
    MULTIRANGE.sampleTime = "00:00:30"
    MULTIRANGE.Get_MultiRange

    Debug.Print 1

End Sub

