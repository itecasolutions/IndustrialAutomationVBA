Attribute VB_Name = "DataRange_Example"
Sub DataRange_Example()

    Dim DATARANGE As New DATARANGE
    DATARANGE.piTag = ""
    DATARANGE.startTime = "1/1/2020  1:38:50 PM"
    DATARANGE.endTime = "1/1/2020  1:39:02 PM"
    DATARANGE.Get_DataRange

    Debug.Print 1

End Sub

