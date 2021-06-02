Attribute VB_Name = "PcsBatchId_Example"
Sub BatchID_Example1()

    Dim PCSBATCHID As New PCSBATCHID
    PCSBATCHID.piTag = ""
    PCSBATCHID.startTime = "1/1/2020  1:38:50 PM"
    PCSBATCHID.endTime = "1/1/2021  1:39:02 PM"
    PCSBATCHID.Get_PcsBatchId

    Debug.Print 1

End Sub

Sub BatchID_Example2()

    Dim myBatchId As New PCSBATCHID
    myBatchId.piTag = ""
    myBatchId.startTime = "1/1/2020  1:38:50 PM"
    myBatchId.endTime = "1/1/2022  1:39:02 PM"
    myBatchId.GetBatchIdNoSplit


End Sub

