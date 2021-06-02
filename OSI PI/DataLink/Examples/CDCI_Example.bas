Attribute VB_Name = "CDCI_Example"
Sub CDCI_Example1()

    Dim myData As New CDCI
    myDataSet = General_Functions.SheetToDataSet("1130_CIP")
    myData.dataMatrix = myDataSet
    myData.negSlope = -0.7
    myData.posSlope = 1
    myData.zeroSlope = 0.15
    myData.step = 3
    newDataSet = myData.IdentifyChange
    myFunc = General_Functions.OutputToSheet1(newDataSet, "testing", "", , True)
    myData.CreatGraph ("testing")
    'Debug.Print 1
End Sub

Sub CDCI_Example2()

    Dim myData As New CDCI
    myDataSet = General_Functions.SheetToDataSet("")
    myData.dataMatrix = myDataSet
    myData.negSlope = -0.7
    myData.posSlope = 1
    myData.zeroSlope = 0.15
    myData.step = 3
    newDataSet = myData.IdentifyChange
    myFunc = General_Functions.OutputToSheet1(newDataSet, "testing1", "", , True)
    myData.CreatGraph ("testing1")
    Debug.Print 1
End Sub

Sub CDCI_Example3()

    Dim myData As New CDCI
    myDataSet = General_Functions.SheetToDataSet("Sheet14")
    myData.dataMatrix = myDataSet
    myData.negSlope = -1
    myData.posSlope = 1
    myData.zeroSlope = 0.15
    myData.step = 3
    newDataSet = myData.IdentifyChange
    myFunc = General_Functions.OutputToSheet1(newDataSet, "testing3", "", , True)
    myData.CreatGraph ("testing3")
End Sub


Sub CDCI_Example11()

    Dim myData As New CDCI
    myDataSet = General_Functions.SheetToDataSet("")
    myData.dataMatrix = myDataSet
    myData.negSlope = -0.7
    myData.posSlope = 1
    myData.zeroSlope = 0.15
    myData.step = 3
    newDataSet = myData.IdentifyChange
    myFunc = General_Functions.OutputToSheet1(newDataSet, "testing", "", , True)
    'myData.CreatGraph ("testing")
    'Debug.Print 1
End Sub

Sub CDCI_Combine_Example1()

    Dim myData As New CDCI
    myDataSet = General_Functions.SheetToDataSet("testing")
    myData.dataMatrix = myDataSet
    myDataSet = General_Functions.SheetToDataSet("testing1")
    myData.dataMatrix1 = myDataSet
    newDataSet = myData.CombinedData
    myFunc = General_Functions.OutputToSheet1(newDataSet, "testing2", "", , True)
    myData.CombinedGraph ("testing2")
End Sub
