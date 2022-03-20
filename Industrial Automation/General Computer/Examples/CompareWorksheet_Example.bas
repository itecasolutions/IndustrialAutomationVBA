Attribute VB_Name = "CompareWorksheet_Example"
Sub CompareWorksheet()
    Dim dataSet1, dataSet2 As Variant
    Dim workbook1, workbook2 As String
    Dim worksheet1, worksheet2 As String
    Dim rowStart1, rowStart2 As Integer
    Dim GeneralFunc As New GeneralFunc
    
    'Get the data into the dataSet variable using a function that goes through each workbook/sheet
    workbook1 = "dashboard-data-latest1.xlsx"
    worksheet1 = "2. Harmonized Indicators"
    dataSet1 = GeneralFunc.SheetToDataSet(workbook1, worksheet1)
    
    'Get the data into the dataSet variable using a function that goes through each workbook/sheet
    workbook2 = "dashboard-data-latest.xlsx"
    worksheet2 = "2. Harmonized Indicators"
    dataSet2 = GeneralFunc.SheetToDataSet(workbook2, worksheet2)
    
    'Set this do what columns you are interested in comparing
    colPos1 = Array(1, 2, 3)
    colPos2 = Array(1, 2, 3)
    
    'Set for where you want to start 1 would be row 1/now Header.
    rowStart1 = 2
    rowStart2 = 2
    
    'Compares the dataSets
     GeneralFunc.Compare2Sheets dataSet1, dataSet2, colPos1, colPos2, rowStart1, rowStart2


End Sub


