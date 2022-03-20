VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GENERALFUNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function IsInArray(stringToBeFound As String, Arr As Variant) As Boolean
  IsInArray = (UBound(Filter(Arr, stringToBeFound)) > -1)
End Function


Function Compare2Sheets(dataSet1 As Variant, dataSet2 As Variant, colPos1 As Variant, colPos2 As Variant, rowStart1 As Variant, rowStart2 As Variant)
    If UBound(colPos1) = UBound(colPos2) Then
        For I = rowStart1 To UBound(dataSet1, 1)
            For j = rowStart2 To UBound(dataSet2, 1)
                matchFlag = 0
                For k = 0 To UBound(colPos1)
                   If dataSet1(I, colPos1(k)) = dataSet2(j, colPos2(k)) Then
                        matchFlag = matchFlag + 1
                   End If
                Next k
                If matchFlag = (UBound(colPos1) + 1) Then
                    Debug.Print ("Match found in Workbook 1 at row " & I & "  and Workbook 2 at row " & j)
                End If
                
            Next j
        Next I
    End If
End Function

Function SheetToDataSet(workbookName As Variant, worksheetName As Variant) As Variant
    'SET PAGE CHARACTERISTICS
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
    'DECLARE VARIABLE
        Dim x_matrix As Range
        Dim x_copyrange As String
        Dim length, lastColumn As Integer
    'DEFINE VARIABLE
        Workbooks(workbookName).Worksheets(worksheetName).Activate
        length = 0
        lastColumn = 0
        For I = 1 To 10
            If length < Workbooks(workbookName).Worksheets(worksheetName).Cells(Rows.Count, I).End(xlUp).Row Then
                length = Workbooks(workbookName).Worksheets(worksheetName).Cells(Rows.Count, I).End(xlUp).Row
            End If
            If lastColumn < Workbooks(workbookName).Worksheets(worksheetName).Cells(I, Columns.Count).End(xlToLeft).Column Then
                lastColumn = Workbooks(workbookName).Worksheets(worksheetName).Cells(I, Columns.Count).End(xlToLeft).Column + 10
            End If
        Next I
        'Let x_copyrange = .Range(.Cells(1, 1), .Cells(length, lastColumn))
    'Return
        SheetToDataSet = Workbooks(workbookName).Worksheets(worksheetName).Range(Cells(1, 1), Cells(length, lastColumn))
End Function
