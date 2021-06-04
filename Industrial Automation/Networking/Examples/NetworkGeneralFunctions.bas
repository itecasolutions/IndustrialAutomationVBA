Attribute VB_Name = "NetworkGeneralFunctions"
'Covington  Library
'Written 29OCT2020 by Nicholas Stom
'General Functions for simplificaiton of Main Routines


Function SheetToDataSet(worksheetName As String) As Variant
    'SET PAGE CHARACTERISTICS
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
    'DECLARE VARIABLE
        Dim x_matrix As Range
        Dim x_copyrange As String
        Dim length, lastColumn As Integer
    'DEFINE VARIABLE
    Worksheets(worksheetName).Activate
        length = 0
        lastColumn = 0
        For I = 1 To 10
            If length < Worksheets(worksheetName).Cells(Rows.Count, I).End(xlUp).Row Then
                length = Worksheets(worksheetName).Cells(Rows.Count, I).End(xlUp).Row
            End If
            If lastColumn < Worksheets(worksheetName).Cells(I, Columns.Count).End(xlToLeft).Column Then
                lastColumn = Worksheets(worksheetName).Cells(I, Columns.Count).End(xlToLeft).Column + 10
            End If
        Next I
        'Let x_copyrange = .Range(.Cells(1, 1), .Cells(length, lastColumn))
    'Return
        SheetToDataSet = Worksheets(worksheetName).Range(Cells(1, 1), Cells(length, lastColumn))
End Function


Function OutputToSheet(dataMatrix As Variant, worksheetName As String, Optional containsString As String, Optional containCol As Integer, Optional dltSheetData As Boolean = False)
    Dim I, sheet_exists As Integer
    sheet_exists = 0
    For I = 1 To Sheets.Count
       If Sheets(I).Visible = -1 Then
           If Sheets(I).name = worksheetName Then
                sheet_exists = 1
           End If
       End If
    Next
    If sheet_exists = 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).name = worksheetName
    End If
    
    'Delete Data on Sheet
        If dltSheetData = True Then
            Worksheets(worksheetName).Activate
            Worksheets(worksheetName).Cells.Select
            Selection.Delete Shift:=xlUp
        End If
    
    
        Set x_matrix = Worksheets(worksheetName).Range("A1:A2")
        For j = 0 To UBound(dataMatrix, 2)
            currentRow = Worksheets(worksheetName).Cells(Rows.Count, 2).End(xlUp).Row + 1
            For I = 0 To UBound(dataMatrix, 1)
                If IsMissing(containsString) Or containsString = "" Then
                    x_matrix(j + 1, I + 1) = dataMatrix(I, j)
                Else
                   If InStr(dataMatrix(containCol, j), containsString) > 0 Then
                        x_matrix(currentRow, I + 1) = dataMatrix(I, j)
                   End If
                End If
            Next I
        Next j
End Function

Function OutputToSheet_1(dataMatrix As Variant, worksheetName As String, Optional containsString As String, Optional containCol As Integer, Optional dltSheetData As Boolean = False)
    Dim I, sheet_exists As Integer
    sheet_exists = 0
    For I = 1 To Sheets.Count
       If Sheets(I).Visible = -1 Then
           If Sheets(I).name = worksheetName Then
                sheet_exists = 1
           End If
       End If
    Next
    If sheet_exists = 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).name = worksheetName
    End If
    
    'Delete Data on Sheet
        If dltSheetData = True Then
            Worksheets(worksheetName).Activate
            Worksheets(worksheetName).Cells.Select
            Selection.Delete Shift:=xlUp
        End If
    
    
        Set x_matrix = Worksheets(worksheetName).Range("A1:A2")
        For j = 1 To UBound(dataMatrix, 2)
            currentRow = Worksheets(worksheetName).Cells(Rows.Count, 2).End(xlUp).Row + 1
            For I = 1 To UBound(dataMatrix, 1)
                If IsMissing(containsString) Or containsString = "" Then
                    x_matrix(j + 1, I + 1) = dataMatrix(I, j)
                Else
                   If InStr(dataMatrix(containCol, j), containsString) > 0 Then
                        x_matrix(currentRow, I + 1) = dataMatrix(I, j)
                   End If
                End If
            Next I
        Next j
End Function

Function OutputToSheet1(dataMatrix As Variant, worksheetName As String, Optional containsString As String, Optional containCol As Integer, Optional dltSheetData As Boolean = False)
    Dim I, sheet_exists As Integer
    sheet_exists = 0
    For I = 1 To Sheets.Count
       If Sheets(I).Visible = -1 Then
           If Sheets(I).name = worksheetName Then
                sheet_exists = 1
           End If
       End If
    Next
    If sheet_exists = 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).name = worksheetName
    End If
    
    'Delete Data on Sheet
        If dltSheetData = True Then
            Worksheets(worksheetName).Activate
            Worksheets(worksheetName).Cells.Select
            Selection.Delete Shift:=xlUp
        End If
    
    
        Set x_matrix = Worksheets(worksheetName).Range("A1:A2")
        For j = 0 To UBound(dataMatrix, 2)
            For I = 0 To UBound(dataMatrix, 1)
                x_matrix(I + 1, j + 1) = dataMatrix(I, j)
            Next I
        Next j
End Function

Function DCNVWithString(dataMatrix As Variant, arrLeng As Integer, Optional containsString As String, Optional containCol As Integer, Optional addRow As Integer = 0) As Variant
    Dim I, j As Integer
    Dim my_matrix As Variant
    Dim firstPass As Boolean
    Dim LString As String
    Dim lArray() As String
    Debug.Print UBound(dataMatrix, 2)
    Debug.Print UBound(dataMatrix, 1)
    'DEFINE VARIABLES
    ReDim my_matrix(arrLeng, 0)
    For j = 1 To UBound(dataMatrix, 1)
        If InStr(dataMatrix(j, containCol), containsString) > 0 Then
             Debug.Print dataMatrix(j, 2)
             If addRow > 0 Then
                addRowIndex = j + addRow
                If addRowIndex <= UBound(dataMatrix, 1) Then
                    If firstPass = False Then
                        firstPass = True
                        newArrayWidth = 0
                    Else
                        newArrayWidth = UBound(my_matrix, 2) + 1
                    End If
                    ReDim Preserve my_matrix(arrLeng, newArrayWidth)
                    my_matrix(0, newArrayWidth) = dataMatrix(j, 2)
                    my_matrix(1, newArrayWidth) = dataMatrix(j, 1)
                    my_matrix(2, newArrayWidth) = dataMatrix(addRowIndex, 1)
                Else
                
                End If
             End If
        End If
    Next j
    DCNVWithString = my_matrix
End Function



Public Function AddHours(ByVal sTime As String, hourAmount As Integer) As String
    Dim dt As Date

    dt = CDate(sTime)
    dt = DateAdd("h", hourAmount, dt)

    AddHours = Format(dt, "mm/dd/yyyy hh:mm:ss")

End Function

Public Function FilterMtxOnString(dataMatrix As Variant, containsString As String, containCol As Integer) As Variant
     Dim I, j, z As Integer
    Dim my_matrix As Variant
    Dim firstPass As Boolean
    Dim LString As String
    Dim lArray() As String
    'DEFINE VARIABLES
    ReDim my_matrix(UBound(dataMatrix, 1), 0)
    z = -1

    For j = 0 To UBound(dataMatrix, 2)
        If InStr(dataMatrix(containCol, j), containsString) > 0 Then
            z = z + 1
            ReDim Preserve my_matrix(UBound(dataMatrix, 1), z)
            For I = 0 To UBound(dataMatrix, 1)
                my_matrix(I, z) = dataMatrix(I, j)
            Next I
        End If
    Next j
    FilterMtxOnString = my_matrix

End Function

Public Function FilterMtxOnWithOutString(dataMatrix As Variant, containsString As String, containCol As Integer) As Variant
     Dim I, j, z As Integer
    Dim my_matrix As Variant
    Dim firstPass As Boolean
    Dim LString As String
    Dim lArray() As String
    'DEFINE VARIABLES
    ReDim my_matrix(UBound(dataMatrix, 1), 0)
    z = -1

    For j = 0 To UBound(dataMatrix, 2)
        If InStr(dataMatrix(containCol, j), containsString) = 0 Then
            z = z + 1
            ReDim Preserve my_matrix(UBound(dataMatrix, 1), z)
            For I = 0 To UBound(dataMatrix, 1)
                my_matrix(I, z) = dataMatrix(I, j)
            Next I
        End If
    Next j
    FilterMtxOnWithOutString = my_matrix

End Function
