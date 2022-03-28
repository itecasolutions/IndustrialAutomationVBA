VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GANTT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub GanttChart10Day(worksheetName As String, dataMtx As Variant)
   'SET PAGE CHARACTERISTICS
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False

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

    Worksheets(worksheetName).Activate
    'Delete Data
        Cells.Select
        Selection.Delete Shift:=xlUp
    'No Grid Line
        ActiveWindow.DisplayGridlines = False
'    'Create Template
        Range("A5").Select
        ActiveCell.FormulaR1C1 = "Name"
        Range("B5").Select
        ActiveCell.FormulaR1C1 = "Start Time"
        Range("C5").Select
        ActiveCell.FormulaR1C1 = "End Time"


    Dim col, col1, col2, col3, timeInt As Integer
    col = 7
    col1 = 66
    col2 = 7
    col3 = 1446
    For x = 1 To 10
        ActiveSheet.Range(Cells(1, col2), Cells(1, col3)).Select
        Selection.Merge

        'Generate Border
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        col2 = col2 + 1440
        col3 = col3 + 1440
        timeInt = 1
        For I = 1 To 24
            Range(Columns(col), Columns(col1)).Select
            Selection.ColumnWidth = 0.1
            ActiveSheet.Range(Cells(2, col), Cells(5, col1)).Select
            Selection.Merge
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With


            'Insert Time
                ActiveCell.FormulaR1C1 = timeInt & ":00"
                With Selection
                    .HorizontalAlignment = xlGeneral
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 90
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = True
                End With
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 90
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = True
                End With
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = False
                    .Orientation = 90
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = True
                End With
            timeInt = timeInt + 1
            If timeInt = 24 Then
                timeInt = 0
            End If
            col = col + 60
            col1 = col1 + 60
        Next I
    Next x

    For I = 6 To 30
        Rows(I).Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        Selection.Borders(xlEdgeLeft).LineStyle = xlNone
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Next I

    Range("H6:I6").Select
    ActiveWindow.FreezePanes = True


    'Grid
        Range("A1:F30").Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        
    FillGanttChart10Day worksheetName, dataMtx
End Sub


Private Sub FillGanttChart10Day(worksheetName As String, dataMatrix As Variant)
   'SET PAGE CHARACTERISTICS
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        Dim x_matrix, y_matrix, z_matrix As Range
        Dim x_copyrange, y_copyrange, z_copyrange, time_string As String
        Dim x_step, y_step, z_step, x_fnl_row, y_fnl_row, z_fnl_row, time_1, time_2, time_3 As Integer
        
    Dim startPos1 As Integer
    startPos1 = 6
    Set x_matrix = Worksheets(worksheetName).Range("A1:A2")
    For j = 0 To UBound(dataMatrix, 1)
        x_matrix(startPos1 + j, 1) = dataMatrix(j, 0)
        x_matrix(startPos1 + j, 2) = dataMatrix(j, 1)
        x_matrix(startPos1 + j, 3) = dataMatrix(j, 2)
    Next j
    
    
    y_fnl_row = Worksheets(worksheetName).Cells(Rows.Count, 1).End(xlUp).Row
    Let y_copyrange = "a" & 1 & ":" & "p" & y_fnl_row
    Set y_matrix = Worksheets(worksheetName).Range(y_copyrange)
    
    time_1 = Int(Worksheets(worksheetName).Cells(6, 2))
    
    y_step = 6
    Do While y_step <= y_fnl_row
        time_2 = Round(((y_matrix(y_step, 2) - time_1) * 1440) + 4, 0)
        time_3 = Round(((y_matrix(y_step, 3) - time_1) * 1440) + 4, 0)
        Worksheets(worksheetName).Range(Cells(y_step, time_2), Cells(y_step, time_3)).Interior.ColorIndex = 33
        y_step = y_step + 1
    Loop
    
    Columns("A:C").Select
    Columns("A:C").EntireColumn.AutoFit
    
        
End Sub
