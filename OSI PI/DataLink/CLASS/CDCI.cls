VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDCI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written 22MARCH2021 by Nicholas Stom
'Continuous Data Change Identifier
'Used for three positions, Increase, Decrease, Flat
'Uses Slope to create Events
'Array needs to be 6 wide, columns are Time (including seconds), PV, and Event(created by class)



Option Explicit
' These are properties
    Public dataMatrix As Variant
    Public dataMatrix1 As Variant
    Public posSlope As Double
    Public negSlope As Double
    Public zeroSlope As Double
    Public step As Integer




Public Function IdentifyChange() As Variant
    Dim currentSlopeAngle, previousSlopeAngle As Integer
    Dim i As Integer
    Dim j As Integer
    Dim negZeroSlope As Double
    Dim eventOut As Integer
    Dim slope, rise, slRun  As Double
    currentSlopeAngle = -2
    previousSlopeAngle = -2
    negZeroSlope = -1 * zeroSlope
    eventOut = 3
    For i = 1 To UBound(dataMatrix, 1)
        'Exit For loop if at end
        If (i + step) > UBound(dataMatrix, 1) Then
            Exit For
        End If
        j = (i + step)
        rise = dataMatrix(j, 2) - dataMatrix(i, 2)
        slRun = (dataMatrix(j, 1) - dataMatrix(i, 1)) * 86400

        slope = rise / slRun
        'Set current Slope
        If slope >= posSlope Then
            currentSlopeAngle = 1
        ElseIf slope <= negSlope Then
            currentSlopeAngle = -1
        ElseIf slope < zeroSlope And slope > (negZeroSlope) Then
            currentSlopeAngle = 0
        End If
        
        'Update current Slope
        If currentSlopeAngle <> previousSlopeAngle Then
            dataMatrix(i, 3) = dataMatrix(i, 2)
'            dataMatrix(i, 4) = currentSlopeAngle
'            dataMatrix(i, 5) = slope
'            dataMatrix(i, 6) = dataMatrix(i, 2)
'            dataMatrix(i, 7) = dataMatrix(j, 2)
            previousSlopeAngle = currentSlopeAngle
        End If
        
        
    Next i
    IdentifyChange = dataMatrix
End Function

Public Function CombinedData() As Variant
    Dim i As Integer
    Dim j As Integer
    Dim x As Integer
    Dim z As Integer
    Dim startTime As Date
    Dim endTime As Date
    Dim timeStep As Double
    j = 1
    'Event Fill
    For i = 1 To UBound(dataMatrix, 1)
        If dataMatrix(i, 3) <> "" Then
            Do While j < UBound(dataMatrix1, 1)
                If dataMatrix1(j, 3) <> "" Then
                    dataMatrix(i, 4) = dataMatrix1(j, 1)
                    dataMatrix(i, 5) = dataMatrix1(j, 2)
                    j = j + 1
                    Exit Do
                Else
                    j = j + 1
                End If
            Loop
        End If
    Next i
    
    'Time BackFill
     For i = 1 To UBound(dataMatrix, 1)
        If (i + 1) < UBound(dataMatrix, 1) Then
            If dataMatrix(i, 4) <> "" And dataMatrix(i + 1, 4) = "" Then
                startTime = dataMatrix(i, 4)
                j = i + 1
                Do While j < UBound(dataMatrix, 1)
                    If dataMatrix(j, 4) <> "" Then
                        endTime = dataMatrix(j, 4)
                        Debug.Print startTime & " " & endTime & " " & j
                        Debug.Print endTime - startTime
                        timeStep = (endTime - startTime) / (j - i)
                        x = i + 1
                        z = j
                        For x = (i + 1) To z
                            dataMatrix(x, 4) = dataMatrix(x - 1, 4) + timeStep
                        Next x
                        j = j + 1
                        Exit Do
                    Else
                        j = j + 1
                    End If
                Loop
            End If
        End If
    Next i
    
    'Data BackFill
     For i = 1 To UBound(dataMatrix, 1)
        If dataMatrix(i, 4) <> "" And dataMatrix(i, 5) = "" Then
            For j = 2 To UBound(dataMatrix1, 1)
                If dataMatrix(i, 4) > dataMatrix1(j - 1, 1) And dataMatrix(i, 4) < dataMatrix1(j + 1, 1) Then
                    dataMatrix(i, 5) = dataMatrix1(j, 2)
                    Exit For
                End If
            Next j
        End If
    Next i
    
    CombinedData = dataMatrix
End Function

Sub CreatGraph(worksheetName As String)
    Dim stringA, stringB, stringC As String
    stringA = "=" & worksheetName & "!$A:$A"
    stringB = "=" & worksheetName & "!$B:$B"
    stringC = "=" & worksheetName & "!$C:$C"
    Worksheets(worksheetName).Activate
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).XValues = stringA
    ActiveChart.FullSeriesCollection(1).Values = stringC
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 6
    End With
    ActiveChart.PlotArea.Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).XValues = stringA
    ActiveChart.FullSeriesCollection(2).Values = stringB
    ActiveChart.FullSeriesCollection(2).Select
    ActiveChart.FullSeriesCollection(2).ChartType = xlXYScatterSmoothNoMarkers
     With ActiveChart.Parent
         .Height = 325 ' resize
         .Width = 2000  ' resize
         .Top = 100    ' reposition
         .Left = 100   ' reposition
     End With
End Sub


Sub CombinedGraph(worksheetName As String)
    Dim stringA, stringB, stringE As String
    stringA = "=" & worksheetName & "!$A:$A"
    stringB = "=" & worksheetName & "!$B:$B"
    stringE = "=" & worksheetName & "!$E:$E"
    Worksheets(worksheetName).Activate

    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Select
    Do Until ActiveChart.SeriesCollection.Count = 0
        ActiveChart.SeriesCollection(1).Delete
    Loop
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).XValues = stringA
    ActiveChart.FullSeriesCollection(1).Values = stringB
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).XValues = stringA
    ActiveChart.FullSeriesCollection(2).Values = stringE

     With ActiveChart.Parent
         .Height = 325 ' resize
         .Width = 2000  ' resize
         .Top = 100    ' reposition
         .Left = 100   ' reposition
     End With
End Sub
