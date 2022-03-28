Attribute VB_Name = "Gantt_Example"
Sub CreateGenGantt()
    Dim genGantt As New GANTT
    'Dim myMtx As Variant
    Dim myMtx As Variant
    ReDim myMtx(10, 2)
    myMtx(0, 0) = "First Line"
    myMtx(0, 1) = "3/27/2022 2:42"
    myMtx(0, 2) = "3/27/2022 5:42"
    myMtx(1, 0) = "Second Line"
    myMtx(1, 1) = "3/27/2022 6:42"
    myMtx(1, 2) = "3/27/2022 10:42"
    genGantt.GanttChart10Day "mySheetName", myMtx
End Sub

