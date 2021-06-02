Attribute VB_Name = "GanttChart_Example"
Sub CreateDeltaVGantt()
    Dim myGantt As New GANTTCHART
    myGantt.FullDeltaVTemplate "TEstindg"
End Sub

Sub CreateDeltaVGantt1()
    Dim myGantt As New GANTTCHART
    myGantt.DeltaVPhaseTemplate "TEstindg"
End Sub

