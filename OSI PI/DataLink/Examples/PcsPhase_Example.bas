Attribute VB_Name = "PcsPhase_Example"
Sub PCSPhase_Example()

    Dim PHASE As New PHASE
    'Multiple Units
        PHASE.piTag = Array("", "")
    'Single Units
        'PHASE.piTag = ""
    PHASE.startTime = "8/1/2020  1:38:50 PM"
    PHASE.endTime = "10/2/2020  1:39:02 PM"
    PHASE.Get_PcsPhase

    Debug.Print 1
End Sub

Sub PCSPhase_Example1()

    Dim PHASE As New PHASE
    'Multiple Units
        'PHASE.piTag = Array("", "")
    'Single Units
        PHASE.piTag = ""
    PHASE.startTime = "12-Apr-21 12:56:36"
    PHASE.endTime = "1/1/2022"
    PHASE.unitParseChar = ":"
    PHASE.unitPosition = 2
    PHASE.GetPhase ("PHASE|")

    Debug.Print 1
End Sub
