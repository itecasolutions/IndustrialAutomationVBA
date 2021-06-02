Attribute VB_Name = "PcsEvent_Example"
Sub PcsEvent_Example()

    Dim EVENTDATA As New EVENTDATA
    'Multiple Units
        'EVENTDATA.piTag = Array("", "")
    'Single Units
        EVENTDATA.piTag = ""
    EVENTDATA.startTime = "9/9/2020  1:38:50 PM"
    EVENTDATA.endTime = "10/2/2020  1:39:02 PM"
    EVENTDATA.Get_Event

End Sub

