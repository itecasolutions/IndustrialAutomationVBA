Attribute VB_Name = "PiVision_Example"
Sub PiVision_Example()

    Dim goToPiVision As New goToPiVision
    'Multiple Units
        goToPiVision.piTag = Array("", "")
    goToPiVision.startTime = "8/1/2020  1:38:50 PM"
    goToPiVision.endTime = "10/2/2020  1:39:02 PM"
    goToPiVision.browserPath = """C:\Windows\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\MicrosoftEdge.exe"""
    goToPiVision.browserPath = """C:\Program Files\Google\Chrome\Application\chrome.exe"""
    goToPiVision.piVisionUrl = ""
    goToPiVision.OpenPiVision
End Sub


