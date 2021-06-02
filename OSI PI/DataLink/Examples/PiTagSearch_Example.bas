Attribute VB_Name = "PiTagSearch_Example"
'Using PI SDK
Sub PiTagSearch_Example()

    Dim myPiTagSearch As New PITAGSEARCH
    myPiTagSearch.piTag = "*yyy*zzz*123*"
	myPiTagSearch.piServer = ""
	myPiTagSearch.printOutput = True
    myPiTagSearch.Get_PiTags

End Sub

