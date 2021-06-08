VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written 04JUN2021 by Nicholas Stom
'Class is used to Get Data from Text

Option Explicit
' These are properties


Function GetTextAsArray(ByVal strFilename As String) As Variant
	Dim FSO As Object, MyFile As Object
	Dim FileName As String, Arr As Variant

	FileName = strFilename
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set MyFile = FSO.OpenTextFile(FileName, 1)
	Arr = Split(MyFile.ReadAll, vbNewLine) 
	GetTextAsArray = Arr

End Function

Function GetTextAsText(ByVal strFilename As String) As String
	Dim strFilename As String: strFilename = strFilename
    Dim strFileContent As String
    Dim iFile As Integer: iFile = FreeFile
    Open strFilename For Input As #iFile
    strFileContent = Input(LOF(iFile), iFile)
    Close #iFile
	GetTextAsText = strFileContent
End Function

