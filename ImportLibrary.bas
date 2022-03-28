Attribute VB_Name = "ImportLibrary"
'Remember to add reference to Microsoft Scripting Runtime
'Remember to trust the VBA Project
Sub ImportLibrary()
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim splitString() As String
    Dim base As String
    Dim baseFolder As Variant
    'Configure baseFolder
	base = ""
    baseFolder = Array("\Industrial Automation\OPC", "\OSI PI\DataLink", "\Industrial Automation\General Computer", "\Industrial Automation\Networking", "\Industrial Automation\EXCEL")
	
	For Each element In baseFolder
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		folder = base & element & "\CLASS\"
		Set oFolder = oFSO.GetFolder(folder)
		For Each oFile In oFolder.Files
			If InStr(oFile.name, ".cls") > 0 Then
				splitString = Split(oFile.name, ".cls")
			ElseIf InStr(oFile.name, ".bas") > 0 Then
				splitString = Split(oFile.name, ".bas")
			End If

		   On Error Resume Next
			ActiveWorkbook.VBProject.VBComponents.Remove ActiveWorkbook.VBProject.VBComponents(splitString(0))
			On Error Resume Next
			ActiveWorkbook.VBProject.VBComponents.Import folder & oFile.name
			On Error Resume Next
		Next oFile
		
		folder = base & element & "\Examples\"
		Set oFolder = oFSO.GetFolder(folder)
		For Each oFile In oFolder.Files
			If InStr(oFile.name, ".cls") > 0 Then
				splitString = Split(oFile.name, ".cls")
			ElseIf InStr(oFile.name, ".bas") > 0 Then
				splitString = Split(oFile.name, ".bas")
			End If

			On Error Resume Next
			ActiveWorkbook.VBProject.VBComponents.Remove ActiveWorkbook.VBProject.VBComponents(splitString(0))
			On Error Resume Next
			ActiveWorkbook.VBProject.VBComponents.Import folder & oFile.name
			On Error Resume Next
		Next oFile
	Next element
End Sub
