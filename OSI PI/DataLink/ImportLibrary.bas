Attribute VB_Name = "ImportLibrary"
'Remember to add reference to Microsoft Scripting Runtime
'Remember to trust the VBA Project
Sub ImportLibrary()
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim splitString() As String
    Dim baseFolder As String
    'Configure baseFolder
    baseFolder = ""
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    folder = baseFolder & "\CLASS\"
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
    
    folder = baseFolder & "\Examples\"
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
End Sub
