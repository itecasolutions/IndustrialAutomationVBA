VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written 04JUN2021 by Nicholas Stom
'Class is used to Get Data from PDF

Option Explicit
' These are properties


Function GetPdfAsText(ByVal strFilename As String) As String
    Dim objAVDoc As New AcroAVDoc
    Dim objPDDoc As New AcroPDDoc
    Dim objPage As AcroPDPage
    Dim objSelection As AcroPDTextSelect
    Dim objHighlight As AcroHiliteList
    Dim pageNum As Long
    Dim strText As String
    Dim tCount As Integer
    On Error GoTo ErrorHandle1
    strText = ""
    If (objAVDoc.Open(strFilename, "")) Then
       Set objPDDoc = objAVDoc.GetPDDoc
       For pageNum = 0 To objPDDoc.GetNumPages() - 1
          Set objPage = objPDDoc.AcquirePage(pageNum)
          Set objHighlight = New AcroHiliteList
          objHighlight.Add 0, 10000 ' Adjust this up if it's not getting all the text on the page
          Set objSelection = objPage.CreatePageHilite(objHighlight)
          If Not objSelection Is Nothing Then
             For tCount = 0 To objSelection.GetNumText - 1
                strText = strText & objSelection.GetText(tCount)
             Next tCount
          End If
       Next pageNum
       objAVDoc.Close 1
    End If
    
    GetPdfAsText = strText

ErrorHandle1:
    If strText = "" Then
        strText = "Error from in reading PDF from VBA PDF CLASS"
    End If
    GetPdfAsText = strText '    Exit Function
End Function

