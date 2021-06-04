Attribute VB_Name = "PdfExample"
Sub GetPdfText()
    Dim myPdfText As String
    Dim myPdf As New PDF
    myPdfText = myPdf.GetPdfAsText("")
    Debug.Print myPdfText
End Sub

