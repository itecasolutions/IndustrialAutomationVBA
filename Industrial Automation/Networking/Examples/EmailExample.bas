Attribute VB_Name = "EmailExample"
Sub EmailExampleSMTP()
    Dim myEmail As New EMAIL
    myEmail.toAddress = ""
    myEmail.subject = ""
    myEmail.fromAddress = ""
    myEmail.textBody = ""
    myEmail.SmtpCdoSend 2, "smtp mail host . com", 25
End Sub


