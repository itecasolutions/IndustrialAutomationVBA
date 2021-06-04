VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EMAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written 04JUN2021 by Nicholas Stom
'Class is used to send Email

Option Explicit
' These are properties
    Public toAddress As String
    Public ccAddress As String
    Public bcAddress As String
    Public fromAddress As String
    Public subject As String
    Public textBody As String
    


Sub SmtpCdoSend(sendUsing As Integer, smtpServer As String, serverPort As Integer)
    Dim iMsg As Object
    Dim iConf As Object
    Dim strbody As String
    Dim Flds As Variant

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")

        iConf.Load -1
        Set Flds = iConf.Fields
        With Flds
            .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = sendUsing
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = serverPort
            .Update
        End With


    With iMsg
        Set .Configuration = iConf
        .To = toAddress
        .CC = ccAddress
        .BCC = bcAddress
        .From = fromAddress
        .subject = subject
        .textBody = textBody
        .Send
    End With

    Set iMsg = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
End Sub

