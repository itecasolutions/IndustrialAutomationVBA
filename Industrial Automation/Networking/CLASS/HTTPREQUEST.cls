VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HTTPREQUEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written 24MARCH2022 by Nicholas Stom

'List of HTTP Header Fields
    'https://en.wikipedia.org/wiki/List_of_HTTP_header_fields


Option Explicit
' These are properties
    Public url As String
    Public reqMeth As String
    Public headerDict As New Scripting.Dictionary
    Public postBodyDict As New Scripting.Dictionary
    
    

' These are methods
    Function XmlHttp() As String
        Dim objHTTP As Object, myUrl As String, postBodyString As String, n As Integer
        Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
        myUrl = url
        postBodyString = ""
        objHTTP.Open reqMeth, myUrl, False
        For n = 0 To headerDict.Count - 1
            objHTTP.setRequestHeader headerDict.Keys(n), headerDict.Items(n)
        Next n
        For n = 0 To headerDict.Count - 1
            postBodyString = postBodyString & postBodyDict.Keys(n) & "=" & postBodyDict.Items(n) & "&"
        Next n
        objHTTP.Send (postBodyString)
        XmlHttp = objHTTP.responseText
    End Function
