Attribute VB_Name = "HttpRequest_Example"
Sub HttpRequest_Example()
    Dim httpReq As New HTTPREQUEST
    httpReq.url = "http://10.0.0.114:5000"
    httpReq.reqMeth = "POST"
    httpReq.headerDict.Add "Content-Type", "application/x-www-form-urlencoded"
    httpReq.postBodyDict.Add "postBodyParameter1", 1
    httpResponse = httpReq.XmlHttp
    Debug.Print httpResponse
End Sub

