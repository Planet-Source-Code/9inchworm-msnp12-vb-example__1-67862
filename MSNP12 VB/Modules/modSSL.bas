Attribute VB_Name = "modSSL"
Public SSL As WinHttp.WinHttpRequest

Public Function SetHTTPLib()
    Set SSL = Nothing
    Set SSL = New WinHttp.WinHttpRequest
    SSL.Option(WinHttpRequestOption_EnableRedirects) = False
End Function

Public Function SendRecvSSL(Method As String, Data As String, _
    Optional ReqHeaderN As String, Optional ReqHeaderD As String) As String
    SSL.Open Method, Data
    If ReqHeaderN <> "" And ReqHeaderD <> "" Then SSL.SetRequestHeader ReqHeaderN, ReqHeaderD
    SSL.Send
    SendRecvSSL = SSL.STATUS & " " & SSL.StatusText & vbCrLf & _
    SSL.GetAllResponseHeaders
End Function

Public Function pKey(AuthKey As String, User As String, Pass As String) As String
    Dim sData As String, sLoginServ As String, sHeader As String
    Call SetHTTPLib
    sHeader = "Passport1.4 OrgVerb=GET,OrgURL=http%3A%2F%2Fmessenger%2Emsn%2Ecom,sign-in=" & _
    Replace$(User, "@", "%40") & ",pwd=" & URLEncode(Pass) & "," & AuthKey
            
    sData = SendRecvSSL("GET", "https://nexus.passport.com/rdr/pprdr.asp")
    If GetBetween(sData, , vbCrLf) = "200 OK" Then
    sLoginServ = "https://" & GetBetween(sData, "DALogin=", ",")
        
ConnectionSSL:
        
        sData = SendRecvSSL("GET", sLoginServ, "Authorization", sHeader)
        
        Select Case GetBetween(sData, , vbCrLf)
            Case "302 Found"
                sLoginServ = GetBetween(sData, "Location: ", vbCrLf)
                GoTo ConnectionSSL
            Case "401 Unauthorized"
                MsgBox "Wrong username / password!": frmMain.sckNS.Close
            Case "200 OK"
                pKey = GetBetween(sData, "from-PP='", "'")
            Case Else
                MsgBox "Received unknown packet from SSL!": frmMain.sckNS.Close
        End Select
    Else
    MsgBox "Could not retrieve data from SSL!": frmMain.sckNS.Close
    End If
End Function




