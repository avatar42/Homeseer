' import BlueIrisLogin def
#Include Secrets.vb

'camName is the short cam name in the camera's config.

' Trigger an alert on a Iris4
Public Sub Iris4(ByVal camName As String)

    callTrigger("10.10.2.47",camName,8077)

End Sub

' Trigger an alert on a Iris3
Public Sub Iris3(ByVal camName As String)

    callTrigger("10.10.2.46",camName,8076)

End Sub

' Trigger an alert on a Iris2
Public Sub Iris2(ByVal camName As String)
    callTrigger("10.10.2.48",camName,8078)

End Sub

' Trigger an alert on a Blue Iris server
Public Sub callTrigger(ByVal host As String, ByVal camName As String, ByVal port As integer)
    Dim label = "callTrigger"
    Dim page
    Try
        hs.WriteLog(label, "http://" & host & ":" & port & "/admin?camera=" & camName & "&trigger" & BlueIrisLogin)
        ' hs.GetURL(IP,URL,TRUE,port)
        page = hs.GetURL(host,"/admin?camera=" & camName & "&trigger" & BlueIrisLogin,TRUE,port)
        If (Instr(page,"camera=NULL") > 0) Then 
            hs.WriteLog(label, "camera " & camName & " not found")
        Else If (Instr(page,"var login_version") > 0) Then 
            hs.WriteLog(label, "secure session keys blocking API calls. Turn off.")
        Else If (Instr(page,"Authorization required") > 0) Then 
            hs.WriteLog(label, "Login info is incorrect")
        Else
            ' sucessful response is
            'signal=greenprofile=-1lock=0camera=longName
            hs.WriteLog(label, "GetURL returned:" & page)
        End If
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message & " " & host & ":" & port)
    End Try
End Sub