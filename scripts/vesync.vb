Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

' import VesyncUsername & VesyncPasswordAsMD5 defs
#Include Secrets.vb

' Note there is a lot more logging in here than is really needed to make sorting issues easier. Set the following to false to turn off extra logging
Dim debug As boolean = false

Const BASE_URL As String = "https://smartapi.vesync.com"


' Get device info from API server to be used by events to call sendOn and sendOff
' Response: [{"deviceName":"ipcam9_pwr","deviceImg":"https://smartapi.vesync.com/v1/app/imgs/wifi/outlet/smart_wifi_outlet.png",
' "cid":"73522109-3df5-4428-9194-8d320dcfabe2","deviceStatus":"on","connectionType":"wifi","connectionStatus":"online","deviceType":"wifi-switch-1.3",
' "model":"wifi-switch","currentFirmVersion":"1.95"}, (repeated for each switch)]
Sub getDevices(ByVal cid As String)
    Dim label = "getDevices"

    Try
        Dim rtn As JArray = Nothing
        Dim respStr As String = sendCmd("/vold/user/devices", "GET")
        If (respStr Is Nothing) Then
            logDebug(label, "respStr: is Nothing")
        Else
            logDebug(label, "respStr: " & respStr)
            rtn = JsonConvert.DeserializeObject(respStr)
            hs.WriteLog(label, "Response: " & rtn.ToString)
        End If

    Catch ex As Exception
        hs.WriteLog("Error", "Error: " & ex.ToString())
    End Try
End Sub

Sub sendOn(ByVal cid As String)
    Dim label = "sendOn"
    Try
        logDebug(label, "cid: " & cid)
        Dim rtn As JObject = Nothing
        Dim respStr As String = sendCmd("/v1/wifi-switch-1.3/" & cid & "/status/on", "PUT")
        If (respStr Is Nothing) Then
            logDebug(label, "respStr: is Nothing")
        Else
            hs.WriteLog(label, "respStr: " & respStr)
        End If
    Catch ex As Exception
        hs.WriteLog("Error", "Error: " & ex.ToString())
    End Try
End Sub

Sub sendOff(ByVal cid As String)
    Dim label = "sendOff"
    Try
        logDebug(label, "cid: " & cid)
        Dim rtn As JObject = Nothing
        Dim respStr As String = sendCmd("/v1/wifi-switch-1.3/" & cid & "/status/off", "PUT")
        If (respStr Is Nothing) Then
            logDebug(label, "respStr: is Nothing")
        Else
            hs.WriteLog(label, "respStr: " & respStr)
        End If
    Catch ex As Exception
        hs.WriteLog("Error", "Error: " & ex.ToString())
    End Try
End Sub

Function logDebug(ByVal label As String, ByVal msg As String)
  if (debug) Then
    hs.WriteLog(label, msg)
  End if
End Function

Function sendCmd(ByVal actionPath As String, ByVal method As String) As String
    Dim respJo As JObject
    Dim label = "sendCmd"
    Dim reader As StreamReader
    Dim readStream As Stream
    Dim httpResponse As Net.HttpWebResponse = Nothing
    Dim respStr As String = ""

    'Make the request to the API
    Try
        respJo = postLogin()
        'Set up the Webrequest
        Dim url = BASE_URL & actionPath
        logDebug(label, "URL:" & url)
        Dim httpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
        logDebug(label, "httpWebRequest:")

        'httpWebRequest.ContentType = "application/x-www-form-urlencoded"
        httpWebRequest.Method = method
        httpWebRequest.Headers("tk") = respJo.Value(Of String)("tk")
        httpWebRequest.Headers("accountID") = respJo.Value(Of String)("accountID")
        logDebug(label, "Headers: " & httpWebRequest.Headers.ToString)

        Try
            ' Get response  
            httpResponse = DirectCast(httpWebRequest.GetResponse(), HttpWebResponse)
            If (httpResponse Is Nothing) Then
                logDebug(label, "httpResponse: is Nothing")
            Else
                logDebug(label, "httpResponse: " & httpResponse.ToString)
                logDebug(label, "httpResponse.StatusCode: " & httpResponse.StatusCode)
                If (method Is "PUT") Then
                    respStr = httpResponse.StatusCode
                Else
                    readStream = httpResponse.GetResponseStream()
                    If (readStream Is Nothing) Then
                        logDebug(label, "readStream: is Nothing")
                    Else
                        ' Get the response stream into a reader  
                        reader = New StreamReader(readStream)

                        If (reader Is Nothing) Then
                            logDebug(label, "reader: is Nothing")
                        Else
                            ' Get response string to return  
                            respStr = reader.ReadToEnd()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            hs.WriteLog("Error", "httpResponse Error: " & ex.ToString())
        End Try
    Catch ex As Exception
        hs.WriteLog("Error", "Error: " & ex.ToString())
    Finally
        If Not httpResponse Is Nothing Then httpResponse.Close()
    End Try

    Return respStr
End Function

' log into API server
' Response: {"tk":"LLXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX-Q==","accountID":"1234567","nickName":"USERNAME","avatarIcon":"","userType":1,"acceptLanguage":"en","termsStatus":true}
Function postLogin() As JObject
    Dim reader As StreamReader
    Dim readStream As Stream
    Dim rtn As JObject = Nothing
    Dim label = "postLogin"
    Dim postdata = "{""account"": """ & VesyncUsername & """,""devToken"": """",""password"": """ & VesyncPasswordAsMD5 & """}"
    logDebug(label, "postdata:" & postdata)

    'Make the request to the API
    Try
        'Set up the Webrequest
        Dim url = BASE_URL & "/vold/user/login"
        logDebug(label, "URL:" & url)
        Dim httpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
        logDebug(label, "httpWebRequest:")


        httpWebRequest.ContentType = "application/x-www-form-urlencoded"
        httpWebRequest.Method = "POST"
        'httpWebRequest.Headers.Add("Authorization", "Bearer XXXXX(FILLED IN)XXXXX")

        Dim encoding As New System.Text.UTF8Encoding
        logDebug(label, "encoding:")

        Dim dataBytes As Byte() = encoding.GetBytes(postdata)
        httpWebRequest.ContentLength = dataBytes.Length
        logDebug(label, "dataBytes.Length:" & dataBytes.Length)

        logDebug(label, "httpWebRequest.GetRequestStream().ToString:" & httpWebRequest.GetRequestStream().ToString)
        Dim myStream = httpWebRequest.GetRequestStream()
        logDebug(label, "myStream:")

        If dataBytes.Length > 0 Then
            myStream.Write(dataBytes, 0, dataBytes.Length)
            myStream.Close()
        End If

        ' Get response  
        Dim httpResponse = DirectCast(httpWebRequest.GetResponse(), HttpWebResponse)
        logDebug(label, "httpResponse: " & httpResponse.ToString)

        readStream = httpResponse.GetResponseStream()
        logDebug(label, "readStream: ")

        ' Get the response stream into a reader  
        reader = New StreamReader(readStream)

        ' Console application output  
        Dim respStr = reader.ReadToEnd()
        hs.WriteLog(label, "Response: " & respStr)
        rtn = JsonConvert.DeserializeObject(respStr)
    Catch ex As Exception
        hs.WriteLog("Error", "Error: " & ex.ToString())
        hs.WriteLog("Error", "Request was: " & postdata)
    End Try
    Return rtn
End Function


