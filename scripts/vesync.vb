Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

' import VesyncUsername & VesyncPasswordAsMD5 defs
#Include Secrets.vb

'load object refs and speech methods
#Include SayIt.vb

' Note there is a lot more logging in here than is really needed to make sorting issues easier. Set the following to false to turn off extra logging
Dim debug As Boolean = True
Dim currentFirmVersion = "2.123"

Const BASE_URL As String = "https://smartapi.vesync.com"


' Get device info from API server to be used by events to call sendOn and sendOff
' Response: [{"deviceName":"ipcam9_pwr","deviceImg":"https://smartapi.vesync.com/v1/app/imgs/wifi/outlet/smart_wifi_outlet.png",
' "cid":"73522109-3df5-4428-9194-8d320dcfabe2","deviceStatus":"on","connectionType":"wifi","connectionStatus":"online","deviceType":"wifi-switch-1.3",
' "model":"wifi-switch","currentFirmVersion":"1.95"}, (repeated for each switch)]
Sub getDevices(ByVal unused As String)
    Dim label = "getDevices"

    Try
        Dim rtn As JArray = getDevicesInfo()
        'hs.WriteLog(label, "ResponseP: " & rtn.ToString)

        For Each item As JObject In rtn
            logDebug(label, "Device: " & item.ToString)
            logDebug(label, "deviceName: " & item("deviceName").ToString)
            logDebug(label, "deviceAddr: " & item("cid").ToString)
            logDebug(label, "deviceStatus: " & item("deviceStatus").ToString)
            logDebug(label, "connectionStatus: " & item("connectionStatus").ToString)
            logDebug(label, "currentFirmVersion: " & item("currentFirmVersion").ToString)
        Next

    Catch ex As Exception
        hs.WriteLog("Error", "Error: " & ex.ToString())
    End Try
End Sub

' Find all the Etekcity plugs with matching names and update their addresses, location2, image, stauts string and status value as needed.
Public Sub fixDevices(ByVal unused As String)
    Dim label As String = "fixDevices"
    Dim typeStr As String = "MyMonitor-plug"
    Dim downDevs As Integer = 0
    Dim writeChg = True
    Try
        Dim rtn As JArray = getDevicesInfo()
        Dim dv As Scheduler.Classes.DeviceClass
        Dim EN As Scheduler.Classes.clsDeviceEnumeration = hs.GetDeviceEnumerator  'Get all devices
        If EN Is Nothing Then
            hs.WriteLog(label, "Error getting Enumerator")
            Exit Sub
        End If
        Do  'check each device that was enumerated
            dv = EN.GetNext
            If dv Is Nothing Then  'No device, so quit
                hs.WriteLog(label, "No devices found")
                Exit Sub
            End If
            ' offline and of type
            Dim objName = dv.Name(Nothing)
            If InStr(objName, "Etekcity.") > 0 Then
                Dim devName = Right(objName, Len(objName) - 9)
                Dim chgd = 0
                Dim bak = "" & dv.Location(Nothing) & "." & dv.Location2(Nothing) & "." & dv.Name(Nothing) & " value:" & dv.devValue(Nothing) & " Address:" & dv.Address(Nothing) & " firmware:" & dv.devString(Nothing) & " type:" & dv.Device_Type_String(Nothing)
                For Each item As JObject In rtn
                    If StrComp(dv.Location2(Nothing), "Remove", 0) <> 0 And (StrComp(devName, item("deviceName").ToString, 0) = 0 Or StrComp(dv.Address(Nothing), item("cid").ToString, 0) = 0) Then
                        If StrComp(devName, item("deviceName").ToString, 0) <> 0 Then
                            logDebug(label, "devName:" & devName & "->" & item("deviceName").ToString & "<")
                            If writeChg Then
                                dv.Name(hs) = "Etekcity." & item("deviceName").ToString
                            End If
                            chgd = chgd + 1
                        End If
                        If StrComp(dv.Image(Nothing), "/images/Devices/Etekcity_small.jpg", 0) <> 0 Then
                            logDebug(label, "Image:" & dv.Image(Nothing) & "->/images/Devices/Etekcity_small.jpg")
                            If writeChg Then
                                dv.Image(hs) = "/images/Devices/Etekcity_small.jpg"
                            End If
                            chgd = chgd + 64
                        End If
                        If StrComp(dv.Location2(Nothing), "Power", 0) <> 0 Then
                            logDebug(label, "Location2:" & dv.Location2(Nothing) & "->Power")
                            If writeChg Then
                                dv.Location2(hs) = "Power"
                            End If
                            chgd = chgd + 2
                        End If
                        If StrComp(dv.Address(Nothing), item("cid").ToString, 0) <> 0 Then
                            logDebug(label, "Address:" & dv.Address(Nothing) & "->" & item("cid").ToString)
                            If writeChg Then
                                dv.Address(hs) = item("cid").ToString
                            End If
                            chgd = chgd + 4
                        End If
                        Dim status = 50
                        If StrComp("online", item("connectionStatus").ToString, 0) = 0 Then
                            If StrComp("on", item("deviceStatus").ToString, 0) = 0 Then
                                status = 100
                            Else
                                status = 0
                            End If
                        End If
                        If dv.devValue(Nothing) <> status Then
                            logDebug(label, "devValue:" & dv.devValue(Nothing) & "->" & status)
                            ' uncomment to set the HS value to match the server value
                            If writeChg Then
                                dv.devValue(hs) = status
                            End If
                            chgd = chgd + 8
                        End If
                        If StrComp(currentFirmVersion, item("currentFirmVersion").ToString, 0) <> 0 Then
                            logDebug(label, "Attention:" & currentFirmVersion & "->" & item("currentFirmVersion").ToString)
                            If writeChg Then
                                dv.Attention(hs) = item("currentFirmVersion").ToString
                            End If
                            chgd = chgd + 16
                        End If
                        If StrComp(dv.Device_Type_String(Nothing), typeStr, 0) <> 0 Then
                            logDebug(label, "Device_Type_String:" & dv.Device_Type_String(Nothing) & "->" & typeStr)
                            If writeChg Then
                                dv.Device_Type_String(hs) = typeStr
                            End If
                            chgd = chgd + 32
                        End If
                        Exit For
                    End If
                Next
                If chgd > 0 Then
                    dv.UserNote(hs) = bak
                    downDevs = downDevs + 1
                    logDebug(label, "Updated device " & dv.Location(hs) & "." & dv.Location2(hs) & "." & dv.Name(hs) & " value:" & dv.devValue(hs) & " Address:" & dv.Address(hs) & " firmware:" & dv.Attention(hs) & " type:" & dv.Device_Type_String(hs) & " chg:" & chgd)
                End If
            End If
        Loop Until EN.Finished
        Dim msg = " devices would be fixed."
        If writeChg Then
            msg = " devices are fixed."
        End If
        sayString("" & downDevs & msg)
        logDebug(label, "" & downDevs & msg)
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub

Function getDevicesInfo()
    Dim label = "getDevicesInfo"
    Dim rtn As JArray = Nothing

    Try
        Dim respStr As String = sendCmd("/vold/user/devices", "GET")
        If (respStr Is Nothing) Then
            logDebug(label, "respStr: is Nothing")
        Else
            logDebug(label, "respStr: " & respStr)
            'rtn = JsonConvert.DeserializeObject(respStr)
            'hs.WriteLog(label, "ResponseD: " & rtn.ToString)

            rtn = Newtonsoft.Json.Linq.JArray.Parse(respStr)
        End If

    Catch ex As Exception
        hs.WriteLog("Error", "Error: " & ex.ToString())
    End Try

    Return rtn
End Function

Sub resetRef(ByVal dvRef As String)
    Dim label = "resetRef"
    Dim dv
    dv = hs.GetDeviceByRef(dvRef)
    sayString("doing a reset of " & betterName(dv))
    reset(dv.Address(Nothing))
End Sub

Sub reset(ByVal cid As String)
    Dim label = "reset"
    sendOff(cid)
    'system.threading.thread.sleep(60000)
    hs.waitsecs(60)
    sendOn(cid)
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
    If (debug) Then
        hs.WriteLog(label, msg)
    End If
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