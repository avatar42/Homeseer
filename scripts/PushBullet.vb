'load object refs and speech methods
#Include SayIt.vb

' import headers def
#Include Secrets.vb

' Methods for sending messages to Pushbullet

Sub Main(ByVal params As String) 
    SendMsg(params)
End Sub

'Refs used
'Dim AlarmOff = 0
'Dim AlarmHome = 1
'Dim AlarmPerm = 2
'Dim AlarmAway = 3
'Dim AlarmLevel2899Ref = 2899

'params in format of title#message you want to send

' Send if alarm in away mode
Sub SendMsgIfAway(ByVal params As String) 
    Dim label = "SendMsgIfAway"
    hs.WriteLog(label, "AlarmLevel2899Ref:" & hs.DeviceValue(AlarmLevel2899Ref))

    if hs.DeviceValue(AlarmLevel2899Ref) >= AlarmAway Then
        SendMsg(params & " in away mode")
    Else if hs.DeviceValue(AlarmLevel2899Ref) >= AlarmHome Then
        Dim paramArr()
        paramArr = params.Split("#")
        sayString(paramArr(0) & " " & paramArr(1))
    End if
End Sub

' Send if alarm in perimeter or away mode
Sub SendMsgIfPerm(ByVal params As String) 
    Dim label = "SendMsgIfPerm"
    hs.WriteLog(label, "AlarmLevel2899Ref:" & hs.DeviceValue(AlarmLevel2899Ref))

    if hs.DeviceValue(AlarmLevel2899Ref) >= AlarmAway Then
        SendMsgIfAway(params)
    Else if hs.DeviceValue(AlarmLevel2899Ref) >= AlarmPerm Then
        SendMsg(params & " in perimeter mode")
    Else if hs.DeviceValue(AlarmLevel2899Ref) >= AlarmHome Then
        Dim paramArr()
        paramArr = params.Split("#")
        sayString(paramArr(0) & " " & paramArr(1))
    End if
End Sub

' Send message no matter mode alarm is in
Sub SendMsg(ByVal params As String) 
    Dim label = "SendMsg"
    dim s, data, title, message 
    Dim paramArr() 
    const server_url = "https://api.pushbullet.com/v2/pushes" 

    hs.WriteLog(label,"params:" & params)
    paramArr = params.Split("#") 
    title = paramArr(0) 
    message = paramArr(1) 

    data = "type=note&title="+title+"&body="+message 

    s = hs.URLAction(server_url, "POST", data, headers) 

    sayAlert(title & " " & message)
End Sub
