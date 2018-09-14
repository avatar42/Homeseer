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

'params in format of title#message you want to send OR 
'the device ref of a device that has a status string OR 
'the device ref of a device like a door sensor where a value of 0 is closed and > 0 is open 
'
' if params has # delimiter then return params
' else assume device ref and if has status string return Name Changed#to Status String
' else return Name Changed#to Open if device value > 0 
' else return Name Changed#to Closed
Function ref2params(ByVal params As Object) As String
    if Instr(params,"#") > 0 Then
        return params
    Else
        Dim dv

        dv = hs.GetDeviceByRef(params)
        if dv.devString(Nothing) is Nothing Then
            if (dv.devValue(Nothing) > 0) Then
                return "" & dv.Name(Nothing) & " Changed#to Open"
            Else
                return "" & dv.Name(Nothing) & " Changed#to Closed"
            End if
        Else 
            return "" & dv.Name(Nothing) & " Changed#to " & dv.devString(Nothing).Replace("*","")
        End If
        
    End If
End Function


' Send if alarm in away mode
Sub SendMsgIfAway(ByVal params As String) 
    Dim label = "SendMsgIfAway"
    hs.WriteLog(label, "AlarmLevel2899Ref:" & hs.DeviceValue(AlarmLevel2899Ref))
    params = ref2params(params)
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
    params = ref2params(params)
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
