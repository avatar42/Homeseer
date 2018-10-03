'load object refs and speech methods
#Include SayIt.vb

' import headers def
#Include Secrets.vb

' Methods for sending messages to Pushbullet

Sub Main(ByVal params As String) 
    SendMsg(params)
End Sub

' uses (Partition14702Ref)
' (AlarmLevel2899Ref)
' Const AlarmOff = 0
' Const AlarmHome = 1
' Const AlarmPerm = 2
' Const AlarmAway = 3


'params in format of title#message you want to send OR 
'the device ref of a device that has a status string OR 
'the device ref of a device like a door sensor where a value of 0 is closed and > 0 is open 
'
' if params has # delimiter then return params
' else assume device ref and if has status string return Name # Status String
' else if device ref is Partition14702Ref then decode state from status strings
' else return Name # Open if device value > 0 
' else return Name # Closed
Function ref2params(ByVal params As Object) As String
    if Instr(params,"#") > 0 Then
        return params
    Else
        Dim dv

        dv = hs.GetDeviceByRef(params)
    hs.WriteLog("ref2params", "devString:(" & dv.devValue(Nothing) & ")")
        if String.IsNullOrEmpty(dv.devString(Nothing)) Then
            if (Partition14702Ref = params) Then
                Dim cond
                try
                    Select Case hs.DeviceValue(Partition14702Ref)
                        Case -1005
                            cond = "Disarm"
                        Case -1004
                            cond = "Arm Maximum"
                        Case -1003
                            cond = "Arm Instant"
                        Case -1002
                            cond = "Arm Night-Stay"
                        Case -1001
                            cond = "Arm Away"
                        Case -1000
                            cond = "Arm Stay"
                        Case 0
                            cond = "Unknown"
                        Case 1
                            cond = "Ready"
                        Case 2
                            cond = "Ready to Arm (Zones are Bypassed)"
                        Case 3
                            cond = "Not ready"
                        Case 4
                            cond = "Armed in Stay Mode"
                        Case 5
                            cond = "Armed in Away Mode"
                        Case 6
                            cond = "Armed Instant (Zero Entry Delay - Stay)"
                        Case 7
                            cond = "Exit Delay in Progress"
                        Case 8
                            cond = "In Alarm"
                        Case 9
                            cond = "Alarm has Occured"
                        Case 10
                            cond = "Armed Maximum (Zero Entry Delay - Away)"
                        Case 11
                            cond = "Armed in Night Stay Mode"
                        Case Else
                            cond = "Error"
                    End Select
                    return "" & dv.Name(Nothing) & " # " & cond
                Catch ex As Exception
                    hs.WriteLog("Error", "Exception in script ref2params:  " & ex.Message)
                End Try

            Else if (dv.devValue(Nothing) > 0) Then
                return "" & dv.Name(Nothing) & " # Open"
            Else
                return "" & dv.Name(Nothing) & " # Closed"
            End if
        Else 
            return "" & dv.Name(Nothing) & " # " & dv.devString(Nothing).Replace("*","")
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
