' Methods for sending messages to Pushbullet

Sub sayString(ByVal msg As String)
    hs.RunScriptFunc("SayIt.vb", "sayString", msg, False, False)
End Sub

Public Sub sayGlobal(ByVal name As String)
    hs.RunScriptFunc("SayIt.vb", "sayGlobal", name, False, False)
End Sub

' Checks script complies and globals used, are defined
Sub Main(ByVal ignored As String)
    sayString("Push bullet Script compiled OK")
    sayGlobal("Partition1Ref")
    sayGlobal("AlarmLevelRef")
    sayGlobal("AlarmOff")
    sayGlobal("AlarmHome")
    sayGlobal("AlarmPerm")
    sayGlobal("AlarmAway")
    Dim curTempID
    curTempID = hs.GetVar("AmbientTemperatureRef")
    SendMsg(curTempID)
End Sub

'Get the best name to use for the device
Public Function betterName(ByVal dv As Object) As String
    Dim label As Object = "betterName"
    Dim name As Object = dv.VoiceCommand(Nothing)
    Try
        hs.WriteLog(label, "VoiceCommand: '" & name & "'")
        If String.IsNullOrEmpty(name) Then
            name = dv.Name(Nothing)
            name.Replace(".", " ").Replace("_pwr", " power").Replace("_", " ")
        End If
        hs.WriteLog(label, "name: '" & name & "'")
        Return name
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
        Return "unknown"
    End Try

End Function

'params in format of title#message you want to send OR 
'the device ref of a device that has a status string OR 
'the device ref of a device like a door sensor where a value of 0 is closed and > 0 is open 
'
' if params has # delimiter then return params
' else assume device ref and if has status string return Name # Status String
' else if device ref is Partition1Ref then decode state from status strings
' else return Name # Open if device value > 0 
' else return Name # Closed
Function ref2params(ByVal params As Object) As String
    if Instr(params,"#") > 0 Then
        return params
    Else
        Dim dv
        Dim cat
        Dim name

        dv = hs.GetDeviceByRef(params)
        cat = dv.Location2(Nothing)
        name = betterName(dv)
        hs.WriteLog("ref2params", "devString:(" & dv.devValue(Nothing) & ")")
        if String.IsNullOrEmpty(dv.devString(Nothing)) Then
            If (hs.GetVar("Partition1Ref") = params) Then
                'hs.WriteLog("ref2params", "Partition1Ref = params")
                Dim cond
                Try
                    Select Case hs.DeviceValue(hs.GetVar("Partition1Ref"))
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
                    Return "" & name & " # " & cond
                Catch ex As Exception
                    hs.WriteLog("Error", "Exception in script ref2params:  " & ex.Message)
                End Try
            ElseIf (InStr(cat, "Temperature") > 0) Then
                'hs.WriteLog("ref2params", "InStr(cat, Temperature) > 0")
                Return "" & name & " is " & dv.devValue(Nothing) & "degrees"
            ElseIf (dv.devValue(Nothing) > 0) Then
                'hs.WriteLog("ref2params", "dv.devValue(Nothing) > 0")
                Return "" & name & " # Open"
            Else
                'hs.WriteLog("ref2params", "dv.devValue(Nothing) <= 0")
                Return "" & name & " # Closed"
            End if
        Else 
            'hs.WriteLog("ref2params", "!String.IsNullOrEmpty(dv.devString(Nothing))")
            return "" & name & " # " & dv.devString(Nothing).Replace("*","")
        End If
        
    End If
End Function


' Send if alarm in away mode
Sub SendMsgIfAway(ByVal params As String) 
    Dim label = "SendMsgIfAway"
    hs.WriteLog(label, "AlarmLevelRef:" & hs.DeviceValue(hs.GetVar("AlarmLevelRef")))
    hs.WriteLog(label, "AlarmAway:" & hs.GetVar("AlarmAway"))
    params = ref2params(params)
    If hs.DeviceValue(hs.GetVar("AlarmLevelRef")) >= hs.GetVar("AlarmAway") Then
        SendMsg(params & " in away mode")
    ElseIf hs.DeviceValue(hs.GetVar("AlarmLevelRef")) >= hs.GetVar("AlarmHome") Then
        Dim paramArr()
        paramArr = params.Split("#")
        'sayString(paramArr(0) & " " & paramArr(1))
        hs.RunScriptFunc("SayIt.vb", "sayString", paramArr(0) & " " & paramArr(1), False, False)
    End if
End Sub

' Send if alarm in perimeter or away mode
Sub SendMsgIfPerm(ByVal params As String) 
    Dim label = "SendMsgIfPerm"
    hs.WriteLog(label, "AlarmLevel2899Ref:" & hs.DeviceValue(hs.GetVar("AlarmLevelRef")))
    params = ref2params(params)
    If hs.DeviceValue(hs.GetVar("AlarmLevelRef")) >= hs.GetVar("AlarmAway") Then
        SendMsgIfAway(params)
    ElseIf hs.DeviceValue(hs.GetVar("AlarmLevelRef")) >= hs.GetVar("AlarmPerm") Then
        SendMsg(params & " in perimeter mode")
    ElseIf hs.DeviceValue(hs.GetVar("AlarmLevelRef")) >= hs.GetVar("AlarmHome") Then
        Dim paramArr()
        paramArr = params.Split("#")
        'sayString(paramArr(0) & " " & paramArr(1))
        hs.RunScriptFunc("SayIt.vb", "sayString", paramArr(0) & " " & paramArr(1), False, False)
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

    s = hs.URLAction(server_url, "POST", data, hs.GetVar("headers"))

    '    sayAlert(title & " " & message)
    hs.RunScriptFunc("SayIt.vb", "sayAlert", title & " " & message, False, False)

End Sub
