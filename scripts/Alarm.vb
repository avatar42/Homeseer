Sub SendMsgIfAway(ByVal ref As String)
    hs.RunScriptFunc("PushBullet.vb", "SendMsgIfAway", ref, False, False)
End Sub

' Checks script complies and globals used, are defined
Sub Main(ByVal ignored As String)
    hs.speakEx(0, "Alarm Script compiled OK", False)
End Sub

' Run event 'eventName' if all alarm sensors in secure mode
' otherwise announce which sensors need checked
Public Sub chkAlarm(eventName As Object)
    Dim label As String = "chkAlarm"
    Dim openSensors As Integer = 0

    Try
        Dim dv As Scheduler.Classes.DeviceClass
        Dim EN As Scheduler.Classes.clsDeviceEnumeration = hs.GetDeviceEnumerator  'Get all devices
        If EN Is Nothing Then
            hs.WriteLog(label, "Error getting Enumerator")
            hs.speakEx(0, "Error getting Enumerator", False)
            Exit Sub
        End If

        Do  'check each device that was enumerated
            dv = EN.GetNext
            If dv Is Nothing Then  'No device, so quit
                hs.speakEx(0, "No devices found", False)
                hs.WriteLog(label, "No devices found")
                Exit Sub
            End If
            'Only check things in category Alarm
            If InStr(dv.Location2(Nothing), "Alarm") > 0 Then
                ' If it is open say something. Note Homeseer door senesor closed = 23
                If (dv.devValue(Nothing) > 0 And dv.devValue(Nothing) <> 23) Then
                    'hs.RunScriptFunc("PushBullet.vb", "SendMsgIfAway", dv.Ref(Nothing), False, False)
                    Dim name As Object = dv.VoiceCommand(Nothing)
                    Try
                        hs.WriteLog(label, "VoiceCommand: '" & name & "'")
                        If String.IsNullOrEmpty(name) Then
                            name = dv.Name(Nothing)
                        End If
                        hs.WriteLog(label, "name: '" & name & "'")
                    Catch ex As Exception
                        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
                        name = "unknown"
                    End Try
                    'hs.speakEx(0, name & " not ready. Value is " & dv.devValue(Nothing), False)
                    SendMsgIfAway(dv.Ref(Nothing))
                    openSensors = openSensors + 1
                End If
            End If
        Loop Until EN.Finished
        If openSensors = 0 Then
            hs.speakEx(0, "House is secure", False)
            hs.TriggerEvent(eventName)
        Else
            hs.WriteLog(label, "" & openSensors & " sensors are showing open")
            hs.speakEx(0, "Alarm could not be set", False)
            hs.TriggerEvent("Red - Security Alert")
        End If
        hs.WriteLog(label, "" & openSensors & " sensors are showing open")
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub
