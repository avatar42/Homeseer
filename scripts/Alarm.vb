'load object refs and speech methods
#Include PushBullet.vb

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
            Exit Sub
        End If

        Do  'check each device that was enumerated
            dv = EN.GetNext
            If dv Is Nothing Then  'No device, so quit
                hs.WriteLog(label, "No devices found")
                Exit Sub
            End If
            'Only check things in category Alarm
            If InStr(dv.Location2(Nothing), "Alarm") > 0 Then
                ' If it is open say something
                If dv.devValue(Nothing) > 0 Then
                    SendMsgIfAway(dv.Ref(Nothing))
                    openSensors = openSensors + 1
                End If
            End If
        Loop Until EN.Finished
        If openSensors = 0 Then
            sayString("House is secure")
            hs.TriggerEvent(eventName)
        Else
            sayAlert("Alarm could not be set")
            hs.TriggerEvent("Red - Security Alert")
        End If
        hs.WriteLog(label, "" & openSensors & " sensors are showing open")
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub
