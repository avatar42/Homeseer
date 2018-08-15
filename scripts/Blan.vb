'load object refs and speech methods
#Include SayIt.vb

' Hack to make Blan devices have a status value
' call from a set or value changed event
Public Sub statusBlan(ByVal refID As Object)
    Dim status As String = hs.DeviceString(refID)
    Dim oldVal = hs.DeviceValue(refID)

    If (InStr(status, "Offline") > 0) Then
        hs.SetDeviceValueByRef(refID, 0, TRUE)
    ElseIf (InStr(status, "Online") > 0) Then
        hs.SetDeviceValueByRef(refID, 100, TRUE)
    Else
        hs.SetDeviceValueByRef(refID, 50, TRUE)
    End If
    Dim newVal = hs.DeviceValue(refID)

    If Not oldVal = newVal Then
        hs.WriteLog("statusBlan", "Fixed value of BLAN device " & refID & " with status of " & status & " from " & oldVal & " to " & newVal)
    End If
End Sub

' Find all the BLAN devices and call statusBlan() on them
Public Sub fixBlan(Parms As Object)
    Dim label As String = "fixBlan"
    Try
        Dim dv As Scheduler.Classes.DeviceClass
        Dim EN As Scheduler.Classes.clsDeviceEnumeration = hs.GetDeviceEnumerator  

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
            If InStr(UCase(dv.Location2(Nothing)), "BLAN") > 0 Then  'Only work with devices in location BLAN
                'hs.WriteLog(label, "ref:" & dv.ref(Nothing) & ":Address:" & dv.Address(Nothing) & ":value:" & dv.devValue(Nothing) & ":location:" & dv.Location(Nothing) & ":location2:" & dv.Location2(Nothing) & ":device name:" & dv.Name(Nothing) & ":type:" & dv.Device_Type_String(Nothing))
                statusBlan(dv.ref(Nothing))
            End If
        Loop Until EN.Finished
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub


