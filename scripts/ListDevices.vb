#Include Refs.vb
' look at all the devices and set virtutals to number of devices, highest ref number and sum of ref numbers so an event can do a check and announce 
' if something changes.
' Mainly this is to monitor some plugins that seem to recreate devices after updates.
' Todo: add save timestamped list if changes are found for compare
Public Sub DevChk(Parms As Object)
    Dim label As String = "DevChk"
    hs.WriteLog(label, "Starting " & label)
    Try
        Dim total As Double = 0
        Dim ref As Long = 0
        Dim cnt As Double = 0
        Dim max As Double = 0
        Dim EN As Scheduler.Classes.clsDeviceEnumeration = hs.GetDeviceEnumerator  'Get all devices
        Dim dv As Scheduler.Classes.DeviceClass

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
            ref = dv.Ref(Nothing)
            'hs.WriteLog(label, "Ref=" & ref & "Address=" & dv.Address(Nothing) & " Code=" & dv.Code(Nothing) & " Can_Dim=" & dv.Can_Dim(Nothing))
            total = total + ref
            cnt = cnt + 1
            If ref > max Then
                max = ref
            End If
        Loop Until EN.Finished
        hs.WriteLog(label, "cnt:  " & cnt & " max: " & max & " total: " & total)
        hs.SetDeviceValueByRef(DeviceCountRef, cnt, True)
        hs.SetDeviceValueByRef(LastDeviceRef, max , True)
        hs.SetDeviceValueByRef(DeviceChksumRef, total, True)
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub