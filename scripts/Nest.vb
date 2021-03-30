' Say everywhere ignoring AnnounceMode981Ref value
Public Sub sayAlert(ByVal msg As Object)
    hs.RunScriptFunc("SayIt.vb", "sayAlert", msg, False, False)
End Sub

Public Sub sayGlobal(ByVal name As String)
    hs.RunScriptFunc("SayIt.vb", "sayGlobal", name, False, False)
End Sub

' Checks script complies and globals used, are defined
Sub Main(ByVal ignored As String)
    sayAlert("Nest Script compiled OK")
    sayGlobal("AmbientTemperatureRef")
    sayGlobal("TargetTemperatureLowRef")
    sayGlobal("TargetTemperatureHighRef")
    sayGlobal("RmrrRef")
    sayGlobal("HVACModeRef")
    sayGlobal("LivingRoomThermostatRef")
    nestStatus(Nothing)
End Sub

Public Sub nestRange(ByVal delta As String)
    sayAlert(delta & " home temperature is " & hs.DeviceValue(hs.GetVar("AmbientTemperatureRef")) & " degrees" & " range is " & hs.DeviceValue(hs.GetVar("TargetTemperatureLowRef")) & " degrees to " & hs.DeviceValue(hs.GetVar("TargetTemperatureHighRef")) & " degrees")
End Sub

Public Sub nestStatus(ByVal delta As String)
    Dim statStr = " nest mode is "
    If delta Is Nothing Then
        delta = "Current"
    End If

    Select Case hs.DeviceValue(hs.GetVar("RmrrRef"))
        Case 0
            statStr = statStr & "Zero "
        Case 1
            statStr = statStr & "Home "
        Case 2
            statStr = statStr & "Away "
        Case 3
            statStr = statStr & "Auto Away "
        Case Else
            statStr = statStr & "Unknown "
    End Select

    Select Case hs.DeviceValue(hs.GetVar("HVACModeRef"))
        Case 0
            statStr = statStr & ", off "
        Case 1
            statStr = statStr & ", Heat "
        Case 2
            statStr = statStr & ", Cool "
        Case 3
            statStr = statStr & ", Auto "
        Case 4
            statStr = statStr & ", Economy "
        Case Else
            statStr = statStr & ", Unknown "
    End Select

    Select Case hs.DeviceValue(hs.GetVar("LivingRoomThermostatRef"))
        Case 0
            statStr = statStr & " and off line "
        Case 1
            statStr = statStr & " and connected "
        Case Else
            statStr = statStr & " and undefined "
    End Select

    sayAlert(delta & statStr & " this " & TimeOfDay())
    sayAlert(delta & " home temperature is " & hs.DeviceValue(hs.GetVar("AmbientTemperatureRef")) & " degrees")
    sayAlert(delta & " range is " & hs.DeviceValue(hs.GetVar("TargetTemperatureLowRef")) & " degrees to " & hs.DeviceValue(hs.GetVar("TargetTemperatureHighRef")) & " degrees")

End Sub

Public Sub NestBump(ByVal delta As Object)
    Dim curTempID, hsetID, lsetID, hset, lset, curTemp
    curTempID = hs.GetVar("AmbientTemperatureRef")
    lsetID = hs.GetVar("TargetTemperatureLowRef")
    hsetID = hs.GetVar("TargetTemperatureHighRef")

    nestStatus("Current")
    'hs.WriteLog("NestBump", "Range was " & hs.DeviceValue(hs.GetVar("TargetTemperatureLowRef")) & " degrees to " & hs.DeviceValue(hs.GetVar("TargetTemperatureHighRef")))
    curTemp = hs.DeviceValue(hs.GetVar("AmbientTemperatureRef")) + delta ' where we want to be at
    lset = Math.Abs(curTemp - hs.DeviceValue(hs.GetVar("TargetTemperatureLowRef")))
    hset = Math.Abs(curTemp - hs.DeviceValue(hs.GetVar("TargetTemperatureHighRef")))
    If lset < hset Then ' if cur temp closer to low temp
        If hs.DeviceValue(hs.GetVar("TargetTemperatureLowRef")) = curTemp Then
            lset = curTemp + delta
            hs.SetDeviceValueByRef(hs.GetVar("HVACModeRef"), 1, TRUE) ' swtich to heat mode
        Else
            lset = curTemp
        End If

        hset = hs.DeviceValue(hs.GetVar("TargetTemperatureHighRef"))
    Else ' if cur temp closer to high temp
        lset = hs.DeviceValue(hs.GetVar("TargetTemperatureLowRef"))
        If hs.DeviceValue(hs.GetVar("TargetTemperatureLowRef")) = curTemp Then
            hset = curTemp + delta
            hs.SetDeviceValueByRef(hs.GetVar("HVACModeRef"), 2, TRUE) ' switch to cool mode
        Else
            hset = curTemp
        End If
    End If
    hs.SetDeviceValueByRef(hs.GetVar("TargetTemperatureLowRef"), lset, TRUE)
    hs.SetDeviceValueByRef(hs.GetVar("TargetTemperatureHighRef"), hset, TRUE)
    hs.SetDeviceValueByRef(hs.GetVar("RmrrRef"), 1, TRUE) 'Home
    hs.SetDeviceValueByRef(hs.GetVar("HVACModeRef"), 3, TRUE) ' switch to auto mode Hopefully this or previous will break out of eco mode

    hs.WriteLog("NestBump", "Setting range to " & lset & " degrees to " & hset)
    nestStatus("New")
    'sayAlert("New range is " & hs.DeviceValue(hs.GetVar("TargetTemperatureLowRef")) & " degrees to " & hs.DeviceValue(hs.GetVar("TargetTemperatureHighRef")))

End Sub
