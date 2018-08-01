'load object refs and speech methods
#Include SayIt.vb

Public Sub nestRange(ByVal delta As String)
    sayAlert(delta & " home temperature is " & hs.DeviceValue(AmbientTemperatureRef) & " degrees" & " range is " & hs.DeviceValue(TargetTemperatureLowRef) & " degrees to " & hs.DeviceValue(TargetTemperatureHighRef) & " degrees")
End Sub

Public Sub nestStatus(ByVal delta As String)
    Dim statStr = " nest mode is "
    If delta Is Nothing Then
        delta = "Current"
    End If

    Select Case hs.DeviceValue(RmrrRef)
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

    Select Case hs.DeviceValue(HVACModeRef)
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

    Select Case hs.DeviceValue(LivingRoomThermostatRef)
        Case 0
            statStr = statStr & " and off line "
        Case 1
            statStr = statStr & " and connected "
        Case Else
            statStr = statStr & " and undefined "
    End Select

    sayAlert(delta & statStr & " this " & TimeOfDay())
    sayAlert(delta & " home temperature is " & hs.DeviceValue(AmbientTemperatureRef) & " degrees")
    sayAlert(delta & " range is " & hs.DeviceValue(TargetTemperatureLowRef) & " degrees to " & hs.DeviceValue(TargetTemperatureHighRef) & " degrees")

End Sub

Public Sub NestBump(ByVal delta As Object)
    Dim curTempID, hsetID, lsetID, hset, lset, curTemp
    curTempID = AmbientTemperatureRef
    lsetID = TargetTemperatureLowRef
    hsetID = TargetTemperatureHighRef

    nestStatus("Current")
    'hs.WriteLog("NestBump", "Range was " & hs.DeviceValue(TargetTemperatureLowRef) & " degrees to " & hs.DeviceValue(TargetTemperatureHighRef))
    curTemp = hs.DeviceValue(AmbientTemperatureRef) + delta ' where we want to be at
    lset = Math.Abs(curTemp - hs.DeviceValue(TargetTemperatureLowRef))
    hset = Math.Abs(curTemp - hs.DeviceValue(TargetTemperatureHighRef))
    If lset < hset Then ' if cur temp closer to low temp
        If hs.DeviceValue(TargetTemperatureLowRef) = curTemp Then
            lset = curTemp + delta
            hs.SetDeviceValueByRef(HVACModeRef , 1, TRUE) ' swtich to heat mode
        Else
            lset = curTemp
        End If

        hset = hs.DeviceValue(TargetTemperatureHighRef)
    Else ' if cur temp closer to high temp
        lset = hs.DeviceValue(TargetTemperatureLowRef)
        If hs.DeviceValue(TargetTemperatureLowRef) = curTemp Then
            hset = curTemp + delta
            hs.SetDeviceValueByRef(HVACModeRef , 2, TRUE) ' switch to cool mode
        Else
            hset = curTemp
        End If
    End If
    hs.SetDeviceValueByRef(TargetTemperatureLowRef, lset, TRUE)
    hs.SetDeviceValueByRef(TargetTemperatureHighRef, hset, TRUE)
    hs.SetDeviceValueByRef(RmrrRef, 1, TRUE) 'Home
    hs.SetDeviceValueByRef(HVACModeRef , 3, TRUE) ' switch to auto mode Hopefully this or previous will break out of eco mode

    hs.WriteLog("NestBump", "Setting range to " & lset & " degrees to " & hset)
    nestStatus("New")
    'sayAlert("New range is " & hs.DeviceValue(TargetTemperatureLowRef) & " degrees to " & hs.DeviceValue(TargetTemperatureHighRef))

End Sub
