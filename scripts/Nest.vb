'load object refs and speech methods
#Include SayIt.vb
'Uses (AmbientTemperature3176Ref)
'(TargetTemperatureLow3177Ref)
'(TargetTemperatureHigh3178Ref)
'(Rmrr3174Ref)
'(HVACMode3179Ref)
'(LivingRoomThermostat3175Ref)

Public Sub nestRange(ByVal delta As String)
    sayAlert(delta & " home temperature is " & hs.DeviceValue(AmbientTemperature3176Ref) & " degrees" & " range is " & hs.DeviceValue(TargetTemperatureLow3177Ref) & " degrees to " & hs.DeviceValue(TargetTemperatureHigh3178Ref) & " degrees")
End Sub

Public Sub nestStatus(ByVal delta As String)
    Dim statStr = " nest mode is "
    If delta Is Nothing Then
        delta = "Current"
    End If

    Select Case hs.DeviceValue(Rmrr3174Ref)
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

    Select Case hs.DeviceValue(HVACMode3179Ref)
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

    Select Case hs.DeviceValue(LivingRoomThermostat3175Ref)
        Case 0
            statStr = statStr & " and off line "
        Case 1
            statStr = statStr & " and connected "
        Case Else
            statStr = statStr & " and undefined "
    End Select

    sayAlert(delta & statStr & " this " & TimeOfDay())
    sayAlert(delta & " home temperature is " & hs.DeviceValue(AmbientTemperature3176Ref) & " degrees")
    sayAlert(delta & " range is " & hs.DeviceValue(TargetTemperatureLow3177Ref) & " degrees to " & hs.DeviceValue(TargetTemperatureHigh3178Ref) & " degrees")

End Sub

Public Sub NestBump(ByVal delta As Object)
    Dim curTempID, hsetID, lsetID, hset, lset, curTemp
    curTempID = AmbientTemperature3176Ref
    lsetID = TargetTemperatureLow3177Ref
    hsetID = TargetTemperatureHigh3178Ref

    nestStatus("Current")
    'hs.WriteLog("NestBump", "Range was " & hs.DeviceValue(TargetTemperatureLow3177Ref) & " degrees to " & hs.DeviceValue(TargetTemperatureHigh3178Ref))
    curTemp = hs.DeviceValue(AmbientTemperature3176Ref) + delta ' where we want to be at
    lset = Math.Abs(curTemp - hs.DeviceValue(TargetTemperatureLow3177Ref))
    hset = Math.Abs(curTemp - hs.DeviceValue(TargetTemperatureHigh3178Ref))
    If lset < hset Then ' if cur temp closer to low temp
        If hs.DeviceValue(TargetTemperatureLow3177Ref) = curTemp Then
            lset = curTemp + delta
            hs.SetDeviceValueByRef(HVACMode3179Ref, 1, TRUE) ' swtich to heat mode
        Else
            lset = curTemp
        End If

        hset = hs.DeviceValue(TargetTemperatureHigh3178Ref)
    Else ' if cur temp closer to high temp
        lset = hs.DeviceValue(TargetTemperatureLow3177Ref)
        If hs.DeviceValue(TargetTemperatureLow3177Ref) = curTemp Then
            hset = curTemp + delta
            hs.SetDeviceValueByRef(HVACMode3179Ref, 2, TRUE) ' switch to cool mode
        Else
            hset = curTemp
        End If
    End If
    hs.SetDeviceValueByRef(TargetTemperatureLow3177Ref, lset, TRUE)
    hs.SetDeviceValueByRef(TargetTemperatureHigh3178Ref, hset, TRUE)
    hs.SetDeviceValueByRef(Rmrr3174Ref, 1, TRUE) 'Home
    hs.SetDeviceValueByRef(HVACMode3179Ref, 3, TRUE) ' switch to auto mode Hopefully this or previous will break out of eco mode

    hs.WriteLog("NestBump", "Setting range to " & lset & " degrees to " & hset)
    nestStatus("New")
    'sayAlert("New range is " & hs.DeviceValue(TargetTemperatureLow3177Ref) & " degrees to " & hs.DeviceValue(TargetTemperatureHigh3178Ref))

End Sub
