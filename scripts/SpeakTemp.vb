Public Sub sayString(ByVal msg As Object)
 hs.RunScriptFunc("SayIt.vb", "sayString", msg, False, False)
End Sub

'Say the name, value (as temp) and last change time of a device.
Public Sub SpeakTemp(ByVal devRef As Object)  
    if (hs.DeviceExistsRef(devRef) ) Then
        sayString("Sensor " & hs.DeviceName(devRef) & " has a temperature of " & hs.DeviceValue(devRef) & " degrees as of " & hs.DeviceLastChangeRef(devRef))
    Else
        sayString("Sensor " & devRef & " not found " )
    End If
End Sub  
 