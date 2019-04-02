'load object refs and speech methods
#Include SayIt.vb

'Say the name, value (as temp) and last change time of a device.
Public Sub checkWMI(ByVal host As Object)  
    Dim label As String = "checkWMI"

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
            Dim cat = dv.Location2(Nothing)
            Dim dtype = dv.Device_Type_String(Nothing)
            if (InStr(dtype, "WMI") > 0)
                'hs.WriteLog(label, host & ":  " & cat & ":  " & dtype)
                if (InStr(cat, host) > 0)
                    Dim name = dv.Name(Nothing)
                    Dim val = dv.devValue(Nothing)
                    hs.WriteLog(label, name & ":  " & Cstr(val))
                    if (InStr(name, "Load") > 0)
                        if (val > 80)
                            sayValue(dv.Ref(Nothing))
                        End If
                    Else if (InStr(name, "CPU") > 0)
                        if (val > 30)
                            sayValue(dv.Ref(Nothing))
                        End If
                    Else if (InStr(name, "RAM") > 0)
                        if (val > 90)
                            sayValue(dv.Ref(Nothing))
                        End If
                    Else if (InStr(name, "Log Size") > 0)
                        if (val > 10)
                            sayValue(dv.Ref(Nothing))
                        End If
                    Else if (InStr(name, "Updates") > 0)
                        if (val > 3)
                            sayValue(dv.Ref(Nothing))
                        End If
                    Else if ( val < 10) ' drive free space
                        sayValue(dv.Ref(Nothing))
                    End If
                End If
            End If
        Loop Until EN.Finished
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub