Sub sayValue(ByVal ref As String)
    hs.RunScriptFunc("SayIt.vb", "sayValue", ref, False, False)
End Sub

Sub sayString(ByVal msg As String)
        hs.RunScriptFunc("SayIt.vb", "sayString", msg, False, False)
End Sub

' Checks script complies and globals used, are defined
Sub Main(ByVal ignored As String)
    sayString("Check W M I Script compiled OK")
End Sub

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
            If (InStr(dtype, "WMI") > 0) Then
                'hs.WriteLog(label, host & ":  " & cat & ":  " & dtype)
                If (InStr(cat, host) > 0) Then
                    Dim name = dv.Name(Nothing)
                    Dim val = dv.devValue(Nothing)
                    hs.WriteLog(label, name & ":  " & CStr(val))
                    If (InStr(name, "Load") > 0) Then
                        If (val > 80) Then
                            sayValue(dv.Ref(Nothing))
                        End If
                    ElseIf (InStr(name, "CPU") > 0) Then

                        If (val > 30) Then
                            sayValue(dv.Ref(Nothing))
                        End If
                    ElseIf (InStr(name, "RAM") > 0) Then

                        If (val > 90) Then
                            sayValue(dv.Ref(Nothing))
                        End If
                    ElseIf (InStr(name, "Log Size") > 0) Then

                        If (val > 30) Then
                            sayValue(dv.Ref(Nothing))
                        End If
                    ElseIf (InStr(name, "Updates") > 0) Then

                        If (val > 3) Then
                            sayValue(dv.Ref(Nothing))
                        End If
                    ElseIf (val < 10) Then ' drive free space
                        sayValue(dv.Ref(Nothing))
                    End If
                End If
            End If
        Loop Until EN.Finished
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub

