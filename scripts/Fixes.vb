'load object refs and speech methods
#Include SayIt.vb

' Find all the devices where type contains typeStr 
Public Sub setImageByType(ByVal parms As String)
    Dim label As String = "setImageByType"
    Dim args() As String = Split(parms, ",")
    Dim typeStr As String = args(0)
    Dim path As String = args(1)

    Dim chgDevs As Integer = 0
    Try
        If typeStr Is Nothing Then
            hs.WriteLog(label, "No type string passed")
            Exit Sub
        End If

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
            ' offline and of type
            If InStr(dv.Device_Type_String(Nothing), typeStr) > 0 Then
                'dv.Image(hs) = path
                hs.WriteLog(label, "Changed " & dv.Ref(Nothing) & ":" & dv.Name(Nothing) & " From:" & dv.Image(Nothing) & " To:" & path)
                chgDevs = chgDevs + 1
            End If
        Loop Until EN.Finished
        sayString("" & chgDevs & " devices were fixed.")
        hs.WriteLog(label, "" & chgDevs & " devices were fixed.")
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub

' Find all the devices where address contains addrStr 
Public Sub setImageByAddr(ByVal parms As String)
    Dim label As String = "setImageByType"
    Dim args() As String = Split(parms, ",")
    Dim addrStr As String = args(0)
    Dim path As String = args(1)

    Dim chgDevs As Integer = 0
    Try
        If addrStr Is Nothing Then
            hs.WriteLog(label, "No address sub string passed")
            Exit Sub
        End If

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
            ' offline and of type
            If InStr(dv.Address(Nothing), addrStr) > 0 Then
                dv.Image(hs) = path
                hs.WriteLog(label, "Changed " & dv.Ref(Nothing) & ":" & dv.Name(Nothing) & " From:" & dv.Image(Nothing) & " To:" & path)
                chgDevs = chgDevs + 1
            End If
        Loop Until EN.Finished
        sayString("" & chgDevs & " devices were fixed.")
        hs.WriteLog(label, "" & chgDevs & " devices were fixed.")
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub

' Find all the devices where Category (location2) contains cat
Public Sub setImageByCat(ByVal parms As String)
    Dim label As String = "setImageByType"
    Dim args() As String = Split(parms, ",")
    Dim cat As String = args(0)
    Dim path As String = args(1)

    Dim chgDevs As Integer = 0
    Try
        If cat Is Nothing Then
            hs.WriteLog(label, "No address sub string passed")
            Exit Sub
        End If

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
            ' offline and of type
            If InStr(dv.Location2(Nothing), cat) > 0 Then
                dv.Image(hs) = path
                hs.WriteLog(label, "Changed " & dv.Ref(Nothing) & ":" & dv.Name(Nothing) & " From:" & dv.Image(Nothing) & " To:" & path)
                chgDevs = chgDevs + 1
            End If
        Loop Until EN.Finished
        sayString("" & chgDevs & " devices were fixed.")
        hs.WriteLog(label, "" & chgDevs & " devices were fixed.")
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub
