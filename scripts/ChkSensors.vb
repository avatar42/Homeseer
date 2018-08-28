'load object refs and speech methods
#Include SayIt.vb

' Use "Do not update device last change time if device value does not change" flag
Dim noChkFlag = HomeSeerAPI.Enums.dvMISC.SET_DOES_NOT_CHANGE_LAST_CHANGE
Dim offlineLoc = "offline"

' Plugins, device roots and counters often only update values if they changed despite the flag setting. Oddly these seem to not usually set the noChkFlag flag either
' so you will probably need to ignore them and other categories of devices that do not update the last changed value..
' return true if is location that should be checked
Public Function isNotFiltered(ByVal dv As Scheduler.Classes.DeviceClass) As Boolean
    Dim cat = dv.Location2(Nothing)
    Dim rm = dv.Location(Nothing)
    Return InStr(cat, "NoData") = 0 And InStr(cat, "Root") = 0 And InStr(cat, "Counters-Timers") = 0 And InStr(rm, "Remove") = 0 And InStr(cat, "UnusedValues") = 0 And InStr(cat, "UPS") = 0 And InStr(rm, "UltraWeatherWU3 Plugin") = 0
End Function

' true if device should be updated at least once every 24 hours
Public Function chk24(ByVal dv As Scheduler.Classes.DeviceClass) As Boolean
    Dim cat = dv.Location2(Nothing)
    Dim rm = dv.Location(Nothing)
    Return InStr(dv.Location2(Nothing), "Batteries") = 0
End Function

' true if device should be updated at least once every hour
Public Function chk1(ByVal dv As Scheduler.Classes.DeviceClass) As Boolean
    Dim cat = dv.Location2(Nothing)
    Dim rm = dv.Location(Nothing)
    Return InStr(dv.Location2(Nothing), "WMI") > 0 Or InStr(dv.Location2(Nothing), "WirelessTag") > 0 Or InStr(dv.Device_Type_String(Nothing), "Netatmo Update") > 0
End Function

' Mark device offline after storing current location
Public Function markOffline(ByVal dv As Scheduler.Classes.DeviceClass)
    If Not dv.Location2(hs) = offlineLoc Then
        ' store cat for restore
        dv.UserNote(hs) = dv.Location2(hs)
        dv.Location2(hs) = offlineLoc
        'hs.SetDeviceString(dv.Ref(Nothing),"<img src=""/images/HomeSeer/ui/Warning.png"">",TRUE)
    End If
End Function

' restore a device that was marked offline
Public Function restoreDevice(ByVal dv As Scheduler.Classes.DeviceClass)
    If dv.Location2(hs) = offlineLoc Then
        dv.Location2(hs) = dv.UserNote(hs)
        'hs.SetDeviceString(dv.Ref(Nothing),"",TRUE)
    End If
End Function

' Mark a device to be checked by chkSensors()
Public Function mark2Chk(ByVal dv As Scheduler.Classes.DeviceClass)
    dv.MISC_Clear(hs, noChkFlag)
    'clean up BLRadar plugin "No Echo" and "Check Sensor" status strings
    If InStr(dv.devString(Nothing), "No Echo") > 0 Or InStr(dv.devString(Nothing), "Check Sensor") > 0 Then
        hs.SetDeviceString(dv.Ref(Nothing), "", True)
    End If

    ' If noChkFlag is unset and Attention has chkSensors in it remove Attention so we can see which devices have flag unset in the device list without opening them up
    If InStr(dv.Attention(Nothing), "chkSensors") > 0 Then
        dv.Attention(hs) = ""
    End If
End Function

' Mark a device NOT to be checked by chkSensors()
Public Function markNoChk(ByVal dv As Scheduler.Classes.DeviceClass)
    dv.MISC_Set(hs, noChkFlag)

    ' If Attention does not has chkSensors in it  set Attention so we can see which devices have flag set in the device list without opening them up
    If InStr(dv.Attention(Nothing), "chkSensors") = 0 Then
        dv.Attention(hs) = "Not Monitored by chkSensors script"
    End If
    restoreDevice(dv)
End Function

' Look for devices that have stopped updating.
' Sets (CheckedDeviceCount4605Ref)
Public Sub chkSensors(unused As Object)
    Dim label As String = "chkSensors"
    Dim now As DateTime = DateTime.Now
    Dim file As System.IO.StreamWriter
    Dim downDevs As Integer = 0
    Dim chkdTotal As Integer = 0
    Try
        Dim dv As Scheduler.Classes.DeviceClass
        Dim EN As Scheduler.Classes.clsDeviceEnumeration = hs.GetDeviceEnumerator  'Get all devices
        If EN Is Nothing Then
            hs.WriteLog(label, "Error getting Enumerator")
            Exit Sub
        End If
        ' overwrite file (append = False)
        file = My.Computer.FileSystem.OpenTextFileWriter("./html/reports/Device2Chk.csv", False)
        file.WriteLine("Ref,Location2 (cat),Location (room),Name,Value,Last_Change,Attrs, Hours Down")
        Do  'check each device that was enumerated
            dv = EN.GetNext
            If dv Is Nothing Then  'No device, so quit
                hs.WriteLog(label, "No devices found")
                Exit Sub
            End If
            'Only check things with noChkFlag unset
            If Not dv.MISC_Check(hs, noChkFlag) Then
                ' Plugins, device roots and counters often only update values if they changed despite the flag setting. Oddly these seem to not usually set the noChkFlag flag either
                ' so you will probably need to ignore them and other categories of devices that do not update the last changed value..
                If isNotFiltered(dv) Then
                    Dim lastChg As DateTime = dv.Last_Change(Nothing)
                    Dim span = now - lastChg
                    chkdTotal = chkdTotal + 1
                    ' Note while checking down for 24 hours is good for most, if you are checking thinks that update once a day 36 or 48 is better. 
                    If span.TotalHours > 48 Or (span.TotalHours > 24 And chk24(dv)) Or (span.TotalHours > 1 And chk1(dv)) Then
                        Dim attrs As String = ""
                        For i As Integer = 1 To 1048576
                            If dv.MISC_Check(hs, i) Then
                                attrs = "" & 1 & attrs
                            Else
                                attrs = "" & 0 & attrs
                            End If
                            i = i * 2
                        Next
                        file.WriteLine("" & dv.Ref(Nothing) & "," & dv.Location2(Nothing) & "," & dv.Location(Nothing) & "," & dv.Name(Nothing) & "," & dv.devValue(Nothing) & "," & dv.Last_Change(Nothing) & "," & attrs & "," & span.TotalHours)
                        markOffline(dv)
                        downDevs = downDevs + 1
                        If chk1(dv) Then
                            sayString("" & dv.Name(Nothing) & " is off line.")
                        End If
                    Else
                        ' if has recovered since last chk restore the cat to remove from check list
                        restoreDevice(dv)
                    End If
                Else
                    restoreDevice(dv)
                End If
                mark2Chk(dv)
            Else ' noChkFlag set (don't monitor)
                If isNotFiltered(dv) Then
                    markNoChk(dv)
                End If
                restoreDevice(dv)
            End If
        Loop Until EN.Finished
        hs.SetDeviceValueByRef(CheckedDeviceCount4605Ref, chkdTotal, True)
        If downDevs > 0 Then
            sayString("" & downDevs & " devices are off line.")
            hs.TriggerEvent("Pink - Homeseer error")
        End If
        hs.WriteLog(label, "" & downDevs & " devices are off line.")
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    Finally
        file.Close()
    End Try
End Sub

' unset noChkFlag on devices not in the filtered locations.
' mainly used to get you started
Public Sub initFlags(Parms As Object)
    Dim label As String = "initFlags"
    Dim downDevs As Integer = 0
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
            'Only check things with noChkFlag (Do not update device last change time if device value does not change) set
            If dv.MISC_Check(hs, noChkFlag) Then
                ' If not in one of the filtered locations then force the noChkFlag off so checker will check
                If isNotFiltered(dv) Then
                    mark2Chk(dv)
                    downDevs = downDevs + 1
                End If
            End If
        Loop Until EN.Finished
        sayString("" & downDevs & " devices had flags cleared.")
        hs.WriteLog(label, "" & downDevs & " devices had flags cleared.")
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub

' Find all the devices where type contains typeStr and Location2 (cat) of offline and set noChkFlag
Public Sub fixFlagsByType(typeStr As String)
    Dim label As String = "fixFlagsByType"
    Dim downDevs As Integer = 0
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
            If InStr(dv.Location2(Nothing), offlineLoc) > 0 And InStr(dv.Device_Type_String(Nothing), typeStr) > 0 Then
                markNoChk(dv)
                downDevs = downDevs + 1
            End If
        Loop Until EN.Finished
        sayString("" & downDevs & " devices are fixed.")
        hs.WriteLog(label, "" & downDevs & " devices are fixed.")
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub

' Find all the devices where Location (room) contains roomStr and Location2 (cat) of offline and set noChkFlag
Public Sub fixFlagsByRoom(roomStr As String)
    Dim label As String = "fixFlagsByRoom"
    Dim downDevs As Integer = 0
    Try
        If roomStr Is Nothing Then
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
            If InStr(dv.Location2(Nothing), offlineLoc) > 0 And InStr(dv.Location(Nothing), roomStr) > 0 Then
                markNoChk(dv)
                downDevs = downDevs + 1
            End If
        Loop Until EN.Finished
        sayString("" & downDevs & " devices are fixed.")
        hs.WriteLog(label, "" & downDevs & " devices are fixed.")
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub
