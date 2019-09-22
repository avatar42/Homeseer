'load object refs and speech methods
#Include SayIt.vb

    '	985	 ONLINE REPLACE_BATTERY PLUGGED BATT UPS	MailServer	UPS15 Status Flags	UPS15_004	Status	Today 8:36:40 AM	
    '   919	10.10.1.1513553 Is Enabled          UPS	    PI2	        UPS13 Name	        UPS13_001	Status	7/21/2018 7:54:06 PM 'Notice still says enabled though offline
    '   922  ONLINE COMM_LOST PLUGGED BATT      UPS	    PI2	        UPS13 Status Flags	UPS13_004	Status	7/21/2018 7:55:52 PM	
    Public Sub chkUPSs(Parms As Object)
        Dim label As String = "chkUPSs"
        Dim downDevs As Integer = 0
         Try
            hs.WriteLog(label, "Running")
            Dim dv As Scheduler.Classes.DeviceClass
            Dim EN As Scheduler.Classes.clsDeviceEnumeration = hs.GetDeviceEnumerator  'Get all devices
            If EN Is Nothing Then
                hs.WriteLog(label, "Error getting Enumerator")
                Exit Sub
            End If
            ' overwrite file
            Do  'check each device that was enumerated
                dv = EN.GetNext
                If dv Is Nothing Then  'No device, so quit
                    hs.WriteLog(label, "No devices found")
                    Exit Sub
                End If
                'Only work with devices with a battery type
                Dim cat = UCase(dv.Location2(Nothing))
                Dim loc = UCase(dv.Location(Nothing))
                If InStr(cat, "UPS") > 0  And InStr(dv.Device_Type_String(Nothing), "Status") > 0 Then
                    'hs.WriteLog(label, "device:" & dv.Ref(Nothing) & ":" & dv.Location2(Nothing) & ":" & dv.Device_Type_String(Nothing) & " " & dv.Location(Nothing) & " (" & dv.Name(Nothing) & ")" & " (" & dv.devString(Nothing) & ")"  & " (" & dv.devValue(Nothing) & ")" & hs.CapiGetStatus(dv.Ref(Nothing)).Status)
                    Dim status = dv.Name(Nothing) + dv.devString(Nothing) + hs.CapiGetStatus(dv.Ref(Nothing)).Status
                    If InStr(status,"Disconnected") > 0 Or InStr(status,"REPLACE_BATTERY") > 0 Or InStr(status,"COMM_LOST") > 0 Or InStr(status,"LOW") > 0 Then
                        Dim attrs As String = ""
                        For i As Integer = 1 To 1048576
                            If dv.MISC_Check(hs, i) Then
                                attrs = "" & 1 & attrs
                            Else
                                attrs = "" & 0 & attrs
                            End If
                            i = i * 2
                        Next
                        hs.WriteLog(label, "device:" & dv.Ref(Nothing) & ":" & dv.Location2(Nothing) & " " & dv.Location(Nothing) & " " & dv.Name(Nothing) & " with value:" & dv.devValue(Nothing) & " updated:" & dv.Last_Change(Nothing) & " attrs:" & attrs)
                        sayString("U P S on " & loc & " is down.")
                        downDevs = downDevs + 1
                    End If
                End If
                If InStr(cat, "UPS") > 0  And InStr(dv.Device_Type_String(Nothing), "Timestamp") > 0 Then

                    Dim lastChg As DateTime = dv.Last_Change(Nothing)
                    Dim span = now - lastChg
                     If InStr(dv.Name(Nothing),"Last Updated") > 0 And span.TotalHours > 1 Then
                        hs.WriteLog(label, "device:" & dv.Ref(Nothing) & ":" & dv.Location2(Nothing) & " " & dv.Location(Nothing) & " " & dv.Name(Nothing) & " with value:" & dv.devValue(Nothing) & " updated:" & dv.Last_Change(Nothing))
                        sayString("U P S on " & loc & " is down.")
                        downDevs = downDevs + 1
                    End If

                End If
            Loop Until EN.Finished
            if downDevs > 0 Then
                hs.TriggerEvent("Pink - Homeseer error")
            End If
            hs.WriteLog(label, "" & downDevs & " UPSs off line.")
        Catch ex As Exception
            hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
        End Try
   End Sub