Sub sayString(ByVal msg As String)
        hs.RunScriptFunc("SayIt.vb", "sayString", msg, False, False)
End Sub

Public Sub sayGlobal(ByVal name As String)
        hs.RunScriptFunc("SayIt.vb", "sayGlobal", name, False, False)
End Sub

' Checks script complies and globals used, are defined
Sub Main(ByVal ignored As String)
    sayString("List Devices Script compiled OK")
    sayGlobal("devicecountref")
    sayGlobal("lastdeviceref")
    sayGlobal("devicechksumref")
End Sub

Const Quote = """"
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
        hs.SetDeviceValueByRef(hs.GetVar("devicecountref"), cnt, True)
        hs.SetDeviceValueByRef(hs.GetVar("lastdeviceref"), max, True)
        hs.SetDeviceValueByRef(hs.GetVar("devicechksumref"), total, True)
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub

Public Function SafeString(ByVal value As Object) As String
    Return Quote & CStr(value).Replace(Quote, "'") & Quote & ","
End Function

' Create a sheet with info about devices including what does not show on the devices management screen
Public Sub exportDevAttrs(Parms As Object)
    Dim label As String = "exportDevAttrs"
    Dim file As System.IO.StreamWriter

    hs.WriteLog(label, "Starting " & label)
    Try

        Dim EN As Scheduler.Classes.clsDeviceEnumeration = hs.GetDeviceEnumerator  'Get all devices
        Dim dv As Scheduler.Classes.DeviceClass

        If EN Is Nothing Then
            hs.WriteLog(label, "Error getting Enumerator")
            Exit Sub
        End If
        file = My.Computer.FileSystem.OpenTextFileWriter("./html/reports/exportDevAttrs." & Now.ToString("yyMMdd") & ".csv", False)
        'Status page stuff
        Dim header As String = "Ref,Image,Status,Cat,Room,Name,Address,Type,Last_Change,Value,Attention,VoiceCommand,"
        'MISC bits
        header = header & "BIT0,BIT1,BIT2,NO_LOG,STATUS_ONLY,HIDDEN,BIT6,INCLUDE_POWERFAIL,SHOW_VALUES,AUTO_VOICE_COMMAND,VOICE_COMMAND_CONFIRM,"
        header = header & "MYHS_DEVICE_CHANGE_NOTIFY,SET_DOES_NOT_CHANGE_LAST_CHANGE,BIT13,BIT14,BIT15,BIT16,"
        header = header & "NO_STATUS_TRIGGER,NO_GRAPHICS_DISPLAY,NO_STATUS_DISPLAY,CONTROL_POPUP,"
        ' Other values
        header = header & "Device_API,SubType,SubType_Description,ImageLarge,Note,"
        file.WriteLine(header)

        Do  'check each device that was enumerated
            dv = EN.GetNext
            If dv Is Nothing Then  'No device, so quit
                hs.WriteLog(label, "No devices found")
                Exit Sub
            End If

            'Status page stuff 
            Dim attrs As String = SafeString(dv.Ref(Nothing)) & SafeString(dv.Image(Nothing)) & SafeString(dv.devString(Nothing)) & SafeString(dv.Location2(Nothing))
            attrs = attrs & SafeString(dv.Location(Nothing)) & SafeString(dv.Name(Nothing)) & SafeString(dv.Address(Nothing)) & SafeString(dv.Device_Type_String(Nothing))
            attrs = attrs & SafeString(dv.Last_Change(Nothing)) & SafeString(dv.devValue(Nothing)) & SafeString(dv.Attention(Nothing)) & SafeString(dv.VoiceCommand(Nothing))
            'MISC bits
            If dv.MISC_Check(hs, 0) Then
                attrs = attrs & 1 & ","
            Else
                attrs = attrs & 0 & ","
            End If
            For i As Integer = 1 To 1048576
                If dv.MISC_Check(hs, i) Then
                    attrs = attrs & 1 & ","
                Else
                    attrs = attrs & 0 & ","
                End If
                i = i * 2
            Next
            ' Other values
            Dim DT As New DeviceTypeInfo
            DT = dv.DeviceType_Get(hs)
            If DT IsNot Nothing Then
                attrs = attrs & SafeString(DT.Device_API) & SafeString(DT.Device_SubType) & SafeString(DT.Device_SubType_Description)
            Else
                attrs = attrs & ",,,"
            End If
            attrs = attrs & SafeString(dv.ImageLarge(Nothing)) & SafeString(dv.UserNote(Nothing))
            file.WriteLine( attrs)
        Loop Until EN.Finished
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    Finally
        file.Close()
    End Try
    hs.WriteLog(label, "Done " & label)
End Sub
