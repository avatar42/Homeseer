' Checks script complies and globals used, are defined
Sub Main(ByVal ignored As String)
    hs.speakEx(0, "Create W M I objects script compiled OK", False)
End Sub
' See https://automation.rmrr42.com/2018/06/use-windows-wmi-and-powershell-to-send.html for example of usage and WMI agent script.
' creates a base set of WMI virtual objects for a hostname
Sub createWMISet(ByVal host As Object)
    ' CPU temp - note not supported on all systems. Often unchanging value
    createMonDev(host & ",CPU,")
    ' CPU load %
    createMonDev(host & ",Load,")
    ' % RAM free
    createMonDev(host & ",RAM,")
    ' % Log size in MB
    createMonDev(host & ",Log,Size")
    ' Updates in pending state
    createMonDev(host & ",Updates,Pending")
    ' % drive space free
    createMonDev(host & ",Drive,C")
End Sub

' Valid typeStrs are: CPU, Drive, Load, RAM, Log, Updates
' parms format is host,typeStr,detail as in
' Hero,CPU,
' Hero,Drive,C
Sub createMonDev(ByVal parms As Object)
    Dim dv As Scheduler.Classes.DeviceClass = Nothing
    Dim args() As String = Split(parms, ",")
    Dim label As String = "createMonDev"
    Dim name As String = "WmiNew"
    Dim host As String = "localhost"
    Dim typeStr As String = "CPU"
    Dim detail As String = ""
    Dim dvRef

    If parms <> "" Then
        host = args(0)
        If args(1) <> "" Then
            typeStr = args(1)
        End If
        name = host & " " & typeStr
        If args(2) <> "" Then
            detail = args(2)
            name = name & " " & detail
        End If
    End If
    hs.WriteLog(label, "name:" & name)

    Try
        dvRef = hs.NewDeviceRef(name)
        hs.WriteLog(label, "Created:" & dvRef)
        fixMonDev(dvRef & "," & typeStr)
        dv = hs.GetDeviceByRef(dvRef)
        dv.Location(hs) = host ' Room in my system
        dv.Location2(hs) = "WMI"
        hs.WriteLog(label, "Updated Location:" & dv.Location(hs) & ":" & dv.Location2(hs))
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try

End Sub

Public Sub fixMonDev(ByVal parms As Object)
    Dim args() As String = Split(parms, ",")
    Dim label As String = "fixMonDev"
    Dim dv As Scheduler.Classes.DeviceClass = Nothing
    Dim level As Integer = 0
    Dim root_dv As Scheduler.Classes.DeviceClass = Nothing
    Dim typeStr As String = "CPU"
    Dim dvRef = 0
    If parms <> "" Then
        dvRef = args(0)
        If args(1) <> "" Then
            typeStr = args(1)
        End If
    End If

    Try
        dv = hs.GetDeviceByRef(dvRef)

        If Not dv Is Nothing Then

            hs.WriteLog(label, "Updating:" & dv.ref(hs))
            dv.Last_Change(hs) = Now
            hs.SetDeviceValueByRef(dvRef, 0, True)
            hs.SetDeviceString(dvRef, "", True)

            'for child ref in X10 world this would be like A1 in Z-wave Q24
            dv.Code(hs) = ""
            hs.WriteLog(label, "Updated Code:" & dv.Code(hs))

            ' set to be monitored by ChkSensors
            dv.MISC_Clear(hs, HomeSeerAPI.Enums.dvMISC.SET_DOES_NOT_CHANGE_LAST_CHANGE)
            dv.MISC_Set(hs, HomeSeerAPI.Enums.dvMISC.SHOW_VALUES)
            hs.WriteLog(label, "Updated flags:" & dvRef)

            ' for easy sorting and filtering in HS3
            dv.Address(hs) = "WMI-" & dvRef
            hs.WriteLog(label, "Updated Address:" & dv.Address(hs))
            dv.Device_Type_String(hs) = "WMI-" & typeStr
            hs.WriteLog(label, "Updated type:" & dv.Device_Type_String(hs))

            'dv.Relationship(hs) = HomeSeerAPI.Enums.eRelationship.Parent_Root
            dv.Relationship(hs) = HomeSeerAPI.Enums.eRelationship.Not_Set 'matches what manual create does.

            ' Setup status values with graphics
            hs.DeviceVSP_ClearAll(dvRef, True)
            hs.DeviceVGP_ClearAll(dvRef, True)

            If String.Equals(typeStr, "CPU") Then
                ' set default thumbnail pic and category
                ' not used in HS4
                dv.Image(hs) = "/images/Devices/PC/cpu-512.png"
                GenRangeCtl(dvRef, -1, 120, "", 4, "Temp ", "C")
                GenSingleStatus(dvRef, -1, "error", "/images/HomeSeer/status/alarm.png")
                GenRangeStatusValue(dvRef, 0, 25.999, "/images/HomeSeer/status/Cool.png")
                GenRangeStatusValue(dvRef, 26, 35.999, "/images/HomeSeer/status/fan-auto.png")
                GenRangeStatusValue(dvRef, 36, 120, "/images/HomeSeer/status/Heat.png")
            ElseIf String.Equals(typeStr, "Drive") Then
                ' set default thumbnail pic and category
                ' not used in HS4
                dv.Image(hs) = "/images/Devices/PC/hard_drive.png"
                GenRangeCtl(dvRef, -1, 100, "", 4, "Free ", "%")
                GenSingleStatus(dvRef, -1, "error", "/images/HomeSeer/status/alarm.png")
                GenRangeStatusValue(dvRef, 0, 10.999, "/images/HomeSeer/status/battery_0.png")
                GenRangeStatusValue(dvRef, 11, 25.999, "/images/HomeSeer/status/battery_25.png")
                GenRangeStatusValue(dvRef, 26, 50.999, "/images/HomeSeer/status/battery_50.png")
                GenRangeStatusValue(dvRef, 51, 75.999, "/images/HomeSeer/status/battery_75.png")
                GenRangeStatusValue(dvRef, 76, 100, "/images/HomeSeer/status/battery_100.png")
            ElseIf String.Equals(typeStr, "RAM") Then
                ' set default thumbnail pic and category
                ' not used in HS4
                dv.Image(hs) = "/images/Devices/PC/ram-icon-drawing_csp50084892.jpg"
                GenRangeCtl(dvRef, -1, 100, "", 4, "Free ", "%")
                GenSingleStatus(dvRef, -1, "error", "/images/HomeSeer/status/alarm.png")
                GenRangeStatusValue(dvRef, 0, 10.999, "/images/HomeSeer/status/battery_0.png")
                GenRangeStatusValue(dvRef, 11, 25.999, "/images/HomeSeer/status/battery_25.png")
                GenRangeStatusValue(dvRef, 26, 50.999, "/images/HomeSeer/status/battery_50.png")
                GenRangeStatusValue(dvRef, 51, 75.999, "/images/HomeSeer/status/battery_75.png")
                GenRangeStatusValue(dvRef, 76, 100, "/images/HomeSeer/status/battery_100.png")
            ElseIf String.Equals(typeStr, "Load") Then
                ' set default thumbnail pic and category
                ' not used in HS4
                dv.Image(hs) = "/images/Devices/PC/meter.png"
                GenRangeCtl(dvRef, -1, 100, "", 4, "Load ", "%")
                GenSingleStatus(dvRef, -1, "error", "/images/HomeSeer/status/alarm.png")
                GenRangeStatusValue(dvRef, 0, 10.999, "/images/HomeSeer/status/battery_100.png")
                GenRangeStatusValue(dvRef, 11, 25.999, "/images/HomeSeer/status/battery_75.png")
                GenRangeStatusValue(dvRef, 26, 50.999, "/images/HomeSeer/status/battery_50.png")
                GenRangeStatusValue(dvRef, 51, 75.999, "/images/HomeSeer/status/battery_25.png")
                GenRangeStatusValue(dvRef, 76, 100, "/images/HomeSeer/status/battery_0.png")
            ElseIf String.Equals(typeStr, "Log") Then
                ' set default thumbnail pic and category
                ' not used in HS4
                dv.Image(hs) = "/images/Devices/PC/log.png"
                GenRangeCtl(dvRef, -1, 100, "", 4, "Log Size ", "MB")
                GenSingleStatus(dvRef, -1, "error", "/images/HomeSeer/status/alarm.png")
                GenRangeStatusValue(dvRef, 0, 10.999, "/images/HomeSeer/status/battery_100.png")
                GenRangeStatusValue(dvRef, 11, 25.999, "/images/HomeSeer/status/battery_75.png")
                GenRangeStatusValue(dvRef, 26, 50.999, "/images/HomeSeer/status/battery_50.png")
                GenRangeStatusValue(dvRef, 51, 75.999, "/images/HomeSeer/status/battery_25.png")
                GenRangeStatusValue(dvRef, 76, 100, "/images/HomeSeer/status/battery_0.png")
            ElseIf String.Equals(typeStr, "Updates") Then
                ' set default thumbnail pic and category
                ' not used in HS4
                dv.Image(hs) = "/images/Devices/PC/downloading-updates.png"
                GenRangeCtl(dvRef, -1, 20, "", 0, "Updates ", "")
                GenSingleStatus(dvRef, -1, "error", "/images/HomeSeer/status/alarm.png")
                GenRangeStatusValue(dvRef, 0, 2, "/images/HomeSeer/status/ok.png")
                GenRangeStatusValue(dvRef, 3, 4, "/images/HomeSeer/status/clockwise-start.png")
                GenRangeStatusValue(dvRef, 5, 20, "/images/HomeSeer/status/clockwise-stop.png")
            Else
                hs.WriteLog(label, "Unknown type:'" & typeStr & "'")
            End If


            ' if want to make child of other device 
            If Not root_dv Is Nothing Then
                root_dv.AssociatedDevice_Add(hs, dvRef)
                dv.Relationship(hs) = HomeSeerAPI.Enums.eRelationship.Child
                dv.AssociatedDevice_Add(hs, root_dv.Ref(hs))
            End If

            Dim DT As New DeviceTypeInfo
            DT.Device_API = DeviceTypeInfo.eDeviceAPI.No_API
            DT.Device_SubType = 0
            DT.Device_SubType_Description = "" '"WMI"
            dv.DeviceType_Set(hs) = DT

            ' TODO make device types work
            ' Dim DT As New DeviceAPI.DeviceTypeInfo
            ' DT.Device_API = DeviceAPI.DeviceTypeInfo.eDeviceAPI.Plug_In
            ' 'DT.Device_Type = CInt(99)
            ' 'DT.Device_SubType = 4
            ' dv.DeviceType_Set(hs) = DT
            hs.WriteLog(label, "Updated type:" & dv.DeviceType_Get(hs).Device_Type_Description)

            hs.SaveEventsDevices
        Else
            hs.WriteLog(label, "Failed to update:" & dvRef)
        End If
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try

End Sub

Public Sub GenSingleStatus(ByVal dvRef As Integer, ByVal val As Double, ByVal status As String, ByVal image As String)
    ' see https://www.homeseer.com/support/homeseer/HS3/HS3Help/vspair.htm
    Dim Pair As VSPair
    ' want ePairStatusControl.Both so can change via URL
    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
    Pair.PairType = VSVGPairType.SingleValue
    Pair.Value = val
    Pair.Status = status
    Pair.Render = Enums.CAPIControlType.Button
    Pair.Render_Location.Row = 1
    Pair.Render_Location.Column = 1
    'Pair.ControlUse = ePairControlUse._Off
    hs.DeviceVSP_AddPair(dvRef, Pair)

    GenSingleStatusValue(dvRef, val, image)
End Sub

Public Sub GenRangeStatus(ByVal dvRef As Integer, ByVal rangeStart As Double, ByVal rangeEnd As Double, ByVal status As String, ByVal image As String, ByVal decimals As Integer, ByVal prefix As String, ByVal suffix As String)

    GenRangeCtl(dvRef, rangeStart, rangeEnd, status, decimals, prefix, suffix)

    GenRangeStatusValue(dvRef, rangeStart, rangeEnd, image)
End Sub

Public Sub GenRangeCtl(ByVal dvRef As Integer, ByVal rangeStart As Double, ByVal rangeEnd As Double, ByVal status As String, ByVal decimals As Integer, ByVal prefix As String, ByVal suffix As String)
    Dim Pair As VSPair
    ' want ePairStatusControl.Both so can change via URL
    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
    Pair.PairType = VSVGPairType.Range
    Pair.RangeStart = rangeStart
    Pair.RangeEnd = rangeEnd
    Pair.RangeStatusDecimals = decimals 'default is 0 (integers only)
    Pair.Render = Enums.CAPIControlType.TextBox_Number
    Pair.Render_Location.Row = 1
    Pair.Render_Location.Column = 1
    Pair.RangeStatusPrefix = prefix
    Pair.RangeStatusSuffix = suffix

    Pair.Status = status
    hs.DeviceVSP_AddPair(dvRef, Pair)

End Sub

Public Sub GenRangeStatusValue(ByVal dvRef As Integer, ByVal rangeStart As Double, ByVal rangeEnd As Double, ByVal image As String)
    Dim GPair As VGPair

    If image <> "" Then
        GPair = New VGPair
        GPair.PairType = VSVGPairType.Range
        GPair.RangeStart = rangeStart
        GPair.RangeEnd = rangeEnd
        GPair.Graphic = image
        hs.DeviceVGP_AddPair(dvRef, GPair)
    End If
End Sub

Public Sub GenSingleStatusValue(ByVal dvRef As Integer, ByVal val As Double, ByVal image As String)
    ' see https://www.homeseer.com/support/homeseer/HS3/HS3Help/vgpair.htm
    Dim GPair As VGPair

    If image <> "" Then
        GPair = New VGPair
        GPair.PairType = VSVGPairType.SingleValue
        GPair.Set_Value = val
        GPair.Graphic = image
        hs.DeviceVGP_AddPair(dvRef, GPair)
    End If
End Sub
