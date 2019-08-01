' See https://homeseer.com/support/homeseer/HS3/SDK/default.htm for API info
' See https://github.com/avatar42/MyMonitor for info on MyMonitor
' Valid typeStrs are: cam, ptz, camA, ptzA, tivo, wu, ssh, cnt, humidity, plug, pressure, rssi, temp and web which is the default
Sub createMonDev(ByVal parms As Object)
    Dim dv As Scheduler.Classes.DeviceClass = Nothing
    Dim args() As String = Split(parms, ",")
    Dim label As String = "createMonDev"
    Dim name As String = "MyMonitorNew"
    Dim typeStr As String = "web"
    Dim dvRef
    
    if parms <> "" Then
        name = args(0)
        if  args(1)  <> "" then
            typeStr = args(1)
        End If
    End If
    hs.WriteLog(label,"name:" & name)

    try
        dvRef = hs.NewDeviceRef(name)
        hs.WriteLog(label,"Created:" & dvRef)
        fixMonDev(dvRef & "," & typeStr)
        dv = hs.GetDeviceByRef(dvRef)
        ' make easy to find in case there is one left over of a diff type
        dv.Location2(hs) = typeStr
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
    Dim typeStr As String = "web"
    Dim dvRef = 0
    if parms <> "" Then
        dvRef = args(0)
        if  args(1)  <> "" then
            typeStr = args(1)
        End If
    End If

    try
        dv = hs.GetDeviceByRef(dvRef)

        If Not dv Is Nothing Then

            hs.WriteLog(label,"Updating:" & dv.ref(hs))
            dv.Last_Change(hs) = Now
            hs.SetDeviceValueByRef(dvRef, 0, TRUE)
            hs.SetDeviceString(dvRef,"",TRUE)
            dv.Location(hs) = "MyMonitor" ' Room in my system
            if String.Equals(typeStr,"cam") Or String.Equals(typeStr,"camA") Then
                dv.Image(hs) = "/images/blueiris/fixedcam.png"
            Else if String.Equals(typeStr,"ptz") Or String.Equals(typeStr,"ptzA") Then
                dv.Image(hs) = "/images/blueiris/ptzcam.png"
            Else if String.Equals(typeStr,"tivo") Then
                dv.Image(hs) = "/images/hspi_ultramon3/tivo_online.png"
            Else if String.Equals(typeStr,"plug") Then
                dv.Image(hs) = "/images/Devices/Etekcity_small.jpg"
                dv.Location2(hs) = "Power" ' Category in my system
            Else if String.Equals(typeStr,"cnt") Then
                dv.Image(hs) = "/images/HomeSeer/status/counters_small.png"
            Else
                dv.Image(hs) = "/images/HomeSeer/status/armed-away_small.png"
            End If

            if String.Equals(typeStr,"plug") Then
                dv.Location2(hs) = "Power" 
            Else
                dv.Location2(hs) = "MyMonitor" 
            End If

            hs.WriteLog(label,"Updated Location:" & dv.Location(hs) & ":" & dv.Location2(hs))

            'for child ref in X10 world this would be like A1 in Z-wave Q24
            dv.Code(hs) = "" 
            hs.WriteLog(label,"Updated Code:" & dv.Code(hs))

            ' set to be monitored by ChkSensors
            dv.MISC_Clear(hs, HomeSeerAPI.Enums.dvMISC.SET_DOES_NOT_CHANGE_LAST_CHANGE)
            dv.MISC_Set(hs, HomeSeerAPI.Enums.dvMISC.SHOW_VALUES)
            hs.WriteLog(label,"Updated flags:" & dvRef)

            ' for easy sorting and filtering
            dv.Address(hs) = "MyMonitor-" & dvRef
            hs.WriteLog(label,"Updated Address:" & dv.Address(hs))
            dv.Device_Type_String(hs) = "MyMonitor-" & typeStr
            hs.WriteLog(label,"Updated type:" & dv.Device_Type_String(hs))

            'dv.Relationship(hs) = HomeSeerAPI.Enums.eRelationship.Parent_Root
            dv.Relationship(hs) = HomeSeerAPI.Enums.eRelationship.Not_Set 'matches what manual create does.

            ' Setup status values with graphics
            hs.DeviceVSP_ClearAll(dvRef, True)
            hs.DeviceVGP_ClearAll(dvRef, True)

            if String.Equals(typeStr,"cam") Or String.Equals(typeStr,"camA") Or String.Equals(typeStr,"ptz") Or String.Equals(typeStr,"ptzA") Then
                GenSingleStatus(dvRef,0,"Offline","/images/HomeSeer/status/modeoff.png")
                ' values set to mostly match match the Blue Iris plugin though it was inconsistent between fixed and PTZ cams
                GenSingleStatus(dvRef,1,"Connected","/images/blueiris/green.png")
                GenSingleStatus(dvRef,2,"Not Connected","/images/blueiris/notconnected.png")
                GenSingleStatus(dvRef,3,"Triggered","/images/HomeSeer/ui/Warning.png")
                GenSingleStatus(dvRef,4,"Recording","/images/blueiris/recording.png")
                GenSingleStatus(dvRef,5,"Motion","/images/blueiris/motion.png")
                GenSingleStatus(dvRef,6,"Paused","/images/blueiris/paused.png")
                GenSingleStatus(dvRef,7,"Hidden","/images/HomeSeer/ui/hide_marked.png")
                GenSingleStatus(dvRef,8,"Temp Full","/images/hspi_ultranetatmo3/battery_full.png")
                GenSingleStatus(dvRef,9,"No Signal","/images/blueiris/red.png")
                GenSingleStatus(dvRef,10,"Disabled","/images/blueiris/disabled.png")
                GenSingleStatus(dvRef, 11, "Low fps", "/images/blueiris/yellow.png")
            Else if String.Equals(typeStr,"cnt") Then
                GenRangeCtl( dvRef, 0, 10000,"",4,"","  μg/m3")
                ' pm2.5 0, 50, 100, 150, 200, 300
                GenRangeStatusValue(dvRef, 0, 12, "/images/hspi_ultranetatmo3/green.png")
                GenRangeStatusValue(dvRef, 12.001, 35.4999, "/images/Nest/cool.png")
                GenRangeStatusValue(dvRef, 35.5, 55.4999, "/images/Nest/warning.png")
                GenRangeStatusValue(dvRef, 55.5, 150.4999, "/images/Nest/emergency.png")
                GenRangeStatusValue(dvRef, 150.5, 250.4999, "/images/hspi_ultranetatmo3/red.png")
                GenRangeStatusValue(dvRef, 250.5, 10000, "/images/hspi_ultranetatmo3/flag_red.png")

                ' pm10.0 0, 50, 100, 150, 200, 300
                'GenRangeStatusValue(dvRef, 0, 54.999, "/images/hspi_ultranetatmo3/green.png")'0
                'GenRangeStatusValue(dvRef, 55, 154.999, "/images/Nest/cool.png")'50
                'GenRangeStatusValue(dvRef, 150, 254.999, "/images/Nest/warning.png")'100
                'GenRangeStatusValue(dvRef, 255, 354.999, "/images/Nest/emergency.png")'150
                'GenRangeStatusValue(dvRef, 355, 424.999, "/images/hspi_ultranetatmo3/red.png")'200
                'GenRangeStatusValue(dvRef, 425, 10000, "/images/hspi_ultranetatmo3/flag_red.png")'300
            Else if String.Equals(typeStr,"humidity") Then
                GenRangeCtl( dvRef, 0, 100,"",1,""," %")
                GenRangeStatusValue(dvRef, 0, 100, "/images/hspi_ultranetatmo3/humidity.png")
            Else if String.Equals(typeStr,"plug") Then
                GenSingleStatus(dvRef,0,"Off","	/images/HomeSeer/status/off.gif")
                GenSingleStatus(dvRef,50,"Offline","/images/HomeSeer/status/unknown.png")
                GenSingleStatus(dvRef,100,"On","/images/HomeSeer/status/on.gif")
            Else if String.Equals(typeStr,"pressure") Then
                GenRangeCtl( dvRef, 0, 1084,"",1,""," inHg")
                GenRangeStatusValue(dvRef, 0, 100, "/images/hspi_ultranetatmo3/pressure.png")
            Else if String.Equals(typeStr,"rssi") Then
                GenRangeCtl( dvRef, -100, 0,"",0,""," db")
                GenRangeStatusValue(dvRef, -100, -90, "/images/hspi_ultranetatmo3/signal_low.png")
                GenRangeStatusValue(dvRef, -89, -70, "/images/hspi_ultranetatmo3/signal_medium.png")
                GenRangeStatusValue(dvRef, -69, -25, "/images/hspi_ultranetatmo3/signal_high.png")
                GenRangeStatusValue(dvRef, -24, -1, "/images/hspi_ultranetatmo3/signal_full.png")
                GenSingleStatusValue(dvRef, 0, "/images/HomeSeer/status/unknown.png")
            Else if String.Equals(typeStr,"temp") Then
                GenRangeCtl( dvRef, -10, 120,"",1,""," °F")
                GenRangeStatusValue(dvRef, -10, 120, "/images/hspi_ultranetatmo3/temperature.png")
            ElseIf String.Equals(typeStr,"tivo") Then
                GenSingleStatus(dvRef,0,"Offline","/images/hspi_ultramon3/tivo_offline.png")
                GenRangeStatus(dvRef,1,199,"Other-Error","/images/hspi_ultramon3/tivo_troubled.png",0,"","")
                GenSingleStatus(dvRef,200,"Online-OK","/images/hspi_ultramon3/tivo_online.png")
                GenRangeStatus(dvRef,201,999,"Other-Error","/images/hspi_ultramon3/tivo_troubled.png",0,"","")
            Else
                GenRangeCtl( dvRef, 0, 999,"",0,"","")
                GenSingleStatusValue(dvRef,0,"/images/HomeSeer/status/modeoff.png")
                GenRangeStatusValue(dvRef,1,199,"/images/HomeSeer/status/alarm.png")
                ' web response codes and custom errors
                GenSingleStatusValue(dvRef,200,"/images/HomeSeer/status/ok_small.png")
                GenRangeStatusValue(dvRef,300,999,"/images/HomeSeer/status/alarmco.png")
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
            DT.Device_SubType_Description = "" '"MyMonitor"
            dv.DeviceType_Set(hs) = DT

            ' TODO make device types work
            ' Dim DT As New DeviceAPI.DeviceTypeInfo
            ' DT.Device_API = DeviceAPI.DeviceTypeInfo.eDeviceAPI.Plug_In
            ' 'DT.Device_Type = CInt(99)
            ' 'DT.Device_SubType = 4
            ' dv.DeviceType_Set(hs) = DT
            hs.WriteLog(label,"Updated type:" & dv.DeviceType_Get(hs).Device_Type_Description)

            hs.SaveEventsDevices 
        Else 
            hs.WriteLog(label,"Failed to update:" & dvRef)
        End If
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try

End Sub

Public Sub GenSingleStatus(ByVal dvRef As Integer,ByVal val As Double,ByVal status As String,ByVal image As String)
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

Public Sub GenRangeStatus(ByVal dvRef As Integer,ByVal rangeStart As Double,ByVal rangeEnd As Double,ByVal status As String,ByVal image As String,ByVal decimals As Integer,ByVal prefix As String,ByVal suffix As String)
    
    GenRangeCtl(dvRef, rangeStart, rangeEnd, status, decimals,prefix,suffix)  

    GenRangeStatusValue(dvRef, rangeStart, rangeEnd, image)
End Sub

Public Sub GenRangeCtl(ByVal dvRef As Integer,ByVal rangeStart As Double,ByVal rangeEnd As Double,ByVal status As String,ByVal decimals As Integer,ByVal prefix As String,ByVal suffix As String)
    Dim Pair As VSPair
    ' want ePairStatusControl.Both so can change via URL
    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
    Pair.PairType = VSVGPairType.Range
    Pair.RangeStart = rangeStart
    Pair.RangeEnd = rangeEnd
    Pair.RangeStatusDecimals=decimals 'default is 0 (integers only)
    Pair.Render = Enums.CAPIControlType.TextBox_Number
    Pair.Render_Location.Row = 1
    Pair.Render_Location.Column = 1
    Pair.RangeStatusPrefix = prefix
    Pair.RangeStatusSuffix = suffix

    Pair.Status = status
    hs.DeviceVSP_AddPair(dvRef, Pair)
    
End Sub

Public Sub GenRangeStatusValue(ByVal dvRef As Integer,ByVal rangeStart As Double,ByVal rangeEnd As Double,ByVal image As String)
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

Public Sub GenSingleStatusValue(ByVal dvRef As Integer,ByVal val As Double,ByVal image As String)
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
