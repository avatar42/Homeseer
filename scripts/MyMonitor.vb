' Checks script complies and globals used, are defined
Sub Main(ByVal ignored As String)
    hs.speakEx(0, "My Monitor Script compiled OK", False)
End Sub

' link to common sayString method
Sub sayString(ByVal msg As String)
        hs.RunScriptFunc("SayIt.vb", "sayString", msg, False, False)
End Sub

' See https://homeseer.com/support/homeseer/HS3/SDK/default.htm for API info
' See https://github.com/avatar42/MyMonitor for info on MyMonitor
' Valid typeStrs are: cam, ptz, camA, ptzA, tivo, ssh, cnt, humidity, plug, pressure, rssi, 
' temp and web which is the default
Sub createMonDev(ByVal parms As Object)
    Dim dv As Scheduler.Classes.DeviceClass = Nothing
    Dim args() As String = Split(parms, ",")
    Dim label As String = "createMonDev"
    ' name used should be unique temp only name tp avoid changing the wrong one.
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
        dvRef = hs.GetDeviceRefByName(name)
        ' only create if there is not one waiting, create one
        if dvRef > 0 Then
            hs.WriteLog(label,"found:" & dvRef)
        Else
            dvRef = hs.NewDeviceRef(name)
            hs.WriteLog(label,"Created:" & dvRef)
        End If
        fixMonDev(dvRef & "," & typeStr)
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try

End Sub

' Get standard location for type typeStr
Public Function type2loc(ByVal typeStr As String) As String
    Dim loc As String = "MyMonitor"
            if String.Equals(typeStr,"battery") Then
                loc = "Batteries" 
                ' Uses pic of object battery is in
            Else if String.Equals(typeStr,"cam") Or String.Equals(typeStr,"camA") Then
                loc = "BlueIris" 
            Else if String.Equals(typeStr,"ptz") Or String.Equals(typeStr,"ptzA") Then
                loc = "BlueIris" 
            Else if String.Equals(typeStr,"tivo") Then
                loc = "Tivo" 
            Else if String.Equals(typeStr,"plug") Then
                loc = "Power" 
            Else if String.Equals(typeStr,"cnt") Then
            Else if String.Equals(typeStr,"motion") Then
                loc = "Motion" 
            Else if String.Equals(typeStr,"pressure") Then
                loc = "Pressure" 
            Else if String.Equals(typeStr,"humidity") Then
                loc = "Humidity" 
            Else if String.Equals(typeStr,"rssi") Then
                loc = "Signal" 
            Else if String.Equals(typeStr,"solar") Then
                loc = "Light" 
            Else if String.Equals(typeStr,"temp") Then
                loc = "Temperature" 
            Else if String.Equals(typeStr,"wind") Then
                loc = "Wind" 
            Else if String.Equals(typeStr,"dir") Then
                loc = "Wind" 
            End If
    return loc
End Function

' Changes device with ref=dvRef to the type typeStr
' parms is expected to be String in format of dvRef,typeStr
Public Sub fixMonDev(ByVal parms As Object)

    Dim args() As String = Split(parms, ",")
    Dim label As String = "fixMonDev"
    Dim dv As Scheduler.Classes.DeviceClass = Nothing
    Dim level As Integer = 0
    Dim root_dv As Scheduler.Classes.DeviceClass = Nothing
    Dim typeStr As String = "web"
    Dim dvRef = 0
    hs.WriteLog(label,"parms:" & parms)
    if parms <> "" Then
        dvRef = args(0)
        if  args(1)  <> "" then
            typeStr = args(1)
        End If
    End If

    try
        hs.WriteLog(label,"Updating:" & dvRef & ":" & typeStr)
        dv = hs.GetDeviceByRef(dvRef)

        If Not dv Is Nothing Then

            hs.WriteLog(label,"Updating:" & dv.ref(hs))
            dv.Last_Change(hs) = Now
            hs.SetDeviceValueByRef(dvRef, 0, TRUE)
            hs.SetDeviceString(dvRef,"",TRUE)
            dv.Location(hs) = "MyMonitor" ' Room in my system

            ' set default thumbnail pic and category
            dv.Location2(hs) = type2loc(typeStr)
            hs.WriteLog("Info","The API interface version of HomeSeer is " & hs.InterfaceVersion.ToString)
            if (hs.InterfaceVersion < 4) Then
                if String.Equals(typeStr,"battery") Then
                    ' Uses pic of object battery is in
                Else if String.Equals(typeStr,"cam") Or String.Equals(typeStr,"camA") Then
                    dv.Image(hs) = "/images/blueiris/fixedcam.png"
                Else if String.Equals(typeStr,"ptz") Or String.Equals(typeStr,"ptzA") Then
                    dv.Image(hs) = "/images/blueiris/ptzcam.png"
                Else if String.Equals(typeStr,"tivo") Then
                    dv.Image(hs) = "/images/hspi_ultramon3/tivo_online.png"
                Else if String.Equals(typeStr,"plug") Then
                    dv.Image(hs) = "/images/Devices/Etekcity_small.jpg"
                Else if String.Equals(typeStr,"cnt") Then
                    dv.Image(hs) = "/images/HomeSeer/status/counters_small.png"
                Else if String.Equals(typeStr,"motion") Then
                    ' Uses pic of object is from
                Else if String.Equals(typeStr,"pressure") Then
                    ' Uses pic of object is from
                Else if String.Equals(typeStr,"humidity") Then
                    ' Uses pic of object is from
                Else if String.Equals(typeStr,"rssi") Then
                    ' Uses pic of object is from
                Else if String.Equals(typeStr,"solar") Then
                    ' Uses pic of object is from
                Else if String.Equals(typeStr,"temp") Then
                    ' Uses pic of object is from
                Else if String.Equals(typeStr,"wind") Then
                    ' Uses pic of object is from
                Else if String.Equals(typeStr,"dir") Then
                    ' Uses pic of object is from
                End If
            Else
                hs.WriteLog(label,"Skipping image add for hs4")
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

            if String.Equals(typeStr,"battery") Then
                GenRangeCtl( dvRef, 0, 100,"",4,"","%")
                GenRangeCtl( dvRef, 101, 254,"",4,""," error")
                GenSingleStatus(dvRef,255,"Battery Low Warning","/images/HomeSeer/status/battery_0.png")
                GenRangeStatusValue(dvRef, 0, 3, "/images/HomeSeer/status/battery_0.png")
                GenRangeStatusValue(dvRef, 4, 36, "/images/HomeSeer/status/battery_25.png")
                GenRangeStatusValue(dvRef, 37, 64, "/images/HomeSeer/status/battery_50.png")
                GenRangeStatusValue(dvRef, 65, 89, "/images/HomeSeer/status/battery_75.png")
                GenRangeStatusValue(dvRef, 90, 100, "/images/HomeSeer/status/battery_100.png")
                GenRangeStatusValue(dvRef, 101, 254, "/images/HomeSeer/status/battery_0.png")

            else if String.Equals(typeStr,"cam") Or String.Equals(typeStr,"camA") Or String.Equals(typeStr,"ptz") Or String.Equals(typeStr,"ptzA") Then
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
                GenSingleStatus(dvRef,0,"Off","/images/HomeSeer/status/off.gif")
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
            ElseIf String.Equals(typeStr,"motion") Then
            'Wireless Tag and Wyze
                GenSingleStatus(dvRef,0,"Disarmed","/images/HomeSeer/status/autolocked.png")
                GenSingleStatus(dvRef,1,"Armed","/images/HomeSeer/status/armed-stay.png")
                GenSingleStatus(dvRef,2,"Moved","/images/HomeSeer/status/Garage-Opening.png")
                GenSingleStatus(dvRef,3,"Opened","/images/HomeSeer/status/open.png")
                GenSingleStatus(dvRef,4,"Closed","/images/HomeSeer/status/closed.png")
                GenSingleStatus(dvRef,5,"DetectedMovement","/images/HomeSeer/status/motion.gif")
                GenSingleStatus(dvRef,6,"TimedOut","/images/HomeSeer/status/unknown.png")

            ElseIf String.Equals(typeStr,"rain") Then
                 GenRangeStatus(dvRef,0,1000,"Rain","/images/hspi_ultraweatherwu3/rain.png",2,""," in")
            ElseIf String.Equals(typeStr,"wind") Then
                 GenRangeStatus(dvRef,0,250 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,""," mph")
            ElseIf String.Equals(typeStr,"solar") Then
                 GenRangeStatus(dvRef,0,2000 ,"","/images/HomeSeer/status/radiation.png",2,"","  W/m^2")
            ElseIf String.Equals(typeStr,"dir") Then
                 GenRangeStatus(dvRef,0,11.24 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"North ","")
                 GenRangeStatus(dvRef,11.25,33.74 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"North-northeast ","")
                 GenRangeStatus(dvRef,33.75,56.24 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"Northeast ","")
                 GenRangeStatus(dvRef,56.25,78.24 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"East-northeast ","")
                 GenRangeStatus(dvRef,78.25,101.24 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"East ","")
                 GenRangeStatus(dvRef,101.25,123.74 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"East-southeast ","")
                 GenRangeStatus(dvRef,123.75,146.24 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"Southeast ","")
                 GenRangeStatus(dvRef,146.25,168.74 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"South-southeast ","")
                 GenRangeStatus(dvRef,168.75,191.24 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"South ","")
                 GenRangeStatus(dvRef,191.25,213.74 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"South-southwest ","")
                 GenRangeStatus(dvRef,213.75,236.24 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"Southwest ","")
                 GenRangeStatus(dvRef,236.25,258.74 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"West-southwest ","")
                 GenRangeStatus(dvRef,258.75,281.24 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"West ","")
                 GenRangeStatus(dvRef,281.25,303.74 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"West-northwest ","")
                 GenRangeStatus(dvRef,303.75,326.24 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"Northwest ","")
                 GenRangeStatus(dvRef,326.25,348.74 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"North-northwest ","")
                 GenRangeStatus(dvRef,348.75,360 ,"Wind","/images/hspi_ultraweatherwu3/wind.png",2,"North ","")
            Else ' web, ssh, wu
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

            ' update any related events with changes
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
    hs.WriteLog("GenSingleStatus","Adding status cntl:" &  val & " as Button " & status)
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
    hs.WriteLog("GenRangeCtl","Adding range cntl:" & prefix & rangeStart & "-" & rangeEnd & "." & decimals & " as TextBox_Number " & suffix)
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
        hs.WriteLog("GenRangeStatusValue","Adding status value:" & rangeStart & "-" & rangeEnd & " as " & image)
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
        hs.WriteLog("GenSingleStatusValue","Adding status value:" & val & " as " & image)
        GPair = New VGPair
        GPair.PairType = VSVGPairType.SingleValue
        GPair.Set_Value = val
        GPair.Graphic = image
        hs.DeviceVGP_AddPair(dvRef, GPair)
    End If
End Sub

' Look for and remove extra MyMonitorNew objects. 
' Nice to have one in reserve to avoid race condition but more than one is not useful.
Public Sub chkDups(ignored As Object)
    Dim label As String = "chkDups"
    Dim found As Integer = 0
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
            If InStr(dv.Name(Nothing), "MyMonitorNew") > 0 Then
                hs.WriteLog(label, "device:" & dv.Ref(Nothing) & ":" & dv.Location2(Nothing) & " " & dv.Location(Nothing) & " " & dv.Name(Nothing) & " with value:" & dv.devValue(Nothing) & " updated:" & dv.Last_Change(Nothing))
                found = found + 1
                If found > 1 Then
                    hs.DeleteDevice(dv.Ref(Nothing))
                End If
            End If
        Loop Until EN.Finished
        If found > 1 Then
            found = found - 1
            sayString("Removed " & found & " extra My Monitor New objects")
        End If
        hs.WriteLog(label, "Removed " & found & " extra My Monitor New objects")
    Catch ex As Exception
            hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Sub
