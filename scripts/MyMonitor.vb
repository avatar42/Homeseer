' See https://homeseer.com/support/homeseer/HS3/SDK/default.htm for API info
' See https://github.com/avatar42/MyMonitor for info on MyMonitor

Sub createMonitorDevice(params As Object)
    Dim label As String = "createMonitorDevice"
    Dim dv As Scheduler.Classes.DeviceClass = Nothing
    Dim name As String = "MyMonitorNew"
    Dim dvRef
    Dim level As Integer = 0
    Dim root_dv As Scheduler.Classes.DeviceClass = Nothing


    if Not params Is Nothing Then
        name = params
    End If
    hs.WriteLog(label,"name:" & name)
    try
        dvRef = hs.NewDeviceRef(name)
        hs.WriteLog(label,"Created:" & dvRef)
        dv = hs.GetDeviceByRef(dvRef)

        If Not dv Is Nothing Then

            hs.WriteLog(label,"Updating:" & dv.ref(hs))
            dv.Last_Change(hs) = Now
            hs.SetDeviceValueByRef(dvRef, 0, TRUE)
            dv.Image(hs) = "/images/HomeSeer/status/armed-away_small.png"

            dv.Location(hs) = "MyMonitor"
            dv.Location2(hs) = "MyMonitor"
            hs.WriteLog(label,"Updated Location:" & dv.Location(hs) & ":" & dv.Location2(hs))

            'unused but here for ref
            dv.Code(hs) = "0"
            hs.WriteLog(label,"Updated Code:" & dv.Code(hs))

            ' set to be monitored by ChkSensors
            dv.MISC_Clear(hs, HomeSeerAPI.Enums.dvMISC.SET_DOES_NOT_CHANGE_LAST_CHANGE)
            hs.WriteLog(label,"Updated flags:" & dvRef)

            ' for easy sorting and filtering
            dv.Address(hs) = "MyMonitor-" & dvRef
            hs.WriteLog(label,"Updated Address:" & dv.Address(hs))
            dv.Device_Type_String(hs) = "MyMonitor"
            hs.WriteLog(label,"Updated type:" & dv.Device_Type_String(hs))

            dv.Relationship(hs) = HomeSeerAPI.Enums.eRelationship.Parent_Root

            GenSingleStatus(dvRef,0,"Offline","/images/HomeSeer/status/modeoff.png")

            GenRangeStatus(dvRef,1,199,"Other-Error","/images/HomeSeer/status/alarm.png")

            GenSingleStatus(dvRef,200,"Online-OK","/images/HomeSeer/status/ok_small.png")

            GenRangeStatus(dvRef,300,999,"Online-Error","/images/HomeSeer/status/alarmco.png")

            ' TODO make device types work
            ' Dim DT As New DeviceAPI.DeviceTypeInfo
            ' DT.Device_API = DeviceAPI.DeviceTypeInfo.eDeviceAPI.Plug_In
            ' 'DT.Device_Type = CInt(99)
            ' 'DT.Device_SubType = 4
            ' dv.DeviceType_Set(hs) = DT
            ' hs.WriteLog(label,"Updated type:" & dv.DeviceType_Get(hs).Device_Type_Description)

            ' if want to make child of other device 
            If Not root_dv Is Nothing Then 
                root_dv.AssociatedDevice_Add(hs, dvRef)
                dv.Relationship(hs) = Enums.eRelationship.Child
                dv.AssociatedDevice_Add(hs, root_dv.Ref(hs))
            End If

            hs.SaveEventsDevices 
        Else 
            hs.WriteLog(label,"Failed to create:" & name)
        End If
    Catch ex As Exception
        hs.WriteLog(label, "Exception in script " & label & ":  " & ex.Message)
    End Try

End Sub

Public Sub GenSingleStatus(ByVal dvRef As Integer,ByVal val As Integer,ByVal status As String,ByVal image As String)
    Dim Pair As VSPair
    Dim GPair As VGPair
    ' want ePairStatusControl.Both so can change via URL
    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
    Pair.PairType = VSVGPairType.SingleValue
    Pair.Value = val
    Pair.Status = status
    hs.DeviceVSP_AddPair(dvRef, Pair)

    GPair = New VGPair
    GPair.PairType = VSVGPairType.SingleValue
    GPair.Set_Value = val
    GPair.Graphic = image
    hs.DeviceVGP_AddPair(dvRef, GPair)
End Sub

Public Sub GenRangeStatus(ByVal dvRef As Integer,ByVal rangeStart As Integer,ByVal rangeEnd As Integer,ByVal status As String,ByVal image As String)
    Dim Pair As VSPair
    Dim GPair As VGPair
    ' want ePairStatusControl.Both so can change via URL
    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
    Pair.PairType = VSVGPairType.Range
    Pair.RangeStart = rangeStart
    Pair.RangeEnd = rangeEnd
    Pair.Status = status
    hs.DeviceVSP_AddPair(dvRef, Pair)
    
    GPair = New VGPair
    GPair.PairType = VSVGPairType.Range
    GPair.RangeStart = rangeStart
    GPair.RangeEnd = rangeEnd
    GPair.Graphic = image
    hs.DeviceVGP_AddPair(dvRef, GPair)
End Sub
