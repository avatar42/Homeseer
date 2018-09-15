'load object refs and speech methods
#Include SayIt.vb
' import BlueIrisLogin def
#Include Secrets.vb

' Note it is also possible to call URLs directly form Blue Iris to set device values in Homeseer
' As in http://HS#IP:PORT/JSON?user=UserName&pass=Password&request=controldevicebyvalue&ref=RefID%&value=Value
' Though be sure the device you are trying to set the value on is set for control and can accept a value in the range you will be sending.

Dim TRIGGER As String = "trigger"
Dim SHOW As String = "show"

' convenience methods for calling camShow(ByVal host As String, ByVal port As integer,ByVal parms As String)
' see camShow
Public Sub Iris2Show(ByVal parms As String)
    camShow("Iris2",8078,parms)
End Sub

' see camShow
Public Sub Iris3Show(ByVal parms As String)
    camShow("Iris3",8076,parms)
End Sub

' see camShow
Public Sub Iris4Show(ByVal parms As String)
    camShow("Iris4",8077,parms)
End Sub

' Set all consoles back to All cameras
Public Sub consoleClear(ByVal notUsed As String)
    callBI("Iris2","All cameras",8078,SHOW)
    callBI("Iris3","All cameras",8076,SHOW)
    callBI("Iris4","All cameras",8077,SHOW)
End Sub

' Set console back to All cameras
Public Sub Iris2Clear(ByVal notUsed As String)
    callBI("Iris2","All cameras",8078,SHOW)
End Sub

' Set console back to All cameras
Public Sub Iris3Clear(ByVal notUsed As String)
    callBI("Iris3","All cameras",8076,SHOW)
End Sub

' Set console back to All cameras
Public Sub Iris4Clear(ByVal notUsed As String)
    callBI("Iris4","All cameras",8077,SHOW)
End Sub

' Triggers an alert on camName And highlights the group / switches main video based on the following
' host = Blue Iris server's name
' port = Blue Iris servr's web API port
' parms = camName,grpName,highlight,switchNow 
' where
' camName = the short name of the camera (or group) to trigger
' grpName = the group name of the cameras to bring forward on the console
' highlight = Y means bring group forward even if Windfilter3864Ref = 1. N means bring group forward only if Windfilter3864Ref = 0
' switchNow = Y means switch main video even if Windfilter3864Ref = 1. N means switch main video only if Windfilter3864Ref = 0
' uses (Windfilter3864Ref)
' (CamSwitchInput4606Ref)
Public Sub camShow(ByVal host As String, ByVal port As integer,ByVal parms As String)
    Dim args() As String = Split(parms, ",")
    Dim camName As String = args(0)
    Dim grpName As String = args(1)
    Dim highlight As String = args(2)
    Dim switchNow As String = args(3)

    callBI(host,camName,port,TRIGGER)
    If hs.DeviceValue(Windfilter3864Ref) = 0 Or (Instr(highlight,"Y") > 0) Then
        callBI(host,grpName,port,SHOW)
        sayString ("Highlighting " & grpName & " on " & host)
    End If
    If hs.DeviceValue(Windfilter3864Ref) = 0 Or (Instr(switchNow,"Y") > 0) Then
        If (Instr(host,"Iris2") > 0) Then
        ' switch video distribution to Blue Iris 2 console (input 2)
            hs.SetDeviceValueByRef(CamSwitchInput4606Ref, 2, TRUE)
            hs.TimerReset("Iris2_motion_delay")
        Else If (Instr(host,"Iris3") > 0) Then 
            ' switch video distribution to Blue Iris 3 console (input 3)
            hs.SetDeviceValueByRef(CamSwitchInput4606Ref, 3, TRUE)
            hs.TimerReset("Iris3_motion_delay")
        Else If (Instr(host,"Iris4") > 0) Then 
            ' switch video distribution to Blue Iris 4 console (input 4)
            hs.SetDeviceValueByRef(CamSwitchInput4606Ref, 4, TRUE)
            hs.TimerReset("Iris4_motion_delay")
        End If
    
    End If
End Sub

' For other things you can do with the Blue Iris API see https://ipcamtalk.com/threads/blue-iris-urls-for-external-streams.24994/ or the Blue Iris help file.
' host can be a name or and IP address but names must be resolvable (ping host works) or use the IP address of the host instead
' With command of trigger = triggers an alert on a camera on the Blue Iris server
' ctlName is the short cam or group name in the camera's config. Must macth exactly.
'
' With command of show = shows the camera group on the console. 
' ctlName is the cam group name (camera names will not work) in the camera's config. Must macth exactly. Toi show all use "All cameras"
' uses (CamSwitchInput4606Ref)
Public Sub callBI(ByVal host As String, ByVal ctlName As String, ByVal port As integer, ByVal command As String)
    Dim label = "callBI"
    Dim page = ""
    Try
        If (Instr(command,TRIGGER) > 0) Then 
            hs.WriteLog(label, "http://" & host & ":" & port & "/admin?camera=" & ctlName & "&trigger" & BlueIrisLogin)
            ' hs.GetURL(IP,URL,TRUE,port)
            page = hs.GetURL(host,"/admin?camera=" & ctlName & "&trigger" & BlueIrisLogin,TRUE,port)
            ' reset timer for host console
            hs.TimerReset(host & "_motion_delay")
        Else If (Instr(command,SHOW) > 0) Then 
            hs.WriteLog(label, "ctlName:" & ctlName & " host:" & host & " switch before:" & hs.DeviceValue(CamSwitchInput4606Ref))
            If (Instr(ctlName,"All cameras") > 0) Then 
                ' switch video distribution to main video (input 1)
                hs.SetDeviceValueByRef(CamSwitchInput4606Ref, 1, TRUE)
            End If
            ' hs.WriteLog(label, " switch after:" & hs.DeviceValue(CamSwitchInput4606Ref))
            hs.WriteLog(label, "http://" & host & ":" & port & "/admin?console=" & ctlName & BlueIrisLogin)
            ' hs.GetURL(IP,URL,TRUE,port)
            page = hs.GetURL(host,"/admin?console=" & ctlName & BlueIrisLogin,TRUE,port)
        Else
            hs.WriteLog(label, "Unsupported command:" & command)
            sayLog( "Unsupported command:" & command)
        End If

        If (Instr(page,"camera=NULL") > 0) Then 
            hs.WriteLog(label, "camera " & ctlName & " not found")
            sayString("camera " & ctlName & " was not found on " & host)
        Else If (Instr(page,"var login_version") > 0) Then 
            'hs.WriteLog(label, "secure session keys blocking API calls. Turn off.")
            sayString("secure session keys blocking API calls on " & host & ". Turn off.")
        Else If (Instr(page,"Authorization required") > 0) Then 
            'hs.WriteLog(label, "Login info is incorrect")
            sayString( "Login info is incorrect for " & host)
        Else If (Instr(page,"signal=greenprofile=-1lock=0") > 0) Then 
            ' sucessful response is
            ' signal=greenprofile=-1lock=0camera=longName
            hs.WriteLog(label, "" & command & " " & ctlName & " was sucessful")
        Else 
            hs.WriteLog(label, "GetURL returned:" & page)
            sayLog( "Failed " & command & " of " & ctlName & " on " & host)
        End If
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message & " " & host & ":" & port)
    End Try
End Sub