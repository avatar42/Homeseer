' Note it is also possible to call URLs directly form Blue Iris to set device values in Homeseer
' As in http://HS#IP:PORT/JSON?user=UserName&pass=Password&request=controldevicebyvalue&ref=RefID%&value=Value
' Though be sure the device you are trying to set the value on is set for control and can accept a value in the range you will be sending.

' currently supported commands
Dim TRIGGER As String = "trigger"
Dim SHOW As String = "show"
Dim PROFILE As String = "profile"

Dim port2 = 8078
Dim port3 = 8076
Dim port4 = 8077
Dim port5 = 8075

Sub sayString(ByVal msg As String)
    hs.RunScriptFunc("SayIt.vb", "sayString", msg, False, False)
End Sub

Public Sub sayGlobal(ByVal name As String)
    hs.RunScriptFunc("SayIt.vb", "sayGlobal", name, False, False)
End Sub

' Checks script complies and globals used, are defined
Sub Main(ByVal ignored As String)
    sayString("Blue Iris Script compiled OK")
    Dim s = hs.GetVar("BlueIrisLogin")
    if (InStr(s, "&pw=") > 0)
        sayString("BlueIrisLogin is defined")
    Else
        sayString("BlueIrisLogin not set")
        hs.WriteLog("Error", "BlueIrisLogin not set correctly")
    End If
    sayGlobal("FocusedcamRef")
    sayGlobal("WindfilterRef")
    sayGlobal("CamSwitchInputRef")
End Sub

' convenience methods for calling camShow(ByVal host As String, ByVal port As integer,ByVal parms As String)
' see camShow
Public Sub Iris2Show(ByVal parms As String)
    camShow("Iris6",port2,parms)
End Sub

' see camShow
Public Sub Iris3Show(ByVal parms As String)
    camShow("Iris3",port3,parms)
End Sub

' see camShow
Public Sub Iris4Show(ByVal parms As String)
    camShow("Iris4",port4,parms)
End Sub

' see camShow
Public Sub Iris5Show(ByVal parms As String)
    camShow("Iris5",port5,parms)
End Sub

' Set all consoles back to All cameras
Public Sub consoleClear(ByVal notUsed As String)
    callBI("Iris6","All cameras",port2,SHOW)
    callBI("Iris3","All cameras",port3,SHOW)
    callBI("Iris4","All cameras",port4,SHOW)
    callBI("Iris5","All cameras",port5,SHOW)
    hs.SetDeviceString(hs.GetVar("FocusedcamRef"), "All consoles cleared", True)
    hs.SetDeviceValueByRef(hs.GetVar("FocusedcamRef"), 0, True)
End Sub

' Set console back to All cameras
Public Sub Iris2Clear(ByVal notUsed As String)
    callBI("Iris6","All cameras",port2,SHOW)
End Sub

' Set console back to All cameras
Public Sub Iris3Clear(ByVal notUsed As String)
    callBI("Iris3","All cameras",port3,SHOW)
End Sub

' Set console back to All cameras
Public Sub Iris4Clear(ByVal notUsed As String)
    callBI("Iris4","All cameras",port4,SHOW)
End Sub

' Set console back to All cameras
Public Sub Iris5Clear(ByVal notUsed As String)
    callBI("Iris5","All cameras",port5,SHOW)
End Sub

' Set console back to All cameras
Public Sub Iris4Profile(ByVal pid As String)
    callBI("Iris4",pid,port4,PROFILE)
End Sub

' Triggers an alert on camName And highlights the group / switches main video based on the following
' host = Blue Iris server's name
' port = Blue Iris servr's web API port
' parms = camName,grpName,highlight,switchNow 
' where
' camName = the short name of the camera (or group) to trigger
' grpName = the group name of the cameras to bring forward on the console
' highlight = Y means bring group forward even if WindfilterRef = 1. N means bring group forward only if WindfilterRef = 0
' switchNow = Y means switch main video even if WindfilterRef = 1. N means switch main video only if WindfilterRef = 0
' uses (WindfilterRef) a flag for reducing highlighting during storms.
' (CamSwitchInputRef) control for PiP HDMI switch
Public Sub camShow(ByVal host As String, ByVal port As integer,ByVal parms As String)
    Dim args() As String = Split(parms, ",")
    Dim camName As String = args(0)
    Dim grpName As String = args(1)
    Dim highlight As String = args(2)
    Dim switchNow As String = args(3)

    callBI(host,camName,port,TRIGGER)
    If hs.DeviceValue(hs.GetVar("WindfilterRef")) = 0 Or (Instr(highlight, "Y") > 0) Then
        callBI(host, grpName, port, SHOW)
    End If
    If hs.DeviceValue(hs.GetVar("WindfilterRef")) = 0 Or (Instr(switchNow, "Y") > 0) Then
        If (Instr(host, "Iris6") > 0) Then
            ' switch video distribution to Blue Iris 2 console (input 2)
            hs.SetDeviceValueByRef(hs.GetVar("CamSwitchInputRef"), 2, True)
            hs.TimerReset("Iris2_motion_delay")
        ElseIf (Instr(host, "Iris3") > 0) Then
            ' switch video distribution to Blue Iris 3 console (input 3)
            hs.SetDeviceValueByRef(hs.GetVar("CamSwitchInputRef"), 3, True)
            hs.TimerReset("Iris3_motion_delay")
        ElseIf (Instr(host, "Iris4") > 0) Then
            ' switch video distribution to Blue Iris 4 console (input 4)
            hs.SetDeviceValueByRef(hs.GetVar("CamSwitchInputRef"), 4, True)
            hs.TimerReset("Iris4_motion_delay")
        ElseIf (Instr(host, "Iris5") > 0) Then
            ' switch video distribution to Blue Iris 4 console (input 4)
            hs.SetDeviceValueByRef(hs.GetVar("CamSwitchInputRef"), 5, True)
            hs.TimerReset("Iris5_motion_delay")
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
' uses (CamSwitchInputRef) control for PiP HDMI switch
' (FocusedcamRef) holds last Highlighting message
Public Sub callBI(ByVal host As String, ByVal ctlName As String, ByVal port As integer, ByVal command As String)
    Dim label = "callBI"
    Dim page = ""
    Try
        If (Instr(command,TRIGGER) > 0) Then
            hs.WriteLog(label, "http://" & host & ":" & port & "/admin?camera=" & ctlName & "&trigger" & hs.GetVar("BlueIrisLogin"))
            ' hs.GetURL(IP,URL,TRUE,port)
            page = hs.GetURL(host, "/admin?camera=" & ctlName & "&trigger" & hs.GetVar("BlueIrisLogin"), True, port)
            ' reset timer for host console
            hs.TimerReset(host & "_motion_delay")
        Else If (Instr(command,PROFILE) > 0) Then
            hs.WriteLog(label, "http://" & host & ":" & port & "/admin?profile=" & ctlName & hs.GetVar("BlueIrisLogin"))
            ' hs.GetURL(IP,URL,TRUE,port)
            page = hs.GetURL(host, "/admin?profile=" & ctlName & hs.GetVar("BlueIrisLogin"), True, port)
        ElseIf (Instr(command,SHOW) > 0) Then
            'hs.WriteLog(label, "ctlName:" & ctlName & " host:" & host & " switch before:" & hs.DeviceValue(hs.GetVar("CamSwitchInputRef")))
            Try
                Try
                    hs.SetDeviceString(hs.GetVar("FocusedcamRef"), "Highlighting " & ctlName & " on " & host, True)
                    If (Instr(host,"Iris6") > 0) Then
                        hs.SetDeviceValueByRef(hs.GetVar("FocusedcamRef"), 2, True)
                    ElseIf (Instr(host,"Iris3") > 0) Then
                        hs.SetDeviceValueByRef(hs.GetVar("FocusedcamRef"), 3, True)
                    ElseIf (Instr(host,"Iris4") > 0) Then
                        hs.SetDeviceValueByRef(hs.GetVar("FocusedcamRef"), 4, True)
                    ElseIf (Instr(host,"Iris5") > 0) Then
                        hs.SetDeviceValueByRef(hs.GetVar("FocusedcamRef"), 5, True)
                    Else
                        hs.SetDeviceValueByRef(hs.GetVar("FocusedcamRef"), 1, True)
                    End If

                Catch ex As Exception
                    hs.SetDeviceString(hs.GetVar("FocusedcamRef"), "Focusedcam was not stored", True)
                End Try
            Catch ex As Exception
                hs.WriteLog("callBI", "FocusedcamRef:" & hs.GetVar("FocusedcamRef") & "parm:" & "Highlighting " & ctlName & " on " & host)
            End Try

            If (Instr(ctlName,"All cameras") > 0) Then
                ' switch video distribution to main video (input 1)
                hs.SetDeviceValueByRef(hs.GetVar("CamSwitchInputRef"), 1, True)
            End If
            ' hs.WriteLog(label, " switch after:" & hs.DeviceValue(hs.GetVar("CamSwitchInputRef")))
            hs.WriteLog(label, "http://" & host & ":" & port & "/admin?console=" & ctlName & hs.GetVar("BlueIrisLogin"))
            ' hs.GetURL(IP,URL,TRUE,port)
            page = hs.GetURL(host, "/admin?console=" & ctlName & hs.GetVar("BlueIrisLogin"), True, port)
        Else
            hs.WriteLog(label, "Unsupported command:" & command)
            sayString("Unsupported command:" & command)
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
        Else If (Instr(page,"signal=greenprofile=") > 0) Then 
            ' sucessful responses can be
            ' signal=greenprofile=-1lock=0camera=longName
            ' signal=greenprofile=1lock=2
            hs.WriteLog(label, "" & command & " " & ctlName & " was sucessful")
        Else 
            hs.WriteLog(label, "GetURL returned:" & page)
            sayString("Failed " & command & " of " & ctlName & " on " & host)
        End If
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message & " " & host & ":" & port)
    End Try
End Sub