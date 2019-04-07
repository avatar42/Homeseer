' ref for mapped objects
#Include Refs.vb

'methods for handling speech announcements of various types
'lastSaid:Dim lastSaid2969Ref = 2969

' Say everywhere ignoring AnnounceMode981Ref value
Public Sub sayAlert(ByVal parm As Object)
    Try
        hs.speakEx(0, parm, False)
        setLast(parm)
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script sayAlert:  " & ex.Message)
        hs.speakEx(0, "Failed to speak alert", False)
    End Try

End Sub

Public Sub sayLog(ByVal parm As Object)
    sayString(parm & ". Check Homeseer log")
End Sub

'Get the best name to use for the device
Public Function betterName(ByVal dv As Object) As String
    Dim label = "betterName"
    Dim name = dv.VoiceCommand(Nothing)
    try
        hs.WriteLog(label, "VoiceCommand: '" & name & "'")
        If String.IsNullOrEmpty(name) Then
            name = dv.Name(Nothing)
            name.Replace("." ," ").Replace("_pwr", " power").Replace("_", " ")
        End If
        hs.WriteLog(label, "name: '" & name & "'")
        return name
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
        return "unknown"
    End Try

End Function

' Look at the display options for a value and ge the one that probably sounds best in speech interface
Public Function betterValue(ByVal dv As Object) As String
    Dim label = "betterValue"
    Dim val = dv.devString(Nothing)
    try
        hs.WriteLog(label, "devString: '" & val & "'")
        If String.IsNullOrEmpty(val) Then
            val = hs.CapiGetStatus(dv.Ref(Nothing)).Status
        End If
        hs.WriteLog(label, "val: '" & val & "'")
        If String.IsNullOrEmpty(val) Then
            val = dv.devValue(Nothing)
        End If
        hs.WriteLog(label, "val: '" & val & "'")
        return val
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
        return "unknown"
    End Try

End Function

'Announce that device has changed to new value
Public Sub sayValue(ByVal dvRef As Object)
    Dim dv = hs.GetDeviceByRef(dvRef)
    sayString(betterName(dv) & " " & dv.devValue(Nothing))
End Sub

'Announce that device has changed to new String value
Public Sub sayVString(ByVal dvRef As Object)
    Dim dv = hs.GetDeviceByRef(dvRef)
    sayString(betterName(dv) & " " & dv.devString(Nothing))
End Sub

' Say the name and status of a device
Public Sub sayStatus(ByVal dvRef As Object)
    Dim dv = hs.GetDeviceByRef(dvRef)
    sayString(betterName(dv) & " is " & betterValue(dv))
End Sub

Public Sub sayWhyReset(ByVal dvRef As Object)
    Dim dv = hs.GetDeviceByRef(dvRef)
    sayString(betterName(dv) & " is " & betterValue(dv) & " so doing a reset")
End Sub

Public Sub announce(ByVal parm As Object)
    hs.speak(parm)
    setLast(parm)
End Sub

Public Sub sayLast(ByVal parm As Object)
    hs.speak(hs.DeviceString(lastSaid2969Ref))
End Sub

Public Sub setLast(ByVal parm As Object)
    Try
        Try
            hs.SetDeviceString(lastSaid2969Ref,parm,TRUE)
        Catch ex As Exception
            hs.SetDeviceString(lastSaid2969Ref,"Last was not stored",TRUE)
        End Try
    Catch ex As Exception
        hs.WriteLog("setLast", "lastSaid2969Ref:" & lastSaid2969Ref & "parm:" & parm)
    End Try

End Sub

Public Sub sayString(ByVal parm As Object)
    hs.WriteLog("sayString", "passed:" & hs.DeviceValue(AnnounceMode981Ref) & ":" & parm)
    Select Case hs.DeviceValue(AnnounceMode981Ref)
        Case 1
            ' Say only on clients
            hs.speak(parm)
            hs.WriteLog("sayString", "1:" & parm)
            setLast(parm)
        Case 2
            ' Say on Sonos group
            hs.speakEx(0, parm, False, "$SONOS$TTS1$")
            hs.speakEx(0, parm, False, "$SONOS$Bedroom$")
            hs.WriteLog("sayString", "2:" & parm)
            setLast(parm)
        Case 3
            ' Say everywhere
            hs.speakEx(0, parm, False)
            'hs.speakEx(0, parm, False, "$SONOS$TTS1$")
            'hs.speakEx(0, parm, False, "$SONOS$Bedroom$")
            'hs.speakEx(0, parm, False, "$HS3$Bedroom$")
            hs.WriteLog("sayString", "3:" & parm)
            setLast(parm)
    End Select
End Sub

Function myTime() As String
    Return Now.ToString("h:mm tt").Replace("PM", "P.M.").Replace("AM", "A.M.").Replace(":0", ", o")
End Function

Function TimeOfDay() As String
    Select Case Now.Hour
        Case Is < 12
            Return "Morning"
        Case 12 To 17
            Return "Afternoon"
        Case 18 To 24
            Return "Evening"
        Case Else
            Return "Error"
    End Select
End Function