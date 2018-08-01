' ref for mapped objects
#Include Refs.vb

'methods for handling speech announcements of various tyes
'lastSaid:Dim lastSaidRef = 2969
' Say everywhere ignoring AnnounceModeRef value
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

Public Sub announce(ByVal parm As Object)
    hs.speak(parm)
    setLast(parm)
End Sub

Public Sub sayLast(ByVal parm As Object)
    hs.speak(hs.DeviceString(lastSaidRef))
End Sub

Public Sub setLast(ByVal parm As Object)
    Try
        Try
            hs.SetDeviceString(lastSaidRef,parm,TRUE)
        Catch ex As Exception
            hs.SetDeviceString(lastSaidRef,"Last was not stored",TRUE)
        End Try
    Catch ex As Exception
        hs.WriteLog("setLast", "lastSaidRef:" & lastSaidRef & "parm:" & parm)
    End Try

End Sub

Public Sub sayString(ByVal parm As Object)
    hs.WriteLog("sayString", "passed:" & hs.DeviceValue(AnnounceModeRef) & ":" & parm)
    Select Case hs.DeviceValue(AnnounceModeRef)
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
