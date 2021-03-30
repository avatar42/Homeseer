' This is the startup script
' It is run once when HomeSeer starts up
' 
' You may also have Startup.vb and it will be run instead of this script.
'
Sub Main(Parm As Object)
    
    hs.WriteLog("Startup", "(Startup.vb script) Scripting is OK and is now running Startup.vb")
	
	Dim SpeakClients As String = ""
	SpeakClients = hs.GetInstanceList
 	If SpeakClients Is Nothing Then SpeakClients = ""
    If String.IsNullOrEmpty(SpeakClients.Trim) Then
        hs.WriteLog("Startup", "(Startup.vb script) No speaker clients detected, waiting up to 30 seconds for any to connect...")
        ' There are no connected speaker clients, so let's wait 30 seconds
        '	(which is the default re-connect interval for a speaker client)
        '	to see if we can get one connected, otherwise the two speak
        '	commands below will not be heard.
        hs.WaitSecs(1)
        Dim Start As Date = Now()
        Do
            SpeakClients = hs.GetInstanceList
            If SpeakClients Is Nothing Then SpeakClients = ""
            If String.IsNullOrEmpty(SpeakClients.Trim) Then
                hs.WaitSecs(1)
            Else
                Exit Do
            End If
        Loop Until Now.Subtract(Start).TotalSeconds > 30
    End If
	
	    
	' Speak - comment the next line if you do not want HomeSeer to speak at startup.
	hs.Speak("Welcome to Home-Seer", True)
		
	' speak the port the web server is running on
    Dim port As String = hs.GetINISetting("Settings", "gWebSvrPort", "")
    If port <> "80" Then
        hs.Speak("Web server port number is " & port)
    End If

    ' You may add your own commands to this script.
    ' See the scripting section of the HomeSeer help system for more information.
    ' You may access help by going to your HomeSeer website and clicking the HELP button,
    ' or by pointing your browser to the /help page of your HomeSeer system.
    ' setting 4597 to 0 did not seem to always work to trigger an event
    'hs.SetDeviceValueByRef(4597, 0, True)
    'hs.WriteLog("Startup", "Set 4597 to 0")
    hs.TriggerEvent("On reboot")
End Sub
