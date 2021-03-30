
Public Sub setGlobal(ByVal name As String, ByVal val As Object)
    'just in case new. If not is ignored
    hs.CreateVar(name)
    hs.SaveVar(name, val)
End Sub

' Loads the global vars used my my Homeseer scripts. The object refs need to be changed to match the ones in your system.
Sub Main(parm As Object)
' Alarm levels
    setGlobal("AlarmOff", 0)
    setGlobal("AlarmHome", 1)
    setGlobal("AlarmPerm", 2)
    setGlobal("AlarmAway", 3)

    ' common objects
    setGlobal("AnnounceModeRef", 981)
    setGlobal("lastSaidRef", 2969)
    setGlobal("DeviceCountRef", 1522)
    setGlobal("LastDeviceRef",2975)
    setGlobal("WindfilterRef",3864)
    setGlobal("DeviceChksumRef",4597)
    setGlobal("CheckedDeviceCountRef",4605)

    ' Alarm.vb objects
    setGlobal("Partition1Ref", 4702)
    setGlobal("AlarmLevelRef", 2899)

    ' Weather.vb if 0 is ignored / skipped
    setGlobal("WeatherAlertsRef", 0)
    setGlobal("S1TemperatureHighNormalRef", 0)
    setGlobal("S1TemperatureLowNormalRef", 0)
    setGlobal("S1TemperatureFeelsLikeRef", 0)
    setGlobal("S2TemperatureFeelsLikeRef", 0)
    setGlobal("S1GustSpeedRef", 2934)
    setGlobal("S1WindDirectionRef", 2935)
    setGlobal("S1RainTodayRef", 1176)
    setGlobal("S1CurrentConditionRef", 0)
    setGlobal("S2CurrentConditionRef", 0)
    setGlobal("S1TodayHighRef", 1168)
    setGlobal("S1TodayLowRef", 1167)
    setGlobal("S1TodayPredictionRef", 0)
    setGlobal("S1LastAlertMessageRef", 0)

    ' BlueIris.vb
    setGlobal("FocusedcamRef", 4720)
    setGlobal("CamSwitchInputRef", 4606)
    setGlobal("WindfilterRef", 3864)

' Media objects and controls
    setGlobal("Harm1activityRef",2876)
    setGlobal("Harm2activityRef",3393)
    setGlobal("CamSwitchInputRef",4606)
    setGlobal("ZettaguardAVSwitchRef",4677)
    setGlobal("Samsung60Ref",2974)

' Nest.vb objects
    setGlobal("RmrrRef",3174)
    setGlobal("LivingRoomThermostatRef",3175)
    setGlobal("AmbientTemperatureRef",3176)
    setGlobal("TargetTemperatureLowRef",3177)
    setGlobal("TargetTemperatureHighRef",3178)
    setGlobal("HVACModeRef",3179)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Below need to be changed to the login info for your accounts
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' See https://docs.pushbullet.com/v2/ on how to get a token
    ' curl -u XXXXXXXXXXXXXXXXXXXXXX: https://api.pushbullet.com/v2/pushes -d type=note -d title="Alert" -d body="push test"
    setGlobal("headers", "Access-Token: XXXXXXXXXXXXXXXXXXXXXX")

    ' login parms for Blue Iris API calls
    setGlobal("BlueIrisLogin", "&user=LOGIN&pw=PASSWORD")

    ' username for Vesync (etekcity)
    setGlobal("VesyncUsername", "yourEmailUser@yourEmail.domain")

    ' MD5 encoded password for Vesync (etekcity) Look at 
    ' https://github.com/avatar42/MyMonitor/blob/master/src/main/java/dea/monitor/tools/EncodeString.java 
    ' for an example of how to MD5 a string
    setGlobal("VesyncPasswordAsMD5", "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX")

    hs.RunScriptFunc("SayIt.vb", "sayString", "Globals loaded", False, False)
End Sub

