Sub sayString(ByVal msg As String)
    hs.RunScriptFunc("SayIt.vb", "sayString", msg, False, False)
End Sub

Public Sub sayGlobal(ByVal name As String)
    hs.RunScriptFunc("SayIt.vb", "sayGlobal", name, False, False)
End Sub

' Checks script complies and globals used, are defined
Sub Main(ByVal ignored As String)
    sayString("Blue Iris Script compiled OK")
    sayGlobal("S1LastAlertMessageRef")
    sayGlobal("S1WindDirectionRef")
    sayGlobal("S1GustSpeedRef")
    sayGlobal("S1RainTodayRef")
    sayGlobal("S1TodayHighRef")
    sayGlobal("S1TemperatureHighNormalRef")
    sayGlobal("S1TodayLowRef")
    sayGlobal("S1TemperatureLowNormalRef")
    sayGlobal("S1TodayPredictionRef")
    sayGlobal("WeatherAlertsRef")
    sayGlobal("WeatherAlertsRef")
    sayGlobal("S1CurrentConditionRef")
    weatherStatus("")
End Sub


Public Sub weatherAlert(ByVal notUsed As Object)
    hs.RunScriptFunc("SayIt.vb", "sayString", "Last alert was " & hs.DeviceString(hs.GetVar("S1LastAlertMessageRef")), False, False)
End Sub

Public Sub weatherStatus(ByVal notUsed As Object)
    Dim label As String = "weatherStatus"
    Try
        If (hs.GetVar("S1CurrentConditionRef") = 0) Then
            hs.RunScriptFunc("SayIt.vb", "sayAlert", "Currently condition is " & hs.DeviceString(hs.GetVar("S2CurrentConditionRef")) & " With a feels like temperature of " & hs.DeviceValue(hs.GetVar("S2TemperatureFeelsLikeRef")) & " degrees", False, False)
        End If

        If (hs.GetVar("S1CurrentConditionRef") = 0) Then
            SayCurrentCondition()
        End If

        'wind spelled wend to sound right
        hs.RunScriptFunc("SayIt.vb", "sayAlert", "wend is from the " & windDirection(hs.DeviceValue(hs.GetVar("S1WindDirectionRef"))) & " gusting to " & hs.DeviceValue(hs.GetVar("S1GustSpeedRef")) & " miles per hour", False, False)
        hs.RunScriptFunc("SayIt.vb", "sayAlert", "Rain since midnight is  " & hs.DeviceValue(hs.GetVar("S1RainTodayRef")) & " inches", False, False)
        If (hs.GetVar("S1TemperatureHighNormalRef") = 0) Then
            hs.RunScriptFunc("SayIt.vb", "sayAlert", "High so far was  " & hs.DeviceValue(hs.GetVar("S1TodayHighRef")) & " degrees", False, False)
            hs.RunScriptFunc("SayIt.vb", "sayAlert", "Low so far was  " & hs.DeviceValue(hs.GetVar("S1TodayLowRef")) & " degrees", False, False)
        Else
            hs.RunScriptFunc("SayIt.vb", "sayAlert", "Expected high is  " & hs.DeviceValue(hs.GetVar("S1TodayHighRef")) & " compared to the normal of " & hs.DeviceValue(hs.GetVar("S1TemperatureHighNormalRef")) & " degrees", False, False)
            hs.RunScriptFunc("SayIt.vb", "sayAlert", "Expected low is  " & hs.DeviceValue(hs.GetVar("S1TodayLowRef")) & " compared to the normal of " & hs.DeviceValue(hs.GetVar("S1TemperatureLowNormalRef")) & " degrees", False, False)
        End If
        If (hs.GetVar("S1TodayPredictionRef") > 0) Then
            hs.RunScriptFunc("SayIt.vb", "sayAlert", "Forecast is  " & hs.DeviceString(hs.GetVar("S1TodayPredictionRef")), False, False)
        End If
        If (hs.GetVar("WeatherAlertsRef") > 0) Then
            hs.RunScriptFunc("SayIt.vb", "sayAlert", "There are " & hs.DeviceValue(hs.GetVar("WeatherAlertsRef")) & " alerts in effect", False, False)
            If (hs.GetVar("S1LastAlertMessageRef") > 0) Then
                If hs.DeviceValue(hs.GetVar("WeatherAlertsRef")) > 0 Is Nothing Then
                    hs.RunScriptFunc("SayIt.vb", "sayAlert", "Last alert was " & hs.DeviceString(hs.GetVar("S1LastAlertMessageRef")), False, False)
                End If
            End If
        End If
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try

End Sub

Public Function SayCurrentCondition() As String
    Dim label As String = "SayCurrentCondition"
    Dim cond As String = "unknown"
    Try
        Select Case hs.DeviceValue(hs.GetVar("S1CurrentConditionRef"))
            Case -25
                cond = "Scattered Clouds"
            Case -24
                cond = "Overcast"
            Case -23
                cond = "Thunderstorm"
            Case -22
                cond = "Thunderstorms"
            Case -21
                cond = "Sunny"
            Case -20
                cond = "Snow"
            Case -19
                cond = "Sleet"
            Case -18
                cond = "Rain"
            Case -17
                cond = "Freezing Rain"
            Case -16
                cond = "Partly Sunny"
            Case -15
                cond = "Partly Cloudy"
            Case -14
                cond = "Mostly Sunny"
            Case -13
                cond = "Mostly Cloudy"
            Case -12
                cond = "Haze"
            Case -11
                cond = "Fog"
            Case -10
                cond = "Flurries"
            Case -9
                cond = "Cloudy"
            Case -8
                cond = "Clear"
            Case -7
                cond = "Chance of a Thunderstorm"
            Case -6
                cond = "Chance of Thunderstorms"
            Case -5
                cond = "Chance of Snow"
            Case -4
                cond = "Chance of Sleet"
            Case -3
                cond = "Chance of Freezing Rain"
            Case -2
                cond = "Chance of Rain"
            Case -1
                cond = "Chance of Flurries"
            Case 0
                cond = "Unknown"
            Case 1
                cond = "Chance of Flurries"
            Case 2
                cond = "Chance of Rain"
            Case 3
                cond = "Chance of Freezing Rain"
            Case 4
                cond = "Chance of Sleet"
            Case 5
                cond = "Chance of Snow"
            Case 6
                cond = "Chance of Thunderstorms"
            Case 7
                cond = "Chance of a Thunderstorm"
            Case 8
                cond = "Clear"
            Case 9
                cond = "Cloudy"
            Case 10
                cond = "Flurries"
            Case 11
                cond = "Fog"
            Case 12
                cond = "Haze"
            Case 13
                cond = "Mostly Cloudy"
            Case 14
                cond = "Mostly Sunny"
            Case 15
                cond = "Partly Cloudy"
            Case 16
                cond = "Partly Sunny"
            Case 17
                cond = "Freezing Rain"
            Case 18
                cond = "Rain"
            Case 19
                cond = "Sleet"
            Case 20
                cond = "Snow"
            Case 21
                cond = "Sunny"
            Case 22
                cond = "Thunderstorms"
            Case 23
                cond = "Thunderstorm"
            Case 24
                cond = "Overcast"
            Case 25
                cond = "Scattered Clouds"
            Case Else
                cond = "Error"
        End Select
        hs.RunScriptFunc("SayIt.vb", "sayAlert", "Currently condition is " & cond & " With a feels like temperature of " & hs.DeviceValue(hs.GetVar("S1TemperatureFeelsLikeRef")) & " degrees", False, False)
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
    Return cond
End Function

Public Function windDirection(ByVal degrees As Object) As String
    Dim label As String = "windDirection"
    Dim dir As String = "unknown"
    If TypeOf degrees Is String Then
        Return windDirectionDir(degrees)
    End If
    Try
        If (degrees >= 0 And degrees <= 23) Then
            Return "north"
        ElseIf (degrees > 23 And degrees <= 68) Then
            Return "north east"
        ElseIf (degrees > 68 And degrees <= 113) Then
            Return "east"
        ElseIf (degrees > 113 And degrees <= 158) Then
            Return "south east"
        ElseIf (degrees > 158 And degrees <= 203) Then
            Return "south"
        ElseIf (degrees > 203 And degrees <= 248) Then
            Return "south west"
        ElseIf (degrees > 248 And degrees <= 293) Then
            Return "west"
        ElseIf (degrees > 293 And degrees <= 338) Then
            Return "north west"
        ElseIf (degrees > 338 And degrees <= 360) Then
        End If
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
    Return "error"
End Function

Public Function windDirectionDir(ByVal degrees As Object) As String
    Dim label As String = "windDirection"
    Dim dir As String = "unknown"
    Try
        If (String.Compare(degrees, "N") = 0) Then
            Return "north"
        ElseIf (String.Compare(degrees, "NE") = 0) Then
            Return "north east"
        ElseIf (String.Compare(degrees, "ENE") = 0) Then
            Return "east north east"
        ElseIf (String.Compare(degrees, "E") = 0) Then
            Return "east"
        ElseIf (String.Compare(degrees, "SE") = 0) Then
            Return "south east"
        ElseIf (String.Compare(degrees, "ESE") = 0) Then
            Return "east south east"
        ElseIf (String.Compare(degrees, "S") = 0) Then
            Return "south"
        ElseIf (String.Compare(degrees, "SW") = 0) Then
            Return "south west"
        ElseIf (String.Compare(degrees, "WSW") = 0) Then
            Return "west south west"
        ElseIf (String.Compare(degrees, "W") = 0) Then
            Return "west"
        ElseIf (String.Compare(degrees, "NW") = 0) Then
            Return "north west"
        ElseIf (String.Compare(degrees, "WNW") = 0) Then
            Return "west north west"
        Else
            Return degrees
        End If
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
    Return degrees
End Function

