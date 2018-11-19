'load object refs and speech methods
#Include SayIt.vb
'Uses (S1LastAlertMessage1355Ref)
'(S1WindDirection1215Ref)
'(S1GustSpeed1212Ref)
'(S1RainToday1216Ref)
'(S1TodayHigh1243Ref)
'(S1TemperatureHighNormal1182Ref)
'(S1TodayLow1244Ref)
'(S1TemperatureLowNormalRef)
'(S1TodayPrediction1245Ref)
'(WeatherAlerts1179Ref)
'(WeatherAlerts1179Ref)
'(S1CurrentCondition1233Ref)


Public Sub weatherAlert(ByVal notUsed As Object)
    sayString("Last alert was "& hs.DeviceString(S1LastAlertMessage1355Ref))
End Sub

Public Sub weatherStatus(ByVal notUsed As Object)
    Dim label As String = "weatherStatus"
    try
        'sayAlert("Currently condition is " & hs.DeviceString(S2CurrentCondition1238Ref) & " With a feels like temperature of " & hs.DeviceValue(S2TemperatureFeelsLike1202Ref) & " degrees")
        SayCurrentCondition()
        'wind spelled wend to sound right
        sayAlert("wend is from the " & windDirection(hs.DeviceValue(S1WindDirection1215Ref)) & " gusting to " & hs.DeviceValue(S1GustSpeed1212Ref) & " miles per hour")
        sayAlert("Rain since midnight is  " & hs.DeviceValue(S1RainToday1216Ref) & " inches")
        sayAlert("Expected high is  " & hs.DeviceValue(S1TodayHigh1243Ref) & " compared to the normal of " & hs.DeviceValue(S1TemperatureHighNormal1182Ref) & " degrees")
        sayAlert("Expected low is  " & hs.DeviceValue(S1TodayLow1244Ref) & " compared to the normal of " & hs.DeviceValue(S1TemperatureLowNormal1185Ref) & " degrees")
        sayAlert("Forecast is  " & hs.DeviceString(S1TodayPrediction1245Ref))
        sayAlert("There are "& hs.DeviceValue(WeatherAlerts1179Ref) & " alerts in effect")
        If hs.DeviceValue(WeatherAlerts1179Ref) > 0 Is Nothing Then
            sayAlert("Last alert was "& hs.DeviceString(S1LastAlertMessage1355Ref))
        End If
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try

End Sub

Public Function SayCurrentCondition()
    Dim label As String = "SayCurrentCondition"
    Dim cond As String = "unknown"
    try
        Select Case hs.DeviceValue(S1CurrentCondition1233Ref)
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
        sayAlert("Currently condition is " & cond & " With a feels like temperature of " & hs.DeviceValue(S1TemperatureFeelsLike1190Ref) & " degrees")
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try
End Function

Public Function windDirection(ByVal degrees As Object) As String
    Dim label As String = "windDirection"
    Dim dir As String = "unknown"
    try
    if (degrees >= 0 and degrees <=23) Then
        return "north"
    Else if (degrees > 23 and degrees <=68) Then
        return "north east"
    Else if (degrees > 68 and degrees <=113) Then
        return "east"
    Else if (degrees > 113 and degrees <=158) Then
        return "south east"
    Else if (degrees > 158 and degrees <=203) Then
        return "south"
    Else if (degrees > 203 and degrees <=248) Then
        return "south west"
    Else if (degrees > 248 and degrees <=293) Then
        return "west"
    Else if (degrees > 293 and degrees <=338) Then
        return "north west"
    Else if (degrees > 338 and degrees <=360) Then
        return "north"
    End If
    Catch ex As Exception
        hs.WriteLog("Error", "Exception in script " & label & ":  " & ex.Message)
    End Try

End Function

