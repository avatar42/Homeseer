﻿' should probably change AlarmLevel2899Ref usages to use Partition14702Ref instead and the matching values
Const AlarmOff = 0
Const AlarmHome = 1
Const AlarmPerm = 2
Const AlarmAway = 3

Const AnnounceMode981Ref = 981
Const WeatherAlerts1179Ref = 1179
Const S1TemperatureHighNormal1182Ref = 1182
Const S1TemperatureLowNormal1185Ref = 1185
Const S1TemperatureFeelsLike1190Ref = 1190
Const S2TemperatureFeelsLike1202Ref = 1202
Const S1GustSpeed1212Ref = 1212
Const S1WindDirection1215Ref = 1215
Const S1RainToday1216Ref = 1216
Const S1CurrentCondition1233Ref = 1233
Const S2CurrentCondition1238Ref = 1238
Const S1TodayHigh1243Ref = 1243
Const S1TodayLow1244Ref = 1244
Const S1TodayPrediction1245Ref = 1245
Const S1LastAlertMessage1355Ref = 1355
Const DeviceCount1522Ref = 1522
Const Harm1activity2876Ref = 2876
Const AlarmLevel2899Ref = 2899
Const lastSaid2969Ref = 2969
Const Samsung602974Ref = 2974
Const LastDevice2975Ref = 2975
Const Rmrr3174Ref = 3174
Const LivingRoomThermostat3175Ref = 3175
Const AmbientTemperature3176Ref = 3176
Const TargetTemperatureLow3177Ref = 3177
Const TargetTemperatureHigh3178Ref = 3178
Const HVACMode3179Ref = 3179
Const Harm2activity3393Ref = 3393
Const Windfilter3864Ref = 3864
Const DeviceChksum4597Ref = 4597
Const CheckedDeviceCount4605Ref = 4605
Const CamSwitchInput4606Ref = 4606
Const ZettaguardAVSwitch4677Ref = 4677
Const Partition14702Ref = 4702
Const Focusedcam4720Ref = 4720

Public Sub RefsLoaded(ByVal delta As Object)
   hs.WriteLog("Refs","Refs loaded:" & 35 & " device refs")
End Sub