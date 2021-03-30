# Homeseer
Homeseer scripts and tools<br>

## Note these work under Windows but not mono (Linux) because of issues with include directives in mono. Going to redo them all in near future.

Script | Description
---- | ----
scripts/Alarm.vb | Check all sensors in location2:Alarm are ready|
scripts/Blan.vb|Hack to make Blan plugin devices have a status value to make events work
scripts/BlueIris.vb|Methods for controlling Blue Iris console, my video distrubtion system and marking video based on alarm mode
scripts/ChkSensors.vb|Chk that sensors and devices have been sending values and notifying about any that might need attention
scripts/ChkUPS.vb|Script for alerting when a UPS needs attention. Either bad battery or lost connection.
scripts/Events.vb|Lists out Event info so you can find events using by searching the generated spreadsheet.
scripts/Fixes.vb|Script for making bulk image changes by location2, address base or type.
scripts/ListDevices.vb|Stores last ref ID, device count and a checksum so you can trigger an event to let you know if a plugin or such changed your devices.
scripts/MyMonitor.vb|methods for creating virtual devices to be used with MyMonitor
scripts/Nest.vb|Some Nest convience methods
scripts/PushBullet.vb|Methods for sending messages to Pushbullet as in if alarm set then Pushbullet else announce
scripts/SayIt.vb|Make annuncements based on factors like alarm and debug levels.
scripts/SpeakTemp.vb|Speak the name, temperature and last change time of a sensor.
scripts/Weather.vb|Speak weather info converting wind direction and S1CurrentCondition from numbers to descriptions.
scripts/index.vb|Used by reports/index.aspx
reports/index.aspx|Folder index example
scripts/vesync.vb|Control Vesync Etekcity plugs via their cloud API directly, bypassing the currently non working IFTTT interface.
scripts/WMI.vb|Check WMI status object to see if they are out of range
scripts/WMIObj.vb|Create WMI status objects for a host
