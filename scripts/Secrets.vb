' See https://docs.pushbullet.com/v2/ on how to get a token
' curl -u XXXXXXXXXXXXXXXXXXXXXX: https://api.pushbullet.com/v2/pushes -d type=note -d title="Alert" -d body="push test"
const headers="Access-Token: XXXXXXXXXXXXXXXXXXXXXX" 

' login parms for Blue Iris API calls
const BlueIrisLogin="&user=LOGIN&pw=PASSWORD" 

' username for Vesync (etekcity)
const VesyncUsername = "yourEmailUser@yourEmail.domain"

' MD5 encoded password for Vesync (etekcity) Look at 
' https://github.com/avatar42/MyMonitor/blob/master/src/main/java/dea/monitor/tools/EncodeString.java 
' for an example of how to MD5 a string
const VesyncPasswordAsMD5 = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" 

Public Sub loaded(ByVal delta As Object)
    hs.WriteLog("Secrets","Secrets loaded")
End Sub