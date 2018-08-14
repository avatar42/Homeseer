' See https://docs.pushbullet.com/v2/ on how to get a token
''curl -u XXXXXXXXXXXXXXXXXXXXXX: https://api.pushbullet.com/v2/pushes -d type=note -d title="Alert" -d body="push test"
const headers="Access-Token: XXXXXXXXXXXXXXXXXXXXXX" 

'login parms for Blue Iris API calls
const BlueIrisLogin="&user=LOGIN&pw=PASSWORD" 

Public Sub loaded(ByVal delta As Object)
    hs.WriteLog("Secrets","Secrets loaded")
End Sub