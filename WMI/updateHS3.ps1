## set these to your servers values
$Hs3ip="10.10.1.45"
$Hs3user="wmi"
$Hs3pass="wmipass"
$tempRef="4588"
$loadRef="4587"
$pctFreeRef="4586"
$logLenRef="4595"
$updateRef="4596"
$BIRef="5713"
## turn trun on extra logging set $debugOn = 1 and off is $debugOn = 0
$debugOn = 0
## also update disk list in local drives section below

$Logfile = "C:\logs\$(gc env:computername)." + (Get-Date).toString("yyyy-MM") + ".log"

## If you find the script is stopping at each error try uncommenting the following line for a run to change ErrorActionPreference
#$ErrorActionPreference = "Continue"

## if you need help figuring out why something is not working try
## removing # from any place you see #Write in the method with the problem.

$baseURL="http://"+$Hs3ip+"/JSON?user="+$Hs3user+"&pass="+$Hs3pass+"&request"

## for even fancier logging see https://gist.github.com/barsv/85c93b599a763206f47aec150fb41ca0
Function Write-Log($Message) {

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Message"
    ## to stop logging put a # in front of next line
    Add-Content $Logfile -Value $Line
    ##to stop console output put a # in front of next line
    Write-Output $Line
}

Function Error-Log($Message,$errAt) {
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp ERROR: $Message " + $Message.InvocationInfo.PositionMessage
    ## to stop logging put a # in front of next line
    Add-Content $Logfile -Value $Line 
    ##to stop console output put a # in front of next line
    #Write-Error $Line
    Write-Output $Line
}

Function Debug-Log($Message) {
    if($debugOn) {
        Write-Log($Message)
    }
}

function sendData($ref,$value) {
Debug-Log ("ref="+$ref)
Debug-Log ("value="+$value)
## Note the try catch is not always triggered when WMI calls fail so
## if we get a null value we know something went wrong so signal that to Homeseer with a -1 value
    if( $value -eq $null) {
        $value = -1
    }
    $url=("$baseURL=controldevicebyvalue&ref=$ref&value=$value")
    Write-Log ($url)
    $resp = (New-Object System.Net.WebClient).DownloadString($url);
    Debug-Log ($resp)

    $respObj = ConvertFrom-Json($resp)
    Debug-Log ($respObj)

    if (! $respObj.Name ) {
       Error-Log ("Error:"+ $resp)
       Error-Log ($url)
    }
}

function sendError($ref,$err) {
    Error-Log ("Error for:$ref :$err",$err.InvocationInfo.PositionMessage)
    sendData $ref "-1"

    $url=("$baseURL=setdeviceproperty&ref=$ref&property=Attention&value=$err")
    Write-Log ($url)
    $resp = (New-Object System.Net.WebClient).DownloadString($url);
    Debug-Log ($resp)

    $respObj = ConvertFrom-Json($resp)
    Debug-Log ($respObj)

    if ((! $respObj.Response) -or (! $respObj.Response -eq "ok")) {
       Error-Log ("Error:"+ $resp)
       Error-Log ($url)
    }
}

function updateDiskFree($diskLetter,$diskRef) {
    try {
        $filter = "DeviceID='"+$diskLetter+"'"

        $disk = get-wmiobject -class "Win32_LogicalDisk" -namespace "root\CIMV2" -Filter "$filter" | Select-Object Size,FreeSpace
        $size = [math]::round($disk.Size/1GB, 2)
        $free = [math]::round($disk.FreeSpace/1GB, 2)
        $freePercent=[math]::round(($free/$size * 100), 2)
        Debug-Log ($diskLetter+" "+$diskRef+" size="+$size+" free="+$free+" free="+$freePercent+"%")

        sendData $diskRef $freePercent
    } catch {

        sendError $diskRef $_
    }
}

## other temp options:
## Get-WmiObject Win32_TemperatureProbe -Namespace "root/cimv2" CurrentReading
function Get-Temperature {
    try {
        $t = @( Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace "root/wmi" )
        $err=$error[0]
        Debug-Log $t
        $cores = 0
        $tempTotal=0
        foreach ($temp in $t)
        {
            $cores += 1
            $currentTempKelvin = $temp.CurrentTemperature / 10
            $currentTempCelsius = $currentTempKelvin - 273.15
            $tempTotal += $currentTempCelsius 
            $currentTempFahrenheit = (9/5) * $currentTempCelsius + 32

            Write-Log ("Core "+$cores+":"+$currentTempCelsius.ToString() + " C : " + $currentTempFahrenheit.ToString() + " F : " + $currentTempKelvin + "K")
        }
        if($t.Count -gt 0) {
            sendData $tempRef ($tempTotal / $t.Count)
        } else {
            sendError $tempRef $err
        }
    } catch {
        sendError $tempRef $_
    }
}

Get-Temperature

 
## CPU load
try {
    $load=Get-WmiObject win32_processor | select LoadPercentage
    $err=$error[0]
    if($load.LoadPercentage -gt 0) {
        sendData $loadRef $load.LoadPercentage
    } else {
        sendError $loadRef $err
    }
} catch {
    sendError $loadRef $_
}

## BlueIris load
try {
    $load=(Get-Counter "\Process(blueiris#1)\% Processor Time").CounterSamples
    $err=$error[0]
    if($load.CookedValue -gt 0) {
        sendData $BIRef $load.CookedValue
    } else {
        sendError $BIRef $err
    }
} catch {
    sendError $BIRef $_
}

## free RAM
try {
    $os = Get-Ciminstance Win32_OperatingSystem
    $pctFree = [math]::Round(($os.FreePhysicalMemory/$os.TotalVisibleMemorySize)*100,2)
    sendData $pctFreeRef $pctFree
} catch {
    sendError $pctFreeRef $_
}

## watch log length
try {
   Debug-Log (Get-Item $Logfile)
   $logLen = [math]::Round((Get-Item $Logfile).length / 1MB,3)
    sendData $logLenRef $logLen
} catch {
    sendError $logLenRef $_
}

## count updates pending
try {
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateupdateSearcher()
    ## does not want to filter on and DownloadPriority=2 or and DownloadPriority>1 so extra update is listed
    $Updates = @($UpdateSearcher.Search("IsHidden=0 and IsInstalled=0").Updates)
    $Updates| Select-Object Title,IsMandatory,IsDownloaded,RebootRequired,AutoSelection,AutoDownload,MsrcSeverity,DeploymentAction,DownloadPriority
 
    if ($Updates.Count -gt 1) {
        Debug-Log ("Updates pending:"+ $Updates.Count)
    }

    $err=$error[0]
    if($Updates.Count -gt 0) {
        sendData $updateRef $Updates.Count
    } else {
        sendError $updateRef $err
    }
} catch {
    sendError $updateRef $_
}


## local drives
updateDiskFree "C:" 4589
updateDiskFree "E:" 4591
updateDiskFree "F:" 4590
updateDiskFree "K:" 4592
