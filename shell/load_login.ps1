#########################
#        eTweetXL       #
#                       #
# Web Automation Script #
#                       #
#########################
#                       #
# Programmer(s): anz7re #
#                       #
###############################
#                             #
# Latest Revision: 7/22/2022  #
#                             #
###############################
#
#
#
#CLEANUP
$OffsetMTInfo = ''
$PassMTInfo = ''
$PostMTInfo = ''
$MedMTInfo = ''
$UserMTInfo = ''
$RuntimeMTInfo = ''
$RetryMTInfo = ''
$BrowserInfo = ''
$xCt = ''

#PATHWAYS

#ERROR FILES
$RtErr = $home + '\.z7\autokit\etweetxl\debug\rt.err'
$WebErr = $home + '\.z7\autokit\etweetxl\debug\web.err'

#BROWSER FILES
$Browser = $home + '\.z7\autokit\etweetxl\mtsett\webset.txt'
$BrowserCheck = $home + '\.z7\autokit\etweetxl\mtsett\webcheck.txt'

#TARGET FILES
$apiMT = $home + '\.z7\autokit\etweetxl\mtsett\api.mt'
$BlankMT = $home + '\.z7\autokit\etweetxl\mtsett\blank.mt'
$OffsetMT = $home + '\.z7\autokit\etweetxl\mtsett\offset.mt'
$OffsetMTC = $home + '\.z7\autokit\etweetxl\mtsett\offset.mtc'
$PassMT = $home + '\.z7\autokit\etweetxl\mtsett\pass.mt'
$PostMT = $home + '\.z7\autokit\etweetxl\mtsett\post.mt'
$ProfileMT = $home + '\.z7\autokit\etweetxl\mtsett\profile.mt'
$MedMT = $home + '\.z7\autokit\etweetxl\mtsett\med.mt'
$RetryMT = $home + '\.z7\autokit\etweetxl\mtsett\retry.mt'
$RuntimeMT = $home + '\.z7\autokit\etweetxl\mtsett\runtime.mt'
$RtCntrMT = $home + '\.z7\autokit\etweetxl\mtsett\rtcntr.mt'
$UserMT = $home + '\.z7\autokit\etweetxl\mtsett\user.mt'
$ThreadMT = $home + '\.z7\autokit\etweetxl\mtsett\thread.mt'
$ThreadCtMT = $home + '\.z7\autokit\etweetxl\mtsett\threadct.mt'
$IniMT = $home + '\.z7\autokit\etweetxl\mtsett\ini.mt'

#RUNTIME SCRIPTS
$apiScript = $home + '\.z7\autokit\etweetxl\shell\win\send_with_api.vbs'
$CheckURL = $home + '\.z7\autokit\etweetxl\shell\win\check_url.bat'
$MeScript = $home + '\.z7\autokit\etweetxl\shell\win\load_login.bat'
$StScript = $home + '\.z7\autokit\etweetxl\shell\win\start_etweetxl.vbs'
$RuntimeError = $home + '\.z7\autokit\etweetxl\shell\win\runtime_error.vbs'
$RuntimeRfsh = $home + '\.z7\autokit\etweetxl\shell\win\runtime_refresh.vbs'
$LoginError = $home + '\.z7\autokit\etweetxl\shell\win\login_error.vbs'
$IniScript = $home + '\.z7\autokit\etweetxl\shell\win\load_initialize.bat'

$DateTime = Get-Date 

#(1)
#
#INITIALIZING...
Out-File -FilePath $RtErr -InputObject "Initializing..." -Encoding default
Start-Process -FilePath $RuntimeError
Start-Sleep -Seconds 1

$OffsetMTInfo = Get-Content -Path $OffsetMT -Encoding Default
$OffsetMTCInfo = Get-Content -Path $OffsetMTC -Encoding Default
$PassMTInfo = Get-Content -Path $PassMT -Encoding Default
$PostMTInfo = Get-Content -Path $PostMT -Encoding Default
$ProfileMTInfo = Get-Content -Path $ProfileMT -Encoding Default
$MedMTInfo = Get-Content -Path $MedMT -Encoding Default
$UserMTInfo = Get-Content -Path $UserMT -Encoding Default
$BrowserInfo = Get-Content -Path $Browser -Encoding Default 
$RetryMTInfo = Get-Content -Path $RetryMT -Encoding Default

$wshell = New-Object -ComObject wscript.Shell;

##### SCRIPT UTILITY #####
function Close-ActiveBrowser{

#EXIT BROWSER WINDOW
$wshell.SendKeys('^{w}')

#CHECK IF BROWSER STILL OPEN
$DefinedItem = Get-Process -Name $BrowserInfo

#INITIAL CHECK
if($DefinedItem.Count -ne '0'){
$wshell.SendKeys('^{w}')

$ErrMsg = "`rAn error occurred during the browsing session"

#DOUBLE CHECK
if($DefinedItem.Count -ne '0'){
Out-File $WebErr -InputObject $DateTime$ErrMsg  -Append
Stop-Process $DefinedItem

        }
    }

 }

function Send-Tweet{

$wshell.SendKeys('{ESC}')
Start-Sleep -Seconds 1
$wshell.SendKeys('{ESC}')
Start-Sleep -Seconds 1
$wshell.SendKeys('^{ENTER}') 

} 

#CHECK FOR THREAD
if (Test-Path $ThreadMT -PathType Leaf){
$xCt = Get-Content -Path $ThreadMT}

#GET RUNTIME & FIND PLACEMENT
if (Test-Path $RtCntrMT -PathType Leaf){
$RtCntr = Get-Content -Path $RtCntrMT}

if (Test-Path $RuntimeMT -PathType Leaf){
$RuntimeMTInfo = Get-Content $RuntimeMT | Select -Index $RtCntr}

#INITIALIZE PROFILE/USER INFORMATION FOR AUTOMATION
If (!(Test-Path $IniMT -PathType Leaf)){
Start-Process -FilePath $IniScript
Exit
}

#STARTING SLEEP...
Out-File -FilePath $RtErr -InputObject "Starting sleep..." -Encoding default
Start-Process -FilePath $RuntimeError
Start-Sleep -Seconds 1

#WAIT
If ($RuntimeMTInfo -eq ''){

#START WAIT
Start-Sleep -Milliseconds $OffsetMTInfo
                                       
} else {

        $TotalSec = (New-TimeSpan -End $RuntimeMTInfo).TotalSeconds

        #HANDLE NEGATIVE TIME
        if ($TotalSec -le 0){
        $TotalSec = 0
        }

        #REPLACE TIME
        $TotalMil = (New-TimeSpan -Seconds $TotalSec).TotalMilliseconds

        #EXPORT TIME TO OFFSET FILE
        Out-File -FilePath $OffsetMT -InputObject $TotalMil -Encoding default

        #CONVERT RUNTIME TO SECONDS & WAIT
        (New-TimeSpan -End $RuntimeMTInfo).TotalSeconds | Sleep;

}

If($RetryMTInfo -ge 5){
Exit
}

#SEND USING API SCRIPT
if (Test-Path $apiMT -PathType Leaf){

#INCREMENT RUNTIME COUNTER
$nwRtCntr = [int]$RtCntr
$nwRtCntr = ($nwRtCntr + 1)
Out-File -FilePath $RtCntrMT -InputObject $nwRtCntr

#REMOVE RETRY FILE
Remove-Item -Path $RetryMT

#SETTING UP POST...
Out-File -FilePath $RtErr -InputObject "Setting up post..." -Encoding default
Start-Process -FilePath $RuntimeError
Start-Sleep -Seconds 1

#STARTING AUTOMATION...
Out-File -FilePath $RtErr -InputObject "Starting automation..." -Encoding default
Start-Process -FilePath $RuntimeError
Start-Sleep -Seconds 1

#START API SCRIPT
Start-Process -FilePath $apiScript

Start-Sleep -Seconds 10

#SEND TWEET
$wshell.SendKeys('{ENTER}')

Start-Sleep -Seconds 1

taskkill /f /im 'cmd.exe'

#NEXT
Start-Process -FilePath $StScript

#INCREMENT RUNTIME COUNTER
$nwRtCntr = [int]$RtCntr
$nwRtCntr = ($nwRtCntr + 1)
Out-File -FilePath $RtCntrMT -InputObject $nwRtCntr

Exit

}

#START TWITTER USING FIREFOX UNDER PRIVATE BROWSER
$Url = Start-Process -FilePath 'C:\Program Files\Mozilla Firefox\firefox.exe' -ArgumentList @( '-private-window', 'twitter.com\login')

#STARTING AUTOMATION...
Out-File -FilePath $RtErr -InputObject "Starting automation..." -Encoding default
Start-Process -FilePath $RuntimeError

Start-Sleep -Seconds 5

#MAKE BROWSER WINDOW FULLSCREEN
$wshell.SendKeys('{F11}')

Start-Sleep -Seconds 5


#HANDLE FIREFOX BREAKING FROM AUTOMATION
If($RetryMTInfo -eq 3){

#RESOLVING ISSUE...
Out-File -FilePath $RtErr -InputObject "Trying to resolve the issue... Please wait..." -Encoding default
Start-Process -FilePath $RuntimeError

Start-Sleep -Seconds 5
$wshell.SendKeys('{ENTER}')
Start-Sleep -Seconds 5

$nwRetryMTInfo = [int]$RetryMTInfo
$nwRetryMTInfo = ($nwRetryMTInfo + 1)
Out-File -FilePath $RetryMT -InputObject $nwRetryMTInfo

#RESTART SCRIPT
Start-Process -FilePath $MeScript

Exit
}

#FIND USERNAME
For ($xNum=0; $xNum -le 2; $xNum++){
$wshell.SendKeys('{TAB}')
}

#SEND USERNAME
$wshell.SendKeys($UserMTInfo)

#====================================================================================================
#LOGIN CHECK
$wshell.SendKeys('^{a}')
$wshell.SendKeys('^{c}')

$Notepad = Start-Process $BlankMT
Start-Sleep -Seconds 1
taskkill /f /fi "WINDOWTITLE eq blank.mt*"

$CheckUser = Get-Clipboard

if($CheckUser -match $UserMTInfo){
#====================================================================================================

#CHECK FOR LOGIN PAGE REACHED
Start-Process $CheckURL
Start-Sleep -Seconds 2
$URL = Get-Content -Path $BrowserCheck
if($URL.Contains("Login on Twitter / Twitter")){

$wshell.SendKeys('{TAB}')
$wshell.SendKeys($PassMTInfo)
$wshell.SendKeys('{TAB}')
$wshell.SendKeys('{ENTER}')

   } else {

#IFLOW LOGIN
$wshell.SendKeys('{ENTER}')
Start-Sleep -Seconds 2
$wshell.SendKeys($PassMTInfo)
$wshell.SendKeys('{ENTER}')
}

Start-Sleep -Seconds 5

#CHECK IF LOGGED IN
Start-Process $CheckURL
Start-Sleep -Seconds 2
$URL = Get-Content -Path $BrowserCheck
if($URL.Contains("Home / Twitter")){

#REFRESH PAGE
$wshell.SendKeys('^{r}')

Start-Sleep -Seconds 5

#COMPOSE TWEET HOTKEY
$wshell.SendKeys('n')

Start-Sleep -Seconds 1

#PASTE TWEET
$wshell.SendKeys($PostMTInfo) 
                                     
Start-Sleep -Seconds 2

#FIND MEDIA IF ADDED TO POST

#FIND ADD MEDIA BUTTON
if ($MedMTInfo -ne ''){
For ($xNum=0; $xNum -le 1; $xNum++){
$wshell.SendKeys('{TAB}')}

Start-Sleep -Seconds 1

#SELECT MEDIA BUTTON
$wshell.SendKeys('{ENTER}')

Start-Sleep -Seconds 3

#FIND MEDIA
$wshell.SendKeys($MedMTInfo)

Start-Sleep -Seconds 2

#SELECT MEDIA
$wshell.SendKeys('{ENTER}')

Start-Sleep -Seconds 5

}


#LOCATION BUTTON BACKUP CHECK & TAB INDEX RETURN 
#
#CHECKS WHETHER OR NOT LOCATION BUTTON IS ACTIVE FOR A USER & RETURNS NEEDED TAB INDEX
#
$LocCheck = $home + '\.z7\autokit\etweetxl\presets\' + $ProfileMTInfo + '\loccheck.txt'

if (Test-Path $LocCheck -PathType Leaf){
$LocType = Get-Content -Path $LocCheck
    if ($LocType -eq 'True'){

        $MedIndex1 = 3
        $MedIndex2 = 2
        $MedIndex3 = 4

        $PostIndex1 = 7
        $PostIndex2 = 6
        $PostIndex3 = 8

    } else {

        $MedIndex1 = 2
        $MedIndex2 = 1
        $MedIndex3 = 3

        $PostIndex1 = 6
        $PostIndex2 = 5
        $PostIndex3 = 7

                    }
        } else {

#IF LOCATION NOT CHECKED BEFORE

if ($MedMTInfo -eq ''){
For ($xNum=0; $xNum -le 6; $xNum++){
Start-Sleep -Milliseconds 100
$wshell.SendKeys('{TAB}')}
} else {
For ($xNum=0; $xNum -le 2; $xNum++){
Start-Sleep -Milliseconds 100
$wshell.SendKeys('{TAB}')}
}

Start-Sleep -Seconds 1

#SELECT LOCATION BUTTON
$wshell.SendKeys('{ENTER}')

#JUMP TO URL
$wshell.SendKeys('^{l}')

Start-Sleep -Seconds 1

#COPY URL
$wshell.SendKeys('^{a}')
$wshell.SendKeys('^{c}')

Start-Sleep -Seconds 1

$LocType = Get-Clipboard

Start-Sleep -Seconds 1

#LOCATION BUTTON FOUND                      
if($LocType.Contains('place_picker')){

$LocBool = 'True'

$MedIndex1 = 3
$MedIndex2 = 2
$MedIndex3 = 4

$PostIndex1 = 7
$PostIndex2 = 6
$PostIndex3 = 8

Out-File -FilePath $LocCheck -InputObject $LocBool

if ($MedMTInfo -eq ''){

#EXIT WINDOW
For ($xNum=0; $xNum -le 3; $xNum++){
$wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 200}
$wshell.SendKeys('{ENTER}')

Start-Sleep -Seconds 1 

if ($xCt -eq ''){
For ($xNum=0; $xNum -le 8; $xNum++){
$wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 200}  
}

} else {

#EXIT WINDOW
For ($xNum=0; $xNum -le 3; $xNum++){
$wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 200}
$wshell.SendKeys('{ENTER}')

Start-Sleep -Seconds 1

if ($xCt -eq ''){
For ($xNum=0; $xNum -le 5; $xNum++){
$wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 200}
}
            }

                } else {

#LOCATION BUTTON NOT FOUND

$LocBool = 'False'

$MedIndex1 = 2
$MedIndex2 = 1
$MedIndex3 = 3

$PostIndex1 = 6
$PostIndex2= 5
$PostIndex3 = 7

Out-File -FilePath $LocCheck -InputObject $LocBool


$wshell.SendKeys('{TAB}')
Start-Sleep -Seconds 1

if ($MedMTInfo -eq ''){

For ($xNum=0; $xNum -le 11; $xNum++){
$wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 200}
$wshell.SendKeys('{ENTER}')    
} else {

For ($xNum=0; $xNum -le 9; $xNum++){
$wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 200}   

            }

                }

                    }


#THREADED POST CHECK
if ($xCt -ne ''){

if ($MedMTInfo -ne ''){
For ($xNum=0; $xNum -le $MedIndex1; $xNum++){
$wshell.SendKeys('{TAB}')}
Start-Sleep -Milliseconds 50
} else {
For ($xNum=0; $xNum -le $PostIndex1; $xNum++){
$wshell.SendKeys('{TAB}')}
Start-Sleep -Milliseconds 50
}

Remove-Item $PostMT

For ($x=2; $x -le $xCt; $x++){

if (Test-Path $PostMT -PathType Leaf){

#FIND THREAD BUTTON (FROM MEDIA)
if ($MedMTInfo -ne ''){
For ($xNum=0; $xNum -le $MedIndex2; $xNum++){
$wshell.SendKeys('{TAB}')}
} else {
#FIND THREAD BUTTON (NO MEDIA)
For ($xNum=0; $xNum -le $PostIndex2; $xNum++){
$wshell.SendKeys('{TAB}')}
Start-Sleep -Milliseconds 50
}
    }


#SELECT THREAD
$wshell.SendKeys('{ENTER}') 

Start-Process -FilePath $StScript

Start-Sleep -Seconds 2

$PostMTInfo = Get-Content -Path $PostMT -Encoding Default
$MedMTInfo = Get-Content -Path $MedMT -Encoding Default

#SEND THREADED POST
$wshell.SendKeys($PostMTInfo) 

Start-Sleep -Seconds 5

#FIND MEDIA IF ADDED TO THREADED POST

#FIND ADD MEDIA BUTTON
if($MedMTInfo -ne ''){
    if($PostMTInfo -ne ' '){
For ($m=0; $m -le 1; $m++){
$wshell.SendKeys('{TAB}')
}
   } else {
    For ($m=0; $m -le 2; $m++){
    $wshell.SendKeys('{TAB}')
        }
      }

Start-Sleep -Seconds 1

#SELECT MEDIA BUTTON
$wshell.SendKeys('{ENTER}')

Start-Sleep -Seconds 3

#FIND THREAD MEDIA
$wshell.SendKeys($MedMTInfo)

Start-Sleep -Seconds 2

#SELECT THREAD MEDIA
$wshell.SendKeys('{ENTER}')

Start-Sleep -Seconds 2

        }

            }

#SET EXIT TIMER (THREAD)
if($MedMTInfo -ne ''){
$SendTimer = 55
    } else {
$SendTimer = 25
}

#SEND TWEET (THREAD)
Send-Tweet

Start-Sleep -Seconds $SendTimer

Close-ActiveBrowser

#INCREMENT RUNTIME COUNTER
$nwRtCntr = [int]$RtCntr
$nwRtCntr = ($nwRtCntr + 1)
Out-File -FilePath $RtCntrMT -InputObject $nwRtCntr

#REMOVE RETRY FILE
Remove-Item -Path $RetryMT
Remove-Item -Path $BrowserCheck

#SETTING UP POST...
Out-File -FilePath $RtErr -InputObject "Setting up post..." -Encoding default
Start-Process -FilePath $RuntimeError
Start-Sleep -Seconds 1

#START SCRIPT
Start-Process -FilePath $StScript

Exit

} 

#SET EXIT TIMER (NOT THREAD)
if($MedMTInfo -ne ''){
$SendTimer = 15
    } else {
$SendTimer = 3
}
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            
#SEND TWEET (NOT THREAD)
Send-Tweet

Start-Sleep -Seconds $SendTimer

Close-ActiveBrowser

#INCREMENT RUNTIME COUNTER
$nwRtCntr = [int]$RtCntr
$nwRtCntr = ($nwRtCntr + 1)
Out-File -FilePath $RtCntrMT -InputObject $nwRtCntr

#REMOVE SESSION FILES
Remove-Item -Path $IniMT
Remove-Item -Path $RetryMT
Remove-Item -Path $BrowserCheck

#SETTING UP POST...
Out-File -FilePath $RtErr -InputObject "Setting up post..." -Encoding default
Start-Process -FilePath $RuntimeError
Start-Sleep -Seconds 1

#START SCRIPT
Start-Process -FilePath $StScript

Exit

    }
     
       }

        

#IF LOGIN PAGE FAILS AGAIN FORCE CLOSE & RETRY (LIKELY AN ISSUE DURING THE WEB LOAD)
#KO BROWSER
taskkill /f /im 'firefox.exe'

#ERROR DURING LOGIN...
Out-File -FilePath $RtErr -InputObject "Error during login... Attempting to retry..." -Encoding default
Start-Process -FilePath $RuntimeError
Start-Sleep -Seconds 1

If ($RetryMTInfo -le 2){
$nwRetryMTInfo = [int]$RetryMTInfo
$nwRetryMTInfo = ($nwRetryMTInfo + 1)
Out-File -FilePath $RetryMT -InputObject $nwRetryMTInfo

#CLONE OLD OFFSET IN CASE OF SUCCESS
Out-File -FilePath $OffsetMTC -InputObject $OffsetMTInfo -Encoding default

#RESTART AFTER 5 SECONDS
Out-File -FilePath $OffsetMT -InputObject 5000 -Encoding default

Start-Sleep -Seconds 20

#START RUNTIMER
Start-Process -FilePath $RuntimeRfsh

#RESTARTING SCRIPT...
Out-File -FilePath $RtErr -InputObject "Restarting script..." -Encoding default
Start-Process -FilePath $RuntimeError

Start-Sleep -Seconds 1

Start-Process -FilePath $MeScript

Start-Sleep -Seconds 5

} else {

#REMOVE RETRY FILE
Remove-Item -Path $RetryMT 

#ERROR DURING LOGIN...
Out-File -FilePath $RtErr -InputObject "Unresolved error during login..." -Encoding default
Start-Process -FilePath $RuntimeError
Start-Sleep -Seconds 1

Start-Process -FilePath $LoginError

}
