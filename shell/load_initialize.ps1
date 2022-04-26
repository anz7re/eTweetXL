###############################
#        eTweetXL             #
#                             #
# User Initialization Script  #
#                             #
###############################
#                       #
# Programmer(s): anz7re #
#                       #
###############################
#                             #
# Latest Revision: 2/23/2022  #
#                             #
###############################
#
#
#
#CLEANUP
$OffsetMTInfo = ''
$PassMTInfo = ''
$PostMTInfo = ''
$UserMTInfo = ''
$RuntimeMTInfo = ''
$RetryMTInfo = ''
$WebsetInfo = ''

#PATHWAYS

#MT FILES
$BlankMT = $home + '\.z7\autokit\etweetxl\mtsett\blank.mt'
$OffsetMT = $home + '\.z7\autokit\etweetxl\mtsett\offset.mt'
$OffsetMTCopy = $home + '\.z7\autokit\etweetxl\mtsett\offset.mtc'
$PassMT = $home + '\.z7\autokit\etweetxl\mtsett\pass.mt'
$PostMT = 'Initializing...'
$ProfileMT = $home + '\.z7\autokit\etweetxl\mtsett\profile.mt'
$RetryMT = $home + '\.z7\autokit\etweetxl\mtsett\retry.mt'
$RuntimeMT = $home + '\.z7\autokit\etweetxl\mtsett\runtime.mt'
$RtCntrMT = $home + '\.z7\autokit\etweetxl\mtsett\rtcntr.mt'
$UserMT = $home + '\.z7\autokit\etweetxl\mtsett\user.mt'
$IniMT = $home + '\.z7\autokit\etweetxl\mtsett\ini.mt'

#WEB SETTINGS
$Webset = $home + '\.z7\autokit\etweetxl\mtsett\webset.txt'
$Webcheck = $home + '\.z7\autokit\etweetxl\mtsett\webcheck.txt'

#SCRIPTS
$CheckURL = $home + '\.z7\autokit\etweetxl\shell\win\check_url.bat'
$MeScript = $home + '\.z7\autokit\etweetxl\shell\win\load_initialize.bat'
$LoadScript = $home + '\.z7\autokit\etweetxl\shell\win\load_login.bat'
$RuntimeError = $home + '\.z7\autokit\etweetxl\shell\win\runtime_error.vbs'
$RuntimeRfsh = $home + '\.z7\autokit\etweetxl\shell\win\runtime_refresh.vbs'
$LoginError = $home + '\.z7\autokit\etweetxl\shell\win\login_error.vbs'

#ERROR FILES
$RtErr = $home + '\.z7\autokit\etweetxl\debug\rt.err'

#INITIALIZING...
Out-File -FilePath $RtErr -InputObject "Initializing..." -Encoding default
Start-Process -FilePath $RuntimeError
Start-Sleep -Seconds 1

$OffsetMTInfo = Get-Content -Path $OffsetMT -Encoding Default
$OffsetMTCopyInfo = Get-Content -Path $OffsetMTCopy -Encoding Default
$PassMTInfo = Get-Content -Path $PassMT -Encoding Default
$PostMTInfo = $PostMT
$ProfileMTInfo = Get-Content -Path $ProfileMT -Encoding Default
$UserMTInfo = Get-Content -Path $UserMT -Encoding Default
$WebsetInfo = Get-Content -Path $Webset -Encoding Default 
$RetryMTInfo = Get-Content -Path $RetryMT -Encoding Default

#CHECK IF PROFILE/USER ALREADY INITIALIZED
$LocCheck = $home + '\.z7\autokit\etweetxl\presets\' + $ProfileMTInfo + '\loccheck.txt'

if (Test-Path $LocCheck -PathType Leaf){

#SEND SUCCESSFUL INITIALIZATION RESULTS
$IniErr = '0'
Out-File -FilePath $IniMT -InputObject $IniErr

#RESTART LOGIN SCRIPT
Start-Process -FilePath $LoadScript

Exit 

    } else {

 }

$Wshell = New-Object -ComObject wscript.shell;

#GET RUNTIME & FIND PLACEMENT
if (Test-Path $RtCntrMT -PathType Leaf){
$RtCntr = Get-Content -Path $RtCntrMT}

if (Test-Path $RuntimeMT -PathType Leaf){
$RuntimeMTInfo = Get-Content $RuntimeMT | Select -Index $RtCntr}

#CHECK FOR MAX LOGIN FAILURE
If($RetryMTInfo -ge 5){
Exit
}

#START TWITTER USING FIREFOX UNDER PRIVATE BROWSER
$FF = Start-Process -FilePath 'C:\Program Files\Mozilla Firefox\firefox.exe' -ArgumentList @( '-private-window', 'twitter.com\login')

#STARTING INITIALIZATION...
Out-File -FilePath $RtErr -InputObject "Starting initialization..." -Encoding default
Start-Process -FilePath $RuntimeError

Start-Sleep -Seconds 5

#MAKE BROWSER WINDOW FULLSCREEN
$Wshell.SendKeys('{F11}')

Start-Sleep -Seconds 5

#HANDLE FIREFOX BREAKING FROM AUTOMATION
If($RetryMTInfo -eq 3){

#RESOLVING ISSUE...
Out-File -FilePath $RtErr -InputObject "Trying to resolve the issue... Please wait..." -Encoding default
Start-Process -FilePath $RuntimeError

Start-Sleep -Seconds 5
$Wshell.SendKeys('{ENTER}')
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
$Wshell.SendKeys('{TAB}')
}

#SEND USERNAME
$Wshell.SendKeys($UserMTInfo)

#====================================================================================================
#Removed 2/11/2022 due to a change in the copy/paste property for the Twitter login boxes.
#
#Initially this would check for the username we sent via SendKeys to match our username saved on file.
#
#Brought this back 2/23/2022 as it seems this property was changed back? Might've missed something...
#
LOGIN CHECK
$Wshell.SendKeys('^{a}')
$Wshell.SendKeys('^{c}')

$Notepad = Start-Process $BlankMT
Start-Sleep -Seconds 1
taskkill /f /fi "WINDOWTITLE eq blank.mt*"

$CheckUser = Get-Clipboard

if($CheckUser -match $UserMTInfo){
#====================================================================================================

#CHECK FOR LOGIN PAGE REACHED
Start-Process $CheckURL
Start-Sleep -Seconds 2
$URL = Get-Content -Path $Webcheck
if($URL.Contains("Login on Twitter / Twitter")){

$Wshell.SendKeys('{TAB}')
$Wshell.SendKeys($PassMTInfo)
$Wshell.SendKeys('{TAB}')
$Wshell.SendKeys('{ENTER}')

   } else {

#IFLOW LOGIN
$Wshell.SendKeys('{ENTER}')
Start-Sleep -Seconds 2
$Wshell.SendKeys($PassMTInfo)
$Wshell.SendKeys('{ENTER}')
}

Start-Sleep -Seconds 5

#CHECK IF LOGGED IN
Start-Process $CheckURL
Start-Sleep -Seconds 2
$URL = Get-Content -Path $Webcheck
if($URL.Contains("Home / Twitter")){

#FIND COMPOSE TWEET BUTTON
For ($xNum=0; $xNum -le 13; $xNum++){
$Wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 50}

#SELECT COMPOSE TWEET BUTTON
$Wshell.SendKeys('{ENTER}')

Start-Sleep -Seconds 1

#PASTE DUMMY TWEET
$Wshell.SendKeys($PostMTInfo) 
                                     
Start-Sleep -Seconds 2

For ($xNum=0; $xNum -le 6; $xNum++){
Start-Sleep -Milliseconds 100
$Wshell.SendKeys('{TAB}')}

Start-Sleep -Seconds 1

#SELECT LOCATION BUTTON IF THERE
$Wshell.SendKeys('{ENTER}')

#JUMP TO URL
$Wshell.SendKeys('^{l}')

Start-Sleep -Seconds 1

#COPY URL
$Wshell.SendKeys('^{a}')
$Wshell.SendKeys('^{c}')

Start-Sleep -Seconds 1

$LocType = Get-Clipboard

Start-Sleep -Seconds 1

#LOCATION BUTTON FOUND                      
if($LocType.Contains('place_picker')){

$LocBool = 'True'

Out-File -FilePath $LocCheck -InputObject $LocBool

                } else {

#LOCATION BUTTON NOT FOUND

$LocBool = 'False'

Out-File -FilePath $LocCheck -InputObject $LocBool 

            }

#EXIT BROWSER WINDOW
$Wshell.SendKeys('^{w}')

#SEND SUCCESSFUL INITIALIZATION RESULTS
$IniErr = '0'
Out-File -FilePath $IniMT -InputObject $IniErr

#RESTART LOGIN SCRIPT
Start-Process -FilePath $LoadScript

Exit
                    }



                        } 
                                

#IF LOGIN PAGE FAILS AGAIN FORCE CLOSE & RETRY (LIKELY AN ISSUE DURING THE WEB LOAD)
#KO BROWSER
taskkill /f /im 'firefox.exe'

#ERROR DURING LOGIN...
Out-File -FilePath $RtErr -InputObject "Error during intialization... Attempting to retry..." -Encoding default
Start-Process -FilePath $RuntimeError
Start-Sleep -Seconds 1

If ($RetryMTInfo -le 2){
$nwRetryMTInfo = [int]$RetryMTInfo
$nwRetryMTInfo = ($nwRetryMTInfo + 1)
Out-File -FilePath $RetryMT -InputObject $nwRetryMTInfo

#CLONE OLD OFFSET IN CASE OF SUCCESS
Out-File -FilePath $OffsetMTCopy -InputObject $OffsetMTInfo -Encoding default

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
Out-File -FilePath $RtErr -InputObject "Unresolved error during initialization..." -Encoding default
Start-Process -FilePath $RuntimeError
Start-Sleep -Seconds 1

Start-Process -FilePath $LoginError

}
