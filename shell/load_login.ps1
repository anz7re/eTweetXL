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
# Latest Revision: 3/31/2022  #
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
$WebsetInfo = ''
$xCt = ''

#PATHWAYS

#MT FILES
$apiMT = $home + '\.z7\autokit\etweetxl\mtsett\api.mt'
$BlankMT = $home + '\.z7\autokit\etweetxl\mtsett\blank.mt'
$OffsetMT = $home + '\.z7\autokit\etweetxl\mtsett\offset.mt'
$OffsetMTCopy = $home + '\.z7\autokit\etweetxl\mtsett\offset.mtc'
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

#WEB SETTINGS
$Webset = $home + '\.z7\autokit\etweetxl\mtsett\webset.txt'
$Webcheck = $home + '\.z7\autokit\etweetxl\mtsett\webcheck.txt'

#SCRIPTS
$apiScript = $home + '\.z7\autokit\etweetxl\shell\win\send_with_api.vbs'
$CheckURL = $home + '\.z7\autokit\etweetxl\shell\win\check_url.bat'
$MeScript = $home + '\.z7\autokit\etweetxl\shell\win\load_login.bat'
$StScript = $home + '\.z7\autokit\etweetxl\shell\win\start_etweetxl.vbs'
$RuntimeError = $home + '\.z7\autokit\etweetxl\shell\win\runtime_error.vbs'
$RuntimeRfsh = $home + '\.z7\autokit\etweetxl\shell\win\runtime_refresh.vbs'
$LoginError = $home + '\.z7\autokit\etweetxl\shell\win\login_error.vbs'
$IniScript = $home + '\.z7\autokit\etweetxl\shell\win\load_initialize.bat'

#ERROR FILES
$RtErr = $home + '\.z7\autokit\etweetxl\debug\rt.err'

#INITIALIZING...
Out-File -FilePath $RtErr -InputObject "Initializing..." -Encoding default
Start-Process -FilePath $RuntimeError
Start-Sleep -Seconds 1

$OffsetMTInfo = Get-Content -Path $OffsetMT -Encoding Default
$OffsetMTCopyInfo = Get-Content -Path $OffsetMTCopy -Encoding Default
$PassMTInfo = Get-Content -Path $PassMT -Encoding Default
$PostMTInfo = Get-Content -Path $PostMT -Encoding Default
$ProfileMTInfo = Get-Content -Path $ProfileMT -Encoding Default
$MedMTInfo = Get-Content -Path $MedMT -Encoding Default
$UserMTInfo = Get-Content -Path $UserMT -Encoding Default
$WebsetInfo = Get-Content -Path $Webset -Encoding Default 
$RetryMTInfo = Get-Content -Path $RetryMT -Encoding Default

$Wshell = New-Object -ComObject wscript.shell;

#CHECK FOR THREAD & FIND COUNT
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
$Wshell.SendKeys('{ENTER}')

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
$FF = Start-Process -FilePath 'C:\Program Files\Mozilla Firefox\firefox.exe' -ArgumentList @( '-private-window', 'twitter.com\login')

#STARTING AUTOMATION...
Out-File -FilePath $RtErr -InputObject "Starting automation..." -Encoding default
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

#PASTE TWEET
$Wshell.SendKeys($PostMTInfo) 
                                     
Start-Sleep -Seconds 2

#FIND MEDIA IF ADDED TO POST

#FIND ADD MEDIA BUTTON
if ($MedMTInfo -ne ''){
For ($xNum=0; $xNum -le 1; $xNum++){
$Wshell.SendKeys('{TAB}')}

Start-Sleep -Seconds 1

#SELECT MEDIA BUTTON
$Wshell.SendKeys('{ENTER}')

Start-Sleep -Seconds 3

#FIND MEDIA
$Wshell.SendKeys($MedMTInfo)

Start-Sleep -Seconds 2

#SELECT MEDIA
$Wshell.SendKeys('{ENTER}')

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
$Wshell.SendKeys('{TAB}')}
} else {
For ($xNum=0; $xNum -le 2; $xNum++){
Start-Sleep -Milliseconds 100
$Wshell.SendKeys('{TAB}')}
}

Start-Sleep -Seconds 1

#SELECT LOCATION BUTTON
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
$Wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 200}
$Wshell.SendKeys('{ENTER}')

Start-Sleep -Seconds 1 

if ($xCt -eq ''){
For ($xNum=0; $xNum -le 8; $xNum++){
$Wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 200}  
}

} else {

#EXIT WINDOW
For ($xNum=0; $xNum -le 3; $xNum++){
$Wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 200}
$Wshell.SendKeys('{ENTER}')

Start-Sleep -Seconds 1

if ($xCt -eq ''){
For ($xNum=0; $xNum -le 5; $xNum++){
$Wshell.SendKeys('{TAB}')
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


$Wshell.SendKeys('{TAB}')
Start-Sleep -Seconds 1

if ($MedMTInfo -eq ''){

For ($xNum=0; $xNum -le 11; $xNum++){
$Wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 200}
$Wshell.SendKeys('{ENTER}')    
} else {

For ($xNum=0; $xNum -le 9; $xNum++){
$Wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 200}   

            }

                }

                    }


#THREADED POST CHECK
if ($xCt -ne ''){

if ($MedMTInfo -ne ''){
For ($xNum=0; $xNum -le $MedIndex1; $xNum++){
$Wshell.SendKeys('{TAB}')}
Start-Sleep -Milliseconds 50
} else {
For ($xNum=0; $xNum -le $PostIndex1; $xNum++){
$Wshell.SendKeys('{TAB}')}
Start-Sleep -Milliseconds 50
}

Remove-Item $PostMT

For ($x=2; $x -le $xCt; $x++){

if (Test-Path $PostMT -PathType Leaf){

#FIND THREAD BUTTON (FROM MEDIA)
if ($MedMTInfo -ne ''){
For ($xNum=0; $xNum -le $MedIndex2; $xNum++){
$Wshell.SendKeys('{TAB}')}
} else {
#FIND THREAD BUTTON (NO MEDIA)
For ($xNum=0; $xNum -le $PostIndex2; $xNum++){
$Wshell.SendKeys('{TAB}')}
Start-Sleep -Milliseconds 50
}
    }


#SELECT THREAD
$Wshell.SendKeys('{ENTER}') 

Start-Process -FilePath $StScript

Start-Sleep -Seconds 2

$PostMTInfo = Get-Content -Path $PostMT -Encoding Default
$MedMTInfo = Get-Content -Path $MedMT -Encoding Default

#SEND THREADED POST
$Wshell.SendKeys($PostMTInfo) 

Start-Sleep -Seconds 5

#FIND MEDIA IF ADDED TO THREADED POST

#FIND ADD MEDIA BUTTON
if ($MedMTInfo -ne ''){
For ($m=0; $m -le 2; $m++){
$Wshell.SendKeys('{TAB}')
}

Start-Sleep -Seconds 1

#SELECT MEDIA BUTTON
$Wshell.SendKeys('{ENTER}')

Start-Sleep -Seconds 3

#FIND THREAD MEDIA
$Wshell.SendKeys($MedMTInfo)

Start-Sleep -Seconds 2

#SELECT THREAD MEDIA
$Wshell.SendKeys('{ENTER}')

Start-Sleep -Seconds 2

        }

            }


if ($MedMTInfo -ne ''){
#FIND SEND TWEET BUTTON (FROM MEDIA)
$xNum = 0
$SendTimer = 55
For ($xNum=0; $xNum -le $MedIndex1; $xNum++){
$Wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 50
}
    } else {

#FIND SEND TWEET BUTTON (NO MEDIA)
$xNum = 0
$SendTimer = 25
For ($xNum=0; $xNum -le ($PostIndex1); $xNum++){
$Wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 50
}
    }

#SEND TWEET THREAD
$Wshell.SendKeys('{ENTER}') 

Start-Sleep -Seconds $SendTimer

#EXIT BROWSER WINDOW
$Wshell.SendKeys('^{w}')

#INCREMENT RUNTIME COUNTER
$nwRtCntr = [int]$RtCntr
$nwRtCntr = ($nwRtCntr + 1)
Out-File -FilePath $RtCntrMT -InputObject $nwRtCntr

#REMOVE RETRY FILE
Remove-Item -Path $RetryMT
Remove-Item -Path $Webcheck

#SETTING UP POST...
Out-File -FilePath $RtErr -InputObject "Setting up post..." -Encoding default
Start-Process -FilePath $RuntimeError
Start-Sleep -Seconds 1

#START SCRIPT
Start-Process -FilePath $StScript

Exit
                } else {

            

if ($MedMTInfo -ne ''){
#FIND SEND TWEET BUTTON (FROM MEDIA NOT THREAD)
$xNum = 0
$SendTimer = 15
For ($xNum=0; $xNum -le ($MedIndex3); $xNum++){
$Wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 50
}
    } else {


#FIND SEND TWEET BUTTON (NO MEDIA NOT THREAD)
$xNum = 0
$SendTimer = 3
For ($xNum=0; $xNum -le ($PostIndex3); $xNum++){
$Wshell.SendKeys('{TAB}')
Start-Sleep -Milliseconds 50
}
    }

        } 
       
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
#SEND TWEET
$Wshell.SendKeys('{ENTER}') 

Start-Sleep -Seconds $SendTimer

#EXIT BROWSER WINDOW
$Wshell.SendKeys('^{w}')

#INCREMENT RUNTIME COUNTER
$nwRtCntr = [int]$RtCntr
$nwRtCntr = ($nwRtCntr + 1)
Out-File -FilePath $RtCntrMT -InputObject $nwRtCntr

#REMOVE SESSION FILES
Remove-Item -Path $IniMT
Remove-Item -Path $RetryMT
Remove-Item -Path $Webcheck

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
Out-File -FilePath $RtErr -InputObject "Unresolved error during login..." -Encoding default
Start-Process -FilePath $RuntimeError
Start-Sleep -Seconds 1

Start-Process -FilePath $LoginError

}
