Attribute VB_Name = "App_Loc"
'/####################################\
'//Application File & Folder Locations\\
'///##################################\\\

'//pers file location
Function xPersFile(persFile) As String

persFile = AppLoc & "\presets\" & xProfile & "\pers\"

End Function
'//API pers file
Function xApiFile(apiFile) As String

apiFile = AppLoc & "\presets\" & xProfile & "\pers\api.pers"

End Function
'//twt file location
Function xTwtFile(twtFile) As String

If ETWEETXLPOST.ProfileListBox.Value = "" Then
twtFile = AppLoc & "\presets\" & xProfile & "\twt\"
    Else
        twtFile = AppLoc & "\presets\" & ETWEETXLPOST.ProfileListBox.Value & "\twt\"
            End If

End Function
'//thr file location
Function xThrFile(thrFile) As String

 On Error Resume Next
 
If ETWEETXLPOST.ProfileListBox.Value = "" Then
thrFile = AppLoc & "\presets\" & xProfile & "\thr\"
    Else
        thrFile = AppLoc & "\presets\" & ETWEETXLPOST.ProfileListBox.Value & "\thr\"
            End If
            
            If Dir(thrFile) = vbNullString Then MkDir (thrFile)

End Function
'//debug folder location
Function debugLoc() As String

debugLoc = AppLoc & "\debug\"

End Function
'//shell folder location
Function xShellWinFldr(xShellWin) As String

xShellWin = AppLoc & "\shell\win\"

End Function
'//app folder location
Function xAppFldr(appFldr) As String

appFldr = AppLoc & "\app\"

End Function
'//temp folder location
Function xTempFldr(tempFldr) As String

tempFldr = AppLoc & "\app\temp\"

End Function
'//MT trigger files
Function xMTapi(mtApi) As String

mtApi = AppLoc & "\mtsett\api.mt"

End Function
Function xMTBlank(mtBlank) As String

mtBlank = AppLoc & "\mtsett\blank.mt"

End Function
Function xMTCheck(mtCheck) As String

mtCheck = AppLoc & "\mtsett\check.mt"

End Function
Function xMTDynOff(mtDynOff) As String

mtDynOff = AppLoc & "\mtsett\dynoff.mt"

End Function
Function xMTini(mtIni) As String

mtIni = AppLoc & "\mtsett\ini.mt"

End Function
Function xMTOffset(mtOffset) As String

mtOffset = AppLoc & "\mtsett\offset.mt"

End Function
Function xMTOffsetCopy(mtOffsetCopy) As String

mtOffsetCopy = AppLoc & "\mtsett\offset.mtc"

End Function
Function xMTUser(mtUser) As String

mtUser = AppLoc & "\mtsett\user.mt"

End Function
Function xMTPass(mtPass) As String

mtPass = AppLoc & "\mtsett\pass.mt"

End Function
Function xMTPost(mtPost) As String

mtPost = AppLoc & "\mtsett\post.mt"

End Function
Function xMTProf(mtProf) As String

mtProf = AppLoc & "\mtsett\profile.mt"

End Function
Function xMTMed(mtMed) As String

mtMed = AppLoc & "\mtsett\med.mt"

End Function
Function xMTRuntime(mtRuntime) As String

mtRuntime = AppLoc & "\mtsett\runtime.mt"

End Function
Function xMTRuntimeCntr(mtRuntimeCntr) As String

mtRuntimeCntr = AppLoc & "\mtsett\rtcntr.mt"

End Function
Function xMTRetryCntr(mtRetryCntr) As String

mtRetryCntr = AppLoc & "\mtsett\retry.mt"

End Function
Function xMTTwt(mtTwt) As String

mtTwt = AppLoc & "\mtsett\twt.mt"

End Function
Function xMTThread(mtThread) As String

mtThread = AppLoc & "\mtsett\thread.mt"

End Function
Function xMTThreadCt(mtThreadCt) As String

mtThreadCt = AppLoc & "\mtsett\threadct.mt"

End Function
'//start application script
Function xApp_StartLink(appStartLink) As String

appStartLink = AppLoc & "\shell\win\load_login.bat"

End Function
'//default flowstrips folder
Function flowStripsLoc(flowStripsFol)

flowStripsFol = Env & "\documents\flowstrips\"

End Function
'//XLFR extract file
Function extractLoc(extractFile) As String

extractFile = AppLoc & "\shell\win\extract.xlfr"

End Function
'//application debug
Function syntaxError(errorFile) As String

errorFile = AppLoc & "\debug\syntax.err"

End Function
'//application backup script
Function backupScript() As String

backupScript = AppLoc & "\shell\win\backup.bat"

End Function
'//python send script
Function apiStartScript() As String

apiStartScript = AppLoc & "\app\send_with_api.py"

End Function
'//OBS data folder
Function OBSDataLoc(OBSDataFol) As String

OBSDataFol = Env & "\.z7\autokit\obs\data\"

End Function
'//python location
Function pyLoc() As String

pyLoc = Env & "\AppData\Local\Programs\Python\Python39\python.exe"

End Function




