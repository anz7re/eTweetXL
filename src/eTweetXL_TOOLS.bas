Attribute VB_Name = "eTweetXL_TOOLS"
'/##################\
'//Application Tools\\
'///################\\\

Sub delAppData()

'//Remove all running application data

    On Error Resume Next
    
    Range("LinkTrig").Value = "0"
    
    Call eTweetXL_LOC.xMTapi(mtApi)
    Call eTweetXL_LOC.xMTBlank(mtBlank)
    Call eTweetXL_LOC.xMTCheck(mtCheck)
    Call eTweetXL_LOC.xMTDynOff(mtDynOff)
    Call eTweetXL_LOC.xMTini(mtIni)
    Call eTweetXL_LOC.xMTOffset(mtOffset)
    Call eTweetXL_LOC.xMTOffsetCopy(mtOffsetCopy)
    Call eTweetXL_LOC.xMTMed(mtMed)
    Call eTweetXL_LOC.xMTPass(mtPass)
    Call eTweetXL_LOC.xMTPost(mtPost)
    Call eTweetXL_LOC.xMTProf(mtProf)
    Call eTweetXL_LOC.xMTThread(mtThread)
    Call eTweetXL_LOC.xMTThreadCt(mtThreadCt)
    Call eTweetXL_LOC.xMTTwt(mtTwt)
    Call eTweetXL_LOC.xMTUser(mtUser)
    Call eTweetXL_LOC.xMTRuntime(mtRuntime)
    Call eTweetXL_LOC.xMTRuntimeCntr(mtRuntimeCntr)
    Call eTweetXL_LOC.xMTRetryCntr(mtRetryCntr)
    
    If Dir(mtApi) <> "" Then Kill (mtApi)
    If Dir(mtBlank) <> "" Then Kill (mtBlank)
    If Dir(mtCheck) <> "" Then Kill (mtCheck)
    If Dir(mtDynOff) <> "" Then Kill (mtDynOff)
    If Dir(mtIni) <> "" Then Kill (mtIni)
    If Dir(mtOffset) <> "" Then Kill (mtOffset)
    If Dir(mtOffsetCopy) <> "" Then Kill (mtOffsetCopy)
    If Dir(mtMed) <> "" Then Kill (mtMed)
    If Dir(mtPass) <> "" Then Kill (mtPass)
    If Dir(mtPost) <> "" Then Kill (mtPost)
    If Dir(mtProf) <> "" Then Kill (mtProf)
    If Dir(mtThread) <> "" Then Kill (mtThread)
    If Dir(mtThreadCt) <> "" Then Kill (mtThreadCt)
    If Dir(mtTwt) <> "" Then Kill (mtTwt)
    If Dir(mtUser) <> "" Then Kill (mtUser)
    If Dir(mtRuntime) <> "" Then Kill (mtRuntime)
    If Dir(mtRuntimeCntr) <> "" Then Kill (mtRuntimeCntr)
    If Dir(mtRetryCntr) <> "" Then Kill (mtRetryCntr)

End Sub
Function fndChar(xChar)

'//Check char entered for Time box
xChar = "a,A,b,B,c,C,d,D,e,E,f,F,g,G,h,H,i,I,j,J,k,K,l,L,m,M,n,N,o,O,p,P,q,Q,r,R,s,S,t,T,u,U,v,V,w,W,x,X,y,Y,z,Z,`,~,!,@,#,$,%,^,&,*,(,),_,-,+,=,[,],{,},\,|,;,',<,>,?,/,."
xChar = xChar & ","","

X = 1
xLetters = Split(xChar, ",")
xLast = UBound(xLetters) - LBound(xLetters)

Do Until X = xLast
If InStr(1, ETWEETXLPOST.TimeBox.Value, xLetters(X)) Then ETWEETXLPOST.TimeBox.Value = Replace(ETWEETXLPOST.TimeBox.Value, xLetters(X), "0"): _
xMsg = 4: Call eTweetXL_MSG.AppMsg(xMsg, errLvl): xChar = "(*Err)": Exit Function
X = X + 1
Loop

End Function
Function fndChar2(xChar)

'//Check char entered for Offset box
xChar = "a,A,b,B,c,C,d,D,e,E,f,F,g,G,h,H,i,I,j,J,k,K,l,L,m,M,n,N,o,O,p,P,q,Q,r,R,s,S,t,T,u,U,v,V,w,W,x,X,y,Y,z,Z,`,~,!,@,#,$,%,^,&,*,(,),_,-,+,=,[,],{,},\,|,;,',<,>,?,/,."
xChar = xChar & ","","

X = 1
xLetters = Split(xChar, ",")
xLast = UBound(xLetters) - LBound(xLetters)

Do Until X = xLast
If InStr(1, ETWEETXLPOST.OffsetBox.Value, xLetters(X)) Then ETWEETXLPOST.OffsetBox.Value = Replace(ETWEETXLPOST.OffsetBox.Value, xLetters(X), "0"): _
xMsg = 4: Call eTweetXL_MSG.AppMsg(xMsg, errLvl): xChar = "(*Err)": Exit Function
X = X + 1
Loop

End Function
Sub enlFlowStrip()

ETWEETXLHOME.xlFlowStrip.Enabled = True
ETWEETXLPOST.xlFlowStrip.Enabled = True
ETWEETXLQUEUE.xlFlowStrip.Enabled = True
ETWEETXLSETUP.xlFlowStrip.Enabled = True

End Sub
Sub disFlowStrip()

ETWEETXLHOME.xlFlowStrip.Enabled = False
ETWEETXLPOST.xlFlowStrip.Enabled = False
ETWEETXLQUEUE.xlFlowStrip.Enabled = False
ETWEETXLSETUP.xlFlowStrip.Enabled = False

End Sub
Sub shtFormat()

Dim lastR, X As Integer
lastR = Cells(Rows.Count, "L").End(xlUp).Row

X = 2
If lastR > X Then

Do Until X = lastR
Range("MainLink").Offset(X, 0).NumberFormat = "General"
Range("UserLink").Offset(X, 0).NumberFormat = "General"
Range("apiLink").Offset(X, 0).NumberFormat = "General"
Range("ProfileLink").Offset(X, 0).NumberFormat = "General"
Range("Draftlink").Offset(X, 0).NumberFormat = "General"
Range("Runtime").Offset(X, 0).NumberFormat = "hh:mm:ss"
X = X + 1
Loop
    End If
    
End Sub
Sub sndAsThread(xLink)

Dim xPostArr(5000), xFinPostArr(5000), xMyMedArr(5000), xWasteArr(5000) As String
Dim xTemp As String
Dim X, Y, Z As Integer

Call eTweetXL_LOC.xMTapi(mtApi)
Call eTweetXL_LOC.xMTThread(mtThread)
Call eTweetXL_LOC.xMTThreadCt(mtThreadCt)
Call eTweetXL_LOC.xMTMed(mtMed)
Call eTweetXL_LOC.xMTPost(mtPost)
Call eTweetXL_LOC.xApp_StartLink(appStartLink)

LLCntr = Range("LinkerCount").Value2: If LLCntr = "" Then LLCntr = 0

'//check for threaded post
If Range("apiLink").Offset(LLCntr + 1, 0).Value = "(*api)" Then GoTo SetupAPI

If Dir(mtThreadCt) = "" Then

'//find thread count if first run through
Open xLink For Input As #1: Do Until EOF(1): Line Input #1, xTemp: Loop: Close #1: _
xTemp = Replace(xTemp, "*-(", vbNullString): _
xTemp = Replace(xTemp, ");", vbNullString)
xTemp = Replace(xTemp, " ", vbNullString)

'//create thread trigger
Open mtThread For Output As #1: Print #1, xTemp: Close #1

'//create thread count file & trigger
Open mtThreadCt For Output As #2: xCt = "1:" & xTemp: Print #2, xCt: Close #2
    Else
        '//we're back...
        Open mtThreadCt For Input As #2: Line Input #2, xCt: Close #2
        
            End If
            

'//find count and total
xCtArr = Split(xCt, ":")

'/######################################################################
'//end of thread
If xCt = CInt(xCtArr(1)) + 1 & ":" & xCtArr(1) Then

'//remove thread files
Kill (mtThread)
Kill (mtThreadCt)

'//increment Linker count
Range("LinkerCount").Value2 = Range("LinkerCount").Value2 + 1

'//deactivate thread trigger
Range("ThreadStatus").Value2 = 0

'//check for empty linker & clear...
If Range("LinkerCount").Value2 = (Range("LinkerTotal").Value2) Then
Call clnLinker2
Exit Sub
    End If
    
    Call eTweetXL_CLICK.StartBtn_Clk
    
    '//Cleanup Queue
    ETWEETXLQUEUE.QueueBox.Clear
    ETWEETXLQUEUE.RuntimeBox.Clear
    ETWEETXLQUEUE.UserBox.Clear
    
    '//Update application state
    Call eTweetXL_TOOLS.updAppState
    Call eTweetXL_GET.getQueueData
    Exit Sub
    End If
                
'/######################################################################
'//
'//open thread file
Open xLink For Input As #3

X = 0

Do Until EOF(3)
    
    '//collect post info
    Do Until InStr(1, xMyPostHldr, "*-(" & xCtArr(0) & ");")
    Line Input #3, xMyPostHldr
    If InStr(1, xMyPostHldr, "*-(" & CInt(xCtArr(0)) - 1 & ");") Then xPostArr(X) = "*/SKIP" Else xPostArr(X) = xMyPostHldr
    X = X + 1
    Loop
    
    '//end of post
    xPostArr(X) = "*/END"
    
    '//find current thread
    X = 0
    Do Until InStr(1, xPostArr(X), "*/SKIP") Or InStr(1, xPostArr(X), "*/END")
    xWasteArr(X) = xPostArr(X)
    X = X + 1
    Loop
    
    '//set to next thread or reset if first run
    If CInt(xCtArr(0)) > 1 Then X = X + 1 Else X = 0
    
    Z = 0
    '//get current post from thread
    Do Until InStr(1, xPostArr(X), "*-(" & xCtArr(0) & ");")
    '//check for media
    If InStr(1, xPostArr(X), "*-") Then xMyMed = xPostArr(X): xMyMed = Replace(xMyMed, "*-", vbNullString) Else xMyPost = xMyPost & xPostArr(X) '//set post
    xFinPostArr(Z) = xPostArr(X)
    X = X + 1
    Loop
    
    
    '//Escape special characters...
    xMyPost = Replace(xMyPost, "{ENTER};", "*/ENTER")
    xMyPost = Replace(xMyPost, "{SPACE};", " ")
    xMyPost = Replace(xMyPost, "{", "{++")
    xMyPost = Replace(xMyPost, "}", "++}")
    xMyPost = Replace(xMyPost, "{++", "{{}")
    xMyPost = Replace(xMyPost, "++}", "{}}")
    xMyPost = Replace(xMyPost, "+", "{+}")
    xMyPost = Replace(xMyPost, "^", "{^}")
    xMyPost = Replace(xMyPost, "%", "{%}")
    xMyPost = Replace(xMyPost, "~", "{~}")
    xMyPost = Replace(xMyPost, "(", "{(}")
    xMyPost = Replace(xMyPost, ")", "{)}")
    xMyPost = Replace(xMyPost, "[", "{[}")
    xMyPost = Replace(xMyPost, "]", "{]}")
    xMyPost = Replace(xMyPost, "*/ENTER", "{ENTER}")
    xMyMed = Replace(xMyMed, "{", "{++")
    xMyMed = Replace(xMyMed, "}", "++}")
    xMyMed = Replace(xMyMed, "{++", "{{}")
    xMyMed = Replace(xMyMed, "++}", "{}}")
    xMyMed = Replace(xMyMed, "+", "{+}")
    xMyMed = Replace(xMyMed, "^", "{^}")
    xMyMed = Replace(xMyMed, "%", "{%}")
    xMyMed = Replace(xMyMed, "~", "{~}")
    xMyMed = Replace(xMyMed, "(", "{(}")
    xMyMed = Replace(xMyMed, ")", "{)}")
    xMyMed = Replace(xMyMed, "[", "{[}")
    xMyMed = Replace(xMyMed, "]", "{]}")
   
'//Replace markers w/ enter...
xMyPost = Replace(xMyPost, "*-;", vbCrLf)
             
    Open mtMed For Output As #7
    Print #7, xMyMed
    Close #7
    
    Open mtPost For Output As #8
    Print #8, xMyPost
    Close #8

'//increment thread counter
Open mtThreadCt For Output As #2: xCt = CInt(xCtArr(0)) + 1 & ":" & xCtArr(1): Print #2, xCt: Close #2

Call clnOpenFiles
   
If Range("ThreadStatus").Value2 = 1 Then Exit Sub

On Error Resume Next

GoTo CloseOut

'//setup post for sending a thread using Twitter API
SetupAPI:
'/######################################################################
'//check for empty linker
If Range("LinkerCount").Value2 = (Range("LinkerTotal").Value2) Then
Call clnLinker2
Exit Sub
End If
                
'/######################################################################
'//
'//open thread file
Open xLink For Input As #3

X = 0

Do Until EOF(3)
    
    '//collect post info
    Line Input #3, xMyPostHldr
    xPostArr(X) = xMyPostHldr
    X = X + 1
    Loop
    Close #3
    
    '//end of thread
    xPostArr(X) = "*/END"
    
    '//find end of thread
    X = 0
    Do Until InStr(1, xPostArr(X), "*/END")
    xWasteArr(X) = xPostArr(X)
    X = X + 1
    Loop
    
    X = 0
    Y = 0
    Z = 0
    '//
    Do Until InStr(1, xPostArr(X), "*/END")
    '//check for media
    If InStr(1, xPostArr(X), "*-") Then xMyMedArr(Y) = xPostArr(X): xMyMedArr(Y) = Replace(xMyMedArr(Y), "*-", vbNullString): Y = Y + 1 '//set post
    xFinPostArr(Z) = xPostArr(X)
    
    '//Escape special characters...
    xPostArr(X) = Replace(xPostArr(X), "{ENTER};", "*/ENTER")
    xPostArr(X) = Replace(xPostArr(X), "{SPACE};", " ")
    xPostArr(X) = Replace(xPostArr(X), "{", "{++")
    xPostArr(X) = Replace(xPostArr(X), "}", "++}")
    xPostArr(X) = Replace(xPostArr(X), "{++", "{{}")
    xPostArr(X) = Replace(xPostArr(X), "++}", "{}}")
    xPostArr(X) = Replace(xPostArr(X), "+", "{+}")
    xPostArr(X) = Replace(xPostArr(X), "^", "{^}")
    xPostArr(X) = Replace(xPostArr(X), "%", "{%}")
    xPostArr(X) = Replace(xPostArr(X), "~", "{~}")
    xPostArr(X) = Replace(xPostArr(X), "(", "{(}")
    xPostArr(X) = Replace(xPostArr(X), ")", "{)}")
    xPostArr(X) = Replace(xPostArr(X), "[", "{[}")
    xPostArr(X) = Replace(xPostArr(X), "]", "{]}")
    xPostArr(X) = Replace(xPostArr(X), "*/ENTER", "{ENTER}")
    xMyMedArr(Y) = Replace(xMyMedArr(Y), "{", "{++")
    xMyMedArr(Y) = Replace(xMyMedArr(Y), "}", "++}")
    xMyMedArr(Y) = Replace(xMyMedArr(Y), "{++", "{{}")
    xMyMedArr(Y) = Replace(xMyMedArr(Y), "++}", "{}}")
    xMyMedArr(Y) = Replace(xMyMedArr(Y), "+", "{+}")
    xMyMedArr(Y) = Replace(xMyMedArr(Y), "^", "{^}")
    xMyMedArr(Y) = Replace(xMyMedArr(Y), "%", "{%}")
    xMyMedArr(Y) = Replace(xMyMedArr(Y), "~", "{~}")
    xMyMedArr(Y) = Replace(xMyMedArr(Y), "(", "{(}")
    xMyMedArr(Y) = Replace(xMyMedArr(Y), ")", "{)}")
    xMyMedArr(Y) = Replace(xMyMedArr(Y), "[", "{[}")
    xMyMedArr(Y) = Replace(xMyMedArr(Y), "]", "{]}")

'//Replace markers w/ enter...
xPostArr(X) = Replace(xPostArr(X), "*-;", vbCrLf)

If xMyMedArr(Y) <> vbNullString Then
    Open mtMed For Append As #7
    Print #7, xMyMedArr(Y)
    Close #7
        End If
        
    Open mtPost For Append As #8
    Print #8, xPostArr(X)
    Close #8

X = X + 1

Loop

Call clnOpenFiles

On Error Resume Next

CloseOut:
         '//Refresh progress bar...
        Call eTweetXL_TOOLS.updProgBar
        Call fxsUnfreeze
        
        ETWEETXLHOME.ActiveUser.Caption = xUser
        ETWEETXLSETUP.ActiveUser.Caption = xUser
        ETWEETXLPOST.ActiveUser.Caption = xUser
        ETWEETXLQUEUE.ActiveUser.Caption = xUser
        
        ETWEETXLHOME.xlFlowStrip.Value = "Starting script..."
        ETWEETXLSETUP.xlFlowStrip.Value = "Starting script..."
        ETWEETXLPOST.xlFlowStrip.Value = "Starting script..."
        ETWEETXLQUEUE.xlFlowStrip.Value = "Starting script..."
        
'//Turn Linker active...
ETWEETXLHOME.AppStatus.Caption = "ON"
ETWEETXLHOME.AppStatus.ForeColor = vbGreen
ETWEETXLHOME.AppStatus.BackColor = vbWhite
If Range("AppState").Value2 <> 1 Then Range("AppState").Value2 = 1

SendWithAPI:
'//Send with api method...
If Range("apiLink").Offset(LLCntr + 1, 0).Value = "(*api)" Then
    API_LINK = 1
    Call eTweetXL_TOOLS.bldAPIScript2
    Open mtApi For Output As #7
    Print #7, ""
    Close #7
    Shell (appStartLink), vbMinimizedNoFocus
    
    '//increment Linker count
    Range("LinkerCount").Value2 = Range("LinkerCount").Value2 + 1
    
        Else
        '//Send with default method...
           If Dir(mtApi) <> "" Then Kill (mtApi)
           Shell (appStartLink), vbMinimizedNoFocus
                End If
    
    '//activate thread trigger
    Range("ThreadStatus").Value2 = 1
    
    '//Cleanup Queue
    ETWEETXLQUEUE.QueueBox.Clear
    ETWEETXLQUEUE.RuntimeBox.Clear
    ETWEETXLQUEUE.UserBox.Clear
    
    '//Update application state
    Call eTweetXL_TOOLS.updAppState
    Call eTweetXL_GET.getQueueData
    
Exit Sub

Loop
    
    
End Sub
Sub bldAPIScript()

'//For sending singular posts through the Twitter API w/ Python+Tweepy+xlwings

Dim xTemp, xTxt, xTxtHdlr, xMed As String
Dim X, xLast As Integer

Call eTweetXL_LOC.xMTapi(mtApi)
Call eTweetXL_LOC.xMTPost(mtPost)
Call eTweetXL_LOC.xMTMed(mtMed)
Call eTweetXL_LOC.xApiFile(apiFile)
Call eTweetXL_LOC.xAppFldr(appFldr)
Call eTweetXL_LOC.xTempFldr(tempFldr)

'//Create temp folder if not one already
On Error Resume Next
If Dir(tempFldr) = "" Then MkDir tempFldr
Err.Clear

'//Create start API .bat script
Open eTweetXL_LOC.apiStartScript For Output As #1

Print #1, "@echo off"
Print #1, "start " & apiScript
Print #1, "exit"

Close #1

'//Get API key information
If Dir(apiFile) <> "" Then
Open apiFile For Input As #1

Line Input #1, apiKey
Line Input #1, apiSecret
Line Input #1, acctoken
Line Input #1, accSecret

Close #1
End If

'//Check for text
If Dir(mtPost) <> "" Then
Open mtPost For Input As #2
Do Until EOF(2)
Line Input #2, xTxt
xTxt = Replace(xTxt, "{ENTER};", "\n")
xTxt = Replace(xTxt, Chr(10), "\n")
xTxt = Replace(xTxt, Chr(13), "\n")
xTxt = Replace(xTxt, "{SPACE};", " ")
xTxtHldr = xTxtHldr & xTxt
Loop
Close #2
End If

'//Check for media
X = 1
If Dir(mtMed) <> "" Then
Open mtMed For Input As #3
Do Until EOF(3)
Line Input #3, xMed
xMed = Replace(xMed, """", "")
xExt = xMed
If xExt <> "" Then Call eTweetXL_GET.getExt(xExt) '//find file extension
xTemp = tempFldr & "temp_" & X & "." & xExt

On Error GoTo SplitMed
FileCopy xMed, xTemp '//copy media to shell folder
xTemp = Replace(xTemp, tempFldr, "") '//remove file path from string
xTemp = "'" & xTemp & "'" '//add single quotes
X = X + 1
Loop
Close #3

SplitMed:
Close #3
'//Multiple media found
X = 1
xMedArr = Split(xMed, "C:\")

xLast = UBound(xMedArr) - LBound(xMedArr) + 1

If xLast < X Then GoTo SkipHere '//

Do Until X = xLast
xExt = xMedArr(X)
If xExt <> "" Then Call eTweetXL_GET.getExt(xExt) '//find extension
xTemp = tempFldr & "temp_" & X & "." & xExt

FileCopy "C:\" & xMedArr(X), xTemp '//copy to temp folder
xTemp = Replace(xTemp, tempFldr, "") '//remove file path from string
xTemp = "'" & xTemp & "'" '//add single quotes
If X > 1 Then xTempHldr = xTempHldr & ", " & xTemp Else: xTempHldr = xTemp
X = X + 1
Loop

End If

SkipHere:
xTemp = xTempHldr
xTxt = xTxtHldr

'//
If xTemp = "" Then GoTo JustText
If xTxt = "" Then GoTo JustMed
GoTo TextMed

'//
If xTemp = "" Then GoTo JustText
If xTxt = "" Then GoTo JustMed
GoTo TextMed

JustText:
'//Create Python API Script (Just send text)
Open eTweetXL_LOC.apiStartScript For Output As #1

Print #1, "import tweepy"
Print #1, "auth = tweepy.OAuthHandler(" & """" & apiKey & """" & ", " & """" & apiSecret & """" & ")"
Print #1, "auth.set_access_token(" & """" & acctoken & """" & ", " & """" & accSecret & """" & ")"
Print #1, "api = tweepy.API(auth)"
Print #1, "tweet = " & """" & xTxt & """"
Print #1, "api.update_status(status=tweet)"
Print #1, "print(" & """" & "Tweet sent" & """" & ")"

Close #1
Exit Sub

JustMed:
'//Create Python API Script (Just send media)
Open eTweetXL_LOC.apiStartScript For Output As #1

Print #1, "import tweepy"
Print #1, "auth = tweepy.OAuthHandler(" & """" & apiKey & """" & ", " & """" & apiSecret & """" & ")"
Print #1, "auth.set_access_token(" & """" & acctoken & """" & ", " & """" & accSecret & """" & ")"
Print #1, "api = tweepy.API(auth)"
Print #1, "Media = api.media_upload(" & xTemp & ")"
Print #1, "tweet = " & """" & xTxt & """"
Print #1, "api.update_status(status=tweet, media_ids=[Media.media_id])"
Print #1, "print(" & """" & "Tweet sent" & """" & ")"

Close #1
Exit Sub

TextMed:
'//Create Python API Script (Send text & media)
Open eTweetXL_LOC.apiStartScript For Output As #1

Print #1, "import tweepy"
Print #1, "auth = tweepy.OAuthHandler(" & """" & apiKey & """" & ", " & """" & apiSecret & """" & ")"
Print #1, "auth.set_access_token(" & """" & acctoken & """" & ", " & """" & accSecret & """" & ")"
Print #1, "api = tweepy.API(auth)"
Print #1, "Media = api.media_upload(" & xTemp & ")"
Print #1, "tweet = " & """" & xTxt & """"
Print #1, "api.update_status(status=tweet, media_ids=[Media.media_id])"
Print #1, "print(" & """" & "Tweet sent" & """" & ")"

Close #1
Exit Sub

End Sub
Sub bldAPIScript2()

'//For sending threads through the Twitter API w/ Python+Tweepy+xlwings

Dim xNwMedArr(5000) As String
Dim xTemp, xTxt, xTxtHldr, xMed, xMedHldr As String
Dim MEDIAFOUND, TEXTFOUND, N, T, X As Integer

Call eTweetXL_LOC.xMTapi(mtApi)
Call eTweetXL_LOC.xMTPost(mtPost)
Call eTweetXL_LOC.xMTMed(mtMed)
Call eTweetXL_LOC.xApiFile(apiFile)
Call eTweetXL_LOC.xAppFldr(appFldr)
Call eTweetXL_LOC.xTempFldr(tempFldr)

'//Create temp folder if not one already
On Error Resume Next
If Dir(tempFldr) = "" Then MkDir tempFldr
Err.Clear

'//Create start API .bat script
Open eTweetXL_LOC.apiStartScript For Output As #1

Print #1, "@echo off"
Print #1, "start " & apiScript
Print #1, "exit"

Close #1

'//Get API key information
If Dir(apiFile) <> "" Then
Open apiFile For Input As #1

Line Input #1, apiKey
Line Input #1, apiSecret
Line Input #1, acctoken
Line Input #1, accSecret

Close #1
End If

'//Check for text
If Dir(mtPost) <> "" Then
Open mtPost For Input As #2
Do Until EOF(2)
FindNext:
Line Input #2, xTxt
If InStr(1, xTxt, "*-") And InStr(1, xTxt, "*-{") = False Then xMed = xTxt: xMedHldr = xMedHldr & xMed: GoTo FindNext
xTxt = Replace(xTxt, "{ENTER};", "\n")
xTxt = Replace(xTxt, Chr(10), "\n")
xTxt = Replace(xTxt, Chr(13), "\n")
xTxt = Replace(xTxt, "{SPACE};", " ")
xTxtHldr = xTxtHldr & xTxt
Loop
Close #2
End If

'//text found trigger
If xTxtHldr <> vbNullString Then TEXTFOUND = 1

'//find & seperate text for threads
xTxtArr = Split(xTxtHldr, "{)};")
xTxtHldr = vbNullString
For X = 0 To UBound(xTxtArr)
xTxtArr(X) = Replace(xTxtArr(X), "*-{(}" & X + 1, vbNullString)
xTxtHldr = xTxtHldr & xTxtArr(X)
Next

'//find & seperate media for threads
If xMedHldr <> vbNullString Then

xMedArr = Split(xMedHldr, "*-")

For X = 0 To UBound(xMedArr)

If xMedArr(X) <> vbNullString Then
'//media found trigger
MEDIAFOUND = 1
xTempArr = Split(xMedArr(X), """")
    For T = 0 To UBound(xTempArr)
    xTempArr(T) = Replace(xTempArr(T), """", vbNullString)
    
    xExt = xTempArr(T)
    If xExt <> "" Then
    Call eTweetXL_GET.getExt(xExt) '//find file extension
    xMedArr(X) = Replace(xMedArr(X), """", vbNullString)
    xTemp = tempFldr & "temp_" & X & "." & xExt
    FileCopy xMedArr(X), xTemp
    xTempArr(T) = xTemp
    End If
   
    If xTempArr(T) <> vbNullString Then xTempArr(T) = "'" & xTempArr(T) & "'"
    If xTempArr(T) <> vbNullString Then If T > 1 Then xTempHldr = xTempHldr & ", " & xTempArr(T) Else: xTempHldr = xTempArr(T)
        Next
        xMedArr(X) = xTempHldr
            End If
                Next
                
        N = 0
        For X = 0 To UBound(xMedArr)
        If xMedArr(X) <> vbNullString Then xNwMedArr(N) = xMedArr(X): _
        xNwMedArr(N) = Replace(xNwMedArr(N), tempFldr, vbNullString): N = N + 1
        Next
    
        xNwMedArr(N) = "*/END"
    
                            End If

'//send text no media
If TEXTFOUND = 1 And MEDIAFOUND <> 1 Then GoTo JustText
'//send media no text
If MEDIAFOUND = 1 And TEXTFOUND <> 1 Then GoTo JustMed
'//send text & media
GoTo TextMed

JustText:
'//Create Python API Script (Just send text)
Open eTweetXL_LOC.apiStartScript For Output As #1

Print #1, "import tweepy"
Print #1, "auth = tweepy.OAuthHandler(" & """" & apiKey & """" & ", " & """" & apiSecret & """" & ")"
Print #1, "auth.set_access_token(" & """" & acctoken & """" & ", " & """" & accSecret & """" & ")"
Print #1, "api = tweepy.API(auth)"
Print #1, "txt1 = " & """" & xTxtArr(0) & """"
Print #1, "tweet1 = api.update_status(status=txt1)"

'//setup reply to tweet(s)
X = 1
Do Until X = UBound(xTxtArr) Or InStr(1, xTxtArr(X), "*/END")
Print #1, "txt" & X + 1 & " = " & """" & xTxtArr(X) & """"
Print #1, "tweet" & X + 1 & " = api.update_status(status=txt" & X + 1 & ", in_reply_to_status_id=tweet" & X & ".id, auto_populate_reply_metadata=True)"
X = X + 1
Loop

Print #1, "print(" & """" & "Tweet sent" & """" & ")"

Close #1
Exit Sub

JustMed:
'//Create Python API Script (Just send media)
Open eTweetXL_LOC.apiStartScript For Output As #2

Print #2, "import tweepy"
Print #2, "auth = tweepy.OAuthHandler(" & """" & apiKey & """" & ", " & """" & apiSecret & """" & ")"
Print #2, "auth.set_access_token(" & """" & acctoken & """" & ", " & """" & accSecret & """" & ")"
Print #2, "api = tweepy.API(auth)"
Print #2, "med1 = api.media_upload(" & xNwMedArr(0) & ")"
Print #2, "txt1 = " & """" & xTxt & """"
Print #2, "tweet1 = api.update_status(status=txt1, media_ids=[med1.media_id])"

'//setup reply to tweet(s)
X = 1
Do Until X = UBound(xNwMedArr) Or InStr(1, xNwMedArr(X), "*/END")
Print #2, "med" & X + 1 & " = api.media_upload(" & xNwMedArr(X) & ")"
Print #2, "txt" & X + 1 & " = " & """" & xTxt & """"
Print #2, "tweet" & X + 1 & " = api.update_status(status=txt" & X + 1 & ", media_ids=[med" & X + 1 & ".media_id]), in_reply_to_status_id=tweet" & X & ".id, auto_populate_reply_metadata=True)"
X = X + 1
Loop

Print #2, "print(" & """" & "Tweet sent" & """" & ")"

Close #2
Exit Sub

TextMed:
'//Create Python API Script (Send text & media)
Open eTweetXL_LOC.apiStartScript For Output As #3

Print #3, "import tweepy"
Print #3, "auth = tweepy.OAuthHandler(" & """" & apiKey & """" & ", " & """" & apiSecret & """" & ")"
Print #3, "auth.set_access_token(" & """" & acctoken & """" & ", " & """" & accSecret & """" & ")"
Print #3, "api = tweepy.API(auth)"
Print #3, "med1 = api.media_upload(" & xNwMedArr(0) & ")"
Print #3, "txt1 = " & """" & xTxtArr(0) & """"
Print #3, "tweet1 = api.update_status(status=txt1, media_ids=[med1.media_id])"

'//setup reply to tweet(s)
X = 1
Do Until X = UBound(xTxtArr) Or InStr(1, xTxtArr(X), "*/END")
Print #3, "med" & X + 1 & " = api.media_upload(" & xNwMedArr(X) & ")"
Print #3, "txt" & X + 1 & " = " & """" & xTxtArr(X) & """"
Print #3, "tweet" & X + 1 & " = api.update_status(status=txt" & X + 1 & ", media_ids=[med" & X + 1 & ".media_id], in_reply_to_status_id=tweet" & X & ".id, auto_populate_reply_metadata=True)"
X = X + 1
Loop

Print #3, "print(" & """" & "Tweet sent" & """" & ")"

Close #3
Exit Sub

End Sub
'//BACKUP TWEETS TO .ZIP
Sub bldZipArchive()

Dim oFSO, oFldr, oSub As Object
Dim PresNm, DestFi, TargFi, MMDDYY As String
Dim nl As Variant
nl = vbNewLine
MMDDYY = Replace(Date, "/", "")

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFldr = oFSO.GetFolder(AppLoc & "\presets")

On Error Resume Next

'//CHECK FOR BACKUP LOCATIONS & CREATE IF NONE EXIST
If Dir(Env & "\.z7\backups") = "" Then MkDir (Env & "\.z7\backups")
If Dir(Env & "\.z7\backups\autokit") = "" Then MkDir (Env & "\.z7\backups\autokit")
If Dir(Env & "\.z7\backups\autokit\etweetxl") = "" Then MkDir (Env & "\.z7\backups\autokit\etweetxl")

TargFi = AppLoc & "\mtsett\mytarg.mt"
DestFi = AppLoc & "\mtsett\mydest.mt"

'//REMOVE TARGET & DESTINATION FILE
If Dir(TargFi) <> "" Then Kill (TargFi)
If Dir(DestFi) <> "" Then Kill (DestFi)

Open TargFi For Append As #1

For Each oSub In oFldr.SubFolders
Print #1, AppLoc & "\presets\" & oSub.name & "\twt"
If Dir(Env & "\.z7\backups\autokit\etweetxl\" & oSub.name & "\") = "" Then
MkDir (Env & "\.z7\backups\autokit\etweetxl\" & oSub.name & "\")
If Dir(Env & "\.z7\backups\autokit\etweetxl\" & oSub.name & "\twt") = "" Then
MkDir (Env & "\.z7\backups\autokit\etweetxl\" & oSub.name & "\twt")
End If
    End If

Open DestFi For Append As #2
Print #2, Env & "\.z7\backups\autokit\etweetxl\" & oSub.name & "\twt\" & oSub.name & "_backup_" & MMDDYY & ".zip"
Close #2

Next
Close #1

'//PRINT LAST BACKUP DATE
Open Env & "\.z7\backups\autokit\etweetxl\info.txt" For Output As #3
Print #3, Date
Close #3

Shell (eTweetXL_LOC.backupScript), vbMinimizedNoFocus

End Sub
Sub runPy()
    
'//Run python script w/ the help of xlwings
    RunPython ("import send_with_api")
    
End Sub
Sub runFlowStrip(xKey)

Call fndEnvironment(appEnv, appBlk)

If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = vbNullString Then _
Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = _
Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinFormLast").Value2

'//xlFlowStrip Keystrokes
Call fndWindow(xWin)

'//Set article to application library & xlFlowStrip text
xArt = "<lib>xtwt;" & xWin.xlFlowStrip.Value

'//Alt key
If xKey = 18 Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 + 18
Exit Sub
End If

'//Ctrl key
If xKey = vbKeyControl Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbKeyControl
Exit Sub
End If
  
'//Shift
If xKey = vbKeyShift Then
Range("xlasKeyCtrl").Value2 = vbKeyShift
Call lexKey(xArt)
Exit Sub
End If
  
'//Key Ctrl+Alt+W
If xKey = vbKeyW Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 35 Then
'//Hide hotkey...
Call eTweetXL_CLICK.HideBtn_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
Exit Sub
End If
End If

Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = ""
        
End Sub
Sub setBoxList(xBox)

On Error GoTo EndMacro

Dim InfoArr(5000) As String
Dim oBox As Object

Call fndWindow(xWin)

If xBox = 1 Then Set oBox = xWin.UserBox
If xBox = 2 Then Set oBox = ETWEETXLPOST.LinkerBox
If xBox = 3 Then Set oBox = ETWEETXLQUEUE.QueueBox
If xBox = 4 Then Set oBox = xWin.RuntimeBox

For X = 0 To oBox.ListCount
InfoArr(X) = oBox.List(X)
oBox.List(X) = vbNullString
InfoRep = Split(InfoArr(X), ") ")
If UBound(InfoRep) > 0 Then InfoArr(X) = "(" & X + 1 & ") " & InfoRep(1)
oBox.List(X) = InfoArr(X)
Next

EndMacro:
End Sub
Sub updAppState()

Call fndWindow(xWin)
Call basSetWinFormPos(xWin, X, Y)

'//Reset internal application error level
Range("AppErrRef").Value = 0

'//Set previous WinForm #
Range("xlasWinFormLast").Value2 = Range("xlasWinForm").Value2

'//Check for links
LLCntr = Range("LinkerCount").Value2

'//Check for UserLink
If Range("UserLink").Offset(LLCntr, 0).Value = "" Then
'//Set active user to last loaded...
xWin.ActiveUser.Caption = Range("User").Value
    Else
'//Set active user to UserLink...
xWin.ActiveUser.Caption = Range("UserLink").Offset(LLCntr, 0).Value
'//Turn Linker active box "ON" if running...
On Error GoTo NotHome
xWin.AppStatus.Caption = "ON"
xWin.AppStatus.ForeColor = vbGreen
xWin.AppStatus.BackColor = vbWhite
End If

'//not updating the home window...
NotHome:
'//Set active user bg color...
If xWin.ActiveUser.Caption <> "" Then
xWin.ActiveUser.BackColor = vbWhite
Else
    xWin.ActiveUser.BackColor = -2147483633
            End If

If Range("AppState").Value2 = 1 Then
Call fxsUnfreeze
    ElseIf Range("AppState").Value2 = 2 Then
        Call fxsFreeze
            Else
                Call dfsFreeze
                    End If
                    
'//Set help wizard
If Range("HelpStatus").Value2 = vbNullString Then Range("HelpStatus").Value2 = 1: xWin.HelpStatus.Caption = "On": Exit Sub
If Range("HelpStatus").Value2 = 0 Then xWin.HelpStatus.Caption = "Off": Exit Sub
If Range("HelpStatus").Value2 = 1 Then xWin.HelpStatus.Caption = "On": Exit Sub

Set xWin = Nothing

End Sub
Sub updProgBar()

LLTotal = Range("LinkerTotal").Value2
LLCntr = Range("LinkerCount").Value2

'//Starting...
If LLCntr = 0 Then ETWEETXLHOME.ProgBar.Width = 0: Exit Sub
If LLCntr = 1 Then ETWEETXLHOME.ProgBar.Width = 5: Exit Sub

ProgBarStatus = (LLCntr * 156) / LLTotal

ETWEETXLHOME.ProgBar.Width = ProgBarStatus
ETWEETXLHOME.ProgRatio.Caption = LLCntr & "/" & LLTotal

End Sub
Public Function dfsHover()

'//default hover state

On Error Resume Next

Call fndWindow(xWin)

If Range("HelpStatus").Value2 <> 1 Then GoTo EndMacro

ETWEETXLPOST.DraftHdr.BackColor = &H8000000B
ETWEETXLPOST.DraftsHdr.BackColor = &H8000000B
ETWEETXLPOST.PostHdr.BackColor = &H8000000B
ETWEETXLPOST.TimeHdr.BackColor = &H8000000B
ETWEETXLPOST.OffsetHdr.BackColor = &H8000000B
ETWEETXLPOST.UsersHdr.BackColor = &H8000000B
ETWEETXLPOST.UserHdr.BackColor = &H8000000B
ETWEETXLPOST.RuntimeHdr.BackColor = &H8000000B
ETWEETXLPOST.LinkerHdr.BackColor = &HFFFF80
ETWEETXLPOST.UserBox.BackColor = &H80000005
ETWEETXLPOST.LinkerBox.BackColor = &H80000005
ETWEETXLPOST.RuntimeBox.BackColor = &H80000005

ETWEETXLQUEUE.PostHdr.BackColor = &H8000000B
ETWEETXLQUEUE.RuntimeBox.BackColor = &H80000005

If Range("HoverStatus").Value2 <> 0 And Range("HoverStatus").Value2 >= 27 Then Range("HoverStatus").Value2 = 0

EndMacro:
xWin.CtrlBoxBtn.BorderStyle = fmBorderStyleNone: xWin.CtrlBoxBtn.SpecialEffect = fmSpecialEffectEtched
xWin.FreezeBtn.BorderStyle = fmBorderStyleNone: xWin.FreezeBtn.SpecialEffect = fmSpecialEffectEtched
xWin.HelpIcon.ForeColor = &H80000011
xWin.HelpStatus.ForeColor = &H80000011
xWin.HideBtn.ForeColor = &H80000011
xWin.AppTag.BorderStyle = fmBorderStyleNone

Set xWin = Nothing

End Function
Public Function fxsHover(xHov)

'//fx hover state

On Error Resume Next

Call fndWindow(xWin)
xColor = vbGreen
'Cornflower - RGB(147, 204, 234)

If Range("HelpStatus").Value2 <> 1 Then GoTo EndMacro

If xHov = 1 Then ETWEETXLPOST.DraftHdr.BackColor = xColor: xColor = vbNullString: Exit Function
If xHov = 2 Then ETWEETXLPOST.DraftsHdr.BackColor = xColor: xColor = vbNullString: Exit Function
If xHov = 3 Then ETWEETXLPOST.OffsetHdr.BackColor = xColor: xColor = vbNullString: Exit Function
If xHov = 4 Then ETWEETXLPOST.PostHdr.BackColor = xColor: ETWEETXLQUEUE.PostHdr.BackColor = xColor: xColor = vbNullString: Exit Function
If xHov = 5 Then ETWEETXLPOST.RuntimeHdr.BackColor = xColor: xColor = vbNullString: Exit Function
If xHov = 6 Then ETWEETXLPOST.TimeHdr.BackColor = xColor: xColor = vbNullString: Exit Function
If xHov = 7 Then ETWEETXLPOST.UserHdr.BackColor = xColor: xColor = vbNullString: Exit Function
If xHov = 8 Then ETWEETXLPOST.UsersHdr.BackColor = xColor: xColor = vbNullString: Exit Function
If xHov = 9 Then ETWEETXLPOST.LinkerHdr.BackColor = xColor: xColor = vbNullString: Exit Function

If xHov = 11 Then ETWEETXLPOST.UserBox.BackColor = xColor: xColor = vbNullString: Exit Function
If xHov = 12 Then ETWEETXLPOST.LinkerBox.BackColor = xColor: xColor = vbNullString: Exit Function
If xHov = 13 Then xWin.RuntimeBox.BackColor = xColor: xColor = vbNullString: Exit Function

EndMacro:
'//These are kept highlighted when help is off...
If xHov = 10 Then xWin.AppTag.BorderStyle = fmBorderStyleSingle: xWin.AppTag.BorderColor = xColor: Exit Function
If xHov = 14 Then xWin.HelpIcon.ForeColor = xColor: xWin.HelpStatus.ForeColor = xColor: Exit Function
If xHov = 15 Then xWin.FreezeBtn.BorderStyle = fmBorderStyleSingle: xWin.FreezeBtn.BorderColor = xColor: Exit Function
If xHov = 16 Then xWin.CtrlBoxBtn.BorderStyle = fmBorderStyleSingle: xWin.CtrlBoxBtn.BorderColor = xColor: Exit Function
If xHov = 17 Then xWin.HideBtn.ForeColor = xColor: Exit Function

End Function
Sub dfsNaviBar()

'//default navigation bar state

On Error Resume Next

Call fndWindow(xWin)

xWin.HomeBtn.ForeColor = vbBlack
xWin.StartBtn.ForeColor = vbBlack
xWin.PostSetupBtn.ForeColor = vbBlack
xWin.ProfileSetupBtn.ForeColor = vbBlack
xWin.QueueBtn.ForeColor = vbBlack
xWin.BreakBtn.ForeColor = vbBlack

End Sub
Public Function fxsNaviBar(xBtn)

'//fx navigation bar state

On Error Resume Next

Call fndWindow(xWin)

xColor = vbGreen
'Cornflower - RGB(147, 204, 234)

'//Home...
If xBtn = 0 Then
xWin.HomeBtn.ForeColor = xColor
xColor = vbNullString
xWin.StartBtn.ForeColor = vbBlack
xWin.PostSetupBtn.ForeColor = vbBlack
xWin.ProfileSetupBtn.ForeColor = vbBlack
xWin.QueueBtn.ForeColor = vbBlack
xWin.BreakBtn.ForeColor = vbBlack
Exit Function

'//Start...
ElseIf xBtn = 1 Then
xWin.StartBtn.ForeColor = xColor
xColor = vbNullString
xWin.HomeBtn.ForeColor = vbBlack
xWin.PostSetupBtn.ForeColor = vbBlack
xWin.ProfileSetupBtn.ForeColor = vbBlack
xWin.QueueBtn.ForeColor = vbBlack
xWin.BreakBtn.ForeColor = vbBlack
Exit Function

'//Profile setup...
ElseIf xBtn = 2 Then
xWin.ProfileSetupBtn.ForeColor = xColor
xColor = vbNullString
xWin.HomeBtn.ForeColor = vbBlack
xWin.PostSetupBtn.ForeColor = vbBlack
xWin.StartBtn.ForeColor = vbBlack
xWin.QueueBtn.ForeColor = vbBlack
xWin.BreakBtn.ForeColor = vbBlack
Exit Function

'//Post setup...
ElseIf xBtn = 3 Then
xWin.PostSetupBtn.ForeColor = xColor
xColor = vbNullString
xWin.HomeBtn.ForeColor = vbBlack
xWin.QueueBtn.ForeColor = vbBlack
xWin.BreakBtn.ForeColor = vbBlack
xWin.ProfileSetupBtn.ForeColor = vbBlack
xWin.StartBtn.ForeColor = vbBlack
Exit Function

'//Queue...
ElseIf xBtn = 4 Then
xWin.QueueBtn.ForeColor = xColor
xColor = vbNullString
xWin.HomeBtn.ForeColor = vbBlack
xWin.PostSetupBtn.ForeColor = vbBlack
xWin.ProfileSetupBtn.ForeColor = vbBlack
xWin.StartBtn.ForeColor = vbBlack
xWin.BreakBtn.ForeColor = vbBlack
Exit Function

'//Break...
ElseIf xBtn = 5 Then
xWin.BreakBtn.ForeColor = xColor
xColor = vbNullString
xWin.HomeBtn.ForeColor = vbBlack
xWin.PostSetupBtn.ForeColor = vbBlack
xWin.ProfileSetupBtn.ForeColor = vbBlack
xWin.StartBtn.ForeColor = vbBlack
xWin.QueueBtn.ForeColor = vbBlack
Exit Function
End If

End Function
Sub undNaviBar(xBtn)

'//underline navigation bar

On Error Resume Next

Call fndWindow(xWin)

'//Home...
If xBtn = 0 Then
xWin.HomeBtn.Font.Underline = True
xWin.StartBtn.Font.Underline = False
xWin.ProfileSetupBtn.Font.Underline = False
xWin.PostSetupBtn.Font.Underline = False
xWin.QueueBtn.Font.Underline = False
xWin.BreakBtn.Font.Underline = False
End If

'//Start...
If xBtn = 1 Then
xWin.StartBtn.Font.Underline = True
xWin.HomeBtn.Font.Underline = False
xWin.ProfileSetupBtn.Font.Underline = False
xWin.PostSetupBtn.Font.Underline = False
xWin.QueueBtn.Font.Underline = False
xWin.BreakBtn.Font.Underline = False
End If

'//Profile setup...
If xBtn = 2 Then
xWin.ProfileSetupBtn.Font.Underline = True
xWin.HomeBtn.Font.Underline = False
xWin.StartBtn.Font.Underline = False
xWin.PostSetupBtn.Font.Underline = False
xWin.QueueBtn.Font.Underline = False
xWin.BreakBtn.Font.Underline = False
End If

'//Post setup...
If xBtn = 3 Then
xWin.PostSetupBtn.Font.Underline = True
xWin.HomeBtn.Font.Underline = False
xWin.StartBtn.Font.Underline = False
xWin.ProfileSetupBtn.Font.Underline = False
xWin.QueueBtn.Font.Underline = False
xWin.BreakBtn.Font.Underline = False
End If

'//Queue setup...
If xBtn = 4 Then
xWin.QueueBtn.Font.Underline = True
xWin.HomeBtn.Font.Underline = False
xWin.StartBtn.Font.Underline = False
xWin.ProfileSetupBtn.Font.Underline = False
xWin.PostSetupBtn.Font.Underline = False
xWin.BreakBtn.Font.Underline = False
End If

'//Break...
If xBtn = 5 Then
xWin.BreakBtn.Font.Underline = True
xWin.HomeBtn.Font.Underline = False
xWin.StartBtn.Font.Underline = False
xWin.ProfileSetupBtn.Font.Underline = False
xWin.PostSetupBtn.Font.Underline = False
xWin.QueueBtn.Font.Underline = False
End If

End Sub
Sub dfsFreeze()

'//default freeze state

        ETWEETXLHOME.FreezeBtn.BackColor = &H8000000B
        ETWEETXLHOME.FreezeBtn.Caption = vbNullString
        ETWEETXLHOME.FreezeBtn.SpecialEffect = fmSpecialEffectEtched
        ETWEETXLPOST.FreezeBtn.BackColor = &H8000000B
        ETWEETXLPOST.FreezeBtn.Caption = vbNullString
        ETWEETXLPOST.FreezeBtn.SpecialEffect = fmSpecialEffectEtched
        ETWEETXLQUEUE.FreezeBtn.BackColor = &H8000000B
        ETWEETXLQUEUE.FreezeBtn.Caption = vbNullString
        ETWEETXLQUEUE.FreezeBtn.SpecialEffect = fmSpecialEffectEtched
        ETWEETXLSETUP.FreezeBtn.BackColor = &H8000000B
        ETWEETXLSETUP.FreezeBtn.Caption = vbNullString
        ETWEETXLSETUP.FreezeBtn.SpecialEffect = fmSpecialEffectEtched
        
End Sub
Sub fxsFreeze()

'//fx freeze state

        ETWEETXLHOME.FreezeBtn.BackColor = vbRed
        ETWEETXLHOME.FreezeBtn.Caption = "| |"
        ETWEETXLHOME.FreezeBtn.SpecialEffect = fmSpecialEffectSunken
        ETWEETXLPOST.FreezeBtn.BackColor = vbRed
        ETWEETXLPOST.FreezeBtn.Caption = "| |"
        ETWEETXLPOST.FreezeBtn.SpecialEffect = fmSpecialEffectSunken
        ETWEETXLQUEUE.FreezeBtn.BackColor = vbRed
        ETWEETXLQUEUE.FreezeBtn.Caption = "| |"
        ETWEETXLQUEUE.FreezeBtn.SpecialEffect = fmSpecialEffectSunken
        ETWEETXLSETUP.FreezeBtn.BackColor = vbRed
        ETWEETXLSETUP.FreezeBtn.Caption = "| |"
        ETWEETXLSETUP.FreezeBtn.SpecialEffect = fmSpecialEffectSunken
        
        
End Sub
Sub fxsUnfreeze()

'//fx unfreeze state

        ETWEETXLHOME.FreezeBtn.BackColor = vbGreen
        ETWEETXLHOME.FreezeBtn.Caption = vbNullString
        ETWEETXLHOME.FreezeBtn.SpecialEffect = fmSpecialEffectEtched
        ETWEETXLPOST.FreezeBtn.BackColor = vbGreen
        ETWEETXLPOST.FreezeBtn.Caption = vbNullString
        ETWEETXLPOST.FreezeBtn.SpecialEffect = fmSpecialEffectEtched
        ETWEETXLQUEUE.FreezeBtn.BackColor = vbGreen
        ETWEETXLQUEUE.FreezeBtn.Caption = vbNullString
        ETWEETXLQUEUE.FreezeBtn.SpecialEffect = fmSpecialEffectEtched
        ETWEETXLSETUP.FreezeBtn.BackColor = vbGreen
        ETWEETXLSETUP.FreezeBtn.Caption = vbNullString
        ETWEETXLSETUP.FreezeBtn.SpecialEffect = fmSpecialEffectEtched
        
End Sub
