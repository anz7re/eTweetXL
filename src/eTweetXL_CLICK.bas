Attribute VB_Name = "eTweetXL_CLICK"
'/###########################\
'//Application Click Features\\
'///#########################\\\

Public Sub AddThreadBtn_Clk()

Dim X, xVal As String

Retry:
Call getWindow(xWin)

On Error GoTo SetForm

lastR = Cells(Rows.Count, "Z").End(xlUp).Row

If xWin.ThreadCt.Caption <> "" Then
If CInt(xWin.ThreadCt.Caption) > 0 Then
lastR = xWin.ThreadCt.Caption
GoTo NextStep
End If
    End If

NextStep:
'//print new post for thread to range
Range("PostThread").Offset(lastR, 0).Value = xWin.PostBox.Value

'//print new media for thread to range
If xWin.MedLinkBox.Value <> vbNullString Then _
Range("MedThread").Offset(lastR, 0).Value = """" & xWin.MedLinkBox.Value & """"

'//relay last thread
If Range("PostThread").Offset(lastR + 1, 0).Value <> vbNullString Then xWin.PostBox.Value = Range("PostThread").Offset(lastR + 1, 0).Value _
Else: xWin.PostBox.Value = vbNullString
If Range("MedThread").Offset(lastR + 1, 0).Value <> vbNullString Then xWin.MedLinkBox.Value = Range("MedThread").Offset(lastR + 1, 0).Value _
Else: xWin.MedLinkBox.Value = vbNullString

'//set total thread count
X = xWin.ThreadCt.Caption
If X = vbNullString Then X = 0 Else X = xWin.ThreadCt.Caption
If CInt(X) <= lastR Then xWin.ThreadCt.Caption = lastR + 1

'//reset media counters
Range("GifCntr").Value = 0
Range("VidCntr").Value = 0

'//reset media scroll
For X = 0 To 3
Range("MediaScroll").Offset(X, 0).Value = vbNullString
Next

Exit Sub

SetForm:
Err.Clear
If Range("xlasWinForm").Value2 <> 13 Then Range("xlasWinForm").Value2 = 13 Else Range("xlasWinForm").Value2 = 14
GoTo Retry

End Sub
Public Sub RmvThreadBtn_Clk()

Call getWindow(xWin)

lastR = Cells(Rows.Count, "Y").End(xlUp).Row

If xWin.ThreadCt.Caption <> "" Then
If CInt(xWin.ThreadCt.Caption) > 0 Then
lastR = xWin.ThreadCt.Caption
GoTo NextStep
End If
    End If
    
NextStep:
'//remove post from thread loc
Range("PostThread").Offset(lastR, 0).Value = vbNullString

'//remove media from thread loc
Range("MedThread").Offset(lastR, 0).Value = vbNullString

'//decrement thread count
xWin.ThreadCt.Caption = lastR - 1

'//show previous post
xWin.PostBox.Value = Range("PostThread").Offset(lastR - 1, 0).Value

End Sub
Public Sub RmvAllThreadBtn_Clk()

Retry:
Call getWindow(xWin)

On Error GoTo SetForm

lastR = Cells(Rows.Count, "Y").End(xlUp).Row
If lastR < 1 Then lastR = Cells(Rows.Count, "Z").End(xlUp).Row

For X = 0 To lastR
Range("PostThread").Offset(X, 0).Value2 = vbNullString
Range("MedThread").Offset(X, 0).Value2 = vbNullString
Next

xWin.ThreadCt.Caption = vbNullString

If Range("xlasSilent").Value2 <> 1 Then _
xWin.xlFlowStrip.Value = "All threads removed..."
        
Exit Sub

SetForm:
Err.Clear
If Range("xlasWinForm").Value2 <> 13 Then Range("xlasWinForm").Value2 = 13 Else Range("xlasWinForm").Value2 = 14
GoTo Retry

End Sub
Public Sub ClrSetupBtn_Clk()

If Range("AppState").Value2 <> 1 Then

'//Cleanup
Call clnMain
Call clnLatch
Call clnLinker
Call clnRuntime
Call clnSpec
Call eTweetXL_TOOLS.delAppData
Range("ConnectTrig").Value2 = 0
Range("LinkTrig").Value = 0
Range("User").Value = vbNullString
ETWEETXLPOST.xlFlowStrip.Enabled = True
ETWEETXLPOST.SendAPI.Value = False
'ETWEETXLPOST.ActiveUser.Caption = vbNullString
ETWEETXLPOST.UserBox.Clear
ETWEETXLPOST.LinkerBox.Clear
ETWEETXLPOST.RuntimeBox.Clear
ETWEETXLPOST.PostBox.Value = vbNullString
ETWEETXLPOST.ProfileListBox.Value = vbNullString
ETWEETXLPOST.UserListBox.Value = vbNullString
ETWEETXLPOST.DraftBox.Value = vbNullString
Call eTweetXL_CLICK.RmvAllThreadBtn_Clk
ETWEETXLPOST.UserHdr.Caption = "User"
ETWEETXLPOST.DraftHdr.Caption = "Draft"
ETWEETXLPOST.RuntimeHdr.Caption = "Runtime"

    If Range("xlasSilent").Value2 <> 1 Then _
    ETWEETXLHOME.xlFlowStrip.Value = "Cleaned..."
    ETWEETXLPOST.xlFlowStrip.Value = "Cleaned..."
    ETWEETXLQUEUE.xlFlowStrip.Value = "Cleaned..."
    ETWEETXLSETUP.xlFlowStrip.Value = "Cleaned..."

    Else
    
    xMsg = 25: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
    
    End If

End Sub
Public Sub DelAllDraftsBtn_Clk()

Call getWindow(xWin)

If xWin.DraftFilterBtn.Caption <> "..." Then
xExt = ".twt": xT = " [•]": If xWin.DraftFilterBtn.Caption <> "•" Then xFil = 0: Call eTweetXL_CLICK.DraftFilterBtn_Clk(xFil)
Call eTweetXL_LOC.xTwtFile(twtFile): xLoc = twtFile
xStr = "Are you sure you wish to delete all single draft posts for '" & ActiveProfile & "'?"
        Else
            xExt = ".thr": xT = " [...]"
            If xWin.DraftFilterBtn.Caption <> "..." Then xFil = 1: Call eTweetXL_CLICK.DraftFilterBtn_Clk(xFil)
            Call eTweetXL_LOC.xThrFile(thrFile): xLoc = thrFile
            xStr = "Are you sure you wish to delete all threaded draft posts for '" & ActiveProfile & "'?"
                End If
                
                
        '//Remove all drafts from a profile
        If ActiveProfile <> "" Then
        
        If Dir(xLoc) <> "" Then
        
        If Range("xlasSilent").Value2 = 1 Then msg = vbYes: GoTo SilentRun
        
        msg = MsgBox(xStr, vbYesNo, eTweetXL_INFO.AppName)
        
SilentRun:
            If msg = vbYes Then
            
            Dim oFSO, oFile, oFldr As Object
            Set oFSO = CreateObject("Scripting.FileSystemObject")
            Set oFldr = oFSO.GetFolder(xLoc)
            
            For Each oFile In oFldr.Files
            Kill (oFile)
            Next
            
            If Range("DraftFilter").Value = 1 Then xType = 0 Else xType = 1
            xType = 0: Call eTweetXL_GET.getPostData(xType)
            
            End If
                End If
                    End If
                    
End Sub
Public Sub HideBtn_Clk()

Dim M As Byte

Call getWindow(xWin)
xWin.Hide
M = MsgBox("Display eTweetXL?", vbOKOnly, eTweetXL_INFO.AppName)
    If M = vbOK Then
        Call xWin.Show
            End If

Set xWin = Nothing

End Sub
Public Sub RuntimeBox_Clk()

'//For editing the time in Runtime boxes

Dim I As Integer: Dim xPos As Integer: Dim X As Integer: Dim LLCntr As Integer
Dim oRuntimeBox As Object

lastR = Cells(Rows.Count, "R").End(xlUp).Row

'//Check for runtime change...(Avoid double run)
If Range("RtChange").Value = 1 Then
Range("RtChange").Value = 0
Exit Sub
End If

'//Check xlFlowStrip Window...
If Range("xlasWinForm").Value2 = 13 Then
Set oRuntimeBox = ETWEETXLPOST.RuntimeBox
GoTo XLPOST
End If

If Range("xlasWinForm").Value2 = 14 Then
Set oRuntimeBox = ETWEETXLQUEUE.RuntimeBox
GoTo XLQUEUE
End If

'//Change runtime in Queue
XLQUEUE:

On Error GoTo ErrMsg

I = oRuntimeBox.ListIndex
RtHldr = oRuntimeBox.Value
RtHldrArr = Split(RtHldr, ") ")
RtHldr = RtHldrArr(1)

'//find position
For X = 0 To I
If oRuntimeBox.Selected(X) = True Then
xPos = X
End If
Next

If RtHldr <> vbNullString Then

NewRt = InputBox("Enter a new runtime:", eTweetXL_INFO.AppName, RtHldr)

    If NewRt <> vbNullString Then
    
    '//Check for time format...
    If InStr(1, NewRt, ":") = False Then GoTo ErrMsg
    
        '//Convert to time...
        ThisRt = Format$(ThisRt, "hh:mm:ss")

        '//Record...
        Range("RtChange").Value = 1
        LLCntr = Range("LinkerCount").Value2: If LLCntr = 0 Then LLCntr = 1
        Range("Runtime").Offset(xPos + LLCntr, 0).Value = NewRt
        NewRt = "(" & xPos + 1 & ") " & NewRt
        oRuntimeBox.List(xPos) = NewRt
            End If
                End If


'//Export to file...
Call eTweetXL_LOC.xMTRuntime(mtRuntime)

Open mtRuntime For Output As #1
For X = 1 To lastR
If Range("Runtime").Offset(X, 0).Value <> "" Then
Print #1, Range("Runtime").Offset(X, 0).Value
End If
Next
Close #1

Set oRuntimeBox = Nothing

Exit Sub

'//Change runtime in post...
XLPOST:

On Error GoTo ErrMsg

I = oRuntimeBox.ListIndex
RtHldr = oRuntimeBox.Value
RtHldrArr = Split(RtHldr, ") ")
RtHldr = RtHldrArr(1)

For X = 0 To I
If oRuntimeBox.Selected(X) = True Then
xPos = X
End If
Next

If RtHldr <> vbNullString Then

On Error GoTo ErrMsg

'//Check for time format...
NewRt = InputBox("Enter a new runtime:", eTweetXL_INFO.AppName, RtHldr)

If NewRt = vbNullString Then Exit Sub

    If InStr(1, NewRt, ":") = False Then GoTo ErrMsg

    If NewRt <> "" Then
    
    '//Convert to time...
    NewRt = Format$(NewRt, "hh:mm:ss")
    NewRt = "(" & xPos + 1 & ") " & NewRt
    '//Record...
    Range("RtChange").Value = 1
    oRuntimeBox.List(xPos) = NewRt: Exit Sub
        End If
            End If
            
Exit Sub

'//Error
ErrMsg:
xMsg = 8: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)

Set oRuntimeBox = Nothing

End Sub
Public Sub BreakBtn_Clk()

'//For breaking/stopping the application incase of an issue during a run, or to refresh

TK = CreateObject("WScript.Shell").Run("taskkill /f /im " & "powershell.exe", 0, True)


'//Cleanup
Call clnMain
Call clnLatch
Call clnLinker
Call clnRuntime
Call clnSpec
Call delAppData
Range("AppState").Value2 = 0
Range("ConnectTrig").Value2 = 0
Range("LinkTrig").Value = 0
Range("User").Value = vbNullString
ETWEETXLHOME.xlFlowStrip.Enabled = True
ETWEETXLPOST.xlFlowStrip.Enabled = True
ETWEETXLQUEUE.xlFlowStrip.Enabled = True
ETWEETXLSETUP.xlFlowStrip.Enabled = True
ETWEETXLHOME.ProgRatio = vbNullString
'ETWEETXLHOME.ActiveUser.Caption = vbNullString
ETWEETXLHOME.ProgBar.Width = 0
ETWEETXLPOST.SendAPI.Value = False
'ETWEETXLPOST.ActiveUser.Caption = vbNullString
ETWEETXLPOST.UserBox.Clear
ETWEETXLPOST.LinkerBox.Clear
ETWEETXLPOST.RuntimeBox.Clear
ETWEETXLPOST.ProfileListBox.Value = vbNullString
ETWEETXLPOST.UserListBox.Value = vbNullString
ETWEETXLPOST.DraftBox.Value = vbNullString
ETWEETXLPOST.UserHdr.Caption = "User"
ETWEETXLPOST.DraftHdr.Caption = "Draft"
ETWEETXLPOST.RuntimeHdr.Caption = "Runtime"
ETWEETXLQUEUE.UserHdr.Caption = "User"
ETWEETXLQUEUE.QueueHdr.Caption = "Queued"
ETWEETXLQUEUE.RuntimeHdr.Caption = "Runtime"
'ETWEETXLQUEUE.ActiveUser.Caption = vbNullString
ETWEETXLQUEUE.QueueBox.Clear
ETWEETXLQUEUE.RuntimeBox.Clear
ETWEETXLQUEUE.UserBox.Clear
'ETWEETXLSETUP.ActiveUser.Caption = vbNullString
Set TK = Nothing

ETWEETXLHOME.AppStatus.Caption = "OFF"
ETWEETXLHOME.AppStatus.ForeColor = vbRed
ETWEETXLHOME.AppStatus.BackColor = -2147483633

Call enlFlowStrip
Call dfsFreeze

ETWEETXLHOME.xlFlowStrip.Value = "Break complete..."
ETWEETXLSETUP.xlFlowStrip.Value = "Break complete..."
ETWEETXLPOST.xlFlowStrip.Value = "Break complete..."
ETWEETXLQUEUE.xlFlowStrip.Value = "Break complete..."

xMsg = 10: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)

If Range("xlasWinForm").Value2 = 11 Then ETWEETXLHOME.Hide
If Range("xlasWinForm").Value2 = 11 Then ETWEETXLHOME.Show

End Sub
Sub StartBtn_Clk()

On Error GoTo ErrMsg

'//check for pause or disable
If Range("AppState").Value2 = 2 Then xMsg = 26: GoTo EndMacro

'//For starting the application automation
Call xlAppScript_xbas.disableWbUpdates

'//Check for set user...
If Range("ActiveUser").Value = "" Then
xMsg = 9: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
Exit Sub
End If

'//Check MainLink...
If Range("MainLink").Offset(1, 0).Value = "" Then xMsg = 21: GoTo EndMacro

ETWEETXLHOME.xlFlowStrip.Value = "Checking for user information..."
ETWEETXLSETUP.xlFlowStrip.Value = "Checking for user information..."
ETWEETXLPOST.xlFlowStrip.Value = "Checking for user information..."
ETWEETXLQUEUE.xlFlowStrip.Value = "Checking for user information..."

Call disFlowStrip

Call eTweetXL_LOC.xMTapi(mtApi)
Call eTweetXL_LOC.xMTBlank(mtBlank)
Call eTweetXL_LOC.xMTPass(mtPass)
Call eTweetXL_LOC.xMTTwt(mtTwt)
Call eTweetXL_LOC.xMTCheck(mtCheck)
Call eTweetXL_LOC.xMTUser(mtUser)
Call eTweetXL_LOC.xMTOffset(mtOffset)
Call eTweetXL_LOC.xMTOffsetCopy(mtOffsetCopy)
Call eTweetXL_LOC.xMTRuntime(mtRuntime)
Call eTweetXL_LOC.xMTRuntimeCntr(mtRuntimeCntr)
Call eTweetXL_LOC.xMTDynOff(mtDynOff)
Call eTweetXL_LOC.xMTMed(mtMed)
Call eTweetXL_LOC.xMTPost(mtPost)
Call eTweetXL_LOC.xMTProf(mtProf)
Call eTweetXL_LOC.xMTRetryCntr(mtRetryCntr)
Call eTweetXL_LOC.xApp_StartLink(appStartLink)
    
'//Clear...
If Range("AppStatus") = 1 Then GoTo SkipClear '//Skip if linkline active
Range("LinkerCount").Value2 = 0
ETWEETXLHOME.ProgBar.Width = 5 '//Refresh progress bar

SkipClear:

'//Export user check...
Open mtCheck For Output As #2
Print #2, Range("ActiveUser").Value
Close #2

'//Export offset...
Open mtOffset For Output As #4
Print #4, Range("ActiveOffset").Value
Close #4

'//Export offset copy...
Open mtOffsetCopy For Output As #4
Print #4, Range("ActiveOffset").Value
Close #4

'//Export retry file...
If Dir(mtRetryCntr) = "" Then
Open mtRetryCntr For Output As #5
Print #5, ""
Close #5
End If

'//Export blank file...
If Dir(mtBlank) = "" Then
Open mtBlank For Output As #5
Print #5, ""
Close #5
End If

'//Export runtime...
lastR = Cells(Rows.Count, "R").End(xlUp).Row
rtTrig = 0

ETWEETXLHOME.xlFlowStrip.Value = "Calculating offset..."
ETWEETXLSETUP.xlFlowStrip.Value = "Calculating offset..."
ETWEETXLPOST.xlFlowStrip.Value = "Calculating offset..."
ETWEETXLQUEUE.xlFlowStrip.Value = "Calculating offset..."

Open mtRuntime For Output As #5
For xNum = 1 To lastR

If Range("R" & xNum).Value <> vbNullString Then

ThisRuntime = Range("R" & xNum).Value
ThisRuntime = Format$(ThisRuntime, "hh:mm:ss")

'//Add offset to runtime...
If Range("ActiveOffset").Value <> "" Then
MyOffset = Range("ActiveOffset").Value
      If InStr(1, MyOffset, "ÿþ") Then
        MyOffset = Range("Offset").Offset(1, 0).Value
            End If
      
MyOffsetSec = (MyOffset / 1000): MyOffsetMin = (MyOffsetSec / 60): MyOffsetHr = (MyOffsetMin / 60)

'//Check for hour...
If MyOffsetHr > 1 Then
'//If hour and something else...
If InStr(1, MyOffsetHr, ".") Then
MyOffsetArr = Split(MyOffsetHr, ".")
MyOffsetHr = Int(MyOffsetArr(0))
'//Find minutes...
MyOffsetMin = Int(MyOffsetMin) - ((MyOffsetHr) * 60)
'//Find seconds...
SubMyOffSec = (MyOffsetMin * 60) + ((MyOffsetHr * 60) * 60)
MyOffsetSec = MyOffsetSec - SubMyOffSec
    Else
        '//If only the hour is present...
        MyOffsetMin = ""
        MyOffsetSec = ""
    End If
        Else
            '//If no hour present...
            MyOffsetHr = ""
            '//If minutes present but no hour...
            If MyOffsetMin > 1 Then
                If InStr(1, MyOffsetMin, ".") Then
                    '//If minute and second present but no hour...
                    MyOffMinArr = Split(MyOffsetMin, ".")
                    MyOffsetSec = MyOffsetSec - (MyOffMinArr(0) * 60)
                    MyOffsetMin = Int(MyOffMinArr(0))
                        Else
                            MyOffsetSec = ""
                                End If
                    Else
                    '//If no hour or minute present (just seconds)...
                        MyOffsetMin = ""
                            End If
                                End If

RtArr = Split(ThisRuntime, ":")
ReCheckHr:
If Int(RtArr(0)) <> 0 Then RtArr(0) = (RtArr(0) + MyOffsetHr): If RtArr(0) > 23 Then RtArr(0) = RtArr(0) - 24
ReCheckMin:
If Int(RtArr(1)) <> 0 Then RtArr(1) = (RtArr(1) + MyOffsetMin)
    If RtArr(1) >= 60 Then
        RtArr(0) = RtArr(0) + 1
        RtArr(1) = RtArr(1) - 60
        If RtArr(0) > 23 Then GoTo ReCheckHr
        End If
ReCheckSec:
If Int(RtArr(2)) <> 0 Then RtArr(2) = (RtArr(2) + MyOffsetSec)
    If RtArr(2) >= 60 Then
        RtArr(1) = RtArr(1) + 1
        RtArr(2) = RtArr(2) - 60
        If RtArr(1) > 60 Then GoTo ReCheckMin
        If RtArr(2) > 60 Then GoTo ReCheckSec
        End If
             
ThisRuntime = RtArr(0) & ":" & RtArr(1) & ":" & RtArr(2)
End If

'//Record runtime...
Print #5, ThisRuntime
Range("R" & xNum).Value = ThisRuntime

'//Record total runtime...
rtTrig = rtTrig + 1
End If
Next
Close #5

'//Export runtime counter...
If Range("RtCntr").Value = 0 Then
Open mtRuntimeCntr For Output As #6
Print #6, 0
Close #6
Range("RtCntr").Value = 1
End If

'//Remove runtime files...
If rtTrig = 0 Then
If Dir(mtRuntime) <> "" Then Kill mtRuntime
If Dir(mtRuntimeCntr) <> "" Then Kill mtRuntimeCntr
End If

'//Check for dynamic offset...
If Range("DynOffsetTrig").Value = 1 Then

Open mtDynOff For Output As #7
Print #7, ""
Close #7
    Else
        If Dir(mtDynOff) <> "" Then
        Kill (mtDynOff)
            End If
        
'//Refresh offset if not dynamic
If Range("DynOffsetTrig").Value <> 1 Then Range("ActiveOffset").Value = vbNullString

                End If

ETWEETXLHOME.xlFlowStrip.Value = "Checking Linker..."
ETWEETXLSETUP.xlFlowStrip.Value = "Checking Linker..."
ETWEETXLPOST.xlFlowStrip.Value = "Checking Linker..."
ETWEETXLQUEUE.xlFlowStrip.Value = "Checking Linker..."

'//Linkline active export...
If Range("LinkTrig").Value = 1 Then

Range("AppStatus").Value = 1

Dim LinkerData(100) As String

'//Export tweet path...
LinkerArr = Split(Range("ActiveTweet").Value, ",")
    
    On Error GoTo NextStep
    xNum = 0
    Open mtTwt For Output As #8
    
    Do Until LinkerArr(xNum) = ""
    Print #8, LinkerArr(xNum)
    xNum = xNum + 1
    Loop
NextStep:
    Close #8
    
    '//Record total links...
    Range("LinkerTotal").Value2 = xNum
    
    '//
    If API_LINK = 1 Then GoTo SendWithAPI '//Avoid double run
'/######################################################################
EmptyLinker:
'//Check for empty Linker & clear...
If Range("LinkerCount").Value2 = (Range("LinkerTotal").Value2 + 1) Then
Call clnLinker2
Exit Sub
End If
'/######################################################################
    If xNum = Range("LinkerCount").Value2 Then
        Range("LinkerCount").Value2 = Range("LinkerCount").Value2 + 1
        GoTo EmptyLinker
            End If
    
    On Error GoTo ErrMsg
    
    '//Open tweet and record information...
    xNum = 1
    LLCntr = Range("LinkerCount").Value2
    If LLCntr = "" Then LLCntr = 0
    Open LinkerArr(LLCntr) For Input As #1
    Do Until EOF(1)
    Line Input #1, LinkerData(xNum)
    xNum = xNum + 1
    Loop
    Close #1
    
ETWEETXLHOME.xlFlowStrip.Value = "Setting up post..."
ETWEETXLSETUP.xlFlowStrip.Value = "Setting up post..."
ETWEETXLPOST.xlFlowStrip.Value = "Setting up post..."
ETWEETXLQUEUE.xlFlowStrip.Value = "Setting up post..."

'/######################################################################
'//Export user...
Dim xUser As String

Open mtUser For Output As #1
Print #1, Range("UserLink").Offset(LLCntr + 1, 0).Value
Close #1

'//Set user active...
Range("User").Value = Range("UserLink").Offset(LLCntr + 1, 0).Value
xUser = Range("User").Value

'//Export target (pass)...
Open mtPass For Output As #4
Print #4, Range("TargetLink").Offset(LLCntr + 1, 0).Value
Close #4

'//Set target active...
Range("ActiveTarget").Value = Range("TargetLink").Offset(LLCntr + 1, 0).Value

'//Export profile...
Open mtProf For Output As #5
Print #5, Range("MainLink").Offset(LLCntr + 1, 0).Value
Close #5

'//Set profile active...
Range("Profile").Value2 = Range("MainLink").Offset(LLCntr + 1, 0).Value

'/######################################################################

        
    '//Check for threaded post
    If InStr(1, LinkerArr(LLCntr), ".thr") Then
    xLink = LinkerArr(LLCntr): Call eTweetXL_TOOLS.sndAsThread(xLink): Exit Sub
            Else
                Range("ThreadStatus").Value2 = 0
                Range("LinkerCount").Value2 = LLCntr + 1
                    End If
    
    xNum = 1
    Do Until InStr(1, LinkerData(xNum), "*-")
    xMyPost = xMyPost + LinkerData(xNum)
    xNum = xNum + 1
    Loop
    
    xMyMed = Replace(LinkerData(xNum + 1), "*-", vbNullString)
    
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

    
    Open mtMed For Output As #7
    Print #7, xMyMed
    Close #7
    
    Open mtPost For Output As #8
    Print #8, xMyPost
    Close #8
        
        Else
        
        '//Single post...
        Open mtTwt For Output As #4
        Print #4, Range("ActiveTweet").Value
        Close #4
       
GoTo NextStep2
        '//Open tweet file & get post...
        xNum = 1
        Open Range("ActiveTweet").Value For Input As #1
'//Grab post...
xMyPostHldr = 0
xMyPost = ""
xMyMedia = ""
endCntr = 0
Do Until xMyPostHldr = "*-;"
'//End if unable to find text after 20 lines...
If endCntr > 20 Then GoTo NextStep2
Line Input #1, xMyPostHldr
If xMyPostHldr <> "" Then
    If xMyPostHldr <> "*-;" Then
    xMyPost = xMyPost & xMyPostHldr & "*-;"
    End If
        End If
        endCntr = endCntr + 1
            Loop

NextStep2:
On Error Resume Next
'//Replace markers w/ enter...
xMyPost = Replace(xMyPost, "*-;", vbCrLf)

'//Grab Media...
Line Input #1, xMyMed
Close #1

        Open mtMed For Output As #7
        Print #7, xMyMed
        Close #7
        
        Open mtPost For Output As #8
        Print #8, xMyPost
        Close #8
        
            End If
              
         '//Rrefresh progress bar...
        Call updProgBar
        Call fxsUnfreeze
        
'        ETWEETXLHOME.ActiveUser.Caption = xUser
'        ETWEETXLSETUP.ActiveUser.Caption = xUser
'        ETWEETXLPOST.ActiveUser.Caption = xUser
'        ETWEETXLQUEUE.ActiveUser.Caption = xUser
        
        ETWEETXLHOME.xlFlowStrip.Value = "Sleeping..."
        ETWEETXLSETUP.xlFlowStrip.Value = "Sleeping..."
        ETWEETXLPOST.xlFlowStrip.Value = "Sleeping..."
        ETWEETXLQUEUE.xlFlowStrip.Value = "Sleeping..."
        
'//Turn Linker active...
ETWEETXLHOME.AppStatus.Caption = "ON"
ETWEETXLHOME.AppStatus.ForeColor = vbGreen
ETWEETXLHOME.AppStatus.BackColor = vbWhite
If Range("AppState").Value2 <> 1 Then Range("AppState").Value2 = 1

SendWithAPI:
'//Send with api method...
If Range("apiLink").Offset(LLCntr + 1, 0).Value = "(*api)" Then
    API_LINK = 1
    Call eTweetXL_TOOLS.bldAPIScript
    Open mtApi For Output As #7
    Print #7, ""
    Close #7
    Shell (appStartLink), vbMinimizedNoFocus
        Else
        '//Send with default method...
           If Dir(mtApi) <> "" Then Kill (mtApi)
           Shell (appStartLink), vbMinimizedNoFocus
                End If
                
'//Cleanup Queue
ETWEETXLQUEUE.QueueBox.Clear
ETWEETXLQUEUE.RuntimeBox.Clear
ETWEETXLQUEUE.UserBox.Clear

'//Update application state
Call eTweetXL_TOOLS.updAppState
Call eTweetXL_GET.getQueueData
Exit Sub

EndMacro:
Call clnOpenFiles
Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
Exit Sub

ErrMsg:
Call clnOpenFiles
xMsg = 27: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
End Sub
Public Sub DraftHdr_Clk()

'//For removing drafts from the Linker

        '//Reset connect trigger...
        If Range("AppState").Value2 <> 1 Then
        Range("ConnectTrig").Value2 = 0
        End If
        
        ETWEETXLPOST.DraftHdr.Caption = "Draft"
        
        '//Remove all drafts from Linker...
        ETWEETXLPOST.LinkerBox.Clear
        
        lastR = Cells(Rows.Count, "P").End(xlUp).Row
        Range("P2:P" & lastR).Value2 = ""
        
        lastR = Cells(Rows.Count, "L").End(xlUp).Row
        Range("L2:L" & lastR).Value2 = ""
        
        If Range("xlasSilent").Value2 <> 1 Then _
        ETWEETXLPOST.xlFlowStrip.Value = "All linked drafts cleared..."
        
        
End Sub
Public Sub LinkerHdr_Clk()

'//For removing all users, drafts, & time from the Linker

On Error GoTo EndMacro

        '//Reset connect trigger...
        If Range("AppState").Value2 <> 1 Then
        Range("ConnectTrig").Value2 = 0
        End If
        
        Call clnLinker
        Call clnSpec
        
        ETWEETXLPOST.UserHdr.Caption = "User"
        ETWEETXLPOST.DraftHdr.Caption = "Draft"
        ETWEETXLPOST.RuntimeHdr.Caption = "Runtime"
        
        '//Remove all users from Linker...
        ETWEETXLPOST.UserBox.Clear
        '//Remove all drafts from Linker...
        ETWEETXLPOST.LinkerBox.Clear
        '//Remove all time from Linker...
        ETWEETXLPOST.RuntimeBox.Clear
        
        If Range("xlasSilent").Value2 <> 1 Then _
        ETWEETXLPOST.xlFlowStrip.Value = "Linker cleared..."
        
        Exit Sub
        
EndMacro:
        If Range("xlasSilent").Value2 <> 1 Then _
        ETWEETXLPOST.xlFlowStrip.Value = "An unknown error occurred while clearing the Linker..."
        
End Sub
Public Sub TimerHdr_Clk()

'//For refresing the Time box

ETWEETXLPOST.TimeBox.Value = "0"
ETWEETXLPOST.TimeBox.Value = ""

If Range("xlasSilent").Value2 <> 1 Then _
ETWEETXLPOST.xlFlowStrip.Value = "Time refreshed..."
        
End Sub
Public Sub UserHdr_Clk()

'//For removing all users from the Linker

        '//Reset connect trigger...
        If Range("AppState").Value2 <> 1 Then
        Range("ConnectTrig").Value2 = 0
        End If
        
        ETWEETXLPOST.UserHdr.Caption = "User"

        '//Remove all users from userbox...
        ETWEETXLPOST.UserBox.Clear
        
        Call clnSpec
        
        If Range("xlasSilent").Value2 <> 1 Then _
        ETWEETXLPOST.xlFlowStrip.Value = "All linked users cleared..."
        
End Sub
Public Sub RuntimeHdr_Clk()

        '//Reset connect trigger
        If Range("AppState").Value2 <> 1 Then
        Range("ConnectTrig").Value2 = 0
        End If
        
        ETWEETXLPOST.RuntimeHdr.Caption = "Runtime"
        
        '//REMOVE ALL TIME FROM RUNTIME$
        ETWEETXLPOST.RuntimeBox.Clear
        
        If Range("xlasSilent").Value2 <> 1 Then _
        ETWEETXLPOST.xlFlowStrip.Value = "All linked times cleared..."
        
End Sub
Public Sub PostHdr_Clk()

        '//default no parameter
        Call getWindow(xWin)
        '//Clear post box
        xWin.PostBox.Value = vbNullString
        
        If Range("xlasSilent").Value2 <> 1 Then _
        xWin.xlFlowStrip.Value = "Post cleared..."
        
        End Sub
Public Function AddPostMedBtn_Clk(ByVal xMed As String) As String

'//For adding media to a post

Dim oMedLinkBox As Object

'//Check xlFlowStrip Window...
If Range("xlasWinForm").Value2 = 13 Then Set oMedLinkBox = ETWEETXLPOST.MedLinkBox
If Range("xlasWinForm").Value2 = 14 Then Set oMedLinkBox = ETWEETXLQUEUE.MedLinkBox

lastR = Cells(Rows.Count, "I").End(xlUp).Row

'//Check if Gif/Video already added...
If Range("GifCntr").Value = 1 Then GoTo ErrMsg1
If Range("VidCntr").Value = 1 Then GoTo ErrMsg1

'//For the first Media...
If Range("MediaScroll").Value = vbNullString Then
lastR = 0
End If

'//Check for 4 Media...
If lastR > 3 Then GoTo ErrMsg2

'//Check if Media already found...
If xMed <> "" Then GoTo SkipOpen

'//Select an Media...
xMed = Application.GetOpenFilename()

SkipOpen:

If xMed = "" Then GoTo EndMacro
If xMed = "False" Then GoTo EndMacro

'//gif
If InStr(1, xMed, ".gif") Then
    If oMedLinkBox.Value <> "" Then GoTo ErrMsg1
If FileLen(xMed) > 15728639 Then
xMsg = 13: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
GoTo EndMacro
    End If
    Range("GifCntr").Value = 1
        End If

'//mp4
If InStr(1, xMed, ".mp4") Then
    If oMedLinkBox.Value <> "" Then GoTo ErrMsg1
If FileLen(xMed) > 51200000000# Then
xMsg = 12: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
GoTo EndMacro
    End If
    Range("VidCntr").Value = 1
        End If
        
'//mov
If InStr(1, xMed, ".mov") Then
    If oMedLinkBox.Value <> "" Then GoTo ErrMsg1
If FileLen(xMed) > 51200000000# Then
xMsg = 12: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
GoTo EndMacro
    End If
    Range("VidCntr").Value = 1
        End If
        
'//flv
If InStr(1, xMed, ".flv") Then
    If oMedLinkBox.Value <> "" Then GoTo ErrMsg1
If FileLen(xMed) > 51200000000# Then
xMsg = 12: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
GoTo EndMacro
    End If
    Range("VidCntr").Value = 1
        End If
                              
                              
Range("MediaScroll").Offset(lastR, 0) = xMed
xMed = "'" & xMed & "'"
xMed = Replace(xMed, "'", """")
       
If oMedLinkBox.Value = "" Then
oMedLinkBox.Value = xMed
    Else
        oMedLinkBox.Value = oMedLinkBox.Value & " " & xMed
            End If
            
Call eTweetXL_GET.getSelMedia
EndMacro:

Set oMedLinkBox = Nothing

Exit Function

'//Debug...
ErrMsg1:
oMedLinkBox.BorderStyle = fmBorderStyleSingle
oMedLinkBox.BorderColor = vbRed
xMsg = 14: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)

Set oMedLinkBox = Nothing

Exit Function

ErrMsg2:
oMedLinkBox.BorderStyle = fmBorderStyleSingle
oMedLinkBox.BorderColor = vbRed
xMsg = 15: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)

Set oMedLinkBox = Nothing

End Function
Public Function AddRuntimeBtn_Clk(ByVal xPos As Long)

'//For adding a time to the Linker

If ETWEETXLPOST.TimeBox <> "" Then

Dim xTot As Long: Dim xPosH As Long
Dim X As Integer: Dim xNum As Integer
Dim xPosArr() As String
Dim xTempArr(5000) As String
Dim xTime As String
Dim EDITMODE As Byte

    
If Len(ETWEETXLPOST.TimeBox.Value) = 7 Or 8 Then

'//Check for invalid characters...
Call eTweetXL_TOOLS.fndChar(xChar)
If xChar = "(*Err)" Then Exit Function

On Error Resume Next

'//check for Linker edit state
    If Range("TimeTrig").Value2 <> 0 Then
        If Range("LinkerIndex").Value2 <> vbNullString Then
        xPos = Range("LinkerIndex").Value2
        '//swap
        If Range("TimeTrig").Value2 = 1 Then
        EDITMODE = 1
        '//above
        ElseIf Range("TimeTrig").Value2 = 2 Then
        EDITMODE = 2
        '//below
        ElseIf Range("TimeTrig").Value2 = 3 Then
        EDITMODE = 3
        End If
            End If
                End If
    
If xPos > 0 Or EDITMODE <> 0 Then

xPosH = xPos

    If ETWEETXLPOST.RuntimeBox.ListCount > 0 Then

    X = 0

    If InStr(1, xPos, ":") Then
    xPosArr = Split(xPos, ":")
    For X = xPosArr(0) To xPosArr(1)
    xTime = "(" & X & ") " & ETWEETXLPOST.TimeBox.Value
    If Left(xTime, 1) = " " Then xTime = Right(xTime, Len(xTime) - 1) '//remove leading space
    ETWEETXLPOST.RuntimeBox.List((X)) = (xTime)
    Next
    Exit Function
    End If
    
    If InStr(1, xPos, ",") Then
    xPosArr = Split(xPos, ",")
    xTot = UBound(xPosArr)
    Do Until X = xTot
    xTime = "(" & X & ") " & ETWEETXLPOST.TimeBox.Value
    If Left(xTime, 1) = " " Then xTime = Right(xTime, Len(xTime) - 1) '//remove leading space
    ETWEETXLPOST.RuntimeBox.List(xPosArr(X)) = (xTime)
    X = X + 1
    Loop
    Exit Function
    End If
    
    
'//swap
    If EDITMODE = 1 Then
    xPosH = xPos + 1
    
'//above
    ElseIf EDITMODE = 2 Then
    
        xPosH = xPos
        
        X = 0
        For xNum = 1 To ETWEETXLPOST.RuntimeBox.ListCount + 1
        
        If X = xPosH Then xTempArr(X) = "(" & X + 1 & ") " & ETWEETXLPOST.TimeBox.Value: _
        xTempArr(X + 1) = ETWEETXLPOST.RuntimeBox.List((X)): X = X + 2: xNum = xNum + 1
        
        xTempArr(X) = ETWEETXLPOST.RuntimeBox.List((xNum - 1))
        X = X + 1
        Next
        
        For X = 0 To ETWEETXLPOST.RuntimeBox.ListCount
        xTime = xTempArr(X)
        xPosArr = Split(xTime, ") ")
        xTime = Trim(xPosArr(1))
        xTime = "(" & X + 1 & ") " & xTime
        
        If X < ETWEETXLPOST.RuntimeBox.ListCount Then
        ETWEETXLPOST.RuntimeBox.List(X) = xTime
            Else
                ETWEETXLPOST.RuntimeBox.AddItem xTime
                    End If
        
        Next
    
        ETWEETXLPOST.RuntimeHdr.Caption = "Runtime" & " (" & ETWEETXLPOST.RuntimeBox.ListCount & ")"
    
        Exit Function
    
'//below
    ElseIf EDITMODE = 3 Then
    
        xPosH = xPos + 1
        
        X = 0
        For xNum = 1 To ETWEETXLPOST.RuntimeBox.ListCount + 1
        
        If X = xPosH Then xTempArr(X) = "(" & X + 1 & ") " & ETWEETXLPOST.TimeBox.Value: _
        xTempArr(X + 1) = ETWEETXLPOST.RuntimeBox.List((X)): X = X + 2: xNum = xNum + 1
        
        xTempArr(X) = ETWEETXLPOST.RuntimeBox.List((xNum - 1))
        X = X + 1
        Next
        
        For X = 0 To ETWEETXLPOST.RuntimeBox.ListCount
        xTime = xTempArr(X)
        xPosArr = Split(xTime, ") ")
        xTime = Trim(xPosArr(1))
        xTime = "(" & X + 1 & ") " & xTime
        
        If X < ETWEETXLPOST.RuntimeBox.ListCount Then
        ETWEETXLPOST.RuntimeBox.List(X) = xTime
            Else
                ETWEETXLPOST.RuntimeBox.AddItem xTime
                    End If
        
        Next
    
        ETWEETXLPOST.RuntimeHdr.Caption = "Runtime" & " (" & ETWEETXLPOST.RuntimeBox.ListCount & ")"
    
        Exit Function
    
        End If
            
            xTime = "(" & xPosH & ") " & ETWEETXLPOST.TimeBox.Value
            If Left(xTime, 1) = " " Then xTime = Right(xTime, Len(xTime) - 1) '//remove leading space
            ETWEETXLPOST.RuntimeBox.List(xPos) = (xTime)
            
            Exit Function
    
                    End If
                        
                        Else
                         
            ETWEETXLPOST.RuntimeHdr.Caption = "Runtime" & " (" & ETWEETXLPOST.RuntimeBox.ListCount + 1 & ")"
            xTime = ETWEETXLPOST.TimeBox.Value
            xTime = ETWEETXLPOST.RuntimeHdr.Caption & " " & xTime
            xTime = Replace(xTime, "Runtime", vbNullString, , , vbTextCompare)
            If Left(xTime, 1) = " " Then xTime = Right(xTime, Len(xTime) - 1) '//remove leading space
            
            ETWEETXLPOST.RuntimeBox.AddItem (xTime)
            
            If Range("xlasSilent").Value2 <> 1 Then
            ETWEETXLPOST.xlFlowStrip.Value = xTime & " linked..."
            End If
                
                End If
                                
                    End If
                        
                        End If
        
End Function
Public Function DynOffset_Clk(ByVal xPos As Byte)

If xPos <> vbNullString Then GoTo SkipHere

If ETWEETXLPOST.DynOffset.Value = True Then
   Range("DynOffsetTrig").Value = 1
    ETWEETXLPOST.DynOffset.Value = True
    If ETWEETXLPOST.OffsetBox.Value = "00:00:00" Then ETWEETXLPOST.OffsetBox.Value = "00:00:01"
        Else
        Range("DynOffsetTrig").Value = 0
        ETWEETXLPOST.DynOffset.Value = False
        If ETWEETXLPOST.OffsetBox.Value = "00:00:01" Then ETWEETXLPOST.OffsetBox.Value = "00:00:00"
            End If
            
                Exit Function
                
SkipHere:
   If xPos = 1 Then
     Range("DynOffsetTrig").Value = 1: ETWEETXLPOST.DynOffset.Value = True
     If ETWEETXLPOST.OffsetBox.Value = "00:00:00" Then ETWEETXLPOST.OffsetBox.Value = "00:00:01"
        ElseIf xPos = 0 Then
        Range("DynOffsetTrig").Value = 0: ETWEETXLPOST.DynOffset.Value = False
        If ETWEETXLPOST.OffsetBox.Value = "00:00:01" Then ETWEETXLPOST.OffsetBox.Value = "00:00:00"
        End If
            
            
End Function
Public Sub xlFlowStripBar_Clk()

'//For extending and restricting the xlFlowStrip bar

If Range("xlasWinForm").Value2 = 11 Then GoTo FlBarHome
If Range("xlasWinForm").Value2 = 12 Then GoTo FlBarSetup
If Range("xlasWinForm").Value2 = 13 Then GoTo FlBarPost
If Range("xlasWinForm").Value2 = 14 Then GoTo FlBarQueue

FlBarHome:
If ETWEETXLHOME.Height = "510" Then
ETWEETXLHOME.Height = "637"
Exit Sub
End If

If ETWEETXLHOME.Height = "637" Then
ETWEETXLHOME.Height = "510"
Exit Sub
End If

FlBarSetup:
If ETWEETXLSETUP.Height = "590.25" Then
ETWEETXLSETUP.Height = "740"
Exit Sub
End If

If ETWEETXLSETUP.Height = "740" Then
ETWEETXLSETUP.Height = "590.25"
Exit Sub
End If

FlBarPost:
If ETWEETXLPOST.Height = "590.25" Then
ETWEETXLPOST.Height = "740"
Exit Sub
End If

If ETWEETXLPOST.Height = "740" Then
ETWEETXLPOST.Height = "590.25"
Exit Sub
End If

FlBarQueue:
If ETWEETXLQUEUE.Height = "540" Then
ETWEETXLQUEUE.Height = "690"
Exit Sub
End If

If ETWEETXLQUEUE.Height = "690" Then
ETWEETXLQUEUE.Height = "540"
Exit Sub
End If

End Sub
Public Function AddDraftBtn_Clk() As Byte

'//For adding drafts to a profile

If ETWEETXLPOST.DraftBox.Value <> "" Then

Call eTweetXL_LOC.xThrFile(thrFile)
Call eTweetXL_LOC.xTwtFile(twtFile)

On Error Resume Next

If ETWEETXLPOST.ThreadTrig.Caption = 1 Then
    If Dir(thrFile & ETWEETXLPOST.DraftBox.Value & ".thr") = "" Then
    Open thrFile & ETWEETXLPOST.DraftBox.Value & ".thr" For Output As #1
    Print #1, ""
    Close #1
            End If
            
                Else
                    
        If Dir(twtFile & ETWEETXLPOST.DraftBox.Value & ".twt") = "" Then
        Open twtFile & ETWEETXLPOST.DraftBox.Value & ".twt" For Output As #2
        Print #2, ""
        Close #2
        End If
                End If
                
                    End If

xType = 0: Call eTweetXL_GET.getPostData(xType)

If Range("xlasSilent").Value2 <> 1 Then _
ETWEETXLPOST.xlFlowStrip.Value = "Draft created..."
        
End Function
Public Function AddLinkBtn_Clk(ByVal xPos As Long)

'//For adding drafts to the Linker

If ETWEETXLPOST.DraftBox <> "" Then

Dim lastRP As Long: Dim xTot As Long: Dim xPosH As Long
Dim X As Integer
Dim xPosArr() As String
Dim xTempArr(5000) As String: Dim xTempArr2() As String
Dim xDraft As String: Dim xExt As String: Dim xT As String
Dim EDITMODE As Byte

On Error Resume Next

xTwt = ETWEETXLPOST.DraftBox.Value

Call getWindow(xWin)

If InStr(1, xTwt, " [...]") = False Then
If xWin.DraftFilterBtn.Caption <> "..." Or InStr(1, xTwt, " [•]") Then
xExt = ".twt": xT = " [•]": If xWin.DraftFilterBtn.Caption <> "•" Then xFil = 0: Call eTweetXL_CLICK.DraftFilterBtn_Clk(xFil)
    Else: GoTo asThread
            End If
        Else
asThread:
            xExt = ".thr": xT = " [...]"
            If xWin.DraftFilterBtn.Caption <> "..." Then xFil = 1: Call eTweetXL_CLICK.DraftFilterBtn_Clk(xFil)
                End If
                
lastRP = Cells(Rows.Count, "P").End(xlUp).Row '//ProfileLink for drafts

'//check for Linker edit state
    If Range("DraftTrig").Value2 <> 0 Then
        If Range("LinkerIndex").Value2 <> vbNullString Then
        xPos = Range("LinkerIndex").Value2
        '//swap
        If Range("DraftTrig").Value2 = 1 Then
        EDITMODE = 1
        '//above
        ElseIf Range("DraftTrig").Value2 = 2 Then
        EDITMODE = 2
        '//below
        ElseIf Range("DraftTrig").Value2 = 3 Then
        EDITMODE = 3
        End If
            End If
                End If
    
If xPos > 0 Or EDITMODE <> 0 Then
    
    If ETWEETXLPOST.LinkerBox.ListCount > 0 Then
    
    X = 0

    If InStr(1, xPos, ":") Then
    xPosArr = Split(xPos, ":")
    For X = xPosArr(0) To xPosArr(1)
    xDraft = "(" & X & ") " & ETWEETXLPOST.DraftBox.Value & " [" & ETWEETXLPOST.DraftFilterBtn.Caption & "]"
    If Left(xDraft, 1) = " " Then xDraft = Right(xDraft, Len(xDraft) - 1) '//remove leading space
    ETWEETXLPOST.LinkerBox.List((X)) = xDraft
    Range("ProfileLink").Offset(X, 0).Value = Range("Profile").Value2
    Range("Draftlink").Offset(lastRP, 0).Value = ETWEETXLPOST.DraftBox.Value
    Next
    Exit Function
    End If
    
    If InStr(1, xPos, ",") Then
    xPosArr = Split(xPos, ",")
    xTot = UBound(xPosArr)
    Do Until X = xTot
    xDraft = "(" & X & ") " & ETWEETXLPOST.DraftBox.Value & " [" & ETWEETXLPOST.DraftFilterBtn.Caption & "]"
    If Left(xDraft, 1) = " " Then xDraft = Right(xDraft, Len(xDraft) - 1) '//remove leading space
    ETWEETXLPOST.LinkerBox.List((X)) = xDraft
    Range("ProfileLink").Offset(X, 0).Value = Range("Profile").Value2
    Range("Draftlink").Offset(lastRP, 0).Value = ETWEETXLPOST.DraftBox.Value
    X = X + 1
    Loop
    Exit Function
    End If
                
'//swap
    If EDITMODE = 1 Then
    xPosH = xPos + 1
    
'//above
    ElseIf EDITMODE = 2 Then
    
        xPosH = xPos
        
        X = 0
        For xNum = 1 To lastRP - 1
        
        If X = xPosH Then xTempArr(X) = ETWEETXLPOST.DraftBox.Value & "[,]" & ActiveProfile: X = X + 1
        
        xTempArr(X) = _
        Range("DraftLink").Offset(xNum, 0).Value & "[,]" & _
        Range("ProfileLink").Offset(xNum, 0).Value
        X = X + 1
        Next
        
        For X = 0 To lastRP - 1
        xTempArr2 = Split(xTempArr(X), "[,]"): xDraft = xTempArr2(0)
        Range("DraftLink").Offset(X + 1, 0).Value = xTempArr2(0)
        Range("ProfileLink").Offset(X + 1, 0).Value = xTempArr2(1)
        
        If X < ETWEETXLPOST.LinkerBox.ListCount Then
        ETWEETXLPOST.LinkerBox.List(X) = "(" & X + 1 & ") " & xDraft
            Else
                ETWEETXLPOST.LinkerBox.AddItem "(" & X + 1 & ") " & xDraft
                    End If
        
        Next
    
        ETWEETXLPOST.DraftHdr.Caption = "Draft" & " (" & ETWEETXLPOST.LinkerBox.ListCount & ")"
    
        Exit Function
    
'//below
    ElseIf EDITMODE = 3 Then
    
        xPosH = xPos + 1
        
        X = 0
        For xNum = 1 To lastRP - 1
        
        If X = xPosH Then xTempArr(X) = ETWEETXLPOST.DraftBox.Value & "[,]" & ActiveProfile: X = X + 1
        
        xTempArr(X) = _
        Range("DraftLink").Offset(xNum, 0).Value & "[,]" & _
        Range("ProfileLink").Offset(xNum, 0).Value
        X = X + 1
        Next
        
        For X = 0 To lastRP - 1
        xTempArr2 = Split(xTempArr(X), "[,]"): xDraft = xTempArr2(0)
        Range("DraftLink").Offset(X + 1, 0).Value = xTempArr2(0)
        Range("ProfileLink").Offset(X + 1, 0).Value = xTempArr2(1)
        
        If X < ETWEETXLPOST.LinkerBox.ListCount Then
        ETWEETXLPOST.LinkerBox.List(X) = "(" & X + 1 & ") " & xDraft
            Else
                ETWEETXLPOST.LinkerBox.AddItem "(" & X + 1 & ") " & xDraft
                    End If
        
        Next
        
        ETWEETXLPOST.DraftHdr.Caption = "Draft" & " (" & ETWEETXLPOST.LinkerBox.ListCount & ")"
    
        Exit Function
    
        End If

    xDraft = "(" & xPosH & ") " & ETWEETXLPOST.DraftBox.Value & " [" & ETWEETXLPOST.DraftFilterBtn.Caption & "]"
    If Left(xDraft, 1) = " " Then xDraft = Right(xDraft, Len(xDraft) - 1) '//remove leading space
    ETWEETXLPOST.LinkerBox.List(xPos) = xDraft
    Range("ProfileLink").Offset(xPosH, 0).Value = Range("Profile").Value2
    Range("Draftlink").Offset(xPosH, 0).Value = ETWEETXLPOST.DraftBox.Value
    
    Exit Function
    
                        End If
                        
                            Else
                            
                            End If
                            
                            ETWEETXLPOST.DraftHdr.Caption = "Draft" & " (" & ETWEETXLPOST.LinkerBox.ListCount + 1 & ")"
                            
                            xDraft = ETWEETXLPOST.DraftHdr.Caption
                            xDraft = Replace(xDraft, "Draft", vbNullString, , , vbTextCompare)
                            xDraft = xDraft & " " & ETWEETXLPOST.DraftBox.Value
                            xDraft = Replace(xDraft, xT, vbNullString)
                            xDraft = xDraft & " [" & ETWEETXLPOST.DraftFilterBtn.Caption & "]"
                            If Left(xDraft, 1) = " " Then xDraft = Right(xDraft, Len(xDraft) - 1) '//remove leading space
                            ETWEETXLPOST.LinkerBox.AddItem (xDraft)
                            
                            Range("ProfileLink").Offset(lastRP, 0).Value = Range("Profile").Value2
                            xTwt = Replace(xTwt, xT, vbNullString)
                            xTwt = xTwt & " [" & ETWEETXLPOST.DraftFilterBtn.Caption & "]"
                            Range("Draftlink").Offset(lastRP, 0).Value = xTwt
                            
                            If Range("xlasSilent").Value2 <> 1 Then
                            xDraft = Replace(xDraft, " [•]", vbNullString)
                            xDraft = Replace(xDraft, " [...]", vbNullString)
                            ETWEETXLPOST.xlFlowStrip.Value = xDraft & " linked..."
                            End If
                            
                                End If
                        
End Function
Public Function NewProfile_Clk(ByVal xInfo As String)

'//automate profile creation w/ xlAppScript
If xInfo <> vbNullString Then

lastR = Cells(Rows.Count, "B").End(xlUp).Row '//last row user column

xInfoArr = Split(xInfo, ",")
xInfoArr(0) = Replace(xInfoArr(0), """", vbNullString)

    '//1 parameter
    If UBound(xInfoArr) = 0 Then
    If Right(xInfoArr(0), 1) = """" Then xInfoArr(0) = Left(xInfoArr(0), Len(xInfoArr(0)) - 1) '//remove ending quote
    If Left(xInfoArr(0), 1) = " " Then xInfoArr(0) = Right(xInfoArr(0), Len(xInfoArr(0)) - 1) '//remove leading space
    ETWEETXLSETUP.ProfileNameBox.Value = xInfoArr(0)
    GoTo CheckName
    End If
    '//2 parameters
    If UBound(xInfoArr) = 1 Then
    If Right(xInfoArr(0), 1) = """" Then xInfoArr(0) = Left(xInfoArr(0), Len(xInfoArr(0)) - 1) '//remove ending quote
    If Left(xInfoArr(0), 1) = " " Then xInfoArr(0) = Right(xInfoArr(0), Len(xInfoArr(0)) - 1) '//remove leading space
    xInfo = xInfoArr(0): Call eTweetXL_CLICK.NewProfile_Clk(xInfo)
    ETWEETXLSETUP.ProfileNameBox.Value = xInfoArr(0)
    If Right(xInfoArr(1), 1) = """" Then xInfoArr(1) = Left(xInfoArr(1), Len(xInfoArr(1)) - 1) '//remove ending quote
    If Left(xInfoArr(1), 1) = " " Then xInfoArr(1) = Right(xInfoArr(1), Len(xInfoArr(1)) - 1) '//remove leading space
    ETWEETXLSETUP.UsernameBox.Value = xInfoArr(1)
    ETWEETXLSETUP.PassBox.Value = "password"
    '//no pin...
    Range("Target").Offset(lastR, 0).Value = ""
    Range("Scure").Offset(lastR, 0).Value = "*"
    GoTo MoveData
    End If
    '//3 parameters
    If UBound(xInfoArr) = 2 Then
    xInfoA = xInfoArr(0)
    xInfoB = xInfoArr(1)
    xInfoC = xInfoArr(2)
    If Right(xInfoA, 1) = """" Then xInfoA = Left(xInfoA, Len(xInfoA) - 1)  '//remove ending quote
    If Left(xInfoA, 1) = " " Then xInfoA = Right(xInfoA, Len(xInfoA) - 1)     '//remove leading space
    xInfo = xInfoA: Call eTweetXL_CLICK.NewProfile_Clk(xInfo)
    ETWEETXLSETUP.ProfileNameBox.Value = xInfoA
    If Right(xInfoB, 1) = """" Then xInfoB = Left(xInfoB, Len(xInfoB) - 1) '//remove ending quote
    If Left(xInfoB, 1) = " " Then xInfoB = Right(xInfoB, Len(xInfoB) - 1) '//remove leading space
    ETWEETXLSETUP.UsernameBox.Value = xInfoB
    If Right(xInfoC, 1) = """" Then xInfoC = Left(xInfoC, Len(xInfoC) - 1) '//remove ending quote
    If Left(xInfoC, 1) = " " Then xInfoC = Right(xInfoC, Len(xInfoC) - 1) '//remove leading space
    ETWEETXLSETUP.PassBox.Value = xInfoC
    '//no pin...
    Range("Target").Offset(lastR, 0).Value = ""
    Range("Scure").Offset(lastR, 0).Value = "*"
    GoTo MoveData
    End If
        '//4 parameters
    If UBound(xInfoArr) = 3 Then
    If Right(xInfoArr(0), 1) = """" Then xInfoArr(0) = Left(xInfoArr(0), Len(xInfoArr(0)) - 1) '//remove ending quote
    If Left(xInfoArr(0), 1) = " " Then xInfoArr(0) = Right(xInfoArr(0), Len(xInfoArr(0)) - 1) '//remove leading space
    xInfoA = xInfoArr(0)
    xInfoB = xInfoArr(1)
    xInfoC = xInfoArr(2)
    If Right(xInfoA, 1) = """" Then xInfoA = Left(xInfoA, Len(xInfoA) - 1) '//remove ending quote
    If Left(xInfoA, 1) = " " Then xInfoA = Right(xInfoA, Len(xInfoA) - 1) '//remove leading space
    xInfo = xInfoA: Call eTweetXL_CLICK.NewProfile_Clk(xInfo)
    ETWEETXLSETUP.ProfileNameBox.Value = xInfoA
    If Right(xInfoB, 1) = """" Then xInfoB = Left(xInfoB, Len(xInfoB) - 1) '//remove ending quote
    If Left(xInfoB, 1) = " " Then xInfoB = Right(xInfoB, Len(xInfoB) - 1) '//remove leading space
    ETWEETXLSETUP.UsernameBox.Value = xInfoB
    If Right(xInfoC, 1) = """" Then xInfoC = Left(xInfoC, Len(xInfoC) - 1) '//remove ending quote
    If Left(xInfoC, 1) = " " Then xInfoC = Right(xInfoC, Len(xInfoC) - 1) '//remove leading space
    ETWEETXLSETUP.PassBox.Value = xInfoC
    PinHldr = xInfoArr(3): X = PinHldr: Call basBinaryHash1(X, xVerify, xHash): PinHldr = """" & xHash & """"
    '//Check edit mode...
    If Range("EditStatus").Value <> 1 Then
    '//Pin for new user...
    Range("Scure").Offset(lastR, 0).Value = "***"
    Range("Target").Offset(lastR, 0).Value = PinHldr '//record
        Else
    '//Replace existing user pin...
    Range("Scure").Offset((lastR - 1), 0).Value = "***"
    Range("Target").Offset((lastR - 1), 0).Value = PinHldr '//record
    End If
    GoTo MoveData
        End If
            
If Right(xInfo, 1) = """" Then xInfo = Left(xInfo, Len(xInfo) - 1) '//remove ending quote
If Left(xInfo, 1) = " " Then xInfo = Right(xInfo, Len(xInfo) - 1) '//remove leading space
ETWEETXLSETUP.ProfileNameBox.Value = xInfo
End If

GoTo CheckName
MoveData:
'(xlas)
       '//Move our data for transfer...
        Range("Profile").Offset(lastR, 0).Value = ETWEETXLSETUP.UsernameBox.Value
        Range("F1").Offset(lastR, 0).Value = ETWEETXLSETUP.PassBox.Value
        
        '//Export data
        Call eTweetXL_POST.pstPersData
        
        '//Refresh data
        Range("DataPullTrig").Value2 = "0"
        '//Import profile information
        Call eTweetXL_GET.getProfileData
        
CheckName:
If ETWEETXLSETUP.ProfileNameBox.Value = "" Then
ETWEETXLSETUP.ProfileNameBox.BorderStyle = fmBorderStyleSingle
ETWEETXLSETUP.ProfileNameBox.BorderColor = vbRed
xMsg = 20: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
Exit Function
End If
'//--------------------------------------------------------------------------------------------------------------------
If ETWEETXLSETUP.ProfileListBox.Value = "" Then ETWEETXLSETUP.ProfileListBox.Value = ETWEETXLSETUP.ProfileNameBox.Value

On Error Resume Next

    If Dir(AppLoc & "\presets\" & ETWEETXLSETUP.ProfileNameBox.Value, vbDirectory) = "" Then
        MkDir (AppLoc & "\presets\" & ETWEETXLSETUP.ProfileNameBox.Value)
        MkDir (AppLoc & "\presets\" & ETWEETXLSETUP.ProfileNameBox.Value & "\pers")
        MkDir (AppLoc & "\presets\" & ETWEETXLSETUP.ProfileNameBox.Value & "\thr")
        MkDir (AppLoc & "\presets\" & ETWEETXLSETUP.ProfileNameBox.Value & "\twt")
            End If
            
        

On Error GoTo ErrMsg
        
        '//REFRESH
        ETWEETXLSETUP.UserListBox.Clear
        ETWEETXLSETUP.ProfileListBox.Value = ETWEETXLSETUP.ProfileNameBox.Value
        '//Import profile information
        If Range("Profile").Value2 = ETWEETXLSETUP.ProfileNameBox.Value Then
        If ETWEETXLSETUP.ProfileListBox.Value = "" Then
           ETWEETXLSETUP.ProfileListBox.Value = ETWEETXLSETUP.ProfileNameBox.Value
                End If
                    End If
        
        If ETWEETXLSETUP.ProfileListBox.Value <> "" Then
        Range("Profile").Value2 = ETWEETXLSETUP.ProfileListBox.Value
        ETWEETXLSETUP.ProfileNameBox.Value = ETWEETXLSETUP.ProfileListBox.Value
        End If
        
        Range("DataPullTrig").Value2 = 0
        
        Call eTweetXL_GET.getProfileData
        
Exit Function

ErrMsg:
Range("GetInfo").Value = vbNullString
Range("EditStatus").Value = 0
              
              
End Function
Public Function NewUser_Clk(ByVal xInfo As String)

On Error GoTo ErrMsg

Dim EDITMODE As Integer: Dim PinHldr As String

lastR = Cells(Rows.Count, "B").End(xlUp).Row '//last row user column

'//automate user creation w/ xlAppScript
If xInfo <> vbNullString Then
xInfoArr = Split(xInfo, ",")
xInfoArr(0) = Replace(xInfoArr(0), """", vbNullString)
    
    '//1 parameter
    If UBound(xInfoArr) = 0 Then
    If Right(xInfoArr(0), 1) = """" Then xInfoArr(0) = Left(xInfoArr(0), Len(xInfoArr(0)) - 1) '//remove ending quote
    If Left(xInfoArr(0), 1) = " " Then xInfoArr(0) = Right(xInfoArr(0), Len(xInfoArr(0)) - 1) '//remove leading space
    ETWEETXLSETUP.UsernameBox.Value = xInfoArr(0)
    ETWEETXLSETUP.PassBox.Value = "password" '//default pass
    '//no pin...
    Range("Target").Offset(lastR, 0).Value = ""
    Range("Scure").Offset(lastR, 0).Value = "*"
    GoTo MoveData
    End If
    '//2 parameters
    If UBound(xInfoArr) = 1 Then
    If Right(xInfoArr(0), 1) = """" Then xInfoArr(0) = Left(xInfoArr(0), Len(xInfoArr(0)) - 1) '//remove ending quote
    If Left(xInfoArr(0), 1) = " " Then xInfoArr(0) = Right(xInfoArr(0), Len(xInfoArr(0)) - 1) '//remove leading space
    ETWEETXLSETUP.UsernameBox.Value = xInfoArr(0)
    If Right(xInfoArr(1), 1) = """" Then xInfoArr(1) = Left(xInfoArr(1), Len(xInfoArr(1)) - 1) '//remove ending quote
    If Left(xInfoArr(1), 1) = " " Then xInfoArr(1) = Right(xInfoArr(1), Len(xInfoArr(1)) - 1) '//remove leading space
    ETWEETXLSETUP.PassBox.Value = xInfoArr(1)
    '//no pin...
    Range("Target").Offset(lastR, 0).Value = ""
    Range("Scure").Offset(lastR, 0).Value = "*"
    GoTo MoveData
    End If
    '//3 parameters
    If UBound(xInfoArr) = 2 Then
    If Right(xInfoArr(0), 1) = """" Then xInfoArr(0) = Left(xInfoArr(0), Len(xInfoArr(0)) - 1) '//remove ending quote
    If Left(xInfoArr(0), 1) = " " Then xInfoArr(0) = Right(xInfoArr(0), Len(xInfoArr(0)) - 1) '//remove leading space
    ETWEETXLSETUP.UsernameBox.Value = xInfoArr(0)
    If Right(xInfoArr(1), 1) = """" Then xInfoArr(1) = Left(xInfoArr(1), Len(xInfoArr(1)) - 1) '//remove ending quote
    If Left(xInfoArr(1), 1) = " " Then xInfoArr(1) = Right(xInfoArr(1), Len(xInfoArr(1)) - 1) '//remove leading space
    ETWEETXLSETUP.PassBox.Value = xInfoArr(1)
    If Right(xInfoArr(2), 1) = """" Then xInfoArr(2) = Left(xInfoArr(2), Len(xInfoArr(2)) - 1) '//remove ending quote
    If Left(xInfoArr(2), 1) = " " Then xInfoArr(2) = Right(xInfoArr(2), Len(xInfoArr(2)) - 1) '//remove leading space
    PinHldr = xInfoArr(2): X = PinHldr: Call basBinaryHash1(X, xVerify, xHash): PinHldr = """" & xHash & """"
    '//Check edit mode...
    If EDITMODE <> 1 Then
    '//Pin for new user...
    Range("Scure").Offset(lastR, 0).Value = "***"
    Range("Target").Offset(lastR, 0).Value = PinHldr '//record
        Else
    '//Replace existing user pin...
    Range("Scure").Offset((lastR - 1), 0).Value = "***"
    Range("Target").Offset((lastR - 1), 0).Value = PinHldr '//record
    End If
    GoTo MoveData
        End If
            End If

'//check user box...
If ETWEETXLSETUP.UsernameBox.Value = "" Then
ETWEETXLSETUP.UsernameBox.BorderStyle = fmBorderStyleSingle
ETWEETXLSETUP.UsernameBox.BorderColor = vbRed
xMsg = 18: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
Exit Function
End If
    
    '//check pass box...
    If ETWEETXLSETUP.PassBox.Value = "" Then
    ETWEETXLSETUP.PassBox.BorderStyle = fmBorderStyleSingle
    ETWEETXLSETUP.PassBox.BorderColor = vbRed
    xMsg = 19: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
    Exit Function
    End If
    
        '//Check for editing mode
        EDITMODE = 0
        
        If Range("EditStatus").Value = 1 Then
            EDITMODE = 1
                ElseIf Range("EditStatus").Value = 2 Then
                    EDITMODE = 2
                            End If

        
        '//If in editing mode...
        If EDITMODE = 2 Then
            '//Check for information...
            If Range("GetInfo").Value = vbNullString Then
                Range("EditStatus").Value = 0
                EDITMODE = 0
                xMsg = 21: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
                xMsg = 22: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
                GoTo ExitEditMode '//Exit edit mode if information not found
                    End If
                    
                    
            GetInfoArr = Split(Range("GetInfo").Value, ",") '//Get edit info
            
            For xPos = 1 To lastR
            If Range("Profile").Offset(xPos, 0) = GetInfoArr(0) Then
            lastR = xPos
                GoTo AskForPin
                    End If
                        Next
                        
                            End If
                            
ExitEditMode:
        For xPos = 1 To lastR
        If ETWEETXLSETUP.UsernameBox.Value = Range("Profile").Offset(xPos, 0) Then GoTo EditUser
        Next
        
        '//If user doesn't exist...
        lastR = Cells(Rows.Count, "A").End(xlUp).Row '//Last row profile column
        
'ASKFORPIN
AskForPin:

    If EDITMODE = 0 Then
        '//Enter pin?
        If Range("xlasSilent") <> 1 Then msg = MsgBox("Would you like to enter a pin for this user?", vbYesNo, eTweetXL_INFO.AppName)
            ElseIf EDITMODE = 1 Then
        '//Enter new pin?
        If Range("xlasSilent") <> 1 Then msg = MsgBox("Would you like to enter a new pin for this user?", vbYesNo, eTweetXL_INFO.AppName)
        '//We've been here, move our data...
                Else: GoTo MoveData
                        End If
                
                '//Exit if nothing entered...
                If msg = "" Then Exit Function
                
        '//Yes
        If msg = vbYes Then
'ENTERPIN
EnterPin:
            If Range("xlasSilent") <> 1 Then msg = InputBox("Enter a 4-digit pin:", eTweetXL_INFO.AppName)
                PinHldr = msg
                
                If PinHldr <> "" Then
                '//Check for 4 digits
                 If Len(PinHldr) = 4 Then GoTo ReEnterPin '//Success
                    If Len(PinHldr) < 4 Then '//Too short
                        If Range("xlasSilent") <> 1 Then MsgBox ("This pin is too short"), vbInformation, eTweetXL_INFO.AppName
                            ElseIf Len(PinHldr) > 4 Then '//Too long
                                If Range("xlasSilent") <> 1 Then MsgBox ("This pin is too long"), vbInformation, eTweetXL_INFO.AppName
                                    End If
                                    ErrCntr = ErrCntr + 1 '//Error counter
                                    If ErrCntr >= 5 Then GoTo ErrMsg '//5 attempts until quit...
                                    GoTo EnterPin
                                    

                                  
'REENTERPIN
ReEnterPin:
                '//Check pin
                 If Range("xlasSilent") <> 1 Then msg = InputBox("Re-enter pin:", eTweetXL_INFO.AppName)
                    If msg = PinHldr Then
                    X = PinHldr: Call basBinaryHash1(X, xVerify, xHash): PinHldr = """" & xHash & """"
                    '//CHheck editing mode...
                        If EDITMODE <> 1 Then
                        '//set pin for new user...
                        Range("Scure").Offset(lastR, 0).Value = "***"
                        Range("Target").Offset(lastR, 0).Value = PinHldr '//RECORD
                            Else
                        '//set pin for existing user...
                        Range("Scure").Offset(xPos, 0).Value = "***"
                        Range("Target").Offset(xPos, 0).Value = PinHldr '//RECORD
                        End If
                            Else
                        '//wrong pin entered...
                        If Range("xlasSilent") <> 1 Then msg = MsgBox("Incorrect pin", vbExclamation, eTweetXL_INFO.AppName)
                                ErrCntr = ErrCntr + 1
                                If ErrCntr >= 5 Then GoTo ErrMsg '//5 attempts until quit...
                                GoTo ReEnterPin
                                    End If
                                        End If
                                        
                                            Else
                                            
                                            '//No pin entered...
                                            Range("Target").Offset(lastR, 0).Value = ""
                                            Range("Scure").Offset(lastR, 0).Value = "*"
                                            
                                                    End If
                                                        
                                                        
If EDITMODE = 1 Then GoTo UpdateInfo

'MOVEDATA
MoveData:

       '//Move our data for transfer...
        Range("Profile").Offset(lastR, 0).Value = ETWEETXLSETUP.UsernameBox.Value
        Range("F1").Offset(lastR, 0).Value = ETWEETXLSETUP.PassBox.Value
        
        '//Export data
        Call eTweetXL_POST.pstPersData
        
        '//Refresh data
        Range("DataPullTrig").Value2 = "0"
        '//Import profile information
        Call eTweetXL_GET.getProfileData
        
                
'//EXITING EDIT MODE...
If EDITMODE = 2 Then
'//Exiting edit mode
xMsg = 22: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
    Else
        End If
        
Range("EditStatus").Value = 0

Exit Function

'//Ask to edit user if user exists
EditUser:
Dim nl As Variant
nl = vbNewLine

xUser = Range("Profile").Offset(xPos, 0).Value '//Get user position
xPass = Range("F1").Offset(xPos, 0).Value '//Get pass position
xPin = Range("Target").Offset(xPos, 0).Value '//Get pin position

xGetInfo = xUser & "," & xPass & "," & xPin
Range("GetInfo").Value = xGetInfo

If Range("xlasSilent") <> 1 Then msg = MsgBox("[ " & xUser & " ]" & nl & nl & _
"Edit information for this user?", vbYesNo, eTweetXL_INFO.AppName)
    If msg = vbYes Then
    Range("EditStatus").Value = 1
    EDITMODE = 1
        GoTo AskForPin '//Ask to setup new pin
            Else: Exit Function
                End If
            
'//Enter new information click '+' to update...
UpdateInfo:
        If Range("xlasSilent") <> 1 Then MsgBox ("Enter new information for '" & xUser & "', then click the '+' button again to update."), vbInformation, eTweetXL_INFO.AppName
        Range("EditStatus").Value = 2
Exit Function

ErrMsg:
If Range("xlasSilent") <> 1 Then MsgBox ("Error adding pin for this user"), vbExclamation, eTweetXL_INFO.AppName
Range("GetInfo").Value = vbNullString
Range("EditStatus").Value = 0

End Function
Public Function AddUserBtn_Clk(ByVal xPos As Long)

Dim lastRA As Long: Dim lastRAM As Long: Dim xTot As Long: Dim xPosH As Long
Dim X As Integer: Dim xNum As Integer
Dim xPosArr() As String
Dim xTempArr(5000) As String: Dim xTempArr2() As String
Dim xPass As String: Dim xType As String: Dim xUser As String
Dim EDITMODE As Byte

'//For adding users to the Linker

If ETWEETXLPOST.UserListBox <> vbNullString Then

'//check for Linker edit state
    If Range("UserTrig").Value2 <> 0 Then
        If Range("LinkerIndex").Value2 <> vbNullString Then
        xPos = Range("LinkerIndex").Value2
        '//swap
        If Range("UserTrig").Value2 = 1 Then
        EDITMODE = 1
        '//above
        ElseIf Range("UserTrig").Value2 = 2 Then
        EDITMODE = 2
        '//below
        ElseIf Range("UserTrig").Value2 = 3 Then
        EDITMODE = 3
        End If
            End If
                End If

If EDITMODE <> 1 Then ETWEETXLPOST.UserHdr.Caption = "User" & " (" & ETWEETXLPOST.UserBox.ListCount + 1 & ")"
    
On Error Resume Next

lastRA = Cells(Rows.Count, "A").End(xlUp).Row
lastRAM = Cells(Rows.Count, "AM").End(xlUp).Row
xUser = Range("User").Value '//focused user
    
If xPos > 0 Or EDITMODE <> 0 Then

xPosH = xPos

    If ETWEETXLPOST.UserBox.ListCount > 0 Then
    
    X = 0

    If InStr(1, xPos, ":") Then
    xPosArr = Split(xPos, ":")
    For X = xPosArr(0) To xPosArr(1)
    xUser = "(" & X & ") " & ETWEETXLPOST.UserListBox.Value
    If Left(xUser, 1) = " " Then xUser = Right(xUser, Len(xUser) - 1) '//remove leading space
    ETWEETXLPOST.UserBox.List((X)) = (xUser)
    Next
    Exit Function
    End If
    
    If InStr(1, xPos, ",") Then
    xPosArr = Split(xPos, ",")
    xTot = UBound(xPosArr)
    Do Until X = xTot
    xUser = "(" & X & ") " & ETWEETXLPOST.UserListBox.Value
    If Left(xUser, 1) = " " Then xUser = Right(xUser, Len(xUser) - 1) '//remove leading space
    ETWEETXLPOST.UserBox.List(xPosArr(X)) = xUser
    X = X + 1
    Loop
    Exit Function
    End If
    
'//swap
    If EDITMODE = 1 Then
    
    xPosH = xPos + 1
    
        For xNum = 1 To lastRA
        If Range("A" & xNum).Value = xUser Then
        Range("TargetLink").Offset(xPosH, 0).Value = Range("F" & xNum).Value
        Range("MainLink").Offset(xPosH, 0).Value = Range("Profile").Value2
        End If
        Next

        '//Check for API send...
         If Range("SendAPI").Value = 1 Then
         Call eTweetXL_LOC.xApiFile(apiFile)
            If Dir(apiFile) = "" Then
                xMsg = 7: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
                Else
                    xUser = xUser & "(*api)"
                        End If
                            End If
                            
        '//Set API send...
        If InStr(1, xUser, "(*api)") Then
        xUser = Replace(ETWEETXLPOST.UserListBox.Value, "(*api)", "")
        xUser = Replace(xUser, Range("Scure").Value, "")
        Range("apiLink").Offset(xPosH, 0).Value = "(*api)"
        Else
            Range("apiLink").Offset(xPosH, 0).Value = "(*)"
                End If
                
        xUser = "(" & xPosH & ") " & ETWEETXLPOST.UserListBox.Value
        xUser = Replace(xUser, Range("Scure").Value, vbNullString)
        If Left(xUser, 1) = " " Then xUser = Right(xUser, Len(xUser) - 1) '//remove leading space
        Range("UserLink").Offset(xPosH, 0).Value = xUser
        ETWEETXLPOST.UserBox.List(xPos) = xUser
        
        If Range("xlasSilent").Value2 <> 1 Then
        ETWEETXLPOST.xlFlowStrip.Value = xUser & " linked..."
        End If
        
'//above
    ElseIf EDITMODE = 2 Then
    
    xPosH = xPos
    
        For xNum = 1 To lastRA
        If Range("A" & xNum).Value = xUser Then xPass = Range("F" & xNum).Value
        Next

        '//Check for API send...
         If Range("SendAPI").Value = 1 Then
         Call eTweetXL_LOC.xApiFile(apiFile)
            If Dir(apiFile) = "" Then
                xMsg = 7: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
                Else
                    xUser = xUser & "(*api)"
                        End If
                            End If
                            
        '//Set API send...
        If InStr(1, xUser, "(*api)") Then
        xUser = Replace(ETWEETXLPOST.UserListBox.Value, "(*api)", "")
        xUser = Replace(xUser, Range("Scure").Value, "")
        xType = "(*api)"
        Else
            xType = "(*)"
                End If
        
        X = 0
        For xNum = 1 To lastRAM - 1
        
        If X = xPosH Then xTempArr(X) = _
        xType & "[,]" & _
        ActiveProfile & "[,]" & _
        xPass & "[,]" & _
        "(0) " & xUser: X = X + 1
        
        xTempArr(X) = _
        Range("apiLink").Offset(xNum, 0).Value & "[,]" & _
        Range("MainLink").Offset(xNum, 0).Value & "[,]" & _
        Range("TargetLink").Offset(xNum, 0).Value & "[,]" & _
        Range("UserLink").Offset(xNum, 0).Value
        X = X + 1
        Next
        
        For X = 0 To lastRAM - 1
        xTempArr2 = Split(xTempArr(X), "[,]")
        Range("apiLink").Offset(X + 1, 0).Value = xTempArr2(0)
        Range("MainLink").Offset(X + 1, 0).Value = xTempArr2(1)
        Range("TargetLink").Offset(X + 1, 0).Value = xTempArr2(2)
        xPosArr = Split(xTempArr2(3), ") "): xUser = xPosArr(1)
        Range("UserLink").Offset(X + 1, 0).Value = "(" & X + 1 & ") " & xUser
        
        If X < ETWEETXLPOST.UserBox.ListCount Then
        ETWEETXLPOST.UserBox.List(X) = "(" & X + 1 & ") " & xUser
            Else
                ETWEETXLPOST.UserBox.AddItem "(" & X + 1 & ") " & xUser
                    End If
        
        Next
        
'//below
    ElseIf EDITMODE = 3 Then
    
    xPosH = xPos + 1
    
        For xNum = 1 To lastRA
        If Range("A" & xNum).Value = xUser Then xPass = Range("F" & xNum).Value
        Next

        '//Check for API send...
         If Range("SendAPI").Value = 1 Then
         Call eTweetXL_LOC.xApiFile(apiFile)
            If Dir(apiFile) = "" Then
                xMsg = 7: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
                Else
                    xUser = xUser & "(*api)"
                        End If
                            End If
                            
        '//Set API send...
        If InStr(1, xUser, "(*api)") Then
        xUser = Replace(ETWEETXLPOST.UserListBox.Value, "(*api)", "")
        xUser = Replace(xUser, Range("Scure").Value, "")
        xType = "(*api)"
        Else
            xType = "(*)"
                End If
        
        X = 0
        For xNum = 1 To lastRAM - 1
        
        If X = xPosH Then xTempArr(X) = _
        xType & "[,]" & _
        ActiveProfile & "[,]" & _
        xPass & "[,]" & _
        "(0) " & xUser: X = X + 1
        
        xTempArr(X) = _
        Range("apiLink").Offset(xNum, 0).Value & "[,]" & _
        Range("MainLink").Offset(xNum, 0).Value & "[,]" & _
        Range("TargetLink").Offset(xNum, 0).Value & "[,]" & _
        Range("UserLink").Offset(xNum, 0).Value
        X = X + 1
        Next
        
        For X = 0 To lastRAM - 1
        xTempArr2 = Split(xTempArr(X), "[,]")
        Range("apiLink").Offset(X + 1, 0).Value = xTempArr2(0)
        Range("MainLink").Offset(X + 1, 0).Value = xTempArr2(1)
        Range("TargetLink").Offset(X + 1, 0).Value = xTempArr2(2)
        xPosArr = Split(xTempArr2(3), ") "): xUser = xPosArr(1)
        Range("UserLink").Offset(X + 1, 0).Value = "(" & X + 1 & ") " & xUser
        
        If X < ETWEETXLPOST.UserBox.ListCount Then
        ETWEETXLPOST.UserBox.List(X) = "(" & X + 1 & ") " & xUser
            Else
                ETWEETXLPOST.UserBox.AddItem "(" & X + 1 & ") " & xUser
                    End If
        
        Next
    
            End If
       
    Exit Function
        
    End If
      
    Else
        
          xUser = Replace(ETWEETXLPOST.UserListBox.Value, Range("Scure").Value, "")
                 
            For xNum = 1 To lastRA
            If Range("A" & xNum).Value = xUser Then
            Range("TargetLink").Offset(lastRAM, 0).Value = Range("F" & xNum).Value
            Range("MainLink").Offset(lastRAM, 0).Value = Range("Profile").Value2
            End If
            Next
            
            '//Check for API send...
             If Range("SendAPI").Value = 1 Then
             Call eTweetXL_LOC.xApiFile(apiFile)
                If Dir(apiFile) = "" Then
                    xMsg = 7: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
                    Else
                        xUser = xUser & "(*api)"
                            End If
                                End If
                                    
            '//Set API send...
            If InStr(1, xUser, "(*api)") Then
            xUser = Replace(ETWEETXLPOST.UserListBox.Value, "(*api)", "")
            xUser = Replace(xUser, Range("Scure").Value, "")
            Range("apiLink").Offset(lastRAM, 0).Value = "(*api)"
                Else
                    Range("apiLink").Offset(lastRAM, 0).Value = "(*)"
                        End If
            xUser = ETWEETXLPOST.UserHdr.Caption & " " & xUser
            xUser = Replace(xUser, "User ", vbNullString, , , vbTextCompare)
            If Left(xUser, 1) = " " Then xUser = Right(xUser, Len(xUser) - 1) '//remove leading space
            Range("UserLink").Offset(lastRAM, 0).Value = xUser
            ETWEETXLPOST.UserBox.AddItem xUser

                If Range("xlasSilent").Value2 <> 1 Then
                ETWEETXLPOST.xlFlowStrip.Value = xUser & " linked..."
                End If
                
                    End If
        
                        End If
                                                       
                                                       
                                
End Function
Public Sub RmvPostMedBtn_Clk()

'//For removing Media from a post

Dim MedArr(5) As String
Dim X As Integer
Dim oMedLinkBox As Object

On Error GoTo SetForm

'//Find running window
Call getWindow(xWin)

Retry:
'//Check xlFlowStrip Window...
If Range("xlasWinForm").Value2 = 13 Then Set oMedLinkBox = ETWEETXLPOST.MedLinkBox
If Range("xlasWinForm").Value2 = 14 Then Set oMedLinkBox = ETWEETXLQUEUE.MedLinkBox

'//Cleanup
oMedLinkBox.Value = ""

'//Get selected media from position
xPos = Range("MediaScroll").Offset(Range("MedScrollPos").Value)

If InStr(1, xTwt, " [...]") = False Then
If xWin.DraftFilterBtn.Caption <> "..." Or InStr(1, xTwt, " [•]") Then
xExt = ".twt": xT = " [•]": If xWin.DraftFilterBtn.Caption <> "•" Then xFil = 0: Call eTweetXL_CLICK.DraftFilterBtn_Clk(xFil)
    Else: GoTo asThread
            End If
        Else
asThread:
            xExt = ".thr": xT = " [...]"
            If xWin.DraftFilterBtn.Caption <> "..." Then xFil = 1: Call eTweetXL_CLICK.DraftFilterBtn_Clk(xFil)
            GoTo RmvThreadMed
                End If
                      
'//
lastR = Cells(Rows.Count, "I").End(xlUp).Row

'//Get media paths from sheet except the one to remove
xNum = 0
For X = 0 To lastR
    If Range("MediaScroll").Offset(X, 0).Value <> xPos Then
        MedArr(xNum) = Range("MediaScroll").Offset(X, 0).Value
        xNum = xNum + 1
            End If
                Next X

'//Clear media area
Range("I1:I" & lastR).ClearContents

'//Print remaining to media area
X = 0
Do Until MedArr(X) = ""
Range("MediaScroll").Offset(X, 0).Value = MedArr(X)
X = X + 1
Loop

'//Print remaining back to Window
X = 0
Do Until MedArr(X) = vbNullString
MedArr(X) = "'" & MedArr(X) & "'"
MedArr(X) = Replace(MedArr(X), "'", """")
If oMedLinkBox.Value = vbNullString Then
oMedLinkBox.Value = MedArr(X)
    Else
        oMedLinkBox.Value = oMedLinkBox.Value & " " & MedArr(X)
            End If
                X = X + 1
                    Loop

If oMedLinkBox.Value = "" Then
Range("GifCntr").Value = 0
Range("VidCntr").Value = 0
End If

Set oMedLinkBox = Nothing

Call eTweetXL_GET.getSelMedia

Exit Sub

RmvThreadMed:

On Error Resume Next

Dim xNwMed As String
xNwMed = vbNullString

lastR = Cells(Rows.Count, "Z").End(xlUp).Row

If xWin.ThreadCt.Caption <> "" Then
If CInt(xWin.ThreadCt.Caption) > 0 Then
lastR = xWin.ThreadCt.Caption
GoTo NextStep
End If
    End If
    
NextStep:
'//remove post from thread loc
xMedArr = Split(Range("MedThread").Offset(lastR, 0).Value, ",")

'//get selected media position
xPos = Range("MedScrollPos").Value

'//remove media from thread loc
xMedArr(xPos) = vbNullString

'//reconnect media
For X = 0 To UBound(xMedArr)
If xMedArr(X) <> vbNullString Then xNwMed = xNwMed & xMedArr(X)
Next

'//print remaining to worksheet
Range("MedThread").Offset(lastR, 0).Value = xNwMed

'//print remaining back to Window
For X = 0 To UBound(xMedArr)
If xMedArr(X) <> vbNullString Then
xMedArr(X) = "'" & xMedArr(X) & "'"
xMedArr(X) = Replace(xMedArr(X), "'", """")
If oMedLinkBox.Value = vbNullString Then
oMedLinkBox.Value = xMedArr(X)
    Else
        oMedLinkBox.Value = oMedLinkBox.Value & " " & xMedArr(X)
            End If
                End If
                    Next
                
                
Exit Sub

SetForm:
Err.Clear
If Range("xlasWinForm").Value2 <> 13 Then Range("xlasWinForm").Value2 = 13 Else Range("xlasWinForm").Value2 = 14
GoTo Retry

End Sub
Public Sub DelDraftBtn_Clk()

'//For removing drafts from a profile

If ETWEETXLPOST.DraftBox.Value <> "" Then

Call getWindow(xWin)
Call eTweetXL_LOC.xThrFile(thrFile)
Call eTweetXL_LOC.xTwtFile(twtFile)

xPos = ETWEETXLPOST.DraftBox.ListIndex
xTwt = ETWEETXLPOST.DraftBox.Value

On Error Resume Next

    If Dir(twtFile & xTwt Or thrFile & xTwt) <> "" Then
    If ETWEETXLPOST.ThreadTrig = 0 Then Kill (twtFile & xTwt & ".twt")
    If ETWEETXLPOST.ThreadTrig = 1 Then Kill (thrFile & xTwt & ".thr")
            
            End If


If InStr(1, xTwt, " [...]") = False Then
If xWin.DraftFilterBtn.Caption <> "..." Or InStr(1, xTwt, " [•]") Then
xExt = ".twt": xT = " [•]": If xWin.DraftFilterBtn.Caption <> "•" Then xFil = 0: Call eTweetXL_CLICK.DraftFilterBtn_Clk(xFil)
    Else: GoTo asThread
            End If
        Else
asThread:
            xExt = ".thr": xT = " [...]"
            If xWin.DraftFilterBtn.Caption <> "..." Then xFil = 1: Call eTweetXL_CLICK.DraftFilterBtn_Clk(xFil)
                End If

                       If xExt = ".twt" Then xType = 0
                       If xExt = ".thr" Then xType = 1
                       
                       Call eTweetXL_GET.getPostData(xType)
                        
                        '//set to next draft
                        ETWEETXLPOST.DraftBox.Value = ETWEETXLPOST.DraftBox.List(xPos)
                        
                            End If
                            
End Sub
Public Sub RmvLinkBtn_Clk()

'//For removing a selected draft from the Linker

Dim xNum As Integer

If ETWEETXLPOST.LinkerBox.ListCount > 0 Then
ETWEETXLPOST.DraftHdr.Caption = "Draft" & " (" & ETWEETXLPOST.LinkerBox.ListCount - 1 & ")"
End If

RtCntr = ETWEETXLPOST.LinkerBox.ListCount

xNum = ETWEETXLPOST.LinkerBox.ListIndex

If xNum < 0 Then
 xNum = (RtCntr - 1)
        If xNum <= 0 Then
                If xNum < 0 Then Exit Sub
                ETWEETXLPOST.LinkerBox.RemoveItem (xNum)
                xBox = 2: Call setBoxList(xBox)
                    Exit Sub
                        End If
                            End If
                            
                If Range("xlasSilent").Value2 <> 1 Then
                ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.LinkerBox.List(xNum) & " unlinked..."
                End If

lastR = Cells(Rows.Count, "P").End(xlUp).Row
Range("DraftLink").Offset(lastR - 1, 0).Value = vbNullString
Range("ProfileLink").Offset(lastR - 1, 0).Value = vbNullString

ETWEETXLPOST.LinkerBox.RemoveItem (xNum)
xBox = 2: Call setBoxList(xBox)

End Sub
Public Sub RmvUserBtn_Clk()

'//For removing a selected user from the Linker

Dim xNum As Integer

If ETWEETXLPOST.UserBox.ListCount > 0 Then
ETWEETXLPOST.UserHdr.Caption = "User" & " (" & ETWEETXLPOST.UserBox.ListCount - 1 & ")"
End If

RtCntr = ETWEETXLPOST.UserBox.ListCount

xNum = ETWEETXLPOST.UserBox.ListIndex

If xNum < 0 Then
 xNum = (RtCntr - 1)
        If xNum <= 0 Then
                If xNum < 0 Then Exit Sub
                ETWEETXLPOST.UserBox.RemoveItem (xNum)
                xBox = 1: Call setBoxList(xBox)
                    Exit Sub
                        End If
                            End If
                            
                If Range("xlasSilent").Value2 <> 1 Then
                ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.UserBox.List(xNum) & " unlinked..."
                End If

lastR = Cells(Rows.Count, "AM").End(xlUp).Row

Range("apiLink").Offset(lastR - 1, 0).Value = vbNullString
Range("MainLink").Offset(lastR - 1, 0).Value = vbNullString
Range("TargetLink").Offset(lastR - 1, 0).Value = vbNullString
Range("UserLink").Offset(lastR - 1, 0).Value = vbNullString

ETWEETXLPOST.UserBox.RemoveItem (xNum)
xBox = 1: Call setBoxList(xBox)

End Sub
Public Sub RmvRuntimeBtn_Clk()

Dim xNum As Integer

If ETWEETXLPOST.RuntimeBox.ListCount > 0 Then
ETWEETXLPOST.RuntimeHdr.Caption = "Runtime" & " (" & ETWEETXLPOST.RuntimeBox.ListCount - 1 & ")"
End If

RtCntr = ETWEETXLPOST.RuntimeBox.ListCount

xNum = ETWEETXLPOST.RuntimeBox.ListIndex
If xNum < 0 Then
    xNum = (RtCntr - 1)
        If xNum < 0 Then Exit Sub
            End If
       
                If Range("xlasSilent").Value2 <> 1 Then
                ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.RuntimeBox.List(xNum) & " unlinked..."
                End If
                
ETWEETXLPOST.RuntimeBox.RemoveItem (xNum)
xBox = 4: Call setBoxList(xBox)

End Sub
Public Sub DelAllUsersBtn_Clk()

Dim xProf, xMsg As String

If ETWEETXLSETUP.ProfileNameBox.Value <> vbNullString Or Range("Profile").Value2 <> vbNullString Then

If ETWEETXLSETUP.ProfileNameBox.Value <> vbNullString Then xProf = ETWEETXLSETUP.ProfileNameBox.Value Else _
xProf = Range("Profile").Value2

If Range("xlasSilent").Value2 = 1 Then xMsg = vbYes: GoTo SilentRun

xMsg = MsgBox("Are you sure you wish to remove all users for '" & xProf & "'?", vbYesNo, eTweetXL_INFO.AppName)

SilentRun:
If xMsg = vbYes Then

lastR = Cells(Rows.Count, "A").End(xlUp).Row

For X = 1 To lastR
xInfo = Range("Profile").Offset(X, 0).Value
Call eTweetXL_CLICK.DelUserBtn_Clk(xInfo)
Next
End If

End If
End Sub
Public Sub DelUserBtn_Clk(ByVal xInfo As String)

If xInfo <> vbNullString Then
If Right(xInfo, 1) = """" Then xInfo = Left(xInfo, Len(xInfo) - 1) '//remove ending quote
If Left(xInfo, 1) = """" Then xInfo = Right(xInfo, Len(xInfo) - 1) '//remove leading quote
If Left(xInfo, 1) = " " Then xInfo = Right(xInfo, Len(xInfo) - 1) '//remove leading space
ETWEETXLSETUP.UsernameBox.Value = xInfo
End If

'//For removing a specific user from a profile
If ETWEETXLSETUP.UsernameBox.Value <> "" Then
        
        lastR = Cells(Rows.Count, "A").End(xlUp).Row
        
        For xNum = 1 To lastR
            If Range("Profile").Offset(xNum, 0).Value2 = ETWEETXLSETUP.UsernameBox.Value Then
            Range("Profile").Offset(xNum, 0).Value2 = vbNullString
            Range("F1").Offset(xNum, 0).Value2 = vbNullString
            Range("Browser").Offset(xNum, 0).Value = vbNullString
            End If
                Next
        
            '//Export data
            Call eTweetXL_POST.pstPersData
            
            '//Refresh
            Range("DataPullTrig").Value2 = "0"
            Call eTweetXL_GET.getProfileData
        
                End If
                
End Sub
Public Sub SaveLinkerBtn_Clk()

'//For creating a .link file (connection state)

'//Enter name for file
xName = InputBox("Enter a name for your link:", eTweetXL_INFO.AppName)

If xName <> "" Then

Dim X As Integer
X = 1

lastR = Cells(Rows.Count, "L").End(xlUp).Row

'//Format sheet
Call eTweetXL_TOOLS.shtFormat

'//Select file save location
xPath = Application.GetSaveAsFilename(xName, ".link, *.link")

Open xPath For Output As #1

Do Until X = lastR
DraftArr = Split(ETWEETXLPOST.LinkerBox.List(X - 1), ") "): xDraft = DraftArr(1)

Print #1, Range("MainLink").Offset(X, 0).Value & "," _
& Range("UserLink").Offset(X, 0).Value _
& Range("apiLink").Offset(X, 0).Value & "," _
& Range("ProfileLink").Offset(X, 0).Value & "," _
& xDraft & "," _
& Format(Range("Runtime").Offset(X, 0).Value, "hh:mm:ss")
X = X + 1
Loop

Close #1

If Not InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) Then
ETWEETXLPOST.xlFlowStrip.Value = "Link saved..."
End If

End If

End Sub
Public Sub SaveQueueBtn_Clk(ByVal xName As String, ByVal xPath As String)

'//For creating a .link file (queue state)

If xName & xPath = vbNullString Then

'//Enter name for file
xName = InputBox("Enter a name for your link:", eTweetXL_INFO.AppName)

    If xName = vbNullString Then Exit Sub

    '//Select file save location
    xPath = Application.GetSaveAsFilename(xName, ".link, *.link")

End If
    
Dim X As Integer

X = Range("LinkerCount").Value2

lastR = Cells(Rows.Count, "L").End(xlUp).Row

'//Format sheet
Call eTweetXL_TOOLS.shtFormat
    
Open xPath For Output As #1

Do Until X = lastR

Print #1, Range("MainLink").Offset(X, 0).Value & "," _
& Range("UserLink").Offset(X, 0).Value _
& Range("apiLink").Offset(X, 0).Value & "," _
& Range("ProfileLink").Offset(X, 0).Value & "," _
& Range("DraftLink").Offset(X, 0).Value & "," _
& Format(Range("Runtime").Offset(X, 0).Value, "hh:mm:ss")
X = X + 1
Loop

Close #1

If Not InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) Then
ETWEETXLPOST.xlFlowStrip.Value = "Link saved..."
End If


End Sub
Public Sub SaveBtn_Clk()

'//For saving API key preset information

Dim oBox, oForm As Object

Set oForm = ETWEETXLAPISETUP

'//API Setup

'//Save to .pers file...
Call eTweetXL_LOC.xApiFile(apiFile)

Open apiFile For Output As #2

Set oBox = oForm.apiKeyBox
If oBox <> "" Then '//check for empty box
Print #2, oBox.Value
    Else
        Close #2
        oBox.BorderColor = vbRed
        xMsg = 3: Call eTweetXL_MSG.AppMsg(xMsg, errLvl) '//if box is empty show error & don't save
        Exit Sub
            End If
       
Set oBox = oForm.apiSecretBox
If oBox <> "" Then
Print #2, oBox.Value
    Else
        Close #2
        oBox.BorderColor = vbRed
        xMsg = 3: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
        Exit Sub
            End If
            
Set oBox = oForm.accTokenBox
If oBox <> "" Then
Print #2, oBox.Value
    Else
        Close #2
        oBox.BorderColor = vbRed
        xMsg = 3: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
        Exit Sub
            End If

Set oBox = oForm.accSecretBox
If oBox <> "" Then
Print #2, oBox.Value
    Else
        Close #2
        oBox.BorderColor = vbRed
        xMsg = 3: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
        Exit Sub
            End If
            
Close #2

Set oBox = Nothing
Set oForm = Nothing

xMsg = 16: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)

End Sub
Public Sub AlignLink_Clk(ByVal xPos As Byte)

If xPos = 0 Then

    Range("AlignTrig").Value2 = 1
    
    If Range("LoadLess").Value2 <> 1 Then ETWEETXLPOST.AlignLink.Value = True
    
    ElseIf xPos = 1 Then
 
    Range("AlignTrig").Value2 = 0
    
    If Range("LoadLess").Value2 <> 1 Then ETWEETXLPOST.AlignLink.Value = False
    
    End If
                
End Sub
Public Sub SendAPI_Clk(ByVal xPos As Byte)

If xPos = 1 Then

    Range("SendAPI").Value2 = 1
    
    If Range("LoadLess").Value2 <> 1 Then ETWEETXLPOST.SendAPI.Value = True
    
    ElseIf xPos = 0 Then
 
    Range("SendAPI").Value2 = 0
    
    If Range("LoadLess").Value2 <> 1 Then ETWEETXLPOST.SendAPI.Value = False
    
    End If
                
End Sub
Public Sub SavePostBtn_Clk()

'//For saving a post (tweet)

Dim PostCharCt As Long

Call getWindow(xWin)

'//Get post character count...
PostCharCt = Len(xWin.PostBox.Value)
xWin.CharCt.Caption = PostCharCt

If PostCharCt <= 280 Then

If Range("xlasWinForm").Value2 = 13 Then xTwt = xWin.DraftBox.Value: Call eTweetXL_POST.pstDraftData(xTwt) Else _
xTwt = ETWEETXLPOST.DraftBox.Value: _
Call eTweetXL_POST.pstDraftData(xTwt) '//from queue

'//find current draft filter
If InStr(1, xTwt, " [...]") = False Then
If xWin.DraftFilterBtn.Caption <> "..." Or InStr(1, xTwt, " [•]") Then
xExt = ".twt": xT = " [•]": If xWin.DraftFilterBtn.Caption <> "•" Then xFil = 0: Call eTweetXL_CLICK.DraftFilterBtn_Clk(xFil)
xType = 0
    Else: GoTo asThread
            End If
        Else
asThread:
            xExt = ".thr": xT = " [...]"
            If xWin.DraftFilterBtn.Caption <> "..." Then xFil = 1: Call eTweetXL_CLICK.DraftFilterBtn_Clk(xFil)
            xType = 1
                End If
            
            
Call eTweetXL_GET.getPostData(xType)

If Range("xlasWinForm").Value2 = 13 Then xWin.DraftBox.Value = xTwt

If Range("xlasSilent") <> 1 Then MsgBox ("Post saved"), vbInformation, eTweetXL_INFO.AppName
    Else
        GoTo EndMacro
            End If

                Exit Sub
                
EndMacro:
xMsg = 21: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
                    
End Sub
Public Sub SetActive_Clk(ByVal xUser As String)

'//For setting a user active

Call eTweetXL_LOC.xPersFile(persFile)

If Dir(persFile) = "" Then GoTo ErrMsg
If xUser = "" Then GoTo ErrMsg

xUser = Replace(xUser, Range("Scure").Value, vbNullString) '//Remove lock symbol from name
If ETWEETXLSETUP.ProfileListBox.Value <> vbNullString Then '//Profile
Range("Profile").Value2 = ETWEETXLSETUP.ProfileListBox.Value: End If
'Range("Browser").Value2 = ETWEETXLSETUP.BrowserBox.Value
Range("Browser").Value2 = "Firefox" '//Browser
Range("ActiveUser").Value = xUser '//User
'ETWEETXLHOME.ActiveUser.Caption = xUser
'ETWEETXLHOME.ActiveUser.BackColor = vbWhite
'ETWEETXLSETUP.ActiveUser.Caption = xUser
'ETWEETXLSETUP.ActiveUser.BackColor = vbWhite
'ETWEETXLPOST.ActiveUser.Caption = xUser
'ETWEETXLPOST.ActiveUser.BackColor = vbWhite

X = 1
Do Until Range("Profile").Offset(X, 0).Value = xUser
X = X + 1
If X > 5000 Then GoTo SkipHere
Loop

SkipHere:
Range("ActiveTarget").Value = Range("User").Offset(X, 0).Value

'//Check for pass...
If Range("xlasWinForm").Value2 = 12 Then
If Range("ActiveTarget").Value = vbNullString Then If ETWEETXLSETUP.PassBox.Value <> vbNullString Then _
Range("ActiveTarget").Value = ETWEETXLSETUP.PassBox.Value Else GoTo ErrMsg
End If

Exit Sub


ErrMsg:
xMsg = 5: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)

End Sub
Public Sub ConnectBtn_Clk()

On Error GoTo ErrMsg

'//For connecting drafts (tweets) from the Linker to run

Dim AutoRuntimeArr(5000) As String: Dim DraftArr(5000) As String: Dim RuntimeArr(5000) As String
Dim xDraft As String: Dim xTime As String: Dim xUser As String: Dim NwTwtFile As String
Dim oDynOffsetBox As Object: Dim oDraftBox As Object: Dim oOffsetBox As Object: Dim oRuntimeBox As Object: Dim oUserBox As Object
Dim xPos As Long: Dim xTotal As Long

'//Check xlFlowStrip Window...
If Range("xlasWinForm").Value2 = 13 Then
Set oDynOffsetBox = ETWEETXLPOST.DynOffset
Set oDraftBox = ETWEETXLPOST.LinkerBox
Set oOffsetBox = ETWEETXLPOST.OffsetBox
Set oRuntimeBox = ETWEETXLPOST.RuntimeBox
Set oUserBox = ETWEETXLPOST.UserBox
End If

'//Check for draft-user match...
If oDraftBox.ListCount > 0 Then
If oUserBox.ListCount > 0 Then
If oDraftBox.ListCount = oUserBox.ListCount Then

'//Connecting...
If Range("xlasSilent").Value2 <> 1 Then
ETWEETXLPOST.xlFlowStrip.Value = "Connecting..."
End If

'//Reset count...
Range("LinkerCount").Value2 = 0

'//Set link & connection trigger(s)...
Range("LinkTrig").Value2 = 1
Range("ConnectStatus").Value2 = 1
Range("ConnectTrig").Value2 = 1

'//Check for dynamic offset...
If oDynOffsetBox.Value = True Then
Range("DynOffsetTrig").Value2 = 1
GoTo DynSetup
    Else
        Range("DynOffsetTrig").Value2 = 0
            End If

'//Convert offset to milliseconds...
OffsetArr = Split(oOffsetBox.Value, ":")
OffsetArrCopy = Split(oOffsetBox.Value, ":")
If CDbl(OffsetArr(0)) <> 0 Then OffsetArr(0) = (OffsetArr(0) * 3600 * 1000)
If CDbl(OffsetArr(1)) <> 0 Then OffsetArr(1) = (OffsetArr(1) * 60 * 1000)
If CDbl(OffsetArr(2)) <> 0 Then OffsetArr(2) = (OffsetArr(2) * 1000)
TotalOffset = CDbl(OffsetArr(0)) + CDbl(OffsetArr(1)) + CDbl(OffsetArr(2))
Range("Offset").Value2 = TotalOffset
Range("ActiveOffset").Value2 = TotalOffset

GoTo LinkerSetup
DynSetup:
'//Setup Dynamic Offset

'//Randomize number for offset...
Randomize
RndNum = CLng((100000 - 1 + 1) * Rnd + 1): RndNum2 = CLng((RndNum - 1 + 1) * Rnd + 1)
RndOffset = CLng((RndNum2 - 1 + 1) * Rnd + 1) * 10
'//
OffsetArr = Split(oOffsetBox.Value, ":")
OffsetArrCopy = Split(oOffsetBox.Value, ":")
If CDbl(OffsetArr(0)) <> 0 Then OffsetArr(0) = (OffsetArr(0) * 360 * 1000)
If CDbl(OffsetArr(1)) <> 0 Then OffsetArr(1) = (OffsetArr(1) * 60 * 1000)
If CDbl(OffsetArr(2)) <> 0 Then OffsetArr(2) = (OffsetArr(2) * 100)
'//
TotalOffset = CDbl(OffsetArr(0)) + CDbl(OffsetArr(1)) + CDbl(OffsetArr(2)) + RndOffset
Range("Offset").Value2 = TotalOffset
Range("ActiveOffset").Value2 = TotalOffset

LinkerSetup:

If Range("xlasSilent").Value2 <> 1 Then
'//Update connection state (xlFlowStrip)...
ETWEETXLPOST.xlFlowStrip.Value = "Connecting drafts..."
Art = "<lib>xbas;delayevent(3);$": Call xlas(Art)
    End If
        
'//Get drafts (tweets)...
xTotal = oDraftBox.ListCount
For xPos = 1 To (xTotal)
oDraftBox.ListIndex = xPos - 1
xDraft = oDraftBox.List(xPos - 1)
xDraft = Replace(xDraft, "(" & xPos & ") ", vbNullString)
DraftArr(xPos) = xDraft
If Range("xlasSilent").Value2 <> 1 Then
'//Update connection state (xlFlowStrip)...
ETWEETXLPOST.xlFlowStrip.Value = "Connecting drafts: " & (xPos * 100) / xTotal & "%..."
Art = "<lib>xbas;delayevent(1);$": Call xlas(Art)
    End If
        Next
        
'//Capture total links...
xTotal = xPos

Call eTweetXL_LOC.xTwtFile(twtFile)
Call eTweetXL_LOC.xThrFile(thrFile)
Range("Post").Value2 = twtFile & oDraftBox.Value & ".twt"
Range("ActiveTweet").Value2 = "" '//Clear active tweet range
xPos = 1

    Do Until DraftArr(xPos) = ""
    If InStr(1, DraftArr(xPos), " [•]") Then NwTwtFile = Replace(twtFile, ActiveProfile, Range("ProfileLink").Offset(xPos, 0).Value2): _
    DraftArr(xPos) = Replace(DraftArr(xPos), " [•]", vbNullString): DraftArr(0) = DraftArr(0) & NwTwtFile & DraftArr(xPos) & ".twt" & ","
    If InStr(1, DraftArr(xPos), " [...]") Then NwTwtFile = Replace(thrFile, ActiveProfile, Range("ProfileLink").Offset(xPos, 0).Value2): _
    DraftArr(xPos) = Replace(DraftArr(xPos), " [...]", vbNullString): DraftArr(0) = DraftArr(0) & NwTwtFile & DraftArr(xPos) & ".thr" & ","
    xPos = xPos + 1
    Loop
            
'//Set tweet(s)
Range("ActiveTweet").Value2 = DraftArr(0)
 
 If Range("xlasSilent").Value2 <> 1 Then
'//Update connection state (xlFlowStrip)...
ETWEETXLPOST.xlFlowStrip.Value = "Connecting times..."
Art = "<lib>xbas;delayevent(3);$": Call xlas(Art)
    End If
        
'//Setup runtime interval...
lastR = Cells(Rows.Count, "R").End(xlUp).Row

'//Clear runtime area...
Range("R1:R" & lastR).ClearContents

'//Check for auto offset...
If oRuntimeBox.ListCount = 1 Then
RuntimeArr(1) = oRuntimeBox.List(0)
If TotalOffset <> 0 Then GoTo AutoOffset
End If

'//Print runtime to range...
xTotal = oRuntimeBox.ListCount
For xPos = 1 To (xTotal)
oRuntimeBox.ListIndex = xPos - 1
xTime = oRuntimeBox.List(xPos - 1)
xTime = Replace(xTime, "(" & xPos & ") ", vbNullString)
xTime = Replace(xTime, " ", vbNullString)
RuntimeArr(xPos) = xTime
'//Set runtime(s)
Range("Runtime").Offset(xPos, 0).Value2 = xTime
If Range("xlasSilent").Value2 <> 1 Then
'//Update connection state (xlFlowStrip)...
ETWEETXLPOST.xlFlowStrip.Value = "Connecting times: " & (xPos * 100) / xTotal & "%..."
Art = "<lib>xbas;delayevent(1);$": Call xlas(Art)
    End If
        Next

GoTo SetupRtCntr

'//Automatically offset runtime if only one time linked, & an offset...
AutoOffset:

RuntimeArr(1) = Replace(RuntimeArr(1), "(1) ", vbNullString)

RuntimeArrCopy = Split(RuntimeArr(1), ":")

'//Set original runtime...
AutoRuntimeArr(1) = RuntimeArr(1)

For xPos = 2 To xTotal
If RuntimeArr(xPos) = "" Then
RuntimeArrCopy(0) = Int(RuntimeArrCopy(0)) + Int(OffsetArrCopy(0))
RuntimeArrCopy(1) = Int(RuntimeArrCopy(1)) + Int(OffsetArrCopy(1))
RuntimeArrCopy(2) = Int(RuntimeArrCopy(2)) + Int(OffsetArrCopy(2))
ThisRuntime = RuntimeArrCopy(0) & ":" & RuntimeArrCopy(1) & ":" & RuntimeArrCopy(2)
AutoRuntimeArr(xPos) = ThisRuntime
End If
If Range("xlasSilent").Value2 <> 1 Then
'//Update connection state (xlFlowStrip)...
ETWEETXLPOST.xlFlowStrip.Value = "Connecting times: " & (xPos * 100) / xTotal & "%..."
Art = "<lib>xbas;delayevent(1);$": Call xlas(Art)
    End If
        Next

For xPos = 1 To xTotal - 1
Range("Runtime").Offset(xPos, 0).Value2 = AutoRuntimeArr(xPos)
Next

SetupRtCntr:
'//Setup runtime counter...
If Range("Runtime").Offset(1, 0).Value2 <> "" Then
Range("RtCntr").Value = 0
    Else
        Range("RtCntr").Value2 = vbNullString
            End If


lastR = Cells(Rows.Count, "Q").End(xlUp).Row

If Range("xlasSilent").Value2 <> 1 Then
'//Update connection state (xlFlowStrip)...
ETWEETXLPOST.xlFlowStrip.Value = "Connecting users..."
Art = "<lib>xbas;delayevent(3);$": Call xlas(Art)
    End If
        
'//Clear UserLink Area...
Range("Q1:Q" & lastR).ClearContents

'//Print user to range from Linker...
xTotal = oUserBox.ListCount
For xPos = 1 To (xTotal)
oUserBox.ListIndex = xPos - 1
xUser = oUserBox.List(xPos - 1)
xUser = Replace(xUser, "(" & xPos & ") ", vbNullString)
Range("UserLink").Offset(xPos, 0).Value2 = xUser
If Range("xlasSilent").Value2 <> 1 Then
'//Update connection state (xlFlowStrip)...
ETWEETXLPOST.xlFlowStrip.Value = "Connecting users: " & (xPos * 100) / xTotal & "%..."
Art = "<lib>xbas;delayevent(1);$": Call xlas(Art)
    End If
        Next

'//
ETWEETXLPOST.ProfileListBox.Value = Range("MainLink").Offset(1, 0).Value

'//Save backup connection link to mtsett folder
Call eTweetXL_POST.pstLastLink

'//Connected (ready to send!)...
ETWEETXLPOST.xlFlowStrip.Value = "Finished connecting posts..."

Range("ConnectStatus").Value2 = 0

Set oDynOffsetBox = Nothing
Set oDraftBox = Nothing
Set oOffsetBox = Nothing
Set oDraftBox = Nothing
Set oRuntimeBox = Nothing
Set oUserBox = Nothing
Exit Sub
        
        Else: GoTo ErrMsg '//If no draft linked...
                        End If
                            End If
                                End If
                      
'//Error
ErrMsg:

Range("ConnectStatus").Value2 = 0

xMsg = 17: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)

End Sub
Public Sub ViewMedBtn_Clk()

'//For viewing post media

Call getWindow(xWin)

If xWin.MedLinkBox.Value <> "" Then

medLink = Range("MedScrollLink").Value

'//Check for spaces...
If InStr(1, medLink, " ") Then
    medLink = Replace(medLink, " ", """ """)
        End If

Call eTweetXL_LOC.xShellWinFldr(xShellWin)

Open xShellWin & "view_med.bat" For Output As #1
Print #1, "@echo off"
Print #1, "start " & medLink
Print #1, "exit"
Close #1

Shell (xShellWin & "view_med.bat"), vbMinimizedNoFocus
End If

End Sub
Public Sub RmvAllProfilesBtn_Clk()

Dim oFSO, oFldr, oSubFldr As Object

Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oFldr = oFSO.GetFolder(AppLoc & "\presets")

If Range("xlasSilent").Value2 = 1 Then msg = vbYes: GoTo SilentRun

msg = MsgBox("Are you sure you wish to remove all profiles?", vbYesNo, eTweetXL_INFO.AppName)

SilentRun:
If msg = vbYes Then

For Each oSubFldr In oFldr.SubFolders
xInfo = oSubFldr.name
Call eTweetXL_CLICK.RmvProfileBtn_Clk(xInfo)
Next

End If

End Sub
Public Sub RmvProfileBtn_Clk(ByVal xInfo As String)

If xInfo <> vbNullString Then
If Right(xInfo, 1) = """" Then xInfo = Left(xInfo, Len(xInfo) - 1) '//remove ending quote
If Left(xInfo, 1) = " " Then xInfo = Right(xInfo, Len(xInfo) - 1) '//remove leading space
ETWEETXLSETUP.ProfileNameBox.Value = xInfo
End If

If ETWEETXLSETUP.ProfileNameBox.Value <> "" Then

Dim xPos As Integer
nl = vbNewLine

'//position
xPos = ETWEETXLSETUP.ProfileListBox.ListIndex
'//
xUser = ETWEETXLSETUP.ProfileNameBox.Value

If Range("xlasSilent") = 1 Then YesNo = vbYes: GoTo SilentRun

    YesNo = MsgBox("[ " & xUser & " ]" & nl & nl & "Are you sure you want to remove this profile?", vbYesNo, eTweetXL_INFO.AppName)

SilentRun:
    If YesNo = vbYes Then
        
On Error Resume Next

Dim oFSO, oFile, oTwtFldr, oThrFldr, oPersFldr As Object
Set oFSO = CreateObject("Scripting.FileSystemObject")

If ETWEETXLSETUP.ProfileListBox.Value = "" Then ETWEETXLSETUP.ProfileListBox.Value = ETWEETXLSETUP.ProfileNameBox.Value

'//twt folder (single posts)
Set oTwtFldr = oFSO.GetFolder(AppLoc & "\presets\" & ETWEETXLSETUP.ProfileNameBox.Value & "\twt")

'//remove twt files (single posts)
For Each oFile In oTwtFldr.Files
Kill (oFile)
Next oFile

'//thr folder (threads)
Set oThrFldr = oFSO.GetFolder(AppLoc & "\presets\" & ETWEETXLSETUP.ProfileNameBox.Value & "\thr")

'//remove thr files (threads)
For Each oFile In oThrFldr.Files
Kill (oFile)
Next oFile

'//pers folder (personal)
Set oPersFldr = oFSO.GetFolder(AppLoc & "\presets\" & ETWEETXLSETUP.ProfileNameBox.Value & "\pers")

'//remove pers files (personal)
For Each oFile In oPersFldr.Files
Kill (oFile)
Next oFile

        RmDir (AppLoc & "\presets\" & ETWEETXLSETUP.ProfileNameBox.Value & "\pers")
        RmDir (AppLoc & "\presets\" & ETWEETXLSETUP.ProfileNameBox.Value & "\thr")
        RmDir (AppLoc & "\presets\" & ETWEETXLSETUP.ProfileNameBox.Value & "\twt")
        RmDir (AppLoc & "\presets\" & ETWEETXLSETUP.ProfileNameBox.Value)
      
'//Refresh
Range("DataPullTrig").Value2 = 0
'//Import profile information
Call eTweetXL_GET.getProfileData
ETWEETXLSETUP.ProfileListBox.Value = ETWEETXLSETUP.ProfileListBox.List(xPos)
                End If
                    End If '//yes
                    
                    
End Sub
Public Sub RmvUserBox_DelClk()

'//For removing a user from the Linker using the Delete key

On Error Resume Next

Dim AArr(5000), AArr2(5000), PArr(5000), PArr2(5000), UArr(5000), UArr2(5000), MArr(5000), MArr2(5000) As String
Dim xPos, I, X As Integer

If ETWEETXLPOST.UserBox.ListCount > 0 Then
ETWEETXLPOST.UserHdr.Caption = "User" & " (" & ETWEETXLPOST.UserBox.ListCount - 1 & ")"
End If

xPos = ETWEETXLPOST.UserBox.ListIndex

'//Get info...
X = 1
lastR = Cells(Rows.Count, "AM").End(xlUp).Row
Do Until X > lastR
If Range("apiLink").Offset(X, 0).Value <> vbNullString And _
Range("apiLink").Offset(X, 0).Value <> " " Then AArr(X) = Range("apiLink").Offset(X, 0).Value
If Range("MainLink").Offset(X, 0).Value <> vbNullString And _
Range("MainLink").Offset(X, 0).Value <> " " Then MArr(X) = Range("MainLink").Offset(X, 0).Value
If Range("TargetLink").Offset(X, 0).Value <> vbNullString And _
Range("TargetLink").Offset(X, 0).Value <> " " Then PArr(X) = Range("TargetLink").Offset(X, 0).Value
If Range("UserLink").Offset(X, 0).Value <> vbNullString And _
Range("Userlin").Offset(X, 0).Value <> " " Then UArr(X) = Range("UserLink").Offset(X, 0).Value
X = X + 1
Loop

AArr(X) = "*/HALT"
MArr(X) = "*/HALT"
PArr(X) = "*/HALT"
UArr(X) = "*/HALT"
If xPos >= 0 Then
    AArr(xPos + 1) = "*/SKIP"
    MArr(xPos + 1) = "*/SKIP"
    PArr(xPos + 1) = "*/SKIP"
    UArr(xPos + 1) = "*/SKIP"
        End If

I = 1
Do Until AArr(I) = "*/HALT"
    If AArr(I) = "*/SKIP" Then: xPos = I: AArr(I) = Replace(AArr(I), "*/SKIP", vbNullString): GoTo SkipHere1
        AArr2(I) = AArr(I)
SkipHere1:
    I = I + 1
    Loop
        AArr2(I) = "*/HALT"
    

I = 1
Do Until MArr(I) = "*/HALT"
    If MArr(I) = "*/SKIP" Then: xPos = I: MArr(I) = Replace(MArr(I), "*/SKIP", vbNullString): GoTo SkipHere2
        MArr2(I) = MArr(I)
SkipHere2:
    I = I + 1
    Loop
        MArr2(I) = "*/HALT"
    
I = 1
Do Until PArr(I) = "*/HALT"
    If PArr(I) = "*/SKIP" Then: xPos = I: PArr(I) = Replace(PArr(I), "*/SKIP", vbNullString): GoTo SkipHere3
        PArr2(I) = PArr(I)
SkipHere3:
    I = I + 1
    Loop
        PArr2(I) = "*/HALT"
    
I = 1
Do Until UArr(I) = "*/HALT"
    If UArr(I) = "*/SKIP" Then: xPos = I: UArr(I) = Replace(UArr(I), "*/SKIP", vbNullString): GoTo SkipHere4
        UArr2(I) = UArr(I)
SkipHere4:
    I = I + 1
    Loop
        UArr2(I) = "*/HALT"
    

Call clnUser

I = 1: X = 1
Do Until AArr(I) = "*/HALT"
If AArr2(I) <> vbNullString And AArr2(I) <> "*/HALT" Then Range("apiLink").Offset(X, 0).Value = AArr2(I): X = X + 1
I = I + 1
Loop

I = 1: X = 1
Do Until MArr(I) = "*/HALT"
If MArr2(I) <> vbNullString And MArr2(I) <> "*/HALT" Then Range("MainLink").Offset(X, 0).Value = MArr2(I): X = X + 1
I = I + 1
Loop

I = 1: X = 1
Do Until PArr(I) = "*/HALT"
If PArr2(I) <> vbNullString And PArr2(I) <> "*/HALT" Then Range("TargetLink").Offset(X, 0).Value = PArr2(I): X = X + 1
I = I + 1
Loop

I = 1: X = 1
Do Until UArr(I) = "*/HALT"
If UArr2(I) <> vbNullString And UArr2(I) <> "*/HALT" Then Range("UserLink").Offset(X, 0).Value = UArr2(I): X = X + 1
I = I + 1
Loop
                
If xPos = 0 Then xPos = 1

If Range("xlasSilent").Value2 <> 1 Then
ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.UserBox.List(xPos - 1) & " unlinked..."
End If

ETWEETXLPOST.UserBox.RemoveItem (xPos - 1)
xBox = 1: Call setBoxList(xBox)

End Sub
Public Sub RmvLinkerBox_DelClk()

'//For removing a draft from the Linker using the Delete key

On Error Resume Next

Dim PArr(5000), PArr2(5000), DArr(5000), DArr2(5000) As String
Dim xPos, I, X As Integer

If ETWEETXLPOST.LinkerBox.ListCount > 0 Then
ETWEETXLPOST.DraftHdr.Caption = "Draft" & " (" & ETWEETXLPOST.LinkerBox.ListCount - 1 & ")"
End If

xPos = ETWEETXLPOST.LinkerBox.ListIndex

X = 1
lastR = Cells(Rows.Count, "P").End(xlUp).Row
Do Until X > lastR
If Range("DraftLink").Offset(X, 0).Value <> vbNullString And _
Range("DraftLink").Offset(X, 0).Value <> " " Then DArr(X) = Range("Draftlink").Offset(X, 0).Value
If Range("ProfileLink").Offset(X, 0).Value <> vbNullString And _
Range("ProfileLink").Offset(X, 0).Value <> " " Then PArr(X) = Range("ProfileLink").Offset(X, 0).Value
X = X + 1
Loop

DArr(X) = "*/HALT"
PArr(X) = "*/HALT"
If xPos >= 0 Then
    DArr(xPos + 1) = "*/SKIP"
    PArr(xPos + 1) = "*/SKIP"
        End If
    
I = 1
Do Until DArr(I) = "*/HALT"
    If DArr(I) = "*/SKIP" Then: xPos = I: DArr(I) = Replace(DArr(I), "*/SKIP", vbNullString): GoTo SkipHere1
        DArr2(I) = DArr(I)
SkipHere1:
    I = I + 1
    Loop
        DArr2(I) = "*/HALT"
    

I = 1
Do Until PArr(I) = "*/HALT"
    If PArr(I) = "*/SKIP" Then: xPos = I: PArr(I) = Replace(PArr(I), "*/SKIP", vbNullString): GoTo SkipHere2
        PArr2(I) = PArr(I)
SkipHere2:
    I = I + 1
    Loop
        PArr2(I) = "*/HALT"
    

Call clnDraft

I = 1: X = 1
Do Until DArr(I) = "*/HALT"
If DArr2(I) <> vbNullString And DArr2(I) <> "*/HALT" Then Range("DraftLink").Offset(X, 0).Value = DArr2(I): X = X + 1
I = I + 1
Loop

I = 1: X = 1
Do Until PArr(I) = "*/HALT"
If PArr2(I) <> vbNullString And PArr2(I) <> "*/HALT" Then Range("ProfileLink").Offset(X, 0).Value = PArr2(I): X = X + 1
I = I + 1
Loop

If xPos = 0 Then xPos = 1

If Range("xlasSilent").Value2 <> 1 Then
ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.LinkerBox.List(xPos - 1) & " unlinked..."
End If

ETWEETXLPOST.LinkerBox.RemoveItem (xPos - 1)
xBox = 2: Call setBoxList(xBox)

End Sub
Public Sub RmvRuntimeBtn_DelClk()

'//For removing a time from the Linker using the Delete key

On Error Resume Next

Dim RArr(5000), RArr2(5000) As String
Dim xPos, I, X As Integer

If ETWEETXLPOST.RuntimeBox.ListCount > 0 Then
ETWEETXLPOST.RuntimeHdr.Caption = "Runtime" & " (" & ETWEETXLPOST.RuntimeBox.ListCount - 1 & ")"
End If

xPos = ETWEETXLPOST.RuntimeBox.ListIndex

X = 1
lastR = Cells(Rows.Count, "R").End(xlUp).Row: If X = lastR Then GoTo JustRmv '//runtime's are added after connection
Do Until X > lastR
If Range("Runtime").Offset(X, 0).Value <> vbNullString And _
Range("Runtime").Offset(X, 0).Value <> " " Then RArr(X) = Range("Runtime").Offset(X, 0).Value
X = X + 1
Loop


RArr(X) = "*/HALT"
If xPos >= 0 Then
    RArr(xPos + 1) = "*/SKIP"
        End If

I = 1
Do Until RArr(I) = "*/HALT"
    If RArr(I) = "*/SKIP" Then: xPos = I: RArr(I) = Replace(RArr(I), "*/SKIP", vbNullString): GoTo SkipHere1
        RArr2(I) = RArr(I)
SkipHere1:
    I = I + 1
    Loop
        RArr2(I) = "*/HALT"
    

Call clnRuntime

I = 1: X = 1
Do Until RArr(I) = "*/HALT"
If RArr2(I) <> vbNullString And RArr2(I) <> "*/HALT" Then Range("Runtime").Offset(X, 0).Value = RArr2(I): X = X + 1
I = I + 1
Loop

JustRmv:
If xPos = 0 Then xPos = 1

If Range("xlasSilent").Value2 <> 1 Then
ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.RuntimeBox.List(xPos - 1) & " unlinked..."
End If

ETWEETXLPOST.RuntimeBox.RemoveItem (xPos - 1)
xBox = 4: Call setBoxList(xBox)

End Sub
Public Sub UpHrBtn_Clk(ByVal xCount As Byte)

'//For adding hour to time box
If ETWEETXLPOST.TimeBox.Value = "" Then ETWEETXLPOST.TimeBox.Value = 0: ETWEETXLPOST.TimeBox.Value = ""

If xCount < 0 Then xCount = xCount * -1

For X = 1 To xCount

TimeHldr = ETWEETXLPOST.TimeBox.Value

TimeArr = Split(TimeHldr, ":")

If TimeArr(0) < 23 Then
TimeArr(0) = TimeArr(0) + 1
'//Add zero...
If Len(TimeArr(0)) < 2 Then TimeArr(0) = Str(0) + (TimeArr(0))
    Else
        TimeArr(0) = "00"
            End If
            
ETWEETXLPOST.TimeBox.Value = TimeArr(0) + ":" + TimeArr(1) + ":" + TimeArr(2)

Next

End Sub
Public Sub DwnHrBtn_Clk(ByVal xCount As Byte)

'//For subtracting hour from time box

If ETWEETXLPOST.TimeBox.Value = "" Then ETWEETXLPOST.TimeBox.Value = 0: ETWEETXLPOST.TimeBox.Value = ""

If xCount < 0 Then xCount = (xCount * -1)

For X = 1 To xCount

TimeHldr = ETWEETXLPOST.TimeBox.Value

TimeArr = Split(TimeHldr, ":")

If TimeArr(0) > 0 Then
TimeArr(0) = TimeArr(0) - 1
'//ADD ZERO
If Len(TimeArr(0)) < 2 Then TimeArr(0) = Str(0) + (TimeArr(0))
    Else
        TimeArr(0) = "24"
        If TimeArr(1) > 0 Then
        TimeArr(0) = TimeArr(0) - 1
            Else
                TimeArr(0) = "59"
                    End If
                        End If
            
            
ETWEETXLPOST.TimeBox.Value = TimeArr(0) + ":" + TimeArr(1) + ":" + TimeArr(2)

Next

End Sub
Public Sub UpMinBtn_Clk(ByVal xCount As Byte)

'//For adding minute to time box

If ETWEETXLPOST.TimeBox.Value = "" Then ETWEETXLPOST.TimeBox.Value = 0: ETWEETXLPOST.TimeBox.Value = ""

If xCount < 0 Then xCount = (xCount * -1)

For X = 1 To xCount

TimeHldr = ETWEETXLPOST.TimeBox.Value

TimeArr = Split(TimeHldr, ":")

If TimeArr(1) < 59 Then
TimeArr(1) = TimeArr(1) + 1
'//ADD ZERO
If Len(TimeArr(1)) < 2 Then TimeArr(1) = Str(0) + (TimeArr(1))
    Else
        TimeArr(1) = "00"
        If TimeArr(0) < 23 Then
        TimeArr(0) = TimeArr(0) + 1
            Else
                TimeArr(0) = "00"
                    End If
                        End If
            
            
ETWEETXLPOST.TimeBox.Value = TimeArr(0) + ":" + TimeArr(1) + ":" + TimeArr(2)

Next

End Sub
Public Sub DwnMinBtn_Clk(ByVal xCount As Byte)

'//For subtracting minute from time box

If ETWEETXLPOST.TimeBox.Value = "" Then ETWEETXLPOST.TimeBox.Value = 0: ETWEETXLPOST.TimeBox.Value = ""

If xCount < 0 Then xCount = (xCount * -1)

For X = 1 To xCount

TimeHldr = ETWEETXLPOST.TimeBox.Value

TimeArr = Split(TimeHldr, ":")

If TimeArr(1) > 0 Then
TimeArr(1) = TimeArr(1) - 1
'//ADD ZERO
If Len(TimeArr(1)) < 2 Then TimeArr(1) = Str(0) + (TimeArr(1))
    Else
        TimeArr(1) = "59"
        If TimeArr(0) > 0 Then
        TimeArr(0) = TimeArr(0) - 1
            Else
                TimeArr(0) = "23"
                    End If
                        End If
            
            
ETWEETXLPOST.TimeBox.Value = TimeArr(0) + ":" + TimeArr(1) + ":" + TimeArr(2)

Next

End Sub
Public Sub UpSecBtn_Clk(ByVal xCount As Byte)

'//For adding second to time box

If ETWEETXLPOST.TimeBox.Value = "" Then ETWEETXLPOST.TimeBox.Value = 0: ETWEETXLPOST.TimeBox.Value = ""

If xCount < 0 Then xCount = (xCount * -1)

For X = 1 To xCount

TimeHldr = ETWEETXLPOST.TimeBox.Value

TimeArr = Split(TimeHldr, ":")

If TimeArr(2) < 59 Then
TimeArr(2) = TimeArr(2) + 1
'//Add zero...
If Len(TimeArr(2)) < 2 Then TimeArr(2) = Str(0) + (TimeArr(2))
    Else
        TimeArr(2) = "00"
        If TimeArr(1) < 59 Then
        TimeArr(1) = TimeArr(1) + 1
            Else
                TimeArr(1) = "00"
                    End If
                        End If
            
            
ETWEETXLPOST.TimeBox.Value = TimeArr(0) + ":" + TimeArr(1) + ":" + TimeArr(2)

Next

End Sub
Public Sub DwnSecBtn_Clk(ByVal xCount As Byte)

'//For subtracting second from time box

If ETWEETXLPOST.TimeBox.Value = "" Then ETWEETXLPOST.TimeBox.Value = 0: ETWEETXLPOST.TimeBox.Value = ""

If xCount < 0 Then xCount = (xCount * -1)

For X = 1 To xCount

TimeHldr = ETWEETXLPOST.TimeBox.Value

TimeArr = Split(TimeHldr, ":")

If TimeArr(2) > 0 Then
TimeArr(2) = TimeArr(2) - 1
'//ADD ZERO
If Len(TimeArr(2)) < 2 Then TimeArr(2) = Str(0) + (TimeArr(2))
    Else
        TimeArr(2) = "59"
        If TimeArr(1) < 59 Then
        TimeArr(1) = TimeArr(1) - 1
            Else
                TimeArr(1) = "00"
                    End If
                        End If
            
            
ETWEETXLPOST.TimeBox.Value = TimeArr(0) + ":" + TimeArr(1) + ":" + TimeArr(2)

Next

End Sub
Public Sub DraftFilterBtn_Clk(ByVal xFil As Byte)

Retry:
Call getWindow(xWin)

On Error GoTo SetForm

If xFil = 1 Then
xType = 1
xWin.DraftFilterBtn.Caption = "..."
Range("DraftFilter").Value = 1
Call eTweetXL_GET.getPostData(xType)
    ElseIf xFil = 0 Then
    xType = 0
    Range("DraftFilter").Value = 0
    xWin.DraftFilterBtn.Caption = "•"
    Call eTweetXL_GET.getPostData(xType)
        End If
        
Exit Sub

SetForm:
Err.Clear
If Range("xlasWinForm").Value2 <> 13 Then Range("xlasWinForm").Value2 = 13 Else Range("xlasWinForm").Value2 = 14
GoTo Retry

End Sub
Public Sub FreezeBtn_Clk()

Call getWindow(xWin)

'//inactive = 0
If Range("AppState").Value2 <> 0 Then

'//freeze = 1 (xlas = 0)
If Range("AppState").Value2 = 1 Then
    Range("AppState").Value2 = 2
    Call fxsFreeze
    ETWEETXLHOME.xlFlowStrip.Value = "Application frozen..."
    ETWEETXLPOST.xlFlowStrip.Value = "Application frozen..."
    ETWEETXLQUEUE.xlFlowStrip.Value = "Application frozen..."
    ETWEETXLSETUP.xlFlowStrip.Value = "Application frozen..."
    '//unfreeze = 2 (xlas = 1)
        ElseIf Range("AppState").Value2 = 2 Then
            Range("AppState").Value2 = 1
            Call fxsUnfreeze
            ETWEETXLHOME.xlFlowStrip.Value = "Application unfrozen..."
            ETWEETXLPOST.xlFlowStrip.Value = "Application unfrozen..."
            ETWEETXLQUEUE.xlFlowStrip.Value = "Application unfrozen..."
            ETWEETXLSETUP.xlFlowStrip.Value = "Application unfrozen..."
                End If
                    End If
                
                      
End Sub
Public Sub HelpStatusBtn_Clk(ByVal xPos As Byte)

On Error GoTo ErrMsg

Call getWindow(xWin)

If xPos = 1 Then

'//help wizard on/active
    Range("HelpStatus").Value2 = 1
    xWin.HelpStatus.Caption = "On"
    If Range("xlasSilent").Value2 <> 1 Then
    xWin.xlFlowStrip.Value = "Help wizard is active..."
    End If
    
    ElseIf xPos = 0 Then
    
'//help wizard off/inactive
    Range("HelpStatus").Value2 = 0
    xWin.HelpStatus.Caption = "Off"
    If Range("xlasSilent").Value2 <> 1 Then
    xWin.xlFlowStrip.Value = "Help wizard is inactive..."
    End If
    
    End If

Exit Sub
ErrMsg:
xMsg = 28: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
End Sub
Public Sub LoadPostBtn_Clk(ByVal xName As String, ByVal xPath As String)

Dim YesNo As Integer: Dim YesNo2 As Integer: Dim X As Integer
Dim xFile As String: Dim xStr As String
Dim fd As FileDialog

Call getWindow(xWin)

'//xlas
If xPath <> vbNullString Then
ETWEETXLPOST.DraftBox.Value = xName
Open xPath For Input As #1
Do Until EOF(1)
Line Input #1, xStr
xWin.PostBox.Value = xWin.PostBox.Value & xStr & vbNewLine
    Loop
        Close #1
            Call eTweetXL_CLICK.SavePostBtn_Clk
            Art = "<lib>xtwt;clr(-post);$": Call xlas(Art)
            Exit Sub
                    End If
    
If Range("xlasSilent").Value2 <> 1 Then
YesNo = MsgBox("Would you like to automatically save the imported text as a draft(s)?", vbYesNo, eTweetXL_INFO.AppName)
End If

Set fd = Application.FileDialog(msoFileDialogFilePicker)
fd.AllowMultiSelect = True
xFile = fd.Show

    If xFile = "-1" Then
    
        For X = 1 To fd.SelectedItems.Count
        
        If YesNo = vbYes Then
    
        YesNo2 = MsgBox("Would you like to enter a name for your draft?", vbYesNo, eTweetXL_INFO.AppName)
    
        If YesNo2 = vbYes Then
        xName = InputBox("Enter a name for your draft:", eTweetXL_INFO.AppName)
        ETWEETXLPOST.DraftBox.Value = xName
            End If
                End If
            
        xFile = fd.SelectedItems(X)
    
        Open xFile For Input As #1
        Do Until EOF(1)
        Line Input #1, xStr
        xWin.PostBox.Value = xWin.PostBox.Value & xStr & vbNewLine
        Loop
        Close #1
            
             If YesNo = vbYes Or Range("xlasSilent").Value2 = 1 Then
             Call eTweetXL_CLICK.SavePostBtn_Clk
             Art = "<lib>xtwt;clr(-post);$": Call xlas(Art)
             End If
                
                    Next
                    
                        End If
                        
End Sub
Public Sub SplitPostBtn_Clk()

Dim xSplitStr As String: Dim xStr As String
Dim xSplitArr() As String
Dim xDiv As Long: Dim xLeftover As Long: Dim xStrLen As Long: Dim xStrLenH As Long
Dim xMod As Integer
xMod = 280
xStr = ETWEETXLPOST.PostBox.Value

If Len(xStr) > 280 Then
    
    '//get character count
    xStrLen = Len(xStr)
    '//find amount to split string by based on 280 character limit
    xDiv = xStrLen / 280
    '//check for remaining ungrouped text
    xLeftover = xStrLen - (xDiv * 280)
    
xStrLenH = xStrLen

Do Until xStrLenH = 0

'//get characters from string
xSplitStr = Left(xStr, xMod)

'//get string position
xSplitArr = Split(xStr, xSplitStr): If UBound(xSplitArr) = 1 Then xStr = xSplitArr(1)
    
'//put split string into post box
ETWEETXLPOST.PostBox.Value = xSplitStr
'//add current post to thread
Call eTweetXL_CLICK.AddThreadBtn_Clk

xStrLenH = xStrLenH - 280
xDiv = xDiv - 1
If xDiv < 0 Then If xLeftover <> 0 Then xMod = xMod - xLeftover: xLeftover = 0 Else xStrLenH = 0
    Loop
        End If
            
End Sub
Public Sub TrimPostBtn_Clk()

Dim xStr As String

xStr = ETWEETXLPOST.PostBox.Value

If Len(xStr) > 280 Then
    xStr = Left(xStr, 280)
        ETWEETXLPOST.PostBox.Value = xStr
            End If
            
End Sub
Public Function ReflectBtn_Clk()

If Range("ReflectTrig").Value2 = 0 Then

Range("ReflectTrig").Value2 = 1
ETWEETXLPOST_EX.ReflectBtn.ForeColor = vbGreen

If Range("AppendTrig").Value = 1 Then
    ETWEETXLPOST.PostBox.Value = ETWEETXLPOST.PostBox.Value & vbNewLine & ETWEETXLPOST_EX.PostBox.Value
        Else
            ETWEETXLPOST.PostBox.Value = ETWEETXLPOST_EX.PostBox.Value
                End If

Else

    Range("ReflectTrig").Value2 = 0: ETWEETXLPOST_EX.ReflectBtn.ForeColor = &H80000011
        
        End If
        
End Function
Public Function AddSizeHBtn_Clk()

ETWEETXLPOST_EX.Width = ETWEETXLPOST_EX.Width + 50
ETWEETXLPOST_EX.PostBox.Width = ETWEETXLPOST_EX.PostBox.Width + 50

'//adjust button locations
ETWEETXLPOST_EX.AppendBtn.Left = ETWEETXLPOST_EX.AppendBtn.Left + 50
ETWEETXLPOST_EX.LoadPostBtn.Left = ETWEETXLPOST_EX.LoadPostBtn.Left + 50
ETWEETXLPOST_EX.ReflectBtn.Left = ETWEETXLPOST_EX.ReflectBtn.Left + 50
ETWEETXLPOST_EX.AddSizeHBtn.Left = ETWEETXLPOST_EX.AddSizeHBtn.Left + 50
ETWEETXLPOST_EX.AddSizeVBtn.Left = ETWEETXLPOST_EX.AddSizeVBtn.Left + 50
ETWEETXLPOST_EX.RmvSizeHBtn.Left = ETWEETXLPOST_EX.RmvSizeHBtn.Left + 50
ETWEETXLPOST_EX.RmvSizeVBtn.Left = ETWEETXLPOST_EX.RmvSizeVBtn.Left + 50
ETWEETXLPOST_EX.AddThreadBtn.Left = ETWEETXLPOST_EX.AddThreadBtn.Left + 50
ETWEETXLPOST_EX.RmvThreadBtn.Left = ETWEETXLPOST_EX.RmvThreadBtn.Left + 50
ETWEETXLPOST_EX.RmvAllThreadBtn.Left = ETWEETXLPOST_EX.RmvAllThreadBtn.Left + 50
ETWEETXLPOST_EX.SplitPostBtn.Left = ETWEETXLPOST_EX.SplitPostBtn.Left + 50
ETWEETXLPOST_EX.TrimPostBtn.Left = ETWEETXLPOST_EX.TrimPostBtn.Left + 50
ETWEETXLPOST_EX.SavePostBtn.Left = ETWEETXLPOST_EX.SavePostBtn.Left + 50
ETWEETXLPOST_EX.ExitBtn.Left = ETWEETXLPOST_EX.ExitBtn.Left + 50
ETWEETXLPOST_EX.CharCt.Left = ETWEETXLPOST_EX.CharCt.Left + 50
ETWEETXLPOST_EX.DraftFilterBtn.Left = ETWEETXLPOST_EX.DraftFilterBtn.Left + 50
ETWEETXLPOST_EX.xlFlowStrip.Left = ETWEETXLPOST_EX.xlFlowStrip.Left + 50
ETWEETXLPOST_EX.ThreadCt.Left = ETWEETXLPOST_EX.ThreadCt.Left + 50

End Function
Public Function AddSizeVBtn_Clk()

ETWEETXLPOST_EX.Height = ETWEETXLPOST_EX.Height + 50
ETWEETXLPOST_EX.PostBox.Height = ETWEETXLPOST_EX.PostBox.Height + 50

End Function
Public Function RmvSizeHBtn_Clk()

If ETWEETXLPOST_EX.Width > 200 Then
ETWEETXLPOST_EX.Width = ETWEETXLPOST_EX.Width - 50
ETWEETXLPOST_EX.PostBox.Width = ETWEETXLPOST_EX.PostBox.Width - 50

'//adjust button locations
ETWEETXLPOST_EX.AppendBtn.Left = ETWEETXLPOST_EX.AppendBtn.Left - 50
ETWEETXLPOST_EX.LoadPostBtn.Left = ETWEETXLPOST_EX.LoadPostBtn.Left - 50
ETWEETXLPOST_EX.ReflectBtn.Left = ETWEETXLPOST_EX.ReflectBtn.Left - 50
ETWEETXLPOST_EX.AddSizeHBtn.Left = ETWEETXLPOST_EX.AddSizeHBtn.Left - 50
ETWEETXLPOST_EX.AddSizeVBtn.Left = ETWEETXLPOST_EX.AddSizeVBtn.Left - 50
ETWEETXLPOST_EX.RmvSizeHBtn.Left = ETWEETXLPOST_EX.RmvSizeHBtn.Left - 50
ETWEETXLPOST_EX.RmvSizeVBtn.Left = ETWEETXLPOST_EX.RmvSizeVBtn.Left - 50
ETWEETXLPOST_EX.AddThreadBtn.Left = ETWEETXLPOST_EX.AddThreadBtn.Left - 50
ETWEETXLPOST_EX.RmvThreadBtn.Left = ETWEETXLPOST_EX.RmvThreadBtn.Left - 50
ETWEETXLPOST_EX.RmvAllThreadBtn.Left = ETWEETXLPOST_EX.RmvAllThreadBtn.Left - 50
ETWEETXLPOST_EX.SplitPostBtn.Left = ETWEETXLPOST_EX.SplitPostBtn.Left - 50
ETWEETXLPOST_EX.TrimPostBtn.Left = ETWEETXLPOST_EX.TrimPostBtn.Left - 50
ETWEETXLPOST_EX.SavePostBtn.Left = ETWEETXLPOST_EX.SavePostBtn.Left - 50
ETWEETXLPOST_EX.ExitBtn.Left = ETWEETXLPOST_EX.ExitBtn.Left - 50
ETWEETXLPOST_EX.CharCt.Left = ETWEETXLPOST_EX.CharCt.Left - 50
ETWEETXLPOST_EX.DraftFilterBtn.Left = ETWEETXLPOST_EX.DraftFilterBtn.Left - 50
ETWEETXLPOST_EX.xlFlowStrip.Left = ETWEETXLPOST_EX.xlFlowStrip.Left - 50
ETWEETXLPOST_EX.ThreadCt.Left = ETWEETXLPOST_EX.ThreadCt.Left - 50
End If

End Function
Public Function RmvSizeVBtn_Clk()

If ETWEETXLPOST_EX.Height > 414 Then
ETWEETXLPOST_EX.Height = ETWEETXLPOST_EX.Height - 50
ETWEETXLPOST_EX.PostBox.Height = ETWEETXLPOST_EX.PostBox.Height - 50
End If

End Function

Public Function UserOpt_Clk(xType)

Range("EditStatus").Value2 = 4

If xType = 1 Then

If ETWEETXLPOST.SwapUser.ForeColor <> vbGreen Then
    Range("UserTrig").Value2 = 1
     ETWEETXLPOST.SwapUser.ForeColor = vbGreen
     ETWEETXLPOST.AddUserA.ForeColor = vbBlack
     ETWEETXLPOST.AddUserB.ForeColor = vbBlack
        Else
            Range("UserTrig").Value2 = 0
            Range("EditStatus").Value2 = 0
            ETWEETXLPOST.SwapUser.ForeColor = vbBlack
                End If
                
ElseIf xType = 2 Then

If ETWEETXLPOST.AddUserA.ForeColor <> vbGreen Then
    Range("UserTrig").Value2 = 2
     ETWEETXLPOST.AddUserA.ForeColor = vbGreen
     ETWEETXLPOST.SwapUser.ForeColor = vbBlack
     ETWEETXLPOST.AddUserB.ForeColor = vbBlack
        Else
            Range("UserTrig").Value2 = 0
            Range("EditStatus").Value2 = 0
            ETWEETXLPOST.AddUserA.ForeColor = vbBlack
                End If
                
ElseIf xType = 3 Then

If ETWEETXLPOST.AddUserB.ForeColor <> vbGreen Then
    Range("UserTrig").Value2 = 3
     ETWEETXLPOST.AddUserB.ForeColor = vbGreen
     ETWEETXLPOST.SwapUser.ForeColor = vbBlack
     ETWEETXLPOST.AddUserA.ForeColor = vbBlack
        Else
            Range("UserTrig").Value2 = 0
            Range("EditStatus").Value2 = 0
            ETWEETXLPOST.AddUserB.ForeColor = vbBlack
                End If
                
                    End If

End Function
Public Function DraftOpt_Clk(xType)

Range("EditStatus").Value2 = 4

If xType = 1 Then

If ETWEETXLPOST.SwapDraft.ForeColor <> vbGreen Then
    Range("DraftTrig").Value2 = 1
     ETWEETXLPOST.SwapDraft.ForeColor = vbGreen
     ETWEETXLPOST.AddDraftA.ForeColor = vbBlack
     ETWEETXLPOST.AddDraftB.ForeColor = vbBlack
        Else
            Range("DraftTrig").Value2 = 0
            Range("EditStatus").Value2 = 0
            ETWEETXLPOST.SwapDraft.ForeColor = vbBlack
                End If
                
ElseIf xType = 2 Then

If ETWEETXLPOST.AddDraftA.ForeColor <> vbGreen Then
    Range("DraftTrig").Value2 = 2
     ETWEETXLPOST.AddDraftA.ForeColor = vbGreen
     ETWEETXLPOST.SwapDraft.ForeColor = vbBlack
     ETWEETXLPOST.AddDraftB.ForeColor = vbBlack
        Else
            Range("DraftTrig").Value2 = 0
            Range("EditStatus").Value2 = 0
            ETWEETXLPOST.AddDraftA.ForeColor = vbBlack
                End If
                
ElseIf xType = 3 Then

If ETWEETXLPOST.AddDraftB.ForeColor <> vbGreen Then
    Range("DraftTrig").Value2 = 3
     ETWEETXLPOST.AddDraftB.ForeColor = vbGreen
     ETWEETXLPOST.SwapDraft.ForeColor = vbBlack
     ETWEETXLPOST.AddDraftA.ForeColor = vbBlack
        Else
            Range("DraftTrig").Value2 = 0
            Range("EditStatus").Value2 = 0
            ETWEETXLPOST.AddDraftB.ForeColor = vbBlack
                End If
                
                    End If

End Function
Public Function RuntimeOpt_Clk(xType)

Range("EditStatus").Value2 = 4

If xType = 1 Then

If ETWEETXLPOST.SwapTime.ForeColor <> vbGreen Then
    Range("TimeTrig").Value2 = 1
     ETWEETXLPOST.SwapTime.ForeColor = vbGreen
     ETWEETXLPOST.AddTimeA.ForeColor = vbBlack
     ETWEETXLPOST.AddTimeB.ForeColor = vbBlack
        Else
            Range("TimeTrig").Value2 = 0
            Range("EditStatus").Value2 = 0
            ETWEETXLPOST.SwapTime.ForeColor = vbBlack
                End If
                
ElseIf xType = 2 Then

If ETWEETXLPOST.AddTimeA.ForeColor <> vbGreen Then
    Range("TimeTrig").Value2 = 2
     ETWEETXLPOST.AddTimeA.ForeColor = vbGreen
     ETWEETXLPOST.SwapTime.ForeColor = vbBlack
     ETWEETXLPOST.AddTimeB.ForeColor = vbBlack
        Else
            Range("TimeTrig").Value2 = 0
            Range("EditStatus").Value2 = 0
            ETWEETXLPOST.AddTimeA.ForeColor = vbBlack
                End If
                
ElseIf xType = 3 Then

If ETWEETXLPOST.AddTimeB.ForeColor <> vbGreen Then
    Range("TimeTrig").Value2 = 3
     ETWEETXLPOST.AddTimeB.ForeColor = vbGreen
     ETWEETXLPOST.SwapTime.ForeColor = vbBlack
     ETWEETXLPOST.AddTimeA.ForeColor = vbBlack
        Else
            Range("TimeTrig").Value2 = 0
            Range("EditStatus").Value2 = 0
            ETWEETXLPOST.AddTimeB.ForeColor = vbBlack
                End If
                
                    End If

End Function

