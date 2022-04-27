Attribute VB_Name = "App_CLICK"
'/###########################\
'//Application Click Features\\
'///#########################\\\

Public Sub AddThread_Clk()

Dim X, xVal As String

Retry:
Call App_TOOLS.FindForm(xForm)

On Error GoTo SetForm

lastRw = Cells(Rows.Count, "Z").End(xlUp).Row

If xForm.ThreadCt.Caption <> "" Then
If CInt(xForm.ThreadCt.Caption) > 0 Then
lastRw = xForm.ThreadCt.Caption
GoTo NextStep
End If
    End If

NextStep:
'//print new post for thread to range
Range("PostThread").Offset(lastRw, 0).Value = xForm.PostBox.Value

'//print new media for thread to range
If xForm.MedLinkBox.Value <> vbNullString Then _
Range("MedThread").Offset(lastRw, 0).Value = """" & xForm.MedLinkBox.Value & """"

'//relay last thread
If Range("PostThread").Offset(lastRw + 1, 0).Value <> vbNullString Then xForm.PostBox.Value = Range("PostThread").Offset(lastRw + 1, 0).Value _
Else: xForm.PostBox.Value = vbNullString
If Range("MedThread").Offset(lastRw + 1, 0).Value <> vbNullString Then xForm.MedLinkBox.Value = Range("MedThread").Offset(lastRw + 1, 0).Value _
Else: xForm.MedLinkBox.Value = vbNullString

'//set total thread count
X = xForm.ThreadCt.Caption
If X = vbNullString Then X = 0 Else X = xForm.ThreadCt.Caption
If CInt(X) <= lastRw Then xForm.ThreadCt.Caption = lastRw + 1

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
If Range("xlasWinForm").Value <> 3 Then Range("xlasWinForm").Value = 3 Else Range("xlasWinForm").Value = 4
GoTo Retry

End Sub
Public Sub RmvThread_Clk()

Call App_TOOLS.FindForm(xForm)

lastRw = Cells(Rows.Count, "Z").End(xlUp).Row

If xForm.ThreadCt.Caption <> "" Then
If CInt(xForm.ThreadCt.Caption) > 0 Then
lastRw = xForm.ThreadCt.Caption
GoTo NextStep
End If
    End If
    
NextStep:
'//remove post from thread loc
Range("PostThread").Offset(lastRw, 0).Value = vbNullString

'//remove media from thread loc
Range("MedThread").Offset(lastRw, 0).Value = vbNullString

'//decrement thread count
xForm.ThreadCt.Caption = lastRw - 1

'//show previous post
xForm.PostBox.Value = Range("PostThread").Offset(lastRw - 1, 0).Value

End Sub
Public Sub RmvAllThread_Clk()

Retry:
Call App_TOOLS.FindForm(xForm)

On Error GoTo SetForm

lastRw = Cells(Rows.Count, "Y").End(xlUp).Row
If lastRw < 1 Then lastRw = Cells(Rows.Count, "Z").End(xlUp).Row

For X = 0 To lastRw
Range("PostThread").Offset(X, 0).Value = vbNullString
Range("MedThread").Offset(X, 0).Value = vbNullString
Next

xForm.ThreadCt.Caption = vbNullString

If InStr(1, xForm.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then _
xForm.xlFlowStrip.Value = "All threads removed..."
        
Exit Sub

SetForm:
Err.Clear
If Range("xlasWinForm").Value <> 3 Then Range("xlasWinForm").Value = 3 Else Range("xlasWinForm").Value = 4
GoTo Retry

End Sub
Public Sub ClrSetup_Clk()

If Range("AppActive").Value <> 1 Then

'//Cleanup
Call Cleanup.ClnMainSpace
Call Cleanup.ClnLatchSpace
Call Cleanup.ClnLinkerSpace
Call Cleanup.ClnRuntimeSpace
Call Cleanup.ClnSpecSpace
Call App_TOOLS.DataKillSwitch
Range("ConnectTrig").Value = 0
Range("LinkTrig").Value = 0
Range("User").Value = vbNullString
ETWEETXLPOST.xlFlowStrip.Enabled = True
ETWEETXLPOST.SendAPI.Value = False
ETWEETXLPOST.ActivePresetBox.Caption = vbNullString
ETWEETXLPOST.UserBox.Clear
ETWEETXLPOST.LinkerBox.Clear
ETWEETXLPOST.RuntimeBox.Clear
ETWEETXLPOST.PostBox.Value = vbNullString
ETWEETXLPOST.ProfileListBox.Value = vbNullString
ETWEETXLPOST.UserListBox.Value = vbNullString
ETWEETXLPOST.DraftBox.Value = vbNullString
Call App_CLICK.RmvAllThread_Clk
ETWEETXLPOST.UserHdr.Caption = "User"
ETWEETXLPOST.DraftHdr.Caption = "Draft"
ETWEETXLPOST.RuntimeHdr.Caption = "Runtime"

    If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then _
    ETWEETXLHOME.xlFlowStrip.Value = "Cleaned..."
    ETWEETXLPOST.xlFlowStrip.Value = "Cleaned..."
    ETWEETXLQUEUE.xlFlowStrip.Value = "Cleaned..."
    ETWEETXLSETUP.xlFlowStrip.Value = "Cleaned..."

    Else
    
    xMsg = 25: App_MSG.AppMsg (xMsg)
    
    End If

End Sub
Public Sub DeleteAllDraft_Clk()

Call FindForm(xForm)

If xForm.DraftFilterBtn.Caption <> "..." Then
xExt = ".twt": xT = " [•]": If xForm.DraftFilterBtn.Caption <> "•" Then xFil = 0: Call App_CLICK.DraftFilterBtn_Clk(xFil)
Call App_Loc.xTwtFile(twtFile): xLoc = twtFile
xStr = "Are you sure you wish to delete all single draft posts for '" & xProfile & "'?"
        Else
            xExt = ".thr": xT = " [...]"
            If xForm.DraftFilterBtn.Caption <> "..." Then xFil = 1: Call App_CLICK.DraftFilterBtn_Clk(xFil)
            Call App_Loc.xThrFile(thrFile): xLoc = thrFile
            xStr = "Are you sure you wish to delete all threaded draft posts for '" & xProfile & "'?"
                End If
                
                
        '//Remove all drafts from a profile
        If xProfile <> "" Then
        
        If Dir(xLoc) <> "" Then
        
        If Range("xlasSilent").Value = 1 Then msg = vbYes: GoTo SilentRun
        
        msg = MsgBox(xStr, vbYesNo, AppTag)
        
SilentRun:
            If msg = vbYes Then
            
            Dim oFSO, oFile, oFldr As Object
            Set oFSO = CreateObject("Scripting.FileSystemObject")
            Set oFldr = oFSO.GetFolder(xLoc)
            
            For Each oFile In oFldr.Files
            Kill (oFile)
            Next
            
            If Range("DraftFilter").Value = 1 Then xType = 0 Else xType = 1
            xType = 0: Call App_IMPORT.MyTweetData(xType)
            
            End If
                End If
                    End If
                    
End Sub
Public Sub RuntimeBox_Clk()

'//For editing the time in Runtime boxes

'//Check for runtime change...(Avoid double run)
If Range("RtChange").Value = 1 Then
Range("RtChange").Value = 0
Exit Sub
End If

Dim oRuntimeBox As Object

lastRw = Cells(Rows.Count, "R").End(xlUp).Row
LLCntr = Range("LinkerCount").Value

'//Check xlFlowStrip Window...
If Range("xlasWinForm").Value = 3 Then
Set oRuntimeBox = ETWEETXLPOST.RuntimeBox
GoTo XLPOST
End If

If Range("xlasWinForm").Value = 4 Then
Set oRuntimeBox = ETWEETXLQUEUE.RuntimeBox
GoTo XLQUEUE
End If

'//Change runtime in Queue
XLQUEUE:

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

On Error GoTo ErrMsg

NewRt = InputBox("Enter a new runtime:", AppTag, RtHldr)

    If NewRt <> vbNullString Then
    
    '//Check for time format...
    If InStr(1, NewRt, ":") = False Then GoTo ErrMsg
    
        '//Convert to time...
        ThisRt = Format$(ThisRt, "hh:mm:ss")

        '//Record...
        Range("RtChange").Value = 1
        LLCntr = LLCntr - 2: If LLCntr < 0 Then LLCntr = 0
        Range("Runtime").Offset(xPos + (2 + (LLCntr)), 0).Value = NewRt
        NewRt = "(" & xPos + 1 & ") " & NewRt
        oRuntimeBox.List(xPos) = NewRt
            End If
                End If


'//Export to file...
Call App_Loc.xMTRuntime(mtRuntime)

Open mtRuntime For Output As #1
For X = 1 To lastRw
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
NewRt = InputBox("Enter a new runtime:", AppTag, RtHldr)

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
xMsg = 8: Call App_MSG.AppMsg(xMsg)

Set oRuntimeBox = Nothing

End Sub
Public Sub Break_Clk()

'//For breaking/stopping the application incase of an issue during a run, or to refresh

TK = CreateObject("WScript.Shell").Run("taskkill /f /im " & "powershell.exe", 0, True)


'//Cleanup
Call ClnMainSpace
Call ClnLatchSpace
Call ClnLinkerSpace
Call ClnRuntimeSpace
Call ClnSpecSpace
Call DataKillSwitch
Range("AppActive").Value = 0
Range("ConnectTrig").Value = 0
Range("LinkTrig").Value = 0
Range("User").Value = vbNullString
ETWEETXLHOME.xlFlowStrip.Enabled = True
ETWEETXLPOST.xlFlowStrip.Enabled = True
ETWEETXLQUEUE.xlFlowStrip.Enabled = True
ETWEETXLSETUP.xlFlowStrip.Enabled = True
ETWEETXLHOME.ProgRatio = vbNullString
ETWEETXLHOME.ActivePresetBox.Caption = vbNullString
ETWEETXLHOME.ProgBar.Width = 0
ETWEETXLPOST.SendAPI.Value = False
ETWEETXLPOST.ActivePresetBox.Caption = vbNullString
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
ETWEETXLQUEUE.ActivePresetBox.Caption = vbNullString
ETWEETXLQUEUE.QueueBox.Clear
ETWEETXLQUEUE.RuntimeBox.Clear
ETWEETXLQUEUE.UserBox.Clear
ETWEETXLSETUP.ActivePresetBox.Caption = vbNullString
Set TK = Nothing

ETWEETXLHOME.LinkerActive.Caption = "OFF"
ETWEETXLHOME.LinkerActive.ForeColor = vbRed
ETWEETXLHOME.LinkerActive.BackColor = -2147483633

Call enableFlowStrip
Call NoFreezeFX

ETWEETXLHOME.xlFlowStrip.Value = "Break complete..."
ETWEETXLSETUP.xlFlowStrip.Value = "Break complete..."
ETWEETXLPOST.xlFlowStrip.Value = "Break complete..."
ETWEETXLQUEUE.xlFlowStrip.Value = "Break complete..."

xMsg = 10: Call App_MSG.AppMsg(xMsg)

If Range("xlasWinForm").Value = 1 Then ETWEETXLHOME.Hide
If Range("xlasWinForm").Value = 1 Then ETWEETXLHOME.Show

End Sub
Sub Start_Auto()

Call Start_Clk

End Sub
Sub Start_Clk()

On Error GoTo ErrMsg

'//check for pause or disable
If Range("AppActive").Value = 2 Then xMsg = 26: GoTo EndMacro

'//For starting the application automation
Call App_TOOLS.xDisable

'//Check for set user...
If Range("ActiveUser").Value = "" Then
xMsg = 9: Call App_MSG.AppMsg(xMsg)
Exit Sub
End If

'//Check Mainlink...
If Range("Mainlink").Offset(1, 0).Value = "" Then xMsg = 21: GoTo EndMacro

ETWEETXLHOME.xlFlowStrip.Value = "Checking for user information..."
ETWEETXLSETUP.xlFlowStrip.Value = "Checking for user information..."
ETWEETXLPOST.xlFlowStrip.Value = "Checking for user information..."
ETWEETXLQUEUE.xlFlowStrip.Value = "Checking for user information..."

Call disableFlowStrip

Call App_Loc.xMTapi(mtApi)
Call App_Loc.xMTBlank(mtBlank)
Call App_Loc.xMTPass(mtPass)
Call App_Loc.xMTTwt(mtTwt)
Call App_Loc.xMTCheck(mtCheck)
Call App_Loc.xMTUser(mtUser)
Call App_Loc.xMTOffset(mtOffset)
Call App_Loc.xMTOffsetCopy(mtOffsetCopy)
Call App_Loc.xMTRuntime(mtRuntime)
Call App_Loc.xMTRuntimeCntr(mtRuntimeCntr)
Call App_Loc.xMTDynOff(mtDynOff)
Call App_Loc.xMTMed(mtMed)
Call App_Loc.xMTPost(mtPost)
Call App_Loc.xMTProf(mtProf)
Call App_Loc.xMTRetryCntr(mtRetryCntr)
Call App_Loc.xApp_StartLink(appStartLink)
    
'//Clear...
If Range("LinkerActive") = 1 Then GoTo SkipClear '//Skip if linkline active
Range("LinkerCount").Value = 0
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
lastRw = Cells(Rows.Count, "R").End(xlUp).Row
rtTrig = 0

ETWEETXLHOME.xlFlowStrip.Value = "Calculating offset..."
ETWEETXLSETUP.xlFlowStrip.Value = "Calculating offset..."
ETWEETXLPOST.xlFlowStrip.Value = "Calculating offset..."
ETWEETXLQUEUE.xlFlowStrip.Value = "Calculating offset..."

Open mtRuntime For Output As #5
For xNum = 1 To lastRw

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
If Range("DynOffTrig").Value = 1 Then

Open mtDynOff For Output As #7
Print #7, ""
Close #7
    Else
        If Dir(mtDynOff) <> "" Then
        Kill (mtDynOff)
            End If
        
'//Refresh offset if not dynamic
If Range("DynOffTrig").Value <> 1 Then Range("ActiveOffset").Value = vbNullString

                End If

ETWEETXLHOME.xlFlowStrip.Value = "Checking Linker..."
ETWEETXLSETUP.xlFlowStrip.Value = "Checking Linker..."
ETWEETXLPOST.xlFlowStrip.Value = "Checking Linker..."
ETWEETXLQUEUE.xlFlowStrip.Value = "Checking Linker..."

'//Linkline active export...
If Range("LinkTrig").Value = 1 Then

Range("LinkerActive").Value = 1

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
    Range("LinkerTotal").Value = xNum
    
    '//
    If API_LINK = 1 Then GoTo SendWithAPI '//Avoid double run
'/######################################################################
EmptyLinker:
'//Check for empty Linker & clear...
If Range("LinkerCount").Value = (Range("LinkerTotal").Value + 1) Then
Call App_TOOLS.ClearLinker
Exit Sub
End If
'/######################################################################
    If xNum = Range("LinkerCount").Value Then
        Range("LinkerCount").Value = Range("LinkerCount").Value + 1
        GoTo EmptyLinker
            End If
    
    On Error GoTo ErrMsg
    
    '//Open tweet and record information...
    xNum = 1
    LLCntr = Range("LinkerCount").Value
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
Print #1, Range("Userlink").Offset(LLCntr + 1, 0).Value
Close #1

'//Set user active...
Range("User").Value = Range("Userlink").Offset(LLCntr + 1, 0).Value
xUser = Range("User").Value

'//Export pass...
Open mtPass For Output As #4
Print #4, Range("Passlink").Offset(LLCntr + 1, 0).Value
Close #4

'//Set pass active...
Range("ActivePass").Value = Range("Passlink").Offset(LLCntr + 1, 0).Value

'//Export profile...
Open mtProf For Output As #5
Print #5, Range("Mainlink").Offset(LLCntr + 1, 0).Value
Close #5

'//Set profile active...
Range("Profile").Value = Range("Mainlink").Offset(LLCntr + 1, 0).Value

'/######################################################################

        
    '//Check for threaded post
    If InStr(1, LinkerArr(LLCntr), ".thr") Then
    xLink = LinkerArr(LLCntr): Call App_TOOLS.SendAsThread(xLink): Exit Sub
            Else
                Range("ThreadActive").Value = 0
                Range("LinkerCount").Value = LLCntr + 1
                    End If
    
    xNum = 1
    Do Until InStr(1, LinkerData(xNum), "*-")
    xMyPost = xMyPost + LinkerData(xNum)
    xNum = xNum + 1
    Loop
    
    xMyMed = Replace(LinkerData(xNum + 1), "*-", vbNullString)
    
    '//Escape special characters...
    xMyPost = Replace(xMyPost, "{ENTER};", Chr(10))
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
        Call ProgBarRefresher
        Call UnfreezeFX
        
        ETWEETXLHOME.ActivePresetBox.Caption = xUser
        ETWEETXLSETUP.ActivePresetBox.Caption = xUser
        ETWEETXLPOST.ActivePresetBox.Caption = xUser
        ETWEETXLQUEUE.ActivePresetBox.Caption = xUser
        
        ETWEETXLHOME.xlFlowStrip.Value = "Sleeping..."
        ETWEETXLSETUP.xlFlowStrip.Value = "Sleeping..."
        ETWEETXLPOST.xlFlowStrip.Value = "Sleeping..."
        ETWEETXLQUEUE.xlFlowStrip.Value = "Sleeping..."
        
'//Turn Linker active...
ETWEETXLHOME.LinkerActive.Caption = "ON"
ETWEETXLHOME.LinkerActive.ForeColor = vbGreen
ETWEETXLHOME.LinkerActive.BackColor = vbWhite
If Range("AppActive").Value <> 1 Then Range("AppActive").Value = 1

SendWithAPI:
'//Send with api method...
If Range("apiLink").Offset(LLCntr + 1, 0).Value = "(*api)" Then
    API_LINK = 1
    Call App_TOOLS.CreateAPIScript
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

'//Update active state
Call App_TOOLS.UpdateActive
Call App_IMPORT.MyNextQueue
Exit Sub

EndMacro:
Call CloseStrandedFiles
Call App_MSG.AppMsg(xMsg)
Exit Sub

ErrMsg:
Call CloseStrandedFiles
xMsg = 27: Call App_MSG.AppMsg(xMsg)
End Sub
Public Sub DraftHdr_Clk()

'//For removing drafts from the Linker

        '//Reset connect trigger...
        If Range("AppActive").Value <> 1 Then
        Range("ConnectTrig").Value = 0
        End If
        
        ETWEETXLPOST.DraftHdr.Caption = "Draft"
        
        '//Remove all drafts from Linker...
        ETWEETXLPOST.LinkerBox.Clear
        
        lastRw = Cells(Rows.Count, "P").End(xlUp).Row
        Range("P2:P" & lastRw).Value = ""
        
        lastRw = Cells(Rows.Count, "L").End(xlUp).Row
        Range("L2:L" & lastRw).Value = ""
        
        If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then _
        ETWEETXLPOST.xlFlowStrip.Value = "All linked drafts cleared..."
        
        
End Sub
Public Sub LinkerHdr_Clk()

'//For removing all users, drafts, & time from the Linker

On Error GoTo EndMacro

        '//Reset connect trigger...
        If Range("AppActive").Value <> 1 Then
        Range("ConnectTrig").Value = 0
        End If
        
        Call Cleanup.ClnLinkerSpace
        Call Cleanup.ClnSpecSpace
        
        ETWEETXLPOST.UserHdr.Caption = "User"
        ETWEETXLPOST.DraftHdr.Caption = "Draft"
        ETWEETXLPOST.RuntimeHdr.Caption = "Runtime"
        
        '//Remove all users from Linker...
        ETWEETXLPOST.UserBox.Clear
        '//Remove all drafts from Linker...
        ETWEETXLPOST.LinkerBox.Clear
        '//Remove all time from Linker...
        ETWEETXLPOST.RuntimeBox.Clear
        
        If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then _
        ETWEETXLPOST.xlFlowStrip.Value = "Linker cleared..."
        
        Exit Sub
        
EndMacro:
        If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then _
        ETWEETXLPOST.xlFlowStrip.Value = "An unknown error occurred while clearing the Linker..."
        
End Sub
Public Sub TimerHdr_Clk()

'//For refresing the Time box

ETWEETXLPOST.TimeBox.Value = "0"
ETWEETXLPOST.TimeBox.Value = ""

If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then _
ETWEETXLPOST.xlFlowStrip.Value = "Time refreshed..."
        
End Sub
Public Sub UserHdr_Clk()

'//For removing all users from the Linker

        '//Reset connect trigger...
        If Range("AppActive").Value <> 1 Then
        Range("ConnectTrig").Value = 0
        End If
        
        ETWEETXLPOST.UserHdr.Caption = "User"

        '//Remove all users from userbox...
        ETWEETXLPOST.UserBox.Clear
        
        Call Cleanup.ClnSpecSpace
        
        If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then _
        ETWEETXLPOST.xlFlowStrip.Value = "All linked users cleared..."
        
End Sub
Public Sub RuntimeHdr_Clk()

        '//Reset connect trigger
        If Range("AppActive").Value <> 1 Then
        Range("ConnectTrig").Value = 0
        End If
        
        ETWEETXLPOST.RuntimeHdr.Caption = "Runtime"
        
        '//REMOVE ALL TIME FROM RUNTIME$
        ETWEETXLPOST.RuntimeBox.Clear
        
        If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then _
        ETWEETXLPOST.xlFlowStrip.Value = "All linked times cleared..."
        
End Sub
Public Sub PostHdr_Clk()

        '//default no parameter
        Call FindForm(xForm)
        '//Clear post box
        xForm.PostBox.Value = vbNullString
        
        If InStr(1, xForm.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then _
        xForm.xlFlowStrip.Value = "Post cleared..."
        
        End Sub

Public Sub AddPostMed_Clk(xMed)

'//For adding media to a post

Dim oMedLinkBox As Object

'//Check xlFlowStrip Window...
If Range("xlasWinForm").Value = 3 Then Set oMedLinkBox = ETWEETXLPOST.MedLinkBox
If Range("xlasWinForm").Value = 4 Then Set oMedLinkBox = ETWEETXLQUEUE.MedLinkBox

lastRw = Cells(Rows.Count, "I").End(xlUp).Row

'//Check if Gif/Video already added...
If Range("GifCntr").Value = 1 Then GoTo ErrMsg1
If Range("VidCntr").Value = 1 Then GoTo ErrMsg1

'//For the first Media...
If Range("MediaScroll").Value = vbNullString Then
lastRw = 0
End If

'//Check for 4 Medias...
If lastRw > 3 Then GoTo ErrMsg2

'//Check if Media already found...
If xMed <> "" Then GoTo SkipOpen

'//Select an Media...
'xMed = Application.GetOpenFilename("PNG Files (*.png), *.png")
xMed = Application.GetOpenFilename()

SkipOpen:

If xMed = "" Then GoTo EndMacro
If xMed = "False" Then GoTo EndMacro

'//gif
If InStr(1, xMed, ".gif") Then
    If oMedLinkBox.Value <> "" Then GoTo ErrMsg1
If FileLen(xMed) > 15728639 Then
xMsg = 13: Call App_MSG.AppMsg(xMsg)
GoTo EndMacro
    End If
    Range("GifCntr").Value = 1
        End If

'//mp4
If InStr(1, xMed, ".mp4") Then
    If oMedLinkBox.Value <> "" Then GoTo ErrMsg1
If FileLen(xMed) > 51200000000# Then
xMsg = 12: Call App_MSG.AppMsg(xMsg)
GoTo EndMacro
    End If
    Range("VidCntr").Value = 1
        End If
        
'//mov
If InStr(1, xMed, ".mov") Then
    If oMedLinkBox.Value <> "" Then GoTo ErrMsg1
If FileLen(xMed) > 51200000000# Then
xMsg = 12: Call App_MSG.AppMsg(xMsg)
GoTo EndMacro
    End If
    Range("VidCntr").Value = 1
        End If
        
'//flv
If InStr(1, xMed, ".flv") Then
    If oMedLinkBox.Value <> "" Then GoTo ErrMsg1
If FileLen(xMed) > 51200000000# Then
xMsg = 12: Call App_MSG.AppMsg(xMsg)
GoTo EndMacro
    End If
    Range("VidCntr").Value = 1
        End If
                              
                              
Range("MediaScroll").Offset(lastRw, 0) = xMed
xMed = "'" & xMed & "'"
xMed = Replace(xMed, "'", """")
'xMed = Replace(xMed, """", vbNullString)
       
If oMedLinkBox.Value = "" Then
oMedLinkBox.Value = xMed
    Else
        oMedLinkBox.Value = oMedLinkBox.Value & " " & xMed
            End If
            
Call App_IMPORT.SelectedMedia
EndMacro:

Set oMedLinkBox = Nothing

Exit Sub

'//Debug...
ErrMsg1:
oMedLinkBox.BorderStyle = fmBorderStyleSingle
oMedLinkBox.BorderColor = vbRed
xMsg = 14: Call App_MSG.AppMsg(xMsg)

Set oMedLinkBox = Nothing

Exit Sub

ErrMsg2:
oMedLinkBox.BorderStyle = fmBorderStyleSingle
oMedLinkBox.BorderColor = vbRed
xMsg = 15: Call App_MSG.AppMsg(xMsg)

Set oMedLinkBox = Nothing

End Sub
Public Sub AddRuntime_Clk(xPos)

'//For adding a time to the Linker

If ETWEETXLPOST.TimeBox <> "" Then

Dim X As Integer
    
If Len(ETWEETXLPOST.TimeBox.Value) = 7 Or 8 Then

'//Check for invalid characters...
Call App_TOOLS.CheckForChar(xChar)
If xChar = "(*Err)" Then Exit Sub

On Error Resume Next

If xPos > 0 Then
    
    If ETWEETXLPOST.RuntimeBox.ListCount > 0 Then

    X = 0

    If InStr(1, xPos, ":") Then
    xPosArr = Split(xPos, ":")
    For X = xPosArr(0) To xPosArr(1)
    iTime = "(" & X & ") " & ETWEETXLPOST.TimeBox.Value
    If Left(iTime, 1) = " " Then iTime = Right(iTime, Len(iTime) - 1) '//remove leading space
    ETWEETXLPOST.RuntimeBox.List((X)) = (iTime)
    Next
    Exit Sub
    End If
    
    If InStr(1, xPos, ",") Then
    xPosArr = Split(xPos, ",")
    xTot = UBound(xPosArr)
    Do Until X = xTot
    iTime = "(" & X & ") " & ETWEETXLPOST.TimeBox.Value
    If Left(iTime, 1) = " " Then iTime = Right(iTime, Len(iTime) - 1) '//remove leading space
    ETWEETXLPOST.RuntimeBox.List(xPosArr(X)) = (iTime)
    X = X + 1
    Loop
    Exit Sub
    End If
        
    If xPos = "" Then
    iTime = ETWEETXLPOST.TimeBox.Value
    iTime = ETWEETXLPOST.RuntimeHdr.Caption & " " & iTime
    iTime = Replace(iTime, "Time", vbNullString, , , vbTextCompare)
    If Left(iTime, 1) = " " Then iTime = Right(iTime, Len(iTime) - 1) '//remove leading space
    ETWEETXLPOST.RuntimeBox.AddItem (iTime)
        Else
            
            iTime = "(" & xPos & ") " & ETWEETXLPOST.TimeBox.Value
            If Left(iTime, 1) = " " Then iTime = Right(iTime, Len(iTime) - 1) '//remove leading space
            ETWEETXLPOST.RuntimeBox.List(xPos) = (iTime)
                End If
        Exit Sub
    
                    End If
                        Else
                         
            ETWEETXLPOST.RuntimeHdr.Caption = "Runtime" & " (" & ETWEETXLPOST.RuntimeBox.ListCount + 1 & ")"
            iTime = ETWEETXLPOST.TimeBox.Value
            iTime = ETWEETXLPOST.RuntimeHdr.Caption & " " & iTime
            iTime = Replace(iTime, "Runtime", vbNullString, , , vbTextCompare)
            If Left(iTime, 1) = " " Then iTime = Right(iTime, Len(iTime) - 1) '//remove leading space
            
            ETWEETXLPOST.RuntimeBox.AddItem (iTime)
            
                If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
                ETWEETXLPOST.xlFlowStrip.Value = iTime & " linked..."
                End If
                
                            End If
                                
                                End If
                                    End If
        
End Sub

Public Function DynOffset_Clk(xPos)

If xPos <> vbNullString Then GoTo SkipHere

If ETWEETXLPOST.DynOffset.Value = True Then
   Range("DynOffTrig").Value = 1
    ETWEETXLPOST.DynOffset.Value = True
    If ETWEETXLPOST.OffsetBox.Value = "00:00:00" Then ETWEETXLPOST.OffsetBox.Value = "00:00:01"
        Else
        Range("DynOffTrig").Value = 0
        ETWEETXLPOST.DynOffset.Value = False
        If ETWEETXLPOST.OffsetBox.Value = "00:00:01" Then ETWEETXLPOST.OffsetBox.Value = "00:00:00"
            End If
            
                Exit Function
                
SkipHere:
   If xPos = 1 Then
     Range("DynOffTrig").Value = 1: ETWEETXLPOST.DynOffset.Value = True
     If ETWEETXLPOST.OffsetBox.Value = "00:00:00" Then ETWEETXLPOST.OffsetBox.Value = "00:00:01"
        ElseIf xPos = 0 Then
        Range("DynOffTrig").Value = 0: ETWEETXLPOST.DynOffset.Value = False
        If ETWEETXLPOST.OffsetBox.Value = "00:00:01" Then ETWEETXLPOST.OffsetBox.Value = "00:00:00"
        End If
            
            
End Function
                
                
Public Sub xlFlowStripBar_Clk()

'//For extending and restricting the xlFlowStrip bar

If Range("xlasWinForm").Value = 1 Then GoTo FlBarHome
If Range("xlasWinForm").Value = 2 Then GoTo FlBarSetup
If Range("xlasWinForm").Value = 3 Then GoTo FlBarPost
If Range("xlasWinForm").Value = 4 Then GoTo FlBarQueue

FlBarHome:
If ETWEETXLHOME.Height = "487.5" Then
ETWEETXLHOME.Height = "637"
Exit Sub
End If

If ETWEETXLHOME.Height = "637" Then
ETWEETXLHOME.Height = "487.5"
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
Public Sub AddDraft_Clk()

'//For adding drafts to a profile

If ETWEETXLPOST.DraftBox.Value <> "" Then

Call App_Loc.xThrFile(thrFile)
Call App_Loc.xTwtFile(twtFile)

On Error Resume Next

'    If Dir(twtFile & ETWEETXLPOST.DraftBox.Value & ".twt") = "" Then
'    Open twtFile & ETWEETXLPOST.DraftBox.Value & ".twt" For Output As #1
'    Print #1, ""
'    Close #1
'            End If

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

xType = 0: Call App_IMPORT.MyTweetData(xType)

If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then _
ETWEETXLPOST.xlFlowStrip.Value = "Draft created..."
        
End Sub
Public Sub AddLink_Auto()

xPos = 0
Call AddLink_Clk(xPos)

End Sub
Public Sub AddLink_Clk(xPos)

'//For adding drafts to the Linker

If ETWEETXLPOST.DraftBox <> "" Then

Dim iDraft As String

On Error Resume Next

lastRwP = Cells(Rows.Count, "P").End(xlUp).Row '//Profilelink for drafts
lastRwAL = Cells(Rows.Count, "AL").End(xlUp).Row '//apiLink for drafts

If xPos > 0 Then
    
    If ETWEETXLPOST.LinkerBox.ListCount > 0 Then
    
    Dim X As Integer
    X = 0

    If InStr(1, xPos, ":") Then
    xPosArr = Split(xPos, ":")
    For X = xPosArr(0) To xPosArr(1)
    iDraft = "(" & X & ") " & ETWEETXLPOST.DraftBox.Value & " [" & ETWEETXLPOST.DraftFilterBtn.Caption & "]"
    If Left(iDraft, 1) = " " Then iDraft = Right(iDraft, Len(iDraft) - 1) '//remove leading space
    ETWEETXLPOST.LinkerBox.List((X)) = iDraft
    Range("Profilelink").Offset(X, 0).Value = Range("Profile").Value
    Range("Draftlink").Offset(lastRwP, 0).Value = ETWEETXLPOST.DraftBox.Value
    Next
    Exit Sub
    End If
    
    If InStr(1, xPos, ",") Then
    xPosArr = Split(xPos, ",")
    xTot = UBound(xPosArr)
    Do Until X = xTot
    iDraft = "(" & X & ") " & ETWEETXLPOST.DraftBox.Value & " [" & ETWEETXLPOST.DraftFilterBtn.Caption & "]"
    If Left(iDraft, 1) = " " Then iDraft = Right(iDraft, Len(iDraft) - 1) '//remove leading space
    ETWEETXLPOST.LinkerBox.List((X)) = iDraft
    Range("Profilelink").Offset(X, 0).Value = Range("Profile").Value
    Range("Draftlink").Offset(lastRwP, 0).Value = ETWEETXLPOST.DraftBox.Value
    X = X + 1
    Loop
    Exit Sub
    End If
    
    iDraft = "(" & xPos & ") " & ETWEETXLPOST.DraftBox.Value & " [" & ETWEETXLPOST.DraftFilterBtn.Caption & "]"
    If Left(iDraft, 1) = " " Then iDraft = Right(iDraft, Len(iDraft) - 1) '//remove leading space
    ETWEETXLPOST.LinkerBox.List(xPos) = iDraft
    Range("Profilelink").Offset(xPos, 0).Value = Range("Profile").Value
    Range("Draftlink").Offset(lastRwP, 0).Value = ETWEETXLPOST.DraftBox.Value
        Exit Sub
    
                        End If
                        
                            Else
                            
                            End If
                            
                            ETWEETXLPOST.DraftHdr.Caption = "Draft" & " (" & ETWEETXLPOST.LinkerBox.ListCount + 1 & ")"
                             
                            iDraft = ETWEETXLPOST.DraftHdr.Caption
                            iDraft = Replace(iDraft, "Draft", vbNullString, , , vbTextCompare)
                            iDraft = iDraft & " " & ETWEETXLPOST.DraftBox.Value & " [" & ETWEETXLPOST.DraftFilterBtn.Caption & "]"
                            If Left(iDraft, 1) = " " Then iDraft = Right(iDraft, Len(iDraft) - 1) '//remove leading space
                            ETWEETXLPOST.LinkerBox.AddItem (iDraft)
                            
                            Range("Profilelink").Offset(lastRwP, 0).Value = Range("Profile").Value
                            Range("Draftlink").Offset(lastRwP, 0).Value = ETWEETXLPOST.DraftBox.Value
                            
                            If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
                            iDraft = Replace(iDraft, " [•]", vbNullString)
                            iDraft = Replace(iDraft, " [...]", vbNullString)
                            ETWEETXLPOST.xlFlowStrip.Value = iDraft & " linked..."
                            End If
                            
                                End If
                        
End Sub
Public Function NewProfile_Clk(xInfo)

'//automate profile creation w/ xlAppScript
If xInfo <> vbNullString Then

lastRw = Cells(Rows.Count, "B").End(xlUp).Row '//last row user column

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
    xInfo = xInfoArr(0): Call App_CLICK.NewProfile_Clk(xInfo)
    ETWEETXLSETUP.ProfileNameBox.Value = xInfoArr(0)
    If Right(xInfoArr(1), 1) = """" Then xInfoArr(1) = Left(xInfoArr(1), Len(xInfoArr(1)) - 1) '//remove ending quote
    If Left(xInfoArr(1), 1) = " " Then xInfoArr(1) = Right(xInfoArr(1), Len(xInfoArr(1)) - 1) '//remove leading space
    ETWEETXLSETUP.UsernameBox.Value = xInfoArr(1)
    ETWEETXLSETUP.PassBox.Value = "password"
    '//no pin...
    Range("Target").Offset(lastRw, 0).Value = ""
    Range("Scure").Offset(lastRw, 0).Value = "*"
    GoTo MoveData
    End If
    '//3 parameters
    If UBound(xInfoArr) = 2 Then
    xInfoA = xInfoArr(0)
    xInfoB = xInfoArr(1)
    xInfoC = xInfoArr(2)
    If Right(xInfoA, 1) = """" Then xInfoA = Left(xInfoA, Len(xInfoA) - 1)  '//remove ending quote
    If Left(xInfoA, 1) = " " Then xInfoA = Right(xInfoA, Len(xInfoA) - 1)     '//remove leading space
    xInfo = xInfoA: Call App_CLICK.NewProfile_Clk(xInfo)
    ETWEETXLSETUP.ProfileNameBox.Value = xInfoA
    If Right(xInfoB, 1) = """" Then xInfoB = Left(xInfoB, Len(xInfoB) - 1) '//remove ending quote
    If Left(xInfoB, 1) = " " Then xInfoB = Right(xInfoB, Len(xInfoB) - 1) '//remove leading space
    ETWEETXLSETUP.UsernameBox.Value = xInfoB
    If Right(xInfoC, 1) = """" Then xInfoC = Left(xInfoC, Len(xInfoC) - 1) '//remove ending quote
    If Left(xInfoC, 1) = " " Then xInfoC = Right(xInfoC, Len(xInfoC) - 1) '//remove leading space
    ETWEETXLSETUP.PassBox.Value = xInfoC
    '//no pin...
    Range("Target").Offset(lastRw, 0).Value = ""
    Range("Scure").Offset(lastRw, 0).Value = "*"
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
    xInfo = xInfoA: Call App_CLICK.NewProfile_Clk(xInfo)
    ETWEETXLSETUP.ProfileNameBox.Value = xInfoA
    If Right(xInfoB, 1) = """" Then xInfoB = Left(xInfoB, Len(xInfoB) - 1) '//remove ending quote
    If Left(xInfoB, 1) = " " Then xInfoB = Right(xInfoB, Len(xInfoB) - 1) '//remove leading space
    ETWEETXLSETUP.UsernameBox.Value = xInfoB
    If Right(xInfoC, 1) = """" Then xInfoC = Left(xInfoC, Len(xInfoC) - 1) '//remove ending quote
    If Left(xInfoC, 1) = " " Then xInfoC = Right(xInfoC, Len(xInfoC) - 1) '//remove leading space
    ETWEETXLSETUP.PassBox.Value = xInfoC
    PinHldr = xInfoArr(3)
    '//Check edit mode...
    If Range("SetupEdit").Value <> 1 Then
    '//Pin for new user...
    Range("Scure").Offset(lastRw, 0).Value = "***"
    Range("Target").Offset(lastRw, 0).Value = PinHldr '//record
        Else
    '//Replace existing user pin...
    Range("Scure").Offset((lastRw - 1), 0).Value = "***"
    Range("Target").Offset((lastRw - 1), 0).Value = PinHldr '//record
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
        Range("Profile").Offset(lastRw, 0).Value = ETWEETXLSETUP.UsernameBox.Value
        Range("User").Offset(lastRw, 0).Value = ETWEETXLSETUP.PassBox.Value
        
        '//Export data
        Call App_EXPORT.MyUserData
        
        '//Refresh data
        Range("DataPullTrig").Value = "0"
        '//Import profile information
        Call App_IMPORT.MyProfileData
        
CheckName:
If ETWEETXLSETUP.ProfileNameBox.Value = "" Then
ETWEETXLSETUP.ProfileNameBox.BorderStyle = fmBorderStyleSingle
ETWEETXLSETUP.ProfileNameBox.BorderColor = vbRed
xMsg = 20: Call App_MSG.AppMsg(xMsg)
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
        Range("DataPullTrig").Value = "0"
        ETWEETXLSETUP.UserListBox.Clear
        ETWEETXLSETUP.ProfileListBox.Value = ETWEETXLSETUP.ProfileNameBox.Value
        '//Import profile information
        Call App_IMPORT.MyProfileData
        
Exit Function

ErrMsg:
Range("GetInfo").Value = ""
Range("SetupEdit").Value = 0
              
              
End Function
Public Function NewUser_Clk(xInfo)

On Error GoTo ErrMsg

Dim EDITMODE As Integer

lastRw = Cells(Rows.Count, "B").End(xlUp).Row '//last row user column

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
    Range("Target").Offset(lastRw, 0).Value = ""
    Range("Scure").Offset(lastRw, 0).Value = "*"
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
    Range("Target").Offset(lastRw, 0).Value = ""
    Range("Scure").Offset(lastRw, 0).Value = "*"
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
    PinHldr = xInfoArr(2)
    '//Check edit mode...
    If EDITMODE <> 1 Then
    '//Pin for new user...
    Range("Scure").Offset(lastRw, 0).Value = "***"
    Range("Target").Offset(lastRw, 0).Value = PinHldr '//record
        Else
    '//Replace existing user pin...
    Range("Scure").Offset((lastRw - 1), 0).Value = "***"
    Range("Target").Offset((lastRw - 1), 0).Value = PinHldr '//record
    End If
    GoTo MoveData
        End If
            End If

'//check user box...
If ETWEETXLSETUP.UsernameBox.Value = "" Then
ETWEETXLSETUP.UsernameBox.BorderStyle = fmBorderStyleSingle
ETWEETXLSETUP.UsernameBox.BorderColor = vbRed
xMsg = 18: Call App_MSG.AppMsg(xMsg)
Exit Function
End If
    
    '//check pass box...
    If ETWEETXLSETUP.PassBox.Value = "" Then
    ETWEETXLSETUP.PassBox.BorderStyle = fmBorderStyleSingle
    ETWEETXLSETUP.PassBox.BorderColor = vbRed
    xMsg = 19: Call App_MSG.AppMsg(xMsg)
    Exit Function
    End If
    
        '//Check for editing mode
        EDITMODE = 0
        
        If Range("SetupEdit").Value = 1 Then
            EDITMODE = 1
                ElseIf Range("SetupEdit").Value = 2 Then
                    EDITMODE = 2
                            End If

        
        '//If in editing mode...
        If EDITMODE = 2 Then
            '//Check for information...
            If Range("GetInfo").Value = "" Then
                Range("SetupEdit").Value = 0
                EDITMODE = 0
                xMsg = 21: Call App_MSG.AppMsg(xMsg)
                xMsg = 22: Call App_MSG.AppMsg(xMsg)
                GoTo ExitEditMode '//Exit edit mode if information not found
                    End If
                    
                    
            GetInfoArr = Split(Range("GetInfo").Value, ",") '//Get edit info
            
            For X = 1 To lastRw
            If Range("Profile").Offset(X, 0) = GetInfoArr(0) Then
            lastRw = X
                GoTo AskForPin
                    End If
                        Next
                        
                            End If
                            
ExitEditMode:
        For X = 1 To lastRw
        If ETWEETXLSETUP.UsernameBox.Value = Range("Profile").Offset(X, 0) Then GoTo EditUser
        Next
        
        '//If user doesn't exist...
        lastRw = Cells(Rows.Count, "A").End(xlUp).Row '//Last row profile column
        
'ASKFORPIN
AskForPin:

    If EDITMODE = 0 Then
        '//Enter pin?
        If Range("xlasSilent") <> 1 Then msg = MsgBox("Would you like to enter a pin for this user?", vbYesNo, AppTag)
            ElseIf EDITMODE = 1 Then
        '//Enter new pin?
        If Range("xlasSilent") <> 1 Then msg = MsgBox("Would you like to enter a new pin for this user?", vbYesNo, AppTag)
        '//We've been here, move our data...
                Else: GoTo MoveData
                        End If
                
                '//Exit if nothing entered...
                If msg = "" Then Exit Function
                
        '//Yes
        If msg = vbYes Then
'ENTERPIN
EnterPin:
            If Range("xlasSilent") <> 1 Then msg = InputBox("Enter a 4-digit pin:", AppTag)
                PinHldr = msg
                
                If PinHldr <> "" Then
                '//Check for 4 digits
                 If Len(PinHldr) = 4 Then GoTo ReEnterPin '//Success
                    If Len(PinHldr) < 4 Then '//Too short
                        If Range("xlasSilent") <> 1 Then MsgBox ("This pin is too short"), vbInformation, AppTag
                            ElseIf Len(PinHldr) > 4 Then '//Too long
                                If Range("xlasSilent") <> 1 Then MsgBox ("This pin is too long"), vbInformation, AppTag
                                    End If
                                    ErrCntr = ErrCntr + 1 '//Error counter
                                    If ErrCntr >= 5 Then GoTo ErrMsg '//5 attempts until quit...
                                    GoTo EnterPin
                                    

                                  
'REENTERPIN
ReEnterPin:
                '//Check pin
                 If Range("xlasSilent") <> 1 Then msg = InputBox("Re-enter pin:", AppTag)
                    If msg = PinHldr Then
                    '//CHheck editing mode...
                        If EDITMODE <> 1 Then
                        '//set pin for new user...
                        Range("Scure").Offset(lastRw, 0).Value = "***"
                        Range("Target").Offset(lastRw, 0).Value = PinHldr '//RECORD
                            Else
                        '//set pin for existing user...
                        Range("Scure").Offset((lastRw - 1), 0).Value = "***"
                        Range("Target").Offset((lastRw - 1), 0).Value = PinHldr '//RECORD
                        End If
                            Else
                        '//wrong pin entered...
                        If Range("xlasSilent") <> 1 Then msg = MsgBox("Incorrect pin", vbExclamation, AppTag)
                                ErrCntr = ErrCntr + 1
                                If ErrCntr >= 5 Then GoTo ErrMsg '//5 attempts until quit...
                                GoTo ReEnterPin
                                    End If
                                        End If
                                        
                                            Else
                                            
                                            '//No pin entered...
                                            Range("Target").Offset(lastRw, 0).Value = ""
                                            Range("Scure").Offset(lastRw, 0).Value = "*"
                                            
                                                    End If
                                                        
                                                        
If EDITMODE = 1 Then GoTo UpdateInfo

'MOVEDATA
MoveData:

       '//Move our data for transfer...
        Range("Profile").Offset(lastRw, 0).Value = ETWEETXLSETUP.UsernameBox.Value
        Range("User").Offset(lastRw, 0).Value = ETWEETXLSETUP.PassBox.Value
        
        '//Export data
        Call App_EXPORT.MyUserData
        
        '//Refresh data
        Range("DataPullTrig").Value = "0"
        '//Import profile information
        Call App_IMPORT.MyProfileData
        
                
'//EXITING EDIT MODE...
If EDITMODE = 2 Then
'//Exiting edit mode
xMsg = 22: Call App_MSG.AppMsg(xMsg)
    Else
        End If
        
Range("SetupEdit").Value = 0

Exit Function

'//Ask to edit user if user exists
EditUser:
Dim nl As Variant
nl = vbNewLine

xUser = Range("Profile").Offset(X, 0).Value '//Get user position
xPass = Range("User").Offset(X, 0).Value '//Get pass position
xPin = Range("Target").Offset(X, 0).Value '//Get pin position

xGetInfo = xUser & "," & xPass & "," & xPin
Range("GetInfo").Value = xGetInfo

If Range("xlasSilent") <> 1 Then msg = MsgBox("[ " & xUser & " ]" & nl & nl & _
"Edit information for this user?", vbYesNo, AppTag)
    If msg = vbYes Then
    Range("SetupEdit").Value = 1
    EDITMODE = 1
        GoTo AskForPin '//Ask to setup new pin
            Else: Exit Function
                End If
            
'//Enter new information click '+' to update...
UpdateInfo:
        If Range("xlasSilent") <> 1 Then MsgBox ("Enter new information for '" & xUser & "', then click the '+' button again to update."), vbInformation, AppTag
        Range("SetupEdit").Value = 2
Exit Function

ErrMsg:
If Range("xlasSilent") <> 1 Then MsgBox ("Error adding pin for this user"), vbExclamation, AppTag
Range("GetInfo").Value = ""
Range("SetupEdit").Value = 0

End Function
Public Sub AddUser_Clk(xPos)

'//For adding users to the Linker

If ETWEETXLPOST.UserListBox <> "" Then

ETWEETXLPOST.UserHdr.Caption = "User" & " (" & ETWEETXLPOST.UserBox.ListCount + 1 & ")"
    
On Error Resume Next

lastRwA = Cells(Rows.Count, "A").End(xlUp).Row
lastRwAM = Cells(Rows.Count, "AM").End(xlUp).Row
xUser = Range("User").Value
    
If xPos > 0 Then
    
    For xNum = 1 To lastRwA
    If Range("A" & xNum).Value = xUser Then
    Range("Passlink").Offset(lastRwAL, 0).Value = Range("B" & xNum).Value
    End If
    Next

    If ETWEETXLPOST.UserBox.ListCount > 0 Then
    
    Dim X As Integer
    X = 0

    If InStr(1, xPos, ":") Then
    xPosArr = Split(xPos, ":")
    For X = xPosArr(0) To xPosArr(1)
    iUser = "(" & X & ") " & ETWEETXLPOST.UserListBox.Value
    If Left(iUser, 1) = " " Then iUser = Right(iUser, Len(iUser) - 1) '//remove leading space
    ETWEETXLPOST.UserBox.List((X)) = (iUser)
    Next
    Exit Sub
    End If
    
    If InStr(1, xPos, ",") Then
    xPosArr = Split(xPos, ",")
    xTot = UBound(xPosArr)
    Do Until X = xTot
    iUser = "(" & X & ") " & ETWEETXLPOST.UserListBox.Value
    If Left(iUser, 1) = " " Then iUser = Right(iUser, Len(iUser) - 1) '//remove leading space
    ETWEETXLPOST.UserBox.List(xPosArr(X)) = iUser
    X = X + 1
    Loop
    Exit Sub
    End If
    
    iUser = "(" & xPos & ") " & ETWEETXLPOST.UserListBox.Value
    If Left(iUser, 1) = " " Then iUser = Right(iUser, Len(iUser) - 1) '//remove leading space
    ETWEETXLPOST.UserBox.List(xPos) = iUser
    Exit Sub
    
                    End If
                        Else
                        
                    iUser = Replace(ETWEETXLPOST.UserListBox.Value, Range("Scure").Value, "")
                 
                 '//Check for API send data...
                 If Range("SendAPI").Value = 1 Then
                 Call App_Loc.xApiFile(apiFile)
                    If Dir(apiFile) = "" Then
                        xMsg = 7: Call App_MSG.AppMsg(xMsg)
                        Else
                            iUser = iUser & "(*api)"
                                End If
                                    End If
                 
            For xNum = 1 To lastRwA
            If Range("A" & xNum).Value = xUser Then
            Range("Passlink").Offset(lastRwAM, 0).Value = Range("B" & xNum).Value
            Range("Mainlink").Offset(lastRwAM, 0).Value = Range("Profile").Value
            End If
            Next
            
            '//Check for API send data...
            If InStr(1, iUser, "(*api)") Then
            iUser = Replace(ETWEETXLPOST.UserListBox.Value, "(*api)", "")
            iUser = Replace(iUser, Range("Scure").Value, "")
            Range("apiLink").Offset(lastRwAM, 0).Value = "(*api)"
                Else
                    Range("apiLink").Offset(lastRwAM, 0).Value = "(*)"
                        End If
            iUser = ETWEETXLPOST.UserHdr.Caption & " " & iUser
            iUser = Replace(iUser, "User", vbNullString, , , vbTextCompare)
            If Left(iUser, 1) = " " Then iUser = Right(iUser, Len(iUser) - 1) '//remove leading space
            ETWEETXLPOST.UserBox.AddItem iUser

                If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
                ETWEETXLPOST.xlFlowStrip.Value = iUser & " linked..."
                End If
                
                    End If
                            
                        End If
                                
End Sub
Public Sub RmvPostMed_Clk()

'//For removing Media from a post

Dim MedArr(5) As String
Dim X As Integer
Dim oMedLinkBox As Object

On Error GoTo SetForm

'//Find running window
Call FindForm(xForm)

Retry:
'//Check xlFlowStrip Window...
If Range("xlasWinForm").Value = 3 Then Set oMedLinkBox = ETWEETXLPOST.MedLinkBox
If Range("xlasWinForm").Value = 4 Then Set oMedLinkBox = ETWEETXLQUEUE.MedLinkBox

'//Cleanup
oMedLinkBox.Value = ""

'//Get selected media from position
xPos = Range("MediaScroll").Offset(Range("MedScrollPos").Value)

If InStr(1, xTwt, " [...]") = False Then
If xForm.DraftFilterBtn.Caption <> "..." Or InStr(1, xTwt, " [•]") Then
xExt = ".twt": xT = " [•]": If xForm.DraftFilterBtn.Caption <> "•" Then xFil = 0: Call App_CLICK.DraftFilterBtn_Clk(xFil)
    Else: GoTo asThread
            End If
        Else
asThread:
            xExt = ".thr": xT = " [...]"
            If xForm.DraftFilterBtn.Caption <> "..." Then xFil = 1: Call App_CLICK.DraftFilterBtn_Clk(xFil)
            GoTo RmvThreadMed
                End If
                      
'//
lastRw = Cells(Rows.Count, "I").End(xlUp).Row

'//Get media paths from sheet except the one to remove
xNum = 0
For X = 0 To lastRw
    If Range("MediaScroll").Offset(X, 0).Value <> xPos Then
        MedArr(xNum) = Range("MediaScroll").Offset(X, 0).Value
        xNum = xNum + 1
            End If
                Next X

'//Clear media space
Range("I1:I" & lastRw).ClearContents

'//Print remaining to media space
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

Call App_IMPORT.SelectedMedia

Exit Sub

RmvThreadMed:

On Error Resume Next

Dim xNwMed As String
xNwMed = vbNullString

lastRw = Cells(Rows.Count, "Z").End(xlUp).Row

If xForm.ThreadCt.Caption <> "" Then
If CInt(xForm.ThreadCt.Caption) > 0 Then
lastRw = xForm.ThreadCt.Caption
GoTo NextStep
End If
    End If
    
NextStep:
'//remove post from thread loc
xMedArr = Split(Range("MedThread").Offset(lastRw, 0).Value, ",")

'//get selected media position
xPos = Range("MedScrollPos").Value

'//remove media from thread loc
xMedArr(xPos) = vbNullString

'//reconnect media
For X = 0 To UBound(xMedArr)
If xMedArr(X) <> vbNullString Then xNwMed = xNwMed & xMedArr(X)
Next

'//print remaining to worksheet
Range("MedThread").Offset(lastRw, 0).Value = xNwMed

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
If Range("xlasWinForm").Value <> 3 Then Range("xlasWinForm").Value = 3 Else Range("xlasWinForm").Value = 4
GoTo Retry

End Sub
Public Sub DeleteDraft_Clk()

'//For removing drafts from a profile

If ETWEETXLPOST.DraftBox.Value <> "" Then

Call FindForm(xForm)
Call App_Loc.xThrFile(thrFile)
Call App_Loc.xTwtFile(twtFile)

xPos = ETWEETXLPOST.DraftBox.ListIndex
xTwt = ETWEETXLPOST.DraftBox.Value

On Error Resume Next

    If Dir(twtFile & xTwt Or thrFile & xTwt) <> "" Then
    If ETWEETXLPOST.ThreadTrig = 0 Then Kill (twtFile & xTwt & ".twt")
    If ETWEETXLPOST.ThreadTrig = 1 Then Kill (thrFile & xTwt & ".thr")
            
            End If


If InStr(1, xTwt, " [...]") = False Then
If xForm.DraftFilterBtn.Caption <> "..." Or InStr(1, xTwt, " [•]") Then
xExt = ".twt": xT = " [•]": If xForm.DraftFilterBtn.Caption <> "•" Then xFil = 0: Call App_CLICK.DraftFilterBtn_Clk(xFil)
    Else: GoTo asThread
            End If
        Else
asThread:
            xExt = ".thr": xT = " [...]"
            If xForm.DraftFilterBtn.Caption <> "..." Then xFil = 1: Call App_CLICK.DraftFilterBtn_Clk(xFil)
                End If

                       If xExt = ".twt" Then xType = 0
                       If xExt = ".thr" Then xType = 1
                       
                       Call App_IMPORT.MyTweetData(xType)
                        
                        '//set to next draft
                        ETWEETXLPOST.DraftBox.Value = ETWEETXLPOST.DraftBox.List(xPos)
                        
                            End If
                            
End Sub
Public Sub RmvLink_Auto()

Call RmvLink_Clk

End Sub
Public Sub RmvLink_Clk()

'//For removing a selected draft from the Linker

Dim iNum As Integer

If ETWEETXLPOST.LinkerBox.ListCount > 0 Then
ETWEETXLPOST.DraftHdr.Caption = "Draft" & " (" & ETWEETXLPOST.LinkerBox.ListCount - 1 & ")"
End If

RtCntr = ETWEETXLPOST.LinkerBox.ListCount

iNum = ETWEETXLPOST.LinkerBox.ListIndex

If iNum < 0 Then
 iNum = (RtCntr - 1)
        If iNum <= 0 Then
                If iNum < 0 Then Exit Sub
                ETWEETXLPOST.LinkerBox.RemoveItem (iNum)
                xBox = 2: Call ResetBoxOrder(xBox)
                    Exit Sub
                        End If
                            End If
                            
                If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
                ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.LinkerBox.List(iNum) & " unlinked..."
                End If

lastRw = Cells(Rows.Count, "P").End(xlUp).Row
Range("DraftLink").Offset(lastRw - 1, 0).Value = vbNullString
Range("ProfileLink").Offset(lastRw - 1, 0).Value = vbNullString

ETWEETXLPOST.LinkerBox.RemoveItem (iNum)
xBox = 2: Call ResetBoxOrder(xBox)

End Sub
Public Sub RmvUser_Clk()

'//For removing a selected user from the Linker

Dim iNum As Integer

If ETWEETXLPOST.UserBox.ListCount > 0 Then
ETWEETXLPOST.UserHdr.Caption = "User" & " (" & ETWEETXLPOST.UserBox.ListCount - 1 & ")"
End If

RtCntr = ETWEETXLPOST.UserBox.ListCount

iNum = ETWEETXLPOST.UserBox.ListIndex

If iNum < 0 Then
 iNum = (RtCntr - 1)
        If iNum <= 0 Then
                If iNum < 0 Then Exit Sub
                ETWEETXLPOST.UserBox.RemoveItem (iNum)
                xBox = 1: Call ResetBoxOrder(xBox)
                    Exit Sub
                        End If
                            End If
                            
                If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
                ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.UserBox.List(iNum) & " unlinked..."
                End If

lastRw = Cells(Rows.Count, "AM").End(xlUp).Row

Range("apiLink").Offset(lastRw - 1, 0).Value = vbNullString
Range("MainLink").Offset(lastRw - 1, 0).Value = vbNullString
Range("PassLink").Offset(lastRw - 1, 0).Value = vbNullString
Range("UserLink").Offset(lastRw - 1, 0).Value = vbNullString

ETWEETXLPOST.UserBox.RemoveItem (iNum)
xBox = 1: Call ResetBoxOrder(xBox)

End Sub
Public Sub RmvRuntime_Clk()

Dim iNum As Integer

If ETWEETXLPOST.RuntimeBox.ListCount > 0 Then
ETWEETXLPOST.RuntimeHdr.Caption = "Runtime" & " (" & ETWEETXLPOST.RuntimeBox.ListCount - 1 & ")"
End If

RtCntr = ETWEETXLPOST.RuntimeBox.ListCount

iNum = ETWEETXLPOST.RuntimeBox.ListIndex
If iNum < 0 Then
    iNum = (RtCntr - 1)
        If iNum < 0 Then Exit Sub
            End If
       
                If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
                ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.RuntimeBox.List(iNum) & " unlinked..."
                End If
                
ETWEETXLPOST.RuntimeBox.RemoveItem (iNum)
xBox = 4: Call ResetBoxOrder(xBox)

End Sub
Public Function DelAllUsers_Clk()

Dim xProf, xMsg As String

If ETWEETXLSETUP.ProfileNameBox.Value <> vbNullString Or Range("Profile").Value <> vbNullString Then

If ETWEETXLSETUP.ProfileNameBox.Value <> vbNullString Then xProf = ETWEETXLSETUP.ProfileNameBox.Value Else _
xProf = Range("Profile").Value

If Range("xlasSilent").Value = 1 Then xMsg = vbYes: GoTo SilentRun

xMsg = MsgBox("Are you sure you wish to remove all users for '" & xProf & "'?", vbYesNo, AppTag)

SilentRun:
If xMsg = vbYes Then

lastRw = Cells(Rows.Count, "A").End(xlUp).Row

For X = 1 To lastRw
xInfo = Range("Profile").Offset(X, 0).Value
Call App_CLICK.DelUser_Clk(xInfo)
Next
End If

End If
End Function
Public Function DelUser_Clk(xInfo)

If xInfo <> vbNullString Then
If Right(xInfo, 1) = """" Then xInfo = Left(xInfo, Len(xInfo) - 1) '//remove ending quote
If Left(xInfo, 1) = """" Then xInfo = Right(xInfo, Len(xInfo) - 1) '//remove leading quote
If Left(xInfo, 1) = " " Then xInfo = Right(xInfo, Len(xInfo) - 1) '//remove leading space
ETWEETXLSETUP.UsernameBox.Value = xInfo
End If

'//For removing a specific user from a profile
If ETWEETXLSETUP.UsernameBox.Value <> "" Then
        
        lastRw = Cells(Rows.Count, "A").End(xlUp).Row
        
        For xNum = 1 To lastRw
            If Range("Profile").Offset(xNum, 0).Value = ETWEETXLSETUP.UsernameBox.Value Then
            Range("Profile").Offset(xNum, 0).Value = ""
            Range("User").Offset(xNum, 0).Value = ""
            Range("Browser").Offset(xNum, 0).Value = ""
            End If
                Next
        
            '//Export data
            Call App_EXPORT.MyUserData
            
            '//Refresh
            Range("DataPullTrig").Value = "0"
            Call App_IMPORT.MyProfileData
        
                End If
                
End Function
Public Sub SaveLinkerBtn_Clk()

'//For creating a .link file (connection state)

'//Enter name for file
fiName = InputBox("Enter a name for your link:", AppTag)

If fiName <> "" Then

Dim X As Integer
X = 1

lastRw = Cells(Rows.Count, "L").End(xlUp).Row

'//Format sheet
Call App_TOOLS.ShFormat

'//Select file save location
fiPath = Application.GetSaveAsFilename(fiName, ".link, *.link")

Open fiPath For Output As #1

Do Until X = lastRw
DraftArr = Split(ETWEETXLPOST.LinkerBox.List(X - 1), ") "): iDraft = DraftArr(1)

Print #1, Range("Mainlink").Offset(X, 0).Value & "," _
& Range("Userlink").Offset(X, 0).Value _
& Range("apiLink").Offset(X, 0).Value & "," _
& Range("Profilelink").Offset(X, 0).Value & "," _
& iDraft & "," _
& Format(Range("Runtime").Offset(X, 0).Value, "hh:mm:ss")
X = X + 1
Loop

Close #1

If Not InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) Then
ETWEETXLPOST.xlFlowStrip.Value = "Link saved..."
End If

End If

End Sub
Public Sub SaveBtn_Clk()

'//For saving API key preset information

Dim oBox, oForm As Object

Set oForm = ETWEETXLAPISETUP

'//API Setup

'//Save to .pers file...
Call App_Loc.xApiFile(apiFile)

Open apiFile For Output As #2

Set oBox = oForm.apiKeyBox
If oBox <> "" Then '//check for empty box
Print #2, oBox.Value
    Else
        Close #2
        oBox.BorderColor = vbRed
        xMsg = 3: Call App_MSG.AppMsg(xMsg) '//if box is empty show error & don't save
        Exit Sub
            End If
       
Set oBox = oForm.apiSecretBox
If oBox <> "" Then
Print #2, oBox.Value
    Else
        Close #2
        oBox.BorderColor = vbRed
        xMsg = 3: Call App_MSG.AppMsg(xMsg)
        Exit Sub
            End If
            
Set oBox = oForm.accTokenBox
If oBox <> "" Then
Print #2, oBox.Value
    Else
        Close #2
        oBox.BorderColor = vbRed
        xMsg = 3: Call App_MSG.AppMsg(xMsg)
        Exit Sub
            End If

Set oBox = oForm.accSecretBox
If oBox <> "" Then
Print #2, oBox.Value
    Else
        Close #2
        oBox.BorderColor = vbRed
        xMsg = 3: Call App_MSG.AppMsg(xMsg)
        Exit Sub
            End If
            
Close #2

Set oBox = Nothing
Set oForm = Nothing

xMsg = 16: Call App_MSG.AppMsg(xMsg)

End Sub
Public Sub SendAPI_Clk(xPos)

If xPos = 1 Then

    Range("SendAPI").Value = 1
    
    If Range("LoadLess").Value <> 1 Then ETWEETXLPOST.SendAPI.Value = True
    
    ElseIf xPos = 0 Then
 
    Range("SendAPI").Value = 0
    
    If Range("LoadLess").Value <> 1 Then ETWEETXLPOST.SendAPI.Value = False
    
    End If
                
End Sub
Public Sub SavePost_Auto()

Call SavePost_Clk

End Sub
Public Sub SavePost_Clk()

'//For saving a post (tweet)

Dim PostCharCt As Integer

Call FindForm(xForm)

'//Get post character count...
PostCharCt = Len(xForm.PostBox.Value)
xForm.CharCt.Caption = PostCharCt

If PostCharCt < 280 Then

If Range("xlasWinForm").Value = 3 Then xTwt = xForm.DraftBox.Value: Call App_EXPORT.MyDraftData(xTwt) Else _
xTwt = ETWEETXLPOST.DraftBox.Value: _
Call App_EXPORT.MyDraftData(xTwt) '//from queue

'//find current draft filter
If InStr(1, xTwt, " [...]") = False Then
If xForm.DraftFilterBtn.Caption <> "..." Or InStr(1, xTwt, " [•]") Then
xExt = ".twt": xT = " [•]": If xForm.DraftFilterBtn.Caption <> "•" Then xFil = 0: Call App_CLICK.DraftFilterBtn_Clk(xFil)
xType = 0
    Else: GoTo asThread
            End If
        Else
asThread:
            xExt = ".thr": xT = " [...]"
            If xForm.DraftFilterBtn.Caption <> "..." Then xFil = 1: Call App_CLICK.DraftFilterBtn_Clk(xFil)
            xType = 1
                End If
            
            
Call App_IMPORT.MyTweetData(xType)

If Range("xlasWinForm").Value = 3 Then xForm.DraftBox.Value = xTwt

If Range("xlasSilent") <> 1 Then MsgBox ("Post saved"), vbInformation, AppTag
    Else
        GoTo EndMacro
            End If

                Exit Sub
                
EndMacro:
xMsg = 21: Call App_MSG.AppMsg(xMsg)
                    
End Sub
Public Function SetActive_Clk(xUser)

'//For setting a user active

Call App_Loc.xPersFile(persFile)

If Dir(persFile) = "" Then GoTo ErrMsg
If xUser = "" Then GoTo ErrMsg

xUser = Replace(xUser, Range("Scure").Value, "") '//Remove lock symbol from name
If ETWEETXLSETUP.ProfileListBox.Value <> "" Then '//Profile
Range("Profile").Value = ETWEETXLSETUP.ProfileListBox.Value: End If
'Range("Browser").Value = ETWEETXLSETUP.BrowserBox.Value
Range("Browser").Value = "Firefox" '//Browser
Range("ActiveUser").Value = xUser '//User
ETWEETXLHOME.ActivePresetBox.Caption = xUser
ETWEETXLHOME.ActivePresetBox.BackColor = vbWhite
ETWEETXLSETUP.ActivePresetBox.Caption = xUser
ETWEETXLSETUP.ActivePresetBox.BackColor = vbWhite
ETWEETXLPOST.ActivePresetBox.Caption = xUser
ETWEETXLPOST.ActivePresetBox.BackColor = vbWhite

X = 1
Do Until Range("Profile").Offset(X, 0).Value = xUser
X = X + 1
If X > 5000 Then GoTo SkipHere
Loop

SkipHere:
Range("ActivePass").Value = Range("User").Offset(X, 0).Value

'//Check for pass...
If Range("xlasWinForm").Value = 2 Then
If Range("ActivePass").Value = "" Then If ETWEETXLSETUP.PassBox.Value <> "" Then _
Range("ActivePass").Value = ETWEETXLSETUP.PassBox.Value Else GoTo ErrMsg
End If

Exit Function


ErrMsg:
xMsg = 5: Call App_MSG.AppMsg(xMsg)

End Function
Public Sub ConnectPost_Auto()

Call ConnectPost_Clk

End Sub
Public Sub ConnectPost_Clk()

On Error GoTo ErrMsg

'//For connecting posts from the Linker

Dim AutoRuntimeArr(5000), DraftArr(5000), RuntimeArr(5000) As String
Dim oDynOffsetBox, oDraftBox, oOffsetBox, oRuntimeBox, oUserBox As Object

'//Check xlFlowStrip Window...
If Range("xlasWinForm").Value = 3 Then
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
If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
ETWEETXLPOST.xlFlowStrip.Value = "Connecting..."
End If

'//Cleanup...
Range("LinkerCount").Value = 0

'//Set link trigger...
Range("LinkTrig").Value = 1
Range("ConnectTrig").Value = 1

'//Check for dynamic offset...
If oDynOffsetBox.Value = True Then
Range("DynOffTrig").Value = 1
GoTo DynSetup
    Else
        Range("DynOffTrig").Value = 0
            End If

'//Convert offset to milliseconds...
OffsetArr = Split(oOffsetBox.Value, ":")
OffsetArrCopy = Split(oOffsetBox.Value, ":")
If CDbl(OffsetArr(0)) <> 0 Then OffsetArr(0) = (OffsetArr(0) * 3600 * 1000)
If CDbl(OffsetArr(1)) <> 0 Then OffsetArr(1) = (OffsetArr(1) * 60 * 1000)
If CDbl(OffsetArr(2)) <> 0 Then OffsetArr(2) = (OffsetArr(2) * 1000)
TotalOffset = CDbl(OffsetArr(0)) + CDbl(OffsetArr(1)) + CDbl(OffsetArr(2))
Range("Offset").Offset(1, 0).Value = TotalOffset
Range("ActiveOffset").Value = TotalOffset

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
Range("Offset").Offset(1, 0).Value = TotalOffset
Range("ActiveOffset").Value = TotalOffset

LinkerSetup:

'//Set drafts...
zNum = 1
For xNum = 0 To (oDraftBox.ListCount - 1)
oDraftBox.ListIndex = xNum
iDraft = oDraftBox.List(xNum)
iDraft = Replace(iDraft, "(" & xNum + 1 & ") ", vbNullString)
DraftArr(zNum) = iDraft
zNum = zNum + 1
Next

'//Capture total links...
LHldr = xNum

Dim NwTwtFile As String

Call App_Loc.xTwtFile(twtFile)
Call App_Loc.xThrFile(thrFile)
Range("Post").Value = twtFile & oDraftBox.Value & ".twt"
Range("ActiveTweet").Value = "" '//Clear active tweet range
xNum = 1

    Do Until DraftArr(xNum) = ""
    If InStr(1, DraftArr(xNum), " [•]") Then NwTwtFile = Replace(twtFile, xProfile, Range("ProfileLink").Offset(xNum, 0).Value): DraftArr(xNum) = Replace(DraftArr(xNum), " [•]", vbNullString): Range("ActiveTweet").Value = Range("ActiveTweet").Value & NwTwtFile & DraftArr(xNum) & ".twt" & ","
    If InStr(1, DraftArr(xNum), " [...]") Then NwTwtFile = Replace(thrFile, xProfile, Range("ProfileLink").Offset(xNum, 0).Value): DraftArr(xNum) = Replace(DraftArr(xNum), " [...]", vbNullString): Range("ActiveTweet").Value = Range("ActiveTweet").Value & NwTwtFile & DraftArr(xNum) & ".thr" & ","
    xNum = xNum + 1
    Loop
            
'//Setup runtime interval...
lastRw = Cells(Rows.Count, "R").End(xlUp).Row

'//Clear runtime space...
For xNum = 1 To lastRw
Range("R" & xNum).ClearContents
Next

'//Check for auto offset...
zNum = 1
If oRuntimeBox.ListCount = 1 Then
RuntimeArr(1) = oRuntimeBox.List(0)
If TotalOffset <> 0 Then GoTo AutoOffset
End If

'//Print runtime to range...
For xNum = 0 To (oRuntimeBox.ListCount - 1)
oRuntimeBox.ListIndex = xNum
iTime = oRuntimeBox.List(xNum)
iTime = Replace(iTime, "(" & xNum + 1 & ") ", vbNullString)
iTime = Replace(iTime, " ", vbNullString)

Range("Runtime").Offset(zNum, 0).Value = iTime
RuntimeArr(zNum) = oRuntimeBox.List(xNum)
zNum = zNum + 1
Next

GoTo SetupRtCntr

'//Automatically offset runtime if only one time linked, & an offset...
AutoOffset:

RuntimeArr(1) = Replace(RuntimeArr(1), "(1) ", vbNullString)

RuntimeArrCopy = Split(RuntimeArr(1), ":")

'//Set original runtime...
AutoRuntimeArr(1) = RuntimeArr(1)

For zNum = 2 To LHldr
If RuntimeArr(zNum) = "" Then
RuntimeArrCopy(0) = Int(RuntimeArrCopy(0)) + Int(OffsetArrCopy(0))
RuntimeArrCopy(1) = Int(RuntimeArrCopy(1)) + Int(OffsetArrCopy(1))
RuntimeArrCopy(2) = Int(RuntimeArrCopy(2)) + Int(OffsetArrCopy(2))
ThisRuntime = RuntimeArrCopy(0) & ":" & RuntimeArrCopy(1) & ":" & RuntimeArrCopy(2)
AutoRuntimeArr(zNum) = ThisRuntime
End If
Next

zNum = 1
For xNum = 1 To (LHldr)
Range("Runtime").Offset(zNum, 0).Value = AutoRuntimeArr(zNum)
zNum = zNum + 1
Next

SetupRtCntr:
'//Setup runtime counter...
If Range("Runtime").Offset(1, 0).Value <> "" Then
Range("RtCntr").Value = 0
    Else
        Range("RtCntr").Value = vbNullString
            End If


lastRw = Cells(Rows.Count, "Q").End(xlUp).Row

'//Clear UserLink space...
For xNum = 1 To lastRw
Range("Q" & xNum).ClearContents
Next

'//Print user to range from Linker...
zNum = 1

For xNum = 0 To (oUserBox.ListCount - 1)
oUserBox.ListIndex = xNum
iUser = oUserBox.List(xNum)
iUser = Replace(iUser, "(" & xNum + 1 & ") ", vbNullString)

Range("Userlink").Offset(zNum, 0).Value = iUser
zNum = zNum + 1
Next

'//
ETWEETXLPOST.ProfileListBox.Value = Range("MainLink").Offset(1, 0).Value

'//Save backup connection link to mtsett folder
Call App_TOOLS.SaveBackupLink

'//Connected (ready to send!)...
ETWEETXLPOST.xlFlowStrip.Value = "Finished connecting posts..."

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
xMsg = 17: Call App_MSG.AppMsg(xMsg)

End Sub
Public Sub ViewMedBtn_Clk()

'//For viewing post media

Call App_TOOLS.FindForm(xForm)

If xForm.MedLinkBox.Value <> "" Then

medLink = Range("MedScrollLink").Value

'//Check for spaces...
If InStr(1, medLink, " ") Then
    medLink = Replace(medLink, " ", """ """)
        End If

Call App_Loc.xShellWinFldr(xShellWin)

Open xShellWin & "view_med.bat" For Output As #1
Print #1, "@echo off"
Print #1, "start " & medLink
Print #1, "exit"
Close #1

Shell (xShellWin & "view_med.bat"), vbMinimizedNoFocus
'Kill (xShellWin & "view_med.bat")
    
End If

End Sub
Sub RmvAllProfiles_Clk()

Dim oFSO, oFldr, oSubFldr As Object

Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oFldr = oFSO.GetFolder(AppLoc & "\presets")

If Range("xlasSilent").Value = 1 Then msg = vbYes: GoTo SilentRun

msg = MsgBox("Are you sure you wish to remove all profiles?", vbYesNo, AppTag)

SilentRun:
If msg = vbYes Then

For Each oSubFldr In oFldr.SubFolders
xInfo = oSubFldr.name
Call App_CLICK.RmvProfile_Clk(xInfo)
Next

End If

End Sub
Public Function RmvProfile_Clk(xInfo)

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

    YesNo = MsgBox("[ " & xUser & " ]" & nl & nl & "Are you sure you want to remove this profile?", vbYesNo, AppTag)

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
Range("DataPullTrig").Value = 0
'//Import profile information
Call App_IMPORT.MyProfileData
ETWEETXLSETUP.ProfileListBox.Value = ETWEETXLSETUP.ProfileListBox.List(xPos)
                End If
                    End If '//yes
                    
                    
End Function
Public Function RmvUserBox_EnClk()

'//For removing a user from the Linker using the Enter key

On Error Resume Next

Dim AArr(5000), AArr2(5000), PArr(5000), PArr2(5000), UArr(5000), UArr2(5000), MArr(5000), MArr2(5000) As String
Dim xPos, I, X As Integer

If ETWEETXLPOST.UserBox.ListCount > 0 Then
ETWEETXLPOST.UserHdr.Caption = "User" & " (" & ETWEETXLPOST.UserBox.ListCount - 1 & ")"
End If

xPos = ETWEETXLPOST.UserBox.ListIndex

'//Get info...
X = 1
lastRw = Cells(Rows.Count, "AM").End(xlUp).Row
Do Until X > lastRw
If Range("apiLink").Offset(X, 0).Value <> vbNullString And _
Range("apiLink").Offset(X, 0).Value <> " " Then AArr(X) = Range("apiLink").Offset(X, 0).Value
If Range("Mainlink").Offset(X, 0).Value <> vbNullString And _
Range("Mainlink").Offset(X, 0).Value <> " " Then MArr(X) = Range("Mainlink").Offset(X, 0).Value
If Range("Passlink").Offset(X, 0).Value <> vbNullString And _
Range("Passlink").Offset(X, 0).Value <> " " Then PArr(X) = Range("Passlink").Offset(X, 0).Value
If Range("Userlink").Offset(X, 0).Value <> vbNullString And _
Range("Userlin").Offset(X, 0).Value <> " " Then UArr(X) = Range("Userlink").Offset(X, 0).Value
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
    

Call Cleanup.ClnUserSpace

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
If PArr2(I) <> vbNullString And PArr2(I) <> "*/HALT" Then Range("PassLink").Offset(X, 0).Value = PArr2(I): X = X + 1
I = I + 1
Loop

I = 1: X = 1
Do Until UArr(I) = "*/HALT"
If UArr2(I) <> vbNullString And UArr2(I) <> "*/HALT" Then Range("UserLink").Offset(X, 0).Value = UArr2(I): X = X + 1
I = I + 1
Loop
                
If xPos = 0 Then xPos = 1

If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.UserBox.List(xPos - 1) & " unlinked..."
End If

ETWEETXLPOST.UserBox.RemoveItem (xPos - 1)
xBox = 1: Call ResetBoxOrder(xBox)

End Function
Public Function RmvLinkerBox_EnClk()

'//For removing a draft from the Linker using the Enter key

On Error Resume Next

Dim PArr(5000), PArr2(5000), DArr(5000), DArr2(5000) As String
Dim xPos, I, X As Integer

If ETWEETXLPOST.LinkerBox.ListCount > 0 Then
ETWEETXLPOST.DraftHdr.Caption = "Draft" & " (" & ETWEETXLPOST.LinkerBox.ListCount - 1 & ")"
End If

xPos = ETWEETXLPOST.LinkerBox.ListIndex

X = 1
lastRw = Cells(Rows.Count, "P").End(xlUp).Row
Do Until X > lastRw
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
    

Call Cleanup.ClnDraftSpace

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

If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.LinkerBox.List(xPos - 1) & " unlinked..."
End If

ETWEETXLPOST.LinkerBox.RemoveItem (xPos - 1)
xBox = 2: Call ResetBoxOrder(xBox)

End Function
Public Sub RmvRuntime_EnClk()

'//For removing a time from the Linker using the Enter key

On Error Resume Next

Dim RArr(5000), RArr2(5000) As String
Dim xPos, I, X As Integer

If ETWEETXLPOST.RuntimeBox.ListCount > 0 Then
ETWEETXLPOST.RuntimeHdr.Caption = "Runtime" & " (" & ETWEETXLPOST.RuntimeBox.ListCount - 1 & ")"
End If

xPos = ETWEETXLPOST.RuntimeBox.ListIndex

X = 1
lastRw = Cells(Rows.Count, "R").End(xlUp).Row: If X = lastRw Then GoTo JustRmv '//runtime's are added after connection
Do Until X > lastRw
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
    

Call Cleanup.ClnRuntimeSpace

I = 1: X = 1
Do Until RArr(I) = "*/HALT"
If RArr2(I) <> vbNullString And RArr2(I) <> "*/HALT" Then Range("Runtime").Offset(X, 0).Value = RArr2(I): X = X + 1
I = I + 1
Loop

JustRmv:
If xPos = 0 Then xPos = 1

If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.RuntimeBox.List(xPos - 1) & " unlinked..."
End If

ETWEETXLPOST.RuntimeBox.RemoveItem (xPos - 1)
xBox = 4: Call ResetBoxOrder(xBox)

End Sub
Public Sub UpHrBtn_Clk(xTimes)

'//For adding hour to time box
If ETWEETXLPOST.TimeBox.Value = "" Then ETWEETXLPOST.TimeBox.Value = 0: ETWEETXLPOST.TimeBox.Value = ""

If xTimes < 0 Then xTimes = xTimes * -1

For X = 1 To xTimes

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
Public Sub DwnHrBtn_Clk(xTimes)

'//For subtracting hour from time box

If ETWEETXLPOST.TimeBox.Value = "" Then ETWEETXLPOST.TimeBox.Value = 0: ETWEETXLPOST.TimeBox.Value = ""

If xTimes < 0 Then xTimes = (xTimes * -1)

For X = 1 To xTimes

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
Public Sub UpMinBtn_Clk(xTimes)

'//For adding minute to time box

If ETWEETXLPOST.TimeBox.Value = "" Then ETWEETXLPOST.TimeBox.Value = 0: ETWEETXLPOST.TimeBox.Value = ""

If xTimes < 0 Then xTimes = (xTimes * -1)

For X = 1 To xTimes

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
Public Sub DwnMinBtn_Clk(xTimes)

'//For subtracting minute from time box

If ETWEETXLPOST.TimeBox.Value = "" Then ETWEETXLPOST.TimeBox.Value = 0: ETWEETXLPOST.TimeBox.Value = ""

If xTimes < 0 Then xTimes = (xTimes * -1)

For X = 1 To xTimes

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
Public Sub UpSecBtn_Clk(xTimes)

'//For adding second to time box

If ETWEETXLPOST.TimeBox.Value = "" Then ETWEETXLPOST.TimeBox.Value = 0: ETWEETXLPOST.TimeBox.Value = ""

If xTimes < 0 Then xTimes = (xTimes * -1)

For X = 1 To xTimes

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
Public Sub DwnSecBtn_Clk(xTimes)

'//For subtracting second from time box

If ETWEETXLPOST.TimeBox.Value = "" Then ETWEETXLPOST.TimeBox.Value = 0: ETWEETXLPOST.TimeBox.Value = ""

If xTimes < 0 Then xTimes = (xTimes * -1)

For X = 1 To xTimes

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
Public Sub DraftFilterBtn_Clk(xFil)

Retry:
Call App_TOOLS.FindForm(xForm)

On Error GoTo SetForm

If xFil = 1 Then
xType = 1
xForm.DraftFilterBtn.Caption = "..."
Range("DraftFilter").Value = 1
If Range("xlasWinForm") = 3 Then Call App_IMPORT.MyTweetData(xType)
    ElseIf xFil = 0 Then
    xType = 0
    Range("DraftFilter").Value = 0
    xForm.DraftFilterBtn.Caption = "•"
    If Range("xlasWinForm") = 3 Then Call App_IMPORT.MyTweetData(xType)
        End If
        
Exit Sub

SetForm:
Err.Clear
If Range("xlasWinForm").Value <> 3 Then Range("xlasWinForm").Value = 3 Else Range("xlasWinForm").Value = 4
GoTo Retry

End Sub
Public Sub FreezeBtn_Clk()

Call FindForm(xForm)

'//inactive = 0
If Range("AppActive").Value <> 0 Then

'//freeze = 1 (xlas = 0)
If Range("AppActive").Value = 1 Then
    Range("AppActive").Value = 2
    Call FreezeFX
    ETWEETXLHOME.xlFlowStrip.Value = "Application frozen..."
    ETWEETXLPOST.xlFlowStrip.Value = "Application frozen..."
    ETWEETXLQUEUE.xlFlowStrip.Value = "Application frozen..."
    ETWEETXLSETUP.xlFlowStrip.Value = "Application frozen..."
    '//unfreeze = 2 (xlas = 1)
        ElseIf Range("AppActive").Value = 2 Then
            Range("AppActive").Value = 1
            Call UnfreezeFX
            ETWEETXLHOME.xlFlowStrip.Value = "Application unfrozen..."
            ETWEETXLPOST.xlFlowStrip.Value = "Application unfrozen..."
            ETWEETXLQUEUE.xlFlowStrip.Value = "Application unfrozen..."
            ETWEETXLSETUP.xlFlowStrip.Value = "Application unfrozen..."
                End If
                    End If
                
                    
        
End Sub
Public Sub HelpStatus_Clk(xPos)

On Error GoTo ErrMsg

Call FindForm(xForm)

If xPos = 1 Then

'//help wizard on/active
    Range("HelpActive").Value = 1
    xForm.HelpStatus.Caption = "On"
    If InStr(1, xForm.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
    xForm.xlFlowStrip.Value = "Help wizard is active..."
    End If
    
    ElseIf xPos = 0 Then
    
'//help wizard off/inactive
    Range("HelpActive").Value = 0
    xForm.HelpStatus.Caption = "Off"
    If InStr(1, xForm.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
    xForm.xlFlowStrip.Value = "Help wizard is inactive..."
    End If
    
    End If

Exit Sub
ErrMsg:
xMsg = 28: Call App_MSG.AppMsg(xMsg)
End Sub
