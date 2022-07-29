Attribute VB_Name = "eTweetXL_GET"
'/############################\
'// Application Get Features  \\
'///##########################\\\

Sub getAPIData()

Call eTweetXL_LOC.xApiFile(apiFile)

If Dir(apiFile) <> "" Then

On Error Resume Next

Open apiFile For Input As #1

Line Input #1, apiKey
Line Input #1, apiSecret
Line Input #1, acctoken
Line Input #1, accSecret

Close #1

ETWEETXLAPISETUP.apiKeyBox.Value = apiKey
ETWEETXLAPISETUP.apiSecretBox.Value = apiSecret
ETWEETXLAPISETUP.accTokenBox.Value = acctoken
ETWEETXLAPISETUP.accSecretBox.Value = accSecret

End If

End Sub
Public Function getLink(ByVal xLink As String)

Dim xProfArr(5000) As String: Dim xProfArr2(5000) As String
Dim xUserArr(5000) As String: Dim xDraftArr(5000) As String
Dim xRuntimeArr(5000) As String
Dim fd As FileDialog
Dim ThisLink As String
Dim xCntr As Integer: Dim x As Integer
x = 1

If Right(xLink, 1) = """" Then xLink = Left(xLink, Len(xLink) - 1) '//remove ending quote
If Left(xLink, 1) = " " Then xLink = Right(xLink, Len(xLink) - 1) '//remove leading space
    
On Error GoTo EndMacro

If xLink = vbNullString Then

'//Get file
Set fd = Application.FileDialog(msoFileDialogFilePicker): fd.AllowMultiSelect = True: ThisLink = fd.Show

    Else
    
        ThisLink = xLink
        
                End If

If ThisLink <> vbNullString Then

If InStr(1, ThisLink, "False") = False Then Range("RemLink").Value2 = ThisLink Else GoTo EndMacro: If xLink <> vbNullString Then ThisLink = xLink

'//Open multiple links for output to Linker
If ThisLink = "-1" Then

Range("RemLink").Value2 = vbNullString

For xCntr = 1 To fd.SelectedItems.Count

ThisLink = fd.SelectedItems(xCntr)
         
Range("RemLink").Value2 = Range("RemLink").Value2 & "," & ThisLink

Open ThisLink For Input As #1
                
Do Until EOF(1)
Line Input #1, xData
xLinkerArr = Split(xData, ",")
xProfArr(x) = xLinkerArr(0)
xUserArr(x) = xLinkerArr(1)
xProfArr2(x) = xLinkerArr(2)
xDraftArr(x) = xLinkerArr(3)
xRuntimeArr(x) = xLinkerArr(4)
x = x + 1
Loop
Close #1
Next
    
    Else
 '//
'//Open single link for output to Linker
Open ThisLink For Input As #1

Do Until EOF(1)
Line Input #1, xData
xLinkerArr = Split(xData, ",")
xProfArr(x) = xLinkerArr(0)
xUserArr(x) = xLinkerArr(1)
xProfArr2(x) = xLinkerArr(2)
xDraftArr(x) = xLinkerArr(3)
xRuntimeArr(x) = xLinkerArr(4)
x = x + 1
Loop
Close #1
    '//
    End If
    
Call xlAppScript_xbas.disableWbUpdates

'//Record data to Linker
xCntr = 1
Do Until xCntr = x

'//Add profile
If ETWEETXLPOST.ProfileListBox.Value <> xProfArr(xCntr) Then ETWEETXLPOST.ProfileListBox.Value = xProfArr(xCntr)

'//Check for send w/ API
If InStr(1, xUserArr(xCntr), "(*api)") Then
ETWEETXLPOST.SendAPI.Value = True
xUserArr(xCntr) = Replace(xUserArr(xCntr), "(*api)", "")
    Else
        xUserArr(xCntr) = Replace(xUserArr(xCntr), "(*)", "")
        ETWEETXLPOST.SendAPI.Value = False
            End If
            
'//Add User
If ETWEETXLPOST.UserListBox.Value <> xUserArr(xCntr) Then ETWEETXLPOST.UserListBox.Value = xUserArr(xCntr)
xPos = 0
Call eTweetXL_CLICK.AddUserBtn_Clk(xPos)

'//Add draft
ETWEETXLPOST.ProfileListBox.Value = xProfArr2(xCntr)
ETWEETXLPOST.DraftBox.Value = xDraftArr(xCntr)
xPos = 0
Call eTweetXL_CLICK.AddLinkBtn_Clk(xPos)

'//Add runtime
ETWEETXLPOST.RuntimeHdr.Caption = "Runtime" & " (" & ETWEETXLPOST.RuntimeBox.ListCount + 1 & ")"
iTime = Replace(ETWEETXLPOST.RuntimeHdr.Caption, "Runtime", vbNullString)
iTime = iTime & " " & xRuntimeArr(xCntr)
Range("Runtime").Offset(xCntr, 0).Value = iTime
ETWEETXLPOST.RuntimeBox.AddItem (iTime)
xCntr = xCntr + 1
Loop

If Not InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) Then
ETWEETXLHOME.xlFlowStrip.Value = "Import successful..."
ETWEETXLPOST.xlFlowStrip.Value = "Import successful..."
ETWEETXLQUEUE.xlFlowStrip.Value = "Import successful..."
ETWEETXLSETUP.xlFlowStrip.Value = "Import successful..."
End If
   
Exit Function
End If

EndMacro:
ETWEETXLHOME.xlFlowStrip.Value = "Import unsuccessful..."
ETWEETXLPOST.xlFlowStrip.Value = "Import unsuccessful..."
ETWEETXLQUEUE.xlFlowStrip.Value = "Import unsuccessful..."
ETWEETXLSETUP.xlFlowStrip.Value = "Import unsuccessful..."

End Function
Sub getProfileNames()
        
Dim ProfDbArr(100) As String
Dim oForm, oFSO, oSubFldr, oFldr As Object
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFldr = oFSO.GetFolder(AppLoc & "\presets\")

xNum = 1

Call getWindow(xWin)

'//Get profile names...
For Each oSubFldr In oFldr.SubFolders

   ProfDbArr(xNum) = oSubFldr.name
 
   xNum = xNum + 1
 
   Next oSubFldr

'//Add to profile box...
xNum = 1
Do Until ProfDbArr(xNum) = ""
xWin.ProfileListBox.AddItem (ProfDbArr(xNum))
xNum = xNum + 1
Loop
    
End Sub
Sub getProfileData()

If Range("DataPullTrig").Value = 0 Then

On Error GoTo EndMacro

Dim lastR, x As Integer
Dim ProfDbArr(5000) As String
Dim oFSO, oSubFldr, oFldr, xWin As Object
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFldr = oFSO.GetFolder(AppLoc & "\presets\")

Call getWindow(xWin)

xNum = 1

'//Get profile names
For Each oSubFldr In oFldr.SubFolders

   ProfDbArr(xNum) = oSubFldr.name
 
   xNum = xNum + 1
 
   Next oSubFldr

'//Cleanup
Call clnProf
If Range("DataPullTrig").Value = 1 Then Exit Sub '//Handle triggering twice...
'//clear post boxes...
If Range("xlasWinForm").Value2 = 12 Or Range("xlasWinForm").Value2 = 13 Or Range("xlasWinForm").Value2 = 15 Then
lastR = Cells(Rows.Count, "A").End(xlUp).Row
xProf = xWin.ProfileListBox.Value
xUser = xWin.UserListBox.Value
xWin.ProfileListBox.Clear
xWin.ProfileListBox.Value = xProf
xWin.UserListBox.Clear

'//check for user in current profile
For x = 1 To lastR
If Range("Profile").Offset(x, 0).Value = xUser Then xWin.UserListBox.Value = xUser
Next
End If

'//Set first profile found active if none already set...
If Range("Profile").Value2 = "" Then
Range("Profile").Value2 = ProfDbArr(1)
End If

If Range("xlasWinForm").Value2 = 12 Or 13 Or 15 Then

'//Add profile names
xNum = 1
Do Until ProfDbArr(xNum) = ""
xWin.ProfileListBox.AddItem (ProfDbArr(xNum))
xNum = xNum + 1
Loop
   
'//Get personal information
Call eTweetXL_LOC.xPersFile(persFile)
'///
xNum = 1

'//
TargetFile = persFile & xProf & ".pers"

'//check for pers file
If Dir(TargetFile) <> "" Then

lastR = Cells(Rows.Count, "A").End(xlUp).Row

'//open .pers target file
Open TargetFile For Input As #1

On Error Resume Next

xCntr = 1
Do Until EOF(1)
Line Input #1, xPers
xPersArr = Split(xPers, ";")
'//Name...
If xPersArr(0) <> "" Then
Range("A" & lastR + xCntr).Value = xPersArr(0)
    '//Lock enabled...
    If xPersArr(3) = "***" Then
    xWin.UserListBox.AddItem (xPersArr(0) & Range("Scure").Value)
    Range("U" & lastR + xCntr).Value = 0 '//Ready lock
        Else
            xWin.UserListBox.AddItem (xPersArr(0))
            Range("U" & lastR + xCntr).Value = 1 '//Set to unlock
                End If
    Else: GoTo SkipBlank
        End If
'//Target...
Range("F" & lastR + xCntr).Value = xPersArr(1)
'//Browser...
Range("B" & lastR + xCntr).Value = xPersArr(2)
'//Lock...
Range("S" & lastR + xCntr).Value = xPersArr(3)
'//Security...
Range("T" & lastR + xCntr).Value = xPersArr(4)
SkipBlank:
xCntr = xCntr + 1
Loop

Close #1

    End If

        xNum = xNum + 1

                End If

                    End If


 '//Set data pull trigger...
 Range("DataPullTrig").Value = 1

    Exit Sub

'//Error...
EndMacro:
Err.Clear

End Sub
Function getTargetData(ByVal xUser As String) As String

If Range("xlasWinForm").Value2 = 12 Then Set oForm = ETWEETXLSETUP
If Range("xlasWinForm").Value2 = 13 Then Set oForm = ETWEETXLPOST

lastR = Cells(Rows.Count, "A").End(xlUp).Row

oForm.PassBox.Value = ""

For xNum = 1 To lastR

If Range("A" & xNum).Value = xUser Then

'//Check for lock...
If Range("A" & xNum).Offset(0, 18).Value = "***" Then

        Range("Target").Value = Range("A" & xNum).Offset(0, 19).Value
        Range("Ucure").Value = Range("A" & xNum).Offset(0, 20).Address
        
        pLatch = Range("Ucure").Value
        If Range(pLatch).Value = 1 Then GoTo SkipLock
    
        oForm.ShowPassBox.Value = False
        XLPINLOCK.Show
            End If

SkipLock:
'//Hide
If oForm.ShowPassBox.Value = False Then
    PassBoxHldr = Range("A" & xNum).Offset(0, 5).Value
    HideByHldr = Len(PassBoxHldr) + 7
                
    For xCntr = 1 To HideByHldr
    oForm.PassBox.Value = oForm.PassBox.Value & "*"
    Next
    
        End If
        
        '//Show
        If oForm.ShowPassBox.Value = True Then
            oForm.PassBox.Value = Range("A" & xNum).Offset(0, 5).Value
                End If
                
                    GoTo FoundPass
                
                        End If
  
Next
FoundPass:
End Function
Sub getPostData(ByVal xType As Byte)

Dim TweetDbArr(5000) As String
Dim xDraft, xExt, xFile, xThread, myPost, myPostHldr, myMedia, myMediaHldr As String
Dim oFSO, oFile, oFldr As Object
Dim ECntr As Byte

Set oFSO = CreateObject("Scripting.FileSystemObject")

Call xlAppScript_xbas.disableWbUpdates

If xType = 0 Then Call eTweetXL_LOC.xTwtFile(twtFile): xFile = twtFile: xExt = ".twt" ': Range("DraftFilter").Value = 0
If xType = 1 Then Call eTweetXL_LOC.xThrFile(thrFile): xFile = thrFile: xExt = ".thr" ': Range("DraftFilter").Value = 1

'//Check for info
If Dir(xFile) = "" Then
Call clnTwt
ETWEETXLPOST.DraftBox.Clear
Exit Sub
End If

Set oFldr = oFSO.GetFolder(xFile)

'//Cleanup
Call clnTwt
xDraft = ETWEETXLPOST.DraftBox.Value
ETWEETXLPOST.DraftBox.Clear

xNum = 1
zNum = 2
'//Grab drafts
For Each oFile In oFldr.Files
    
    If InStr(1, oFile, xExt) Then
    TweetDbArr(xNum) = oFile.name
    Range("C" & zNum).Value = oFile.name
    ETWEETXLPOST.DraftBox.AddItem (Replace(TweetDbArr(xNum), xExt, ""))
    xNum = xNum + 1
    zNum = zNum + 1
    End If
    xNumHldr = xNum
 
    Next oFile

On Error Resume Next

zNum = 2
For xNum = 1 To xNumHldr

Open xFile & TweetDbArr(xNum) For Input As #1

'//Grab post
myPostHldr = vbNullString
myPost = vbNullString
myMedia = vbNullString
ECntr = 0

postAsThread:

Do Until myPostHldr = "*-;"
'//End if unable to find text after 50 lines
If ECntr > 50 Then GoTo NextStep
Line Input #1, myPostHldr
If myPostHldr <> vbNullString Then
    If myPostHldr <> "*-;" Then
    myPost = myPost & myPostHldr
        End If
            Else
                ECntr = ECntr + 1
                    End If
                        Loop

NextStep:
myPost = myPost & "*-;"

'//Grab media
Line Input #1, myMediaHldr
If myPost & myMediaHldr = vbNullString Then GoTo SkipHere
myMediaHldr = Replace(myMediaHldr, "*-", vbNullString)
myMedia = myMedia & myMediaHldr & "*-;"

'//Thread check
Line Input #1, xThread

If InStr(1, xThread, "*-(") Then
xThread = vbNullString: myPostHldr = vbNullString
GoTo postAsThread: End If

'//Send to range
Range("D" & zNum).Value = myPost
Range("E" & zNum).Value = myMedia

SkipHere:
zNum = zNum + 1
Close #1
Next

End Sub
Sub getQueueData()

Dim RArr(5000), QArr(5000), UArr(5000) As String
Dim x, xNum As Integer

On Error GoTo SkipHere

'//Runtime
lastR = Cells(Rows.Count, "R").End(xlUp).Row
'//MainLink
lastU = Cells(Rows.Count, "M").End(xlUp).Row
'//DraftLink
lastQ = Cells(Rows.Count, "L").End(xlUp).Row

Call eTweetXL_LOC.xMTTwt(mtTwt)

'//Get current linker position
LLCntr = Range("LinkerCount").Value2
'//holder
LLHldr = LLCntr

'//Get queued tweets
xNum = 1
Open mtTwt For Input As #1
For x = 1 To (LLCntr - 1)
Line Input #1, NextTweetPath
Next
Do Until EOF(1)
Line Input #1, NextTweetPath
If InStr(1, NextTweetPath, "\twt\") Then NmArr = Split(NextTweetPath, "\twt\")
If InStr(1, NextTweetPath, "\thr\") Then NmArr = Split(NextTweetPath, "\thr\")
NextTweetName = NmArr(1)
NextTweetName = Replace(NextTweetName, ".twt", " [•]")
NextTweetName = Replace(NextTweetName, ".thr", " [...]")
QArr(xNum) = NextTweetName: xNum = xNum + 1
Loop
Close #1

'//Add queued tweet
For x = 1 To (lastQ - 1)
If QArr(x) <> vbNullString Then ETWEETXLQUEUE.QueueBox.AddItem "(" & LLHldr & ") " & (QArr(x)): LLHldr = LLHldr + 1
Next

ETWEETXLQUEUE.QueueHdr.Caption = "Queued (" & ETWEETXLQUEUE.QueueBox.ListCount & ") "

'//Set current queued tweet
ETWEETXLQUEUE.CurrQueue = QArr(1)

RtHldr = lastR - (Range("LinkerCount").Value2 + 1)

'//Get runtime & user information
For x = LLCntr To (lastR)
ThisR = Range("Runtime").Offset(x, 0).Value
ThisU = Range("UserLink").Offset(x, 0).Value
ThisR = Format$(ThisR, "hh:mm:ss")
RArr(x) = ThisR
UArr(x) = ThisU
Next

'//Add queued runtime
For x = 1 To (lastR - 1)
If RArr(x) <> vbNullString Then ETWEETXLQUEUE.RuntimeBox.AddItem "(" & x & ") " & (RArr(x))
Next

ETWEETXLQUEUE.RuntimeHdr.Caption = "Runtime (" & ETWEETXLQUEUE.RuntimeBox.ListCount & ") "

'//Add queued users
For x = 1 To (lastU - 1)
If UArr(x) <> vbNullString Then ETWEETXLQUEUE.UserBox.AddItem "(" & x & ") " & (UArr(x))
Next

ETWEETXLQUEUE.UserHdr.Caption = "User (" & ETWEETXLQUEUE.UserBox.ListCount & ") "

'//Set current runtime
ETWEETXLQUEUE.CurrRuntime.Value = RArr(LLCntr)

'//Open next tweet in queue
Open NextTweetPath For Input As #1

'//Grab post
myPostHldr = 0
myPost = ""
myMedia = ""
ECntr = 0

Do Until myPostHldr = "*-;"
'//End if unable to find text after 50 lines
If ECntr > 50 Then GoTo NextStep
Line Input #1, myPostHldr
If myPostHldr <> vbNullString Then
    If myPostHldr <> "*-;" Then
    myPost = myPost & myPostHldr
        End If
            Else
                ECntr = ECntr + 1
                    End If
                        Loop

NextStep:
myPost = myPost & "*-;"

'//Grab media
Line Input #1, myMedia
If myPost & myMedia = "" Then GoTo SkipHere

'//Remove spacers
myMedia = Replace(myMedia, "*-", "")
myPost = Replace(myPost, "*-;", "")

'//Send to window
ETWEETXLQUEUE.PostBox.Value = myPost
ETWEETXLQUEUE.MedLinkBox.Value = myMedia

SkipHere:
Close #1

End Sub
Sub getSelMedia()

Dim oForm As Object

'//Check xlFlowStrip Window...
If Range("xlasWinForm").Value2 = 13 Then Set oForm = ETWEETXLPOST
If Range("xlasWinForm").Value2 = 14 Then Set oForm = ETWEETXLQUEUE

On Error Resume Next

lastR = Cells(Rows.Count, "I").End(xlUp).Row

Range("MedScrollPos").Value2 = Range("MedScrollPos").Value2 - 1

'//LEFT
If Range("MedScrollPos").Value2 < 0 Then Range("MedScrollPos").Value2 = lastR - 1

MedLinkHldr = Range("MediaScroll").Offset(Range("MedScrollPos").Value2)
MedLinkHldr = Replace(MedLinkHldr, """", "")

If Range("LoadLess").Value2 = 1 Then Exit Sub

If Dir(MedLinkHldr) <> "" Then
    
    oForm.MedDemo.Picture = LoadPicture(MedLinkHldr)
    oForm.MedDemo.PictureSizeMode = fmPictureSizeModeStretch
    
        End If
        
        oForm.MedCt.Caption = Range("MedScrollPos").Value2 + 1
        Range("MedScrollLink").Value2 = MedLinkHldr
        
End Sub
Function getSelPost(ByVal xTwt As String)

On Error Resume Next

If xTwt <> vbNullString Then

'//remove numbered count
xTwtArr = Split(xTwt, ") ")
xTwt = xTwtArr(1)

Dim lastR, x As Integer

'//Check xlFlowStrip Window...
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
                    
 

If xTwt <> vbNullString Then
        
'//check hover position to find where we're selecting from
If Range("HoverPos") = 44 Then
lastR = Cells(Rows.Count, "P").End(xlUp).Row
For x = 1 To lastR

If InStr(1, xTwt, Range("Draftlink").Offset(x, 0).Value) And x = xWin.LinkerBox.ListIndex + 1 Then
    
If Range("xlasWinForm").Value2 <> 1 Then
    If Range("xlasWinForm").Value2 <> 4 Then
        If xWin.ProfileListBox.Value <> Range("ProfileLink").Offset(x, 0).Value Then _
           xWin.ProfileListBox.Value = Range("ProfileLink").Offset(x, 0).Value
                GoTo FindDraft
                    Else
                    Call eTweetXL_CLICK.RmvAllThreadBtn_Clk
                    Call clnThr
                       If xExt = ".twt" Then xType = 0
                       If xExt = ".thr" Then xType = 1
                       If ETWEETXLQUEUE.QueueBox.Value <> vbNullString Then x = ETWEETXLQUEUE.QueueBox.ListIndex + 1
                       ETWEETXLPOST.ProfileListBox.Value = Range("ProfileLink").Offset(x, 0).Value: _
                       Call eTweetXL_GET.getPostData(xType): Call eTweetXL_GET.getProfileData: GoTo FindDraft
                            End If
                                End If
                                    End If
                                        Next
                                            End If
                                                End If
                            

FindDraft:

xNum = 1
xTwt = Replace(xTwt, xT, vbNullString)
Do Until Range("TweetLoc").Offset(xNum, 0).Value = xTwt & xExt '//Find draft name on sheet
xNum = xNum + 1
If xNum > 5000 Then
If Range("DraftFilter").Value = 0 Then xFil = 1
If Range("DraftFilter").Value = 1 Then xFil = 0
Call eTweetXL_CLICK.DraftFilterBtn_Clk(xFil)
Range("AppErrRef").Value = Range("AppErrRef").Value + 1
If Range("AppErrRef").Value <= 1 Then GoTo FindDraft Else Range("AppErrRef").Value = 0
GoTo SkipHere
End If
Loop
'//
SkipHere:

If Range("TweetLoc").Offset(xNum, 0).Value = xTwt & xExt Then

        '//split post & media & place in threaded loc
        
        xPostArr = Split(Range("Post").Offset(xNum, 0).Value, "*-;")
        xMedArr = Split(Range("Media").Offset(xNum, 0).Value, "*-;")
        
        If UBound(xPostArr) > 1 Then
        
        Call eTweetXL_CLICK.RmvAllThreadBtn_Clk
        Call clnThr
        
        If Range("DraftFilter").Value <> 1 Then xFil = 1: Call eTweetXL_CLICK.DraftFilterBtn_Clk(xFil)
        
        '//set thread trigger
        xWin.ThreadTrig.Caption = 1
        
        x = 0
        Do Until x = UBound(xPostArr)
        
        xPostArr(x) = Replace(xPostArr(x), "{ENTER};", Chr(10))
        xPostArr(x) = Replace(xPostArr(x), "{SPACE};", " ")
        
        xWin.PostBox.Value = xPostArr(x)
        If UBound(xMedArr) > x Then xWin.MedLinkBox.Value = xMedArr(x)
        
        Call eTweetXL_CLICK.AddThreadBtn_Clk
        
        If UBound(xMedArr) > x Then
        
        If InStr(1, xMedArr(x), """ """) Then
        MedArr = Split(xMedArr(x), """ """)
        
        For xNum = 0 To UBound(MedArr)
        MedArr(xNum) = Replace(MedArr(xNum), """", "")
        Range("MediaScroll").Offset(xNum, 0) = MedArr(xNum)
        If InStr(1, MedArr(xNum), ".gif") Then
            Range("GifCntr").Value = 1
            ElseIf InStr(1, MedArr(xNum), ".mp4") Then
            Range("VidCntr").Value = 1
                ElseIf InStr(1, MedArr(xNum), ".mov") Then
                Range("VidCntr").Value = 1
                    ElseIf InStr(1, MedArr(xNum), ".flv") Then
                    Range("VidCntr").Value = 1
                        End If
        Next
            Else
            xNum = 0
            Range("MediaScroll").Offset(xNum, 0) = xWin.MedLinkBox.Value
                End If
                    
                    End If
                    
                        x = x + 1
                            Loop
        
                                    Else
                                    
        Call eTweetXL_CLICK.RmvAllThreadBtn_Clk
        Call clnThr
        
        '//clear thread trigger
        xWin.ThreadTrig.Caption = 0
        
        xWin.PostBox.Value = Replace(Range("Post").Offset(xNum, 0), "*-;", vbCrLf)
        xWin.MedLinkBox.Value = Replace(Range("Media").Offset(xNum, 0), "*-;", vbNullString)
        
        If InStr(1, Range("Media").Offset(xNum, 0), """ """) Then
        MedArr = Split(Range("Media").Offset(xNum, 0), """ """)
        
        For xNum = 0 To UBound(MedArr)
        MedArr(xNum) = Replace(MedArr(xNum), """", "")
        Range("MediaScroll").Offset(xNum, 0) = MedArr(xNum)
        If InStr(1, MedArr(xNum), ".gif") Then
            Range("GifCntr").Value = 1
            ElseIf InStr(1, MedArr(xNum), ".mp4") Then
            Range("VidCntr").Value = 1
                ElseIf InStr(1, MedArr(xNum), ".mov") Then
                Range("VidCntr").Value = 1
                    ElseIf InStr(1, MedArr(xNum), ".flv") Then
                    Range("VidCntr").Value = 1
                        End If
        Next
            Else
            xNum = 0
            Range("MediaScroll").Offset(xNum, 0) = xWin.MedLinkBox.Value
                End If
                    
                    End If
                        End If
                            End If '//nothing
                        
End Function
Function getExt(xExt)

xExtArr = Split(xExt, ".")

xExt = xExtArr(1)

End Function
Function getDate()

getDate = Date
getDate = Replace(getDate, "/", "_")

End Function
Sub getRtState()
 
If InStr(1, Range("RtState").Value, "Starting automation...") Then
Range("ActiveOffset").Value2 = Range("Offset").Value2
GoTo PrintErr
End If

If InStr(1, Range("RtState").Value, "Trying to resolve the issue... Please wait...") Then
Range("ActiveOffset").Value2 = 5000
GoTo PrintErr
End If

Exit Sub

PrintErr:
ETWEETXLHOME.xlFlowStrip.Value = Range("RtState").Value
ETWEETXLPOST.xlFlowStrip.Value = Range("RtState").Value
ETWEETXLQUEUE.xlFlowStrip.Value = Range("RtState").Value
ETWEETXLSETUP.xlFlowStrip.Value = Range("RtState").Value

End Sub

