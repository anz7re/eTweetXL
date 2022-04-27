Attribute VB_Name = "App_IMPORT"
'/############################\
'//Application Import Features\\
'///##########################\\\

Sub MyAPIData()

Call App_Loc.xApiFile(apiFile)

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
Public Function MyLink(xLink)

Dim xProfArr(5000), xProfArr2(5000), xUserArr(5000), xDraftArr(5000), xRuntimeArr(5000) As String
Dim ThisLink As String
Dim X  As Integer
X = 1

If Right(xLink, 1) = """" Then xLink = Left(xLink, Len(xLink) - 1) '//remove ending quote
If Left(xLink, 1) = " " Then xLink = Right(xLink, Len(xLink) - 1) '//remove leading space
    
On Error GoTo EndMacro

If xLink = vbNullString Then

'//Get file
ThisLink = Application.GetOpenFilename("link Files (*.link*), *.link*")

    Else
    
        ThisLink = xLink
        
                End If

If ThisLink <> vbNullString Then

If InStr(1, ThisLink, "False") = False Then Range("RemLink").Value = ThisLink Else GoTo EndMacro: If xLink <> vbNullString Then ThisLink = xLink

'//Open file and save data to array for output to Linker
Open ThisLink For Input As #1

Do Until EOF(1)
Line Input #1, xData
xLinkerArr = Split(xData, ",")
xProfArr(X) = xLinkerArr(0)
xUserArr(X) = xLinkerArr(1)
xProfArr2(X) = xLinkerArr(2)
xDraftArr(X) = xLinkerArr(3)
xRuntimeArr(X) = xLinkerArr(4)
X = X + 1
Loop
Close #1

Call App_TOOLS.xDisable

xCntr = 1
Do Until xCntr = X

'//Add user
ETWEETXLPOST.ProfileListBox.Value = xProfArr(xCntr)

'//Check for send w/ API
If InStr(1, xUserArr(xCntr), "(*api)") Then
ETWEETXLPOST.SendAPI.Value = True
xUserArr(xCntr) = Replace(xUserArr(xCntr), "(*api)", "")
    Else
        xUserArr(xCntr) = Replace(xUserArr(xCntr), "(*)", "")
        ETWEETXLPOST.SendAPI.Value = False
            End If
            
ETWEETXLPOST.UserListBox.Value = xUserArr(xCntr)
xPos = 0
Call App_CLICK.AddUser_Clk(xPos)

'//Add draft
ETWEETXLPOST.ProfileListBox.Value = xProfArr2(xCntr)
ETWEETXLPOST.DraftBox.Value = xDraftArr(xCntr)
xPos = 0
Call App_CLICK.AddLink_Clk(xPos)

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
Sub MyProfileNames()
        
Dim ProfDbArr(100) As String
Dim oForm, oFSO, oSubFldr, oFldr As Object
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFldr = oFSO.GetFolder(Env & AppPath & "\presets\")

xNum = 1

Call App_TOOLS.FindForm(xForm)

'//Get profile names...
For Each oSubFldr In oFldr.SubFolders

   ProfDbArr(xNum) = oSubFldr.name
 
   xNum = xNum + 1
 
   Next oSubFldr

'//Add to profile box...
xNum = 1
Do Until ProfDbArr(xNum) = ""
xForm.ProfileListBox.AddItem (ProfDbArr(xNum))
xNum = xNum + 1
Loop
    
End Sub
Sub MyProfileData()

If Range("DataPullTrig").Value = 0 Then

On Error GoTo EndMacro

Dim lastRw, X As Integer
Dim ProfDbArr(5000) As String
Dim oFSO, oSubFldr, oFldr, xForm As Object
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFldr = oFSO.GetFolder(Env & AppPath & "\presets\")

Call App_TOOLS.FindForm(xForm)

xNum = 1

'//Get profile names
For Each oSubFldr In oFldr.SubFolders

   ProfDbArr(xNum) = oSubFldr.name
 
   xNum = xNum + 1
 
   Next oSubFldr

'//Cleanup
Call Cleanup.ClnProfSpace
If Range("DataPullTrig").Value = 1 Then Exit Sub '//Handle triggering twice...
'//clear post boxes...
If Range("xlasWinForm").Value = 2 Or Range("xlasWinForm").Value = 3 Or Range("xlasWinForm").Value = 5 Then
lastRw = Cells(Rows.Count, "A").End(xlUp).Row
xProf = xForm.ProfileListBox.Value
xUser = xForm.UserListBox.Value
xForm.ProfileListBox.Clear
xForm.ProfileListBox.Value = xProf
xForm.UserListBox.Clear

'//check for user in current profile
For X = 1 To lastRw
If Range("Profile").Offset(X, 0).Value = xUser Then xForm.UserListBox.Value = xUser
Next
End If

'//Set first profile found active if none already set...
If Range("Profile").Value = "" Then
Range("Profile").Value = ProfDbArr(1)
End If

If Range("xlasWinForm").Value = 2 Or 3 Or 5 Then

'//Add profile names
xNum = 1
Do Until ProfDbArr(xNum) = ""
xForm.ProfileListBox.AddItem (ProfDbArr(xNum))
xNum = xNum + 1
Loop
   
'//Get personal information
Call App_Loc.xPersFile(persFile)
'///
xNum = 1

'//
TargetFile = persFile & xProf & ".pers"

'//check for pers file
If Dir(TargetFile) <> "" Then

lastRw = Cells(Rows.Count, "A").End(xlUp).Row

'//open .pers target file
Open TargetFile For Input As #1

On Error Resume Next

xCntr = 1
Do Until EOF(1)
Line Input #1, xPers
xPersArr = Split(xPers, ";")
'//Name...
If xPersArr(0) <> "" Then
Range("A" & lastRw + xCntr).Value = xPersArr(0)
    '//Lock enabled...
    If xPersArr(3) = "***" Then
    xForm.UserListBox.AddItem (xPersArr(0) & Range("Scure").Value)
    Range("U" & lastRw + xCntr).Value = 0 '//Ready lock
        Else
            xForm.UserListBox.AddItem (xPersArr(0))
            Range("U" & lastRw + xCntr).Value = 1 '//Set to unlock
                End If
    Else: GoTo SkipBlank
        End If
'//Pass...
Range("B" & lastRw + xCntr).Value = xPersArr(1)
'//Browser...
Range("C" & lastRw + xCntr).Value = xPersArr(2)
'//Lock...
Range("S" & lastRw + xCntr).Value = xPersArr(3)
'//Security...
Range("T" & lastRw + xCntr).Value = xPersArr(4)
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
Function MyPassData(xUser)

If Range("xlasWinForm").Value = 2 Then Set oForm = ETWEETXLSETUP
If Range("xlasWinForm").Value = 3 Then Set oForm = ETWEETXLPOST

lastRw = Cells(Rows.Count, "A").End(xlUp).Row

oForm.PassBox.Value = ""

For xNum = 1 To lastRw

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
    PassBoxHldr = Range("A" & xNum).Offset(0, 1).Value
    HideByHldr = Len(PassBoxHldr) + 7
                
    For xCntr = 1 To HideByHldr
    oForm.PassBox.Value = oForm.PassBox.Value & "*"
    Next
    
        End If
        
        '//Show
        If oForm.ShowPassBox.Value = True Then
            oForm.PassBox.Value = Range("A" & xNum).Offset(0, 1).Value
                End If
                
                    GoTo FoundPass
                
                        End If
  
Next
FoundPass:
End Function
Function MyPassData2(xProf)

lastRw = Cells(Rows.Count, "A").End(xlUp).Row

ETWEETXLSETUP.PassBox.Value = ""

For xNum = 1 To lastRw

If Range("A" & xNum).Value = xProf Then

'//Check for lock...
If Range("A" & xNum).Offset(0, 296).Value = "***" Then

        Range("Target2").Value = Range("A" & xNum).Offset(0, 297).Value
        Range("Ucure2").Value = Range("A" & xNum).Offset(0, 298).Address
        
        pLatch = Range("Ucure2").Value
        If Range(pLatch).Value = 1 Then GoTo SkipLock
    
        ETWEETXLSETUP.ShowPassBox.Value = False
        XLPINLOCK.Show
            End If

SkipLock:
'//Hide
If ETWEETXLSETUP.ShowPassBox.Value = False Then
    PassBoxHldr = Range("A" & xNum).Offset(0, 1).Value
    HideByHldr = Len(PassBoxHldr) + 7
                
    For xCntr = 1 To HideByHldr
    ETWEETXLSETUP.PassBox.Value = ETWEETXLSETUP.PassBox.Value & "*"
    Next
    
        End If
        
        '//Show
        If ETWEETXLSETUP.ShowPassBox.Value = True Then
            ETWEETXLSETUP.PassBox.Value = Range("A" & xNum).Offset(0, 1).Value
                End If
                
                    GoTo FoundPass
                
                        End If
  
Next
FoundPass:
End Function
Sub MyTweetData(xType)

Dim TweetDbArr(5000) As String
Dim xDraft, xExt, xFile, xThread, myPost, myPostHldr, myMedia, myMediaHldr As String
Dim oFSO, oFile, oFldr As Object
Dim ECntr As Byte

Set oFSO = CreateObject("Scripting.FileSystemObject")

Call App_TOOLS.xDisable

If xType = 0 Then Call App_Loc.xTwtFile(twtFile): xFile = twtFile: xExt = ".twt" ': Range("DraftFilter").Value = 0
If xType = 1 Then Call App_Loc.xThrFile(thrFile): xFile = thrFile: xExt = ".thr" ': Range("DraftFilter").Value = 1

'//Check for info
If Dir(xFile) = "" Then
Call Cleanup.ClnTwtSpace
ETWEETXLPOST.DraftBox.Clear
Exit Sub
End If

Set oFldr = oFSO.GetFolder(xFile)

'//Cleanup
Call Cleanup.ClnTwtSpace
xDraft = ETWEETXLPOST.DraftBox.Value
ETWEETXLPOST.DraftBox.Clear

xNum = 1
zNum = 2
'//Grab drafts
For Each oFile In oFldr.Files
    
    If InStr(1, oFile, xExt) Then
    TweetDbArr(xNum) = oFile.name
    Range("D" & zNum).Value = oFile.name
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
Range("E" & zNum).Value = myPost
Range("F" & zNum).Value = myMedia

SkipHere:
zNum = zNum + 1
Close #1
Next

End Sub
Sub MyNextQueue()

Dim RArr(5000), QArr(5000), UArr(5000) As String
Dim X, xNum As Integer

On Error GoTo SkipHere

'//runtime
lastR = Cells(Rows.Count, "R").End(xlUp).Row
'//mainlink
lastU = Cells(Rows.Count, "M").End(xlUp).Row
'//draftlink
lastQ = Cells(Rows.Count, "L").End(xlUp).Row

Call App_Loc.xMTTwt(mtTwt)

'//Get current linker position
LLCntr = Range("LinkerCount").Value
'//holder
LLHldr = LLCntr

'//Get queued tweets
xNum = 1
Open mtTwt For Input As #1
For X = 1 To (LLCntr - 1)
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
For X = 1 To (lastQ - 1)
If QArr(X) <> vbNullString Then ETWEETXLQUEUE.QueueBox.AddItem "(" & LLHldr & ") " & (QArr(X)): LLHldr = LLHldr + 1
Next

ETWEETXLQUEUE.QueueHdr.Caption = "Queued (" & ETWEETXLQUEUE.QueueBox.ListCount & ") "

'//Set current queued tweet
ETWEETXLQUEUE.CurrQueue = QArr(1)

RtHldr = lastR - (Range("LinkerCount").Value + 1)

'//Get runtime & user information
For X = LLCntr To (lastR)
ThisR = Range("Runtime").Offset(X, 0).Value
ThisU = Range("Userlink").Offset(X, 0).Value
ThisR = Format$(ThisR, "hh:mm:ss")
RArr(X) = ThisR
UArr(X) = ThisU
Next

'//Add queued runtime
For X = 1 To (lastR - 1)
If RArr(X) <> vbNullString Then ETWEETXLQUEUE.RuntimeBox.AddItem "(" & X & ") " & (RArr(X))
Next

ETWEETXLQUEUE.RuntimeHdr.Caption = "Runtime (" & ETWEETXLQUEUE.RuntimeBox.ListCount & ") "

'//Add queued users
For X = 1 To (lastU - 1)
If UArr(X) <> vbNullString Then ETWEETXLQUEUE.UserBox.AddItem "(" & X & ") " & (UArr(X))
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
Sub SelectedMedia()

Dim oForm As Object

'//Check xlFlowStrip Window...
If Range("xlasWinForm").Value = 3 Then Set oForm = ETWEETXLPOST
If Range("xlasWinForm").Value = 4 Then Set oForm = ETWEETXLQUEUE

On Error Resume Next

lastRw = Cells(Rows.Count, "I").End(xlUp).Row

Range("MedScrollPos").Value = Range("MedScrollPos").Value - 1

'//LEFT
If Range("MedScrollPos").Value < 0 Then Range("MedScrollPos").Value = lastRw - 1

MedLinkHldr = Range("MediaScroll").Offset(Range("MedScrollPos").Value)
MedLinkHldr = Replace(MedLinkHldr, """", "")

If Range("LoadLess") <> 1 Then Exit Sub

If Dir(MedLinkHldr) <> "" Then
    
    oForm.MedDemo.Picture = LoadPicture(MedLinkHldr)
    oForm.MedDemo.PictureSizeMode = fmPictureSizeModeStretch
    
        End If
        
        oForm.MedCt.Caption = Range("MedScrollPos").Value + 1
        Range("MedScrollLink").Value = MedLinkHldr
        
End Sub
Function SelectedTweet(xTwt)

On Error Resume Next

If xTwt <> vbNullString Then

'//remove numbered count
xTwtArr = Split(xTwt, ") ")
xTwt = xTwtArr(1)

Dim lastRw, X As Integer

'//Check xlFlowStrip Window...
Call App_TOOLS.FindForm(xForm)
        
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
                    
 

If xTwt <> vbNullString Then
        
'//check hover position to find where we're selecting from
If Range("HoverPos") = 44 Then
lastRw = Cells(Rows.Count, "P").End(xlUp).Row
For X = 1 To lastRw

If InStr(1, xTwt, Range("Draftlink").Offset(X, 0).Value) And X = xForm.LinkerBox.ListIndex + 1 Then
    
If Range("xlasWinForm").Value <> 1 Then
    If Range("xlasWinForm").Value <> 4 Then
        If xForm.ProfileListBox.Value <> Range("Profilelink").Offset(X, 0).Value Then _
           xForm.ProfileListBox.Value = Range("Profilelink").Offset(X, 0).Value
                GoTo FindDraft
                    Else
                    Call App_CLICK.RmvAllThread_Clk
                    Call Cleanup.ClnThrSpace
                       If xExt = ".twt" Then xType = 0
                       If xExt = ".thr" Then xType = 1
                       If ETWEETXLQUEUE.QueueBox.Value <> vbNullString Then X = ETWEETXLQUEUE.QueueBox.ListIndex + 1
                       ETWEETXLPOST.ProfileListBox.Value = Range("Profilelink").Offset(X, 0).Value: _
                       Call App_IMPORT.MyTweetData(xType): Call App_IMPORT.MyProfileData: GoTo FindDraft
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
If Range("DraftFilter").Value <> 1 Then xFil = 1: Call App_CLICK.DraftFilterBtn_Clk(xFil)
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
        
        Call App_CLICK.RmvAllThread_Clk
        Call Cleanup.ClnThrSpace
        
        If Range("DraftFilter").Value <> 1 Then xFil = 1: Call App_CLICK.DraftFilterBtn_Clk(xFil)
        
        '//set thread trigger
        xForm.ThreadTrig.Caption = 1
        
        X = 0
        Do Until X = UBound(xPostArr)
        
        xPostArr(X) = Replace(xPostArr(X), "{ENTER};", Chr(10))
        xPostArr(X) = Replace(xPostArr(X), "{SPACE};", " ")
        
        xForm.PostBox.Value = xPostArr(X)
        If UBound(xMedArr) > X Then xForm.MedLinkBox.Value = xMedArr(X)
        
        Call App_CLICK.AddThread_Clk
        
        If UBound(xMedArr) > X Then
        
        If InStr(1, xMedArr(X), """ """) Then
        MedArr = Split(xMedArr(X), """ """)
        
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
            Range("MediaScroll").Offset(xNum, 0) = xForm.MedLinkBox.Value
                End If
                    
                    End If
                    
                        X = X + 1
                            Loop
        
                                    Else
                                    
        Call App_CLICK.RmvAllThread_Clk
        Call Cleanup.ClnThrSpace
        
        '//clear thread trigger
        xForm.ThreadTrig.Caption = 0
        
        xForm.PostBox.Value = Replace(Range("Post").Offset(xNum, 0), "*-;", vbCrLf)
        xForm.MedLinkBox.Value = Replace(Range("Media").Offset(xNum, 0), "*-;", vbNullString)
        
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
            Range("MediaScroll").Offset(xNum, 0) = xForm.MedLinkBox.Value
                End If
                    
                    End If
                        End If
                            End If '//nothing
                        
End Function
Sub OBSText()

'Dim NPSong As String
'
'Call App_Loc.OBSDataLoc(OBSDataFol)
'
'Open OBSDataFol & "import.txt" For Input As #1
'Line Input #1, NPSong
'Close #1
'
'ETWEETXLPOST.DraftBox.Value = "import"
'ETWEETXLPOST.PostBox.Value = NPSong

End Sub
