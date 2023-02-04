VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ETWEETXLPOST 
   Caption         =   "eTweetXL"
   ClientHeight    =   11220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17280
   OleObjectBlob   =   "ETWEETXLPOST.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ETWEETXLPOST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

'//Check for active app
If Range("AppState").Value2 = 0 Then
Range("ConnectTrig").Value2 = 0
End If

'//Welcome...
If Me.xlFlowStrip.Value = "" Then Me.xlFlowStrip.Value = eTweetXL_INFO.AppWelcome
            
'//Clear linker space
If Range("ConnectTrig").Value2 = 0 Then
        Call clnLinker
        Call clnSpec
            ETWEETXLPOST.LinkerBox.Clear
            ETWEETXLPOST.RuntimeBox.Clear
            ETWEETXLPOST.UserBox.Clear
            End If
            
'//Reset Send API
Range("SendAPI").Value2 = 0

'//Dynamic offset...
If Range("DynOffsetTrig").Value2 = 1 Then
ETWEETXLPOST.DynOffset.Value = True
    Else
         ETWEETXLPOST.DynOffset.Value = False
            End If
            
'//Refresh media scroll...
Range("MedScrollPos").Value2 = 0

Range("DataPullTrig").Value2 = 0
ETWEETXLPOST.OffsetBox.Value = "00:00:00"

End Sub
Private Sub UserForm_Activate()

'//Cleanup
Me.ProfileListBox.Clear
Me.DraftBox.Clear

'//Show runtime action message
Call eTweetXL_GET.getRtState

'//WinForm #
xWin = 13: Call setWindow(xWin)

'//Update application state
Call eTweetXL_TOOLS.updAppState

'//Reset triggers
Range("AlignTrig").Value2 = 1: AlignLink.Value = True
Range("UserTrig").Value2 = 0: SwapUser.ForeColor = vbBlack: AddUserA.ForeColor = vbBlack: AddUserB.ForeColor = vbBlack
Range("DraftTrig").Value2 = 0: SwapDraft.ForeColor = vbBlack: AddDraftA.ForeColor = vbBlack: AddDraftB.ForeColor = vbBlack
Range("TimeTrig").Value2 = 0: SwapTime.ForeColor = vbBlack: AddTimeA.ForeColor = vbBlack: AddTimeB.ForeColor = vbBlack

'//Import profile names
Call eTweetXL_GET.getProfileNames

xFil = 0: Call eTweetXL_CLICK.DraftFilterBtn_Clk(xFil)

'//Window name
If Me.xlFlowStrip.Value = vbNullString Or Range("AppState").Value2 <> 1 Then Me.xlFlowStrip.Value = "Tweet Setup..."

End Sub
Private Sub LinkerBox_Change()

If Range("ConnectStatus").Value2 <> 1 Then

On Error Resume Next

If Range("AlignTrig").Value2 = 1 Then
If Range("EditStatus").Value2 <= 0 Then
RuntimeBox.Selected(LinkerBox.ListIndex) = True
UserBox.Selected(LinkerBox.ListIndex) = True
End If
    End If
    
Range("LinkerIndex").Value2 = LinkerBox.ListIndex

End If

End Sub
Private Sub RuntimeBox_Change()

If Range("ConnectStatus").Value2 <> 1 Then

On Error Resume Next

If Range("AlignTrig").Value2 = 1 Then
If Range("EditStatus").Value2 <= 0 Then
LinkerBox.Selected(RuntimeBox.ListIndex) = True
UserBox.Selected(RuntimeBox.ListIndex) = True
End If
    End If

Range("LinkerIndex").Value2 = RuntimeBox.ListIndex

End If

End Sub
Private Sub UserBox_Change()

If Range("ConnectStatus").Value2 <> 1 Then

On Error Resume Next

If Range("AlignTrig").Value2 = 1 Then
If Range("EditStatus").Value2 <= 0 Then
LinkerBox.Selected(UserBox.ListIndex) = True
RuntimeBox.Selected(LinkerBox.ListIndex) = True
End If
    End If

Range("LinkerIndex").Value2 = UserBox.ListIndex

End If

End Sub
Private Sub HelpIcon_Click()

If Range("HelpStatus").Value2 = 0 And HelpIcon.Caption = "Off" Then Exit Sub
If Range("HelpStatus").Value2 = 1 And HelpIcon.Caption = "On" Then Exit Sub

If Range("HelpStatus").Value2 = 0 Then xPos = 1 Else xPos = 0
If xPos = vbNullString Then Range("HelpStatus").Value2 = 0: xPos = 1

Call eTweetXL_CLICK.HelpStatusBtn_Clk(xPos)

End Sub
Private Sub HelpStatus_Click()

If Range("HelpStatus").Value2 = 0 And HelpIcon.Caption = "Off" Then Exit Sub
If Range("HelpStatus").Value2 = 1 And HelpIcon.Caption = "On" Then Exit Sub

If Range("HelpStatus").Value2 = 0 Then xPos = 1 Else xPos = 0
If xPos = vbNullString Then Range("HelpStatus").Value2 = 0: xPos = 1

Call eTweetXL_CLICK.HelpStatusBtn_Clk(xPos)

End Sub
Private Sub FreezeBtn_Click()

Call eTweetXL_CLICK.FreezeBtn_Clk

End Sub
Private Sub StartBtn_Click()

Call eTweetXL_CLICK.StartBtn_Clk

End Sub
Private Sub BreakBtn_Click()

eTweetXL_CLICK.BreakBtn_Clk

End Sub

Private Sub CtrlBoxBtn_Click()

Call eTweetXL_FOCUS.shw_CTRLBOX

End Sub
Private Sub QueueBtn_Click()

Me.Hide
Call eTweetXL_FOCUS.shw_ETWEETXLQUEUE

End Sub
Private Sub PostSetupBtn_Click()

Me.Hide
Call eTweetXL_FOCUS.shw_ETWEETXLPOST

End Sub
Private Sub ProfileSetupBtn_Click()

Me.Hide
Call eTweetXL_FOCUS.shw_ETWEETXLSETUP

End Sub
Private Sub HideBtn_Click()

Call eTweetXL_CLICK.HideBtn_Clk

End Sub
Private Sub DraftsHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Art = "<lib>xtwt;winform(13);add.draft(*);$" '//xlas
Call xlas(Art)
        
End Sub
Private Sub DraftHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        
Call eTweetXL_CLICK.DraftHdr_Clk
        
End Sub
Private Sub DelAllDraftsBtn_Click()

Call eTweetXL_CLICK.DelAllDraftsBtn_Clk

End Sub
Private Sub ClrSetupBtn_Click()

Call eTweetXL_CLICK.ClrSetupBtn_Clk

End Sub
Private Sub LoadLinkerBtn_Click()

Range("LoadLess").Value = 1
xLink = ""
Call eTweetXL_GET.getLink(xLink)
Range("LoadLess").Value = 0

End Sub
Private Sub LastLinkBtn_Click()

Range("LoadLess").Value = 1
xLink = AppLoc & "\mtsett\lastlink.link"
Call eTweetXL_GET.getLink(xLink)
Range("LoadLess").Value = 0

End Sub
Private Sub LinkerHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call eTweetXL_CLICK.LinkerHdr_Clk

End Sub
Private Sub ReloadLinkerBtn_Click()

Range("LoadLess").Value = 1
xLink = Range("RemLink").Value

If InStr(1, xLink, ",") Then

Dim xLinkArr() As String
Dim X As Integer

xLinkArr = Split(xLink, ",")

For X = 1 To UBound(xLinkArr)
xLink = xLinkArr(X)
Call eTweetXL_GET.getLink(xLink)
Next

    Else

    Call eTweetXL_GET.getLink(xLink)
    Range("LoadLess").Value = 0
        
        End If
        
End Sub
Private Sub RuntimeHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        
Call eTweetXL_CLICK.RuntimeHdr_Clk
            
End Sub
Private Sub OffsetHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

ETWEETXLPOST.OffsetBox.Value = "00:00:00"

If Range("xlasSilent").Value2 <> 1 Then _
ETWEETXLPOST.xlFlowStrip.Value = "Offset refreshed..."
        
End Sub
Private Sub PostHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'//Clear post box
Call eTweetXL_CLICK.PostHdr_Clk

End Sub
Private Sub SaveLinkerBtn_Click()

If Range("ConnectTrig").Value2 = 1 Then
Call eTweetXL_CLICK.SaveLinkerBtn_Clk
    Else
        xMsg = 6: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
            End If

End Sub
Private Sub AlignLink_Click()

If Range("AlignTrig").Value = 0 And AlignLink.Value = False Then Exit Sub
If Range("AlignTrig").Value = 1 And AlignLink.Value = True Then Exit Sub

If Range("AlignTrig").Value = 1 Then xPos = 1 Else xPos = 0
If xPos = vbNullString Then Range("AlignTrig").Value = 1: xPos = 0

Call eTweetXL_CLICK.AlignLink_Clk(xPos)

End Sub
Private Sub SendAPI_Click()

If Range("SendAPI").Value = 0 And SendAPI.Value = False Then Exit Sub
If Range("SendAPI").Value = 1 And SendAPI.Value = True Then Exit Sub

If Range("SendAPI").Value = 0 Then xPos = 1 Else xPos = 0
If xPos = vbNullString Then Range("SendAPI").Value = 0: xPos = 1

Call eTweetXL_CLICK.SendAPI_Clk(xPos)

End Sub
Private Sub UserBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call eTweetXL_CLICK.RmvUserBox_DelClk

End Sub
Private Sub UserHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        
Call eTweetXL_CLICK.UserHdr_Clk
        
End Sub
Private Sub AddDraftBtn_Click()

eTweetXL_CLICK.AddDraftBtn_Clk

End Sub
'//ADD RUNTIME BUTTON
Private Sub AddRuntimeBtn_Click()

xPos = 0
Call eTweetXL_CLICK.AddRuntimeBtn_Clk(xPos)

End Sub
Private Sub AddRuntimeBtn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode.Value = 38 Then
UpHrDwnBtn.SetFocus
End If

End Sub
Private Sub MedDemoScroll_SpinDown()

On Error Resume Next

lastR = Cells(Rows.Count, "I").End(xlUp).Row

Range("MedScrollPos").Value2 = Range("MedScrollPos").Value2 - 1

'//LEFT
If Range("MedScrollPos").Value2 < 0 Then Range("MedScrollPos").Value2 = lastR - 1

MedLinkHldr = Range("MediaScroll").Offset(Range("MedScrollPos").Value2)
MedLinkHldr = Replace(MedLinkHldr, """", "")

If MedLinkHldr <> "" Then

If Dir(MedLinkHldr) <> "" Then
    MedDemo.Picture = LoadPicture(MedLinkHldr)
    MedDemo.PictureSizeMode = fmPictureSizeModeStretch
        
        End If
            Else
                MedDemo.Picture = Nothing
                    End If
        
        MedCt.Caption = Range("MedScrollPos").Value2 + 1
        Range("MedScrollLink").Value2 = MedLinkHldr

End Sub
Private Sub MedDemoScroll_SpinUp()

On Error Resume Next

lastR = Cells(Rows.Count, "I").End(xlUp).Row

Range("MedScrollPos").Value2 = Range("MedScrollPos").Value2 + 1

'//RIGHT
If Range("MedScrollPos").Value2 > 3 Then Range("MedScrollPos").Value2 = 0

MedLinkHldr = Range("MediaScroll").Offset(Range("MedScrollPos").Value2)
MedLinkHldr = Replace(MedLinkHldr, """", "")

If MedLinkHldr <> "" Then

If Dir(MedLinkHldr) <> "" Then
    MedDemo.Picture = LoadPicture(MedLinkHldr)
    MedDemo.PictureSizeMode = fmPictureSizeModeStretch
        
        End If
            Else
                MedDemo.Picture = Nothing
                    End If

        MedCt.Caption = Range("MedScrollPos").Value2 + 1
        Range("MedScrollLink").Value2 = MedLinkHldr

End Sub
Private Sub MedLinkBox_Change()

If Range("LoadLess") = 1 Then Exit Sub

On Error Resume Next

If MedLinkBox.SpecialEffect <> fmSpecialEffectSunken Then
MedLinkBox.SpecialEffect = fmSpecialEffectSunken
End If

MediaHldr = MedLinkBox.Value

'//media check
If InStr(1, MediaHldr, """ """) Then

MedArr = Split(MediaHldr, """ """)
For X = 0 To UBound(MedArr)
Range("MediaScroll").Offset(X, 0).Value2 = MedArr(X)
    Next
    
    MediaHldr = MedArr(0)
    
        Else
    
            MediaHldr = MedLinkBox.Value
    
                End If
        
MediaHldr = Replace(MediaHldr, """", "")

If Dir(MediaHldr) <> "" Then

    MedCt.Caption = 1
    MedDemo.Picture = LoadPicture(MediaHldr)
    MedDemo.PictureSizeMode = fmPictureSizeModeStretch
        End If

Range("MedScrollLink").Value2 = MediaHldr
  
End Sub
Private Sub PostThreadScroll_SpinDown()

On Error Resume Next

lastR = Cells(Rows.Count, "Y").End(xlUp).Row
If lastR < 1 Then lastR = Cells(Rows.Count, "Z").End(xlUp).Row

Call clnMediaScroll

Range("ThreadScrollPos").Value2 = Range("ThreadScrollPos").Value2 - 1

'//LEFT
If Range("ThreadScrollPos").Value2 <= 0 Then Range("ThreadScrollPos").Value2 = lastR - 1

ThreadHldr = Range("PostThread").Offset(Range("ThreadScrollPos").Value2)
MediaHldr = Range("MedThread").Offset(Range("ThreadScrollPos").Value2)

'//media check
If InStr(1, MediaHldr, """ """) Then

MedArr = Split(MediaHldr, """ """)
For X = 0 To UBound(MedArr)
Range("MediaScroll").Offset(X, 0).Value2 = MedArr(X)
    Next
        Else
            Range("MediaScroll").Offset(0, 0).Value2 = MediaHldr
                End If

If ThreadHldr <> vbNullString Then
PostBox.Value = ThreadHldr
If Left(MediaHldr, Len(MediaHldr) - Len(MediaHldr) + 1) = """" Then MediaHldr = Left(MediaHldr, Len(MediaHldr) - 1)
If Right(MediaHldr, Len(MediaHldr) - Len(MediaHldr) - 1) = """" Then MediaHldr = Right(MediaHldr, Len(MediaHldr) - 1)
MedLinkBox.Value = MediaHldr
    Else
        PostBox.Value = vbNullString
        MedLinkBox.Value = vbNullString
            End If
        
ThreadCt.Caption = Range("ThreadScrollPos").Value2
Range("MedScrollPos").Value2 = 0

End Sub
Private Sub PostThreadScroll_SpinUp()

On Error Resume Next

lastR = Cells(Rows.Count, "Y").End(xlUp).Row
If lastR < 1 Then lastR = Cells(Rows.Count, "Z").End(xlUp).Row

Call clnMediaScroll

Range("ThreadScrollPos").Value2 = Range("ThreadScrollPos").Value2 + 1

'//RIGHT
If Range("ThreadScrollPos").Value2 >= lastR Then Range("ThreadScrollPos").Value2 = 1

ThreadHldr = Range("PostThread").Offset(Range("ThreadScrollPos").Value2)
MediaHldr = Range("MedThread").Offset(Range("ThreadScrollPos").Value2)

'//media check
If InStr(1, MediaHldr, """ """) Then

MedArr = Split(MediaHldr, """ """)
For X = 0 To UBound(MedArr)
Range("MediaScroll").Offset(X, 0).Value2 = MedArr(X)
    Next
        Else
            Range("MediaScroll").Offset(0, 0).Value2 = MediaHldr
                End If
        
If ThreadHldr <> vbNullString Then
PostBox.Value = ThreadHldr
If Left(MediaHldr, Len(MediaHldr) - Len(MediaHldr) + 1) = """" Then MediaHldr = Left(MediaHldr, Len(MediaHldr) - 1)
If Right(MediaHldr, Len(MediaHldr) - Len(MediaHldr) - 1) = """" Then MediaHldr = Right(MediaHldr, Len(MediaHldr) - 1)
MedLinkBox.Value = MediaHldr
    Else
        PostBox.Value = vbNullString
        MedLinkBox.Value = vbNullString
            End If
        
ThreadCt.Caption = Range("ThreadScrollPos").Value2
Range("MedScrollPos").Value2 = 0

End Sub
Private Sub LinkerBox_Click()

If Range("EditStatus").Value2 <= 0 Then

If Range("DataPullTrig").Value2 <> 1 Then xNum = LinkerBox.ListIndex: _
ETWEETXLPOST.ProfileListBox.Value = Range("ProfileLink").Offset(xNum + 1, 0).Value2

xTwt = LinkerBox.Value
'//remove numbered count
If xTwt <> vbNullString Then xTwtArr = Split(xTwt, ") "): xTwt = xTwtArr(1)

If Range("DraftTrig").Value2 <> 1 Then
DraftBox.Value = xTwt
xTwt = Replace(xTwt, " [•]", vbNullString)
xTwt = Replace(xTwt, " [...]", vbNullString)
DraftBox.Value = xTwt

If InStr(1, xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
xDraft = LinkerBox.Value
xDraft = Replace(xDraft, " [•]", vbNullString)
xDraft = Replace(xDraft, " [...]", vbNullString)
xlFlowStrip.Value = xDraft & " selected..."
End If
    End If
    If InStr(1, xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
    xDraft = LinkerBox.Value
    xDraft = Replace(xDraft, " [•]", vbNullString)
    xDraft = Replace(xDraft, " [...]", vbNullString)
    xlFlowStrip.Value = xDraft & " selected..."
    End If
        End If
        
End Sub
Private Sub RuntimeBox_Click()

If InStr(1, xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
xlFlowStrip.Value = RuntimeBox.Value & " selected..."
End If

End Sub
Private Sub UserBox_Click()

If InStr(1, xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
xlFlowStrip.Value = UserBox.Value & " selected..."
End If

End Sub
Private Sub PostBox_Change()

'//WinForm #
xWin = 131: Call setWindow(xWin)

Call eTweetXL_CHANGE.PostBox_Chg

'//WinForm #
xWin = 13: Call setWindow(xWin)

End Sub
Private Sub RmvRuntimeBtn_Click()

'//remove runtime from Linker
Call eTweetXL_CLICK.RmvRuntimeBtn_Clk

End Sub
Private Sub RuntimeBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'//remove runtime from Linker (2x click)
eTweetXL_CLICK.RuntimeBox_Clk

End Sub
Private Sub RuntimeBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'//Backspace key
If KeyCode = 8 Then
Call eTweetXL_CLICK.RuntimeBox_Clk
Exit Sub
End If

'//Space key
If KeyCode = 32 Then
xType = 3: xPos = Range("LinkerIndex").Value2
Call eTweetXL_TOOLS.hostSelection(xType, xPos)
Exit Sub
End If

'//Delete key
If KeyCode = 46 Then
'//remove runtime from Linker
Call eTweetXL_CLICK.RmvRuntimeBtn_DelClk
Exit Sub
End If

End Sub
Private Sub UserBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'//Backspace key
If KeyCode = 8 Then
Call eTweetXL_CLICK.RmvUserBtn_Clk
Exit Sub
End If

'//Space key
If KeyCode = 32 Then
xType = 1: xPos = Range("LinkerIndex").Value2
Call eTweetXL_TOOLS.hostSelection(xType, xPos)
Exit Sub
End If

'//Delete key
If KeyCode = 46 Then
'//remove user from Linker
Call eTweetXL_CLICK.RmvUserBox_DelClk
Exit Sub
End If

End Sub
Private Sub OffsetBox_Change()

'//refresh offset
eTweetXL_CHANGE.OffsetBox_Chg

End Sub
Private Sub DelDraftBtn_Click()

'//remove draft from local database
Call eTweetXL_CLICK.DelDraftBtn_Clk

End Sub
Private Sub DynOffset_Click()

Call eTweetXL_CLICK.DynOffset_Clk(xPos)

End Sub
Private Sub PostBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

Range("xlasInputField").Value2 = 99

'//Key Alt
If KeyCode.Value = 18 Then
Range("xlasKeyCtrl").Value2 = Range("xlasKeyCtrl").Value2 + 18
KeyCode.Value = 0
Exit Sub
End If

'//Key Ctrl
If KeyCode.Value = vbKeyControl Then
Range("xlasKeyCtrl").Value2 = 17
KeyCode.Value = 0
Exit Sub
End If

'//Key Enter
If KeyCode.Value = 13 Then
PostBox.EnterKeyBehavior = True
Exit Sub
End If

'//Key Shift
If KeyCode.Value = vbKeyShift Then
Range("xlasKeyCtrl").Value2 = Range("xlasKeyCtrl").Value2 + 16
Exit Sub
End If

'//Key Tab
If KeyCode.Value = 9 Then
PostBox.TabKeyBehavior = True
Exit Sub
End If

'//Key Ctrl+D
If KeyCode.Value = vbKeyD Then
If Range("xlasKeyCtrl").Value2 = 17 Then
PostBox.Value = ""
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
  
'//Key Ctrl+F
If KeyCode.Value = vbKeyF Then
If Range("xlasKeyCtrl").Value2 = 17 Then
Range("xlasWinForm").Value2 = 131
XLFONTBOX.Show
Range("xlasWinForm").Value2 = 13
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+H
If KeyCode.Value = vbKeyH Then
If Range("xlasKeyCtrl").Value2 = 17 Then
Range("xlasWinForm").Value2 = 131
XLREPLACE.Show
Range("xlasWinForm").Value2 = 13
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+S
If KeyCode.Value = vbKeyS Then
If Range("xlasKeyCtrl").Value2 = 17 Then
Call eTweetXL_CLICK.SavePostBtn_Clk
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
'//Key Ctrl+T
If KeyCode.Value = vbKeyT Then
If Range("xlasKeyCtrl").Value2 = 17 Then
Call eTweetXL_CLICK.AddThreadBtn_Clk
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
'//Key Ctrl+R
If KeyCode.Value = vbKeyR Then
If Range("xlasKeyCtrl").Value2 = 17 Then
Call eTweetXL_CLICK.RmvThreadBtn_Clk
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Shift+C
If KeyCode.Value = vbKeyC Then
If Range("xlasKeyCtrl").Value2 = 33 Then
Call eTweetXL_CLICK.TrimPostBtn_Clk
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Shift+D
If KeyCode.Value = vbKeyD Then
If Range("xlasKeyCtrl").Value2 = 33 Then
Call eTweetXL_CLICK.DelDraftBtn_Clk
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
'//Key Ctrl+Shift+R
If KeyCode.Value = vbKeyR Then
If Range("xlasKeyCtrl").Value2 = 33 Then
Call eTweetXL_CLICK.RmvAllThreadBtn_Clk
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Shift+X
If KeyCode.Value = vbKeyX Then
If Range("xlasKeyCtrl").Value2 = 33 Then
Call eTweetXL_CLICK.SplitPostBtn_Clk
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

Range("xlasKeyCtrl").Value2 = vbNullString

End Sub
Private Sub AddThreadBtn_Click()

If ThreadCt.Caption = vbNullString Or 0 Then ThreadCt.Caption = 1

If PostBox.Value <> "" Then
Call eTweetXL_CLICK.AddThreadBtn_Clk
End If

End Sub
Private Sub RmvThreadBtn_Click()

lastR = Cells(Rows.Count, "Y").End(xlUp).Row
If lastR < 1 Then lastR = Cells(Rows.Count, "Z").End(xlUp).Row: If lastR <= 1 Then Exit Sub

Call eTweetXL_CLICK.RmvThreadBtn_Clk

End Sub
Private Sub RmvAllThreadBtn_Click()

Call eTweetXL_CLICK.RmvAllThreadBtn_Clk

End Sub
Private Sub AddMedBtn_Click()

xMed = vbNullString
Call eTweetXL_CLICK.AddPostMedBtn_Clk(xMed)

End Sub
Private Sub RmvMedBtn_Click()

Call eTweetXL_CLICK.RmvPostMedBtn_Clk

End Sub
Private Sub DraftBox_Change()

Call eTweetXL_CHANGE.DraftBox_Chg

End Sub
Private Sub UserListBox_Change()

Range("User").Value = Replace(ETWEETXLPOST.UserListBox.Value, Range("Scure").Value, "")
xUser = Range("User").Value

If xUser <> "" Then
Call eTweetXL_CLICK.SetActive_Clk(xUser)
End If

If Range("xlasSilent").Value2 <> 1 Then
ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.UserListBox.Value & " selected..."
End If

End Sub
'//Add user w/ enter button
Private Sub UserListBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 13 Then
Call AddUserBtn_Click
KeyCode = 0
Exit Sub
End If

'//Key Tab
If KeyCode.Value = vbKeyTab Then
ETWEETXLPOST.DraftBox.SetFocus
KeyCode.Value = 0
Exit Sub
End If
        
End Sub

Private Sub DraftBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'//Key Ctrl
If KeyCode.Value = vbKeyControl Then
Range("xlasKeyCtrl").Value2 = vbKeyControl
KeyCode.Value = 0
Exit Sub
End If

'//Alt key
If KeyCode.Value = 18 Then
Range("xlasKeyCtrl").Value2 = Range("xlasKeyCtrl").Value2 + 18
KeyCode.Value = 0
Exit Sub
End If

'//Key Enter
If KeyCode = 13 Then
Call AddLinkBtn_Click
KeyCode = 0
Exit Sub
End If

'//Key Tab
If KeyCode.Value = vbKeyTab Then
ETWEETXLPOST.PostBox.SetFocus
KeyCode.Value = 0
Exit Sub
End If

'//Key Ctrl+S
If KeyCode.Value = vbKeyS Then
If Range("xlasKeyCtrl").Value2 = 17 Then
Call eTweetXL_CLICK.SavePostBtn_Clk
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
End If
    Exit Sub
        End If

'//Key Ctrl+Alt+D
If KeyCode.Value = vbKeyD Then
If Range("xlasKeyCtrl").Value2 = 35 Then
Call eTweetXL_CLICK.DelDraftBtn_Clk
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

End Sub
Private Sub ProfileListBox_Click()

If InStr(1, xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
xlFlowStrip.Value = ProfileListBox.Value & "..."
End If

'//check if already loaded...
If Range("Profile").Value2 <> ETWEETXLPOST.ProfileListBox.Value Then
Range("Profile").Value2 = ETWEETXLPOST.ProfileListBox.Value
    Else
        End If

Range("DataPullTrig").Value2 = 0

xType = 0: Call eTweetXL_GET.getPostData(xType)
Call eTweetXL_GET.getProfileData

End Sub
Private Sub ProfileListBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'//Key Tab
If KeyCode.Value = vbKeyTab Then
ETWEETXLPOST.UserListBox.SetFocus
KeyCode.Value = 0
Exit Sub
End If
        
End Sub
Private Sub TimeBox_Change()

Call eTweetXL_CHANGE.TimeBox_Chg

End Sub
Private Sub TimeBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'//Add time w/ enter key
If KeyCode.Value = 13 Then
Call AddRuntimeBtn_Click
KeyCode.Value = 0
End If

End Sub
Private Sub TimeHdr_Click()

Call eTweetXL_CLICK.TimerHdr_Clk

End Sub
'///////////////////////////////////////HOUR ADJUSTMENT BUTTON//////////////////////////////////////////////
Private Sub UpHrBtn_Click()

Dim xCount As Integer

xCount = 1

Call eTweetXL_CLICK.UpHrBtn_Clk(xCount)

End Sub
Private Sub UpHrDwnBtn_SpinUp()

Dim xCount As Integer

xCount = 1

Call eTweetXL_CLICK.UpHrBtn_Clk(xCount)

End Sub
Private Sub DwnHrBtn_Click()

Dim xCount As Integer

xCount = 1

Call eTweetXL_CLICK.DwnHrBtn_Clk(xCount)

End Sub
Private Sub UpHrDwnBtn_SpinDown()

Dim xCount As Integer

xCount = 1

Call eTweetXL_CLICK.DwnHrBtn_Clk(xCount)

End Sub
Private Sub UpHrDwnBtn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'//UP
If KeyCode.Value = 38 Then

For xCntr = 1 To 1
Next

Exit Sub

End If

'//DOWN

If KeyCode.Value = 40 Then

For xCntr = 1 To 1
Next

Exit Sub

End If

'//ENTER KEY
If KeyCode.Value = 13 Then

For xCntr = 1 To 1
Call AddRuntimeBtn_Click
Next

Exit Sub

End If

End Sub
'///////////////////////////////////////MINUTE ADJUSTMENT BUTTON//////////////////////////////////////////////
Private Sub UpMinBtn_Click()

Dim xCount As Integer

xCount = 1

Call eTweetXL_CLICK.UpMinBtn_Clk(xCount)

End Sub
Private Sub UpMinDwnBtn_SpinUp()

Dim xCount As Integer

xCount = 1

eTweetXL_CLICK.UpMinBtn_Clk (xCount)

End Sub
Private Sub DwnMinBtn_Click()

Dim xCount As Integer

xCount = 1

Call eTweetXL_CLICK.DwnMinBtn_Clk(xCount)

End Sub
Private Sub UpMinDwnBtn_SpinDown()

Dim xCount As Integer

xCount = 1
Call eTweetXL_CLICK.DwnMinBtn_Clk(xCount)

End Sub
Private Sub UpMinDwnBtn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode.Value = 38 Then

For xCntr = 1 To 1
Next

Exit Sub

End If

'//DOWN

If KeyCode.Value = 40 Then

For xCntr = 1 To 1
Next

Exit Sub

End If

'//ENTER KEY
If KeyCode.Value = 13 Then

For xCntr = 1 To 1
Call AddRuntimeBtn_Click
Next

Exit Sub

End If

End Sub
'///////////////////////////////////////SECOND ADJUSTMENT BUTTON//////////////////////////////////////////////
Private Sub UpSecBtn_Click()

Dim xCount As Integer

xCount = 1

Call eTweetXL_CLICK.UpSecBtn_Clk(xCount)

End Sub
Private Sub UpSecDwnBtn_SpinUp()

Dim xCount As Integer

xCount = 1

Call eTweetXL_CLICK.UpSecBtn_Clk(xCount)

End Sub
Private Sub DwnSecBtn_Click()

Dim xCount As Integer

xCount = 1

Call eTweetXL_CLICK.DwnSecBtn_Clk(xCount)

End Sub
Private Sub UpSecDwnBtn_SpinDown()

Dim xCount As Integer

xCount = 1

Call eTweetXL_CLICK.DwnSecBtn_Clk(xCount)

End Sub
Private Sub UpSecDwnBtn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode.Value = 38 Then

For xCntr = 1 To 1
Next

Exit Sub

End If

'//DOWN

If KeyCode.Value = 40 Then

For xCntr = 1 To 1
Next

Exit Sub

End If

'//ENTER KEY
If KeyCode.Value = 13 Then

For xCntr = 1 To 1
Call AddRuntimeBtn_Click
Next

Exit Sub

End If

End Sub
Private Sub AddLinkBtn_Click()

'//Reset connect trigger
If Range("AppState").Value2 <> 1 Then
Range("ConnectTrig").Value2 = 0
End If
        
xPos = 0

Call eTweetXL_CLICK.AddLinkBtn_Clk(xPos)

End Sub
Private Sub RmvLinkBtn_Click()

'//Reset connect trigger
If Range("AppState").Value2 <> 1 Then
Range("ConnectTrig").Value2 = 0
End If

Call eTweetXL_CLICK.RmvLinkBtn_Clk

End Sub
Private Sub AddUserBtn_Click()

'//Reset connect trigger
If Range("AppState").Value2 <> 1 Then
Range("ConnectTrig").Value2 = 0
End If

xPos = 0
Call eTweetXL_CLICK.AddUserBtn_Clk(xPos)

End Sub
Private Sub RmvUserBtn_Click()

'//Reset connect trigger
If Range("AppState").Value2 <> 1 Then
Range("ConnectTrig").Value2 = 0
End If

Call eTweetXL_CLICK.RmvUserBtn_Clk

End Sub
Private Sub LinkerBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call eTweetXL_CLICK.RmvLinkerBox_DelClk

End Sub
Private Sub LinkerBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'//Backspace key
If KeyCode = 8 Then
Call eTweetXL_CLICK.RmvLinkBtn_Clk
Exit Sub
End If

'//Space key
If KeyCode = 32 Then
xType = 2: xPos = Range("LinkerIndex").Value2
Call eTweetXL_TOOLS.hostSelection(xType, xPos)
End If

'//Delete key
If KeyCode = 46 Then
'//remove draft from Linker
Call eTweetXL_CLICK.RmvLinkerBox_DelClk
End If

End Sub
Private Sub xlFlowStrip_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'//WinForm #
xWin = 13: Call setWindow(xWin)

xKey = KeyCode.Value
Call eTweetXL_TOOLS.runFlowStrip(xKey)
KeyCode.Value = xKey

End Sub
Private Sub xlFlowStripBar_Click()

Call eTweetXL_CLICK.xlFlowStripBar_Clk

End Sub
Private Sub SavePostBtn_Click()

Call eTweetXL_CLICK.SavePostBtn_Clk

End Sub
Private Sub ConnectBtn_Click()

Call xlAppScript_xbas.disableWbUpdates
Call eTweetXL_CLICK.ConnectBtn_Clk

End Sub
Private Sub HomeBtn_Click()

Me.Hide
Call eTweetXL_FOCUS.shw_ETWEETXLHOME

End Sub
Private Sub AppTag_Click()

ETWEETXLPOST.Hide
ETWEETXLHOME.Show

End Sub
Private Sub UsersHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Art = "<lib>xtwt;winform(13);add.user(*);$" '//xlas
Call xlas(Art)

End Sub
Private Sub ViewMedBtn_Click()

eTweetXL_CLICK.ViewMedBtn_Clk

End Sub
Private Sub DetachPostBtn_Click()

ETWEETXLPOST_EX.Show

End Sub
Private Sub DraftFilterBtn_Click()

If DraftFilterBtn.Caption = "..." Then xFil = 0
If DraftFilterBtn.Caption = "•" Then xFil = 1
        
Call eTweetXL_CLICK.DraftFilterBtn_Clk(xFil)

End Sub
Private Sub LoadPostBtn_Click()

Call eTweetXL_CLICK.LoadPostBtn_Clk(xName, xPath)

End Sub
Private Sub SplitPostBtn_Click()

Call eTweetXL_CLICK.SplitPostBtn_Clk

End Sub
Private Sub TrimPostBtn_Click()

Call eTweetXL_CLICK.TrimPostBtn_Clk

End Sub
Private Sub SwapDraft_Click()

xType = 1: Call eTweetXL_CLICK.DraftOpt_Clk(xType)

End Sub
Private Sub SwapUser_Click()

xType = 1: Call eTweetXL_CLICK.UserOpt_Clk(xType)

End Sub
Private Sub SwapTime_Click()
         
xType = 1: Call eTweetXL_CLICK.RuntimeOpt_Clk(xType)

End Sub
Private Sub AddUserA_Click()

xType = 2: Call eTweetXL_CLICK.UserOpt_Clk(xType)

End Sub
Private Sub AddUserB_Click()

xType = 3: Call eTweetXL_CLICK.UserOpt_Clk(xType)

End Sub

Private Sub AddDraftA_Click()

xType = 2: Call eTweetXL_CLICK.DraftOpt_Clk(xType)

End Sub
Private Sub AddDraftB_Click()

xType = 3: Call eTweetXL_CLICK.DraftOpt_Clk(xType)

End Sub
Private Sub AddTimeA_Click()

xType = 2: Call eTweetXL_CLICK.RuntimeOpt_Clk(xType)

End Sub
Private Sub AddTimeB_Click()

xType = 3: Call eTweetXL_CLICK.RuntimeOpt_Clk(xType)

End Sub
'//hover effects
Private Sub HomeBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xBtn = 0
Call eTweetXL_TOOLS.fxsNaviBar(xBtn)
Call eTweetXL_TOOLS.undNaviBar(xBtn)

End Sub
Private Sub StartBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xBtn = 1
Call eTweetXL_TOOLS.fxsNaviBar(xBtn)
Call eTweetXL_TOOLS.undNaviBar(xBtn)

xMsg = 52: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub ProfileSetupBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xBtn = 2
Call eTweetXL_TOOLS.fxsNaviBar(xBtn)
Call eTweetXL_TOOLS.undNaviBar(xBtn)

xMsg = 54: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub PostSetupBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xBtn = 3
Call eTweetXL_TOOLS.fxsNaviBar(xBtn)
Call eTweetXL_TOOLS.undNaviBar(xBtn)

xMsg = 55: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub QueueBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xBtn = 4
Call eTweetXL_TOOLS.fxsNaviBar(xBtn)
Call eTweetXL_TOOLS.undNaviBar(xBtn)

xMsg = 53: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub BreakBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xBtn = 5
Call eTweetXL_TOOLS.fxsNaviBar(xBtn)
Call eTweetXL_TOOLS.undNaviBar(xBtn)


xMsg = 51: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub HelpIcon_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover
xHov = 14: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub HelpStatus_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover
xHov = 14: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub DraftHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 1: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover
xHov = 1: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub DraftsHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 2: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover
xHov = 2: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub OffsetHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 3: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover
xHov = 3: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub PostHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 4: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover
xHov = 4: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub RuntimeHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 5: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover
xHov = 5: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub TimeHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 6: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover
xHov = 6: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub UserHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 7: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover
xHov = 7: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub UsersHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 8: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover
xHov = 8: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub LinkerHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 9: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover
xHov = 9: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub FreezeBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 10: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover
xHov = 15: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub AppTag_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 11: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover
xHov = 10: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub CtrlBoxBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 12: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover
xHov = 16: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub SendAPI_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 13: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub DynOffset_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 14: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub DraftFilterBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 15: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub DelAllDraftsBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 16: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub RmvDraft_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 17: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub AddDraft_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 18: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub xlFlowStripBar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 19: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub ActiveUser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

'On Error Resume Next

'Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
'xMsg = 20: Call HoverHelp(xMsg)

'Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub AddMedBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 29: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub RmvMedBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 30: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub

Private Sub ViewMedBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 31: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub SavePostBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 32: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub AddThreadBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 33: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub RmvThreadBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 34: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub RmvAllThreadBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 35: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub ConnectBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 36: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub AddUserBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 37: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub RmvUserBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 38: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub AddLinkBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 39: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub RmvLinkBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 40: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub AddRuntimeBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 41: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub RmvRuntimeBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 42: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub UserBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 43: Call HoverHelp(xMsg)
Call eTweetXL_TOOLS.dfsHover
xHov = 11: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub LinkerBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 44: Call HoverHelp(xMsg)
Call eTweetXL_TOOLS.dfsHover
xHov = 12: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub RuntimeBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 45: Call HoverHelp(xMsg)
Call eTweetXL_TOOLS.dfsHover
xHov = 13: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub SaveLinkerBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 46: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub LoadLinkerBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 47: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub ReloadLinkerBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 48: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub ClrSetupBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 49: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub LastLinkBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 50: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub PostBg1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub PostBg2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub

Private Sub PostBg3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub

Private Sub PostBg4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub

Private Sub PostBg5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub

Private Sub PostBg6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub HideBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover
xHov = 17: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub UserForm_Terminate()

'//backup current queue state
Call eTweetXL_POST.pstLastQueue

Range("xlasWinForm").Value2 = Range("xlasWinFormLast").Value2

End Sub
