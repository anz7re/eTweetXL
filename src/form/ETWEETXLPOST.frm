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

'//CHECK FOR DYNAMIC OFFSET
If Range("DynOffTrig").Value = 1 Then
ETWEETXLPOST.DynOffset.Value = True
    Else
         ETWEETXLPOST.DynOffset.Value = False
            End If

'//CHECK FOR ACTIVE APP
If Range("AppActive").Value = 0 Then
Range("ConnectTrig").Value = 0
End If

'//WELCOME
If Me.xlFlowStrip.Value = "" Then Me.xlFlowStrip.Value = App_INFO.AppWelcome

'//RESET SEND API
Range("SendAPI").Value = 0

'//CLEAR CONNECT SPACE
If Range("ConnectTrig").Value = 0 Then
        Call Cleanup.ClnLinkerSpace
        Call Cleanup.ClnSpecSpace
            ETWEETXLPOST.LinkerBox.Clear
            ETWEETXLPOST.RuntimeBox.Clear
            ETWEETXLPOST.UserBox.Clear
            End If
            

'//REFRESH Media SCROLL...
Range("MedScrollPos").Value = 0

Range("DataPullTrig").Value = 0
Me.Caption = AppTag
ETWEETXLPOST.OffsetBox.Value = "00:00:00"

End Sub
Private Sub UserForm_Activate()

'//WinForm #
Range("xlasWinForm").Value = 3

'//Cleanup
ETWEETXLPOST.ProfileListBox.Clear
ETWEETXLPOST.DraftBox.Clear

'//Show runtime action message
Call App_TOOLS.ShowRtAction

'//Update active state
Call App_TOOLS.UpdateActive

'//Import profile names
Call App_IMPORT.MyProfileNames

xFil = 0: Call App_CLICK.DraftFilterBtn_Clk(xFil)

'//Window name
If Me.xlFlowStrip.Value = vbNullString Or Range("AppActive").Value <> 1 Then Me.xlFlowStrip.Value = "Tweet Setup..."

End Sub
Private Sub LinkerBox_Change()

On Error Resume Next

RuntimeBox.Selected(LinkerBox.ListIndex) = True
UserBox.Selected(LinkerBox.ListIndex) = True

End Sub

Private Sub RuntimeBox_Change()

On Error Resume Next

LinkerBox.Selected(RuntimeBox.ListIndex) = True
UserBox.Selected(RuntimeBox.ListIndex) = True

End Sub

Private Sub UserBox_Change()

On Error Resume Next

LinkerBox.Selected(UserBox.ListIndex) = True
RuntimeBox.Selected(LinkerBox.ListIndex) = True

End Sub
Private Sub HomeBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 0
Call App_TOOLS.NaviBarColor(xBtn)
Call App_TOOLS.NaviBarUnderline(xBtn)

End Sub
Private Sub StartBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 1
Call App_TOOLS.NaviBarColor(xBtn)
Call App_TOOLS.NaviBarUnderline(xBtn)

xMsg = 52: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub ProfileSetupBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 2
Call App_TOOLS.NaviBarColor(xBtn)
Call App_TOOLS.NaviBarUnderline(xBtn)

xMsg = 54: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub

Private Sub HelpIcon_Click()

If Range("HelpActive").Value = 0 And HelpIcon.Caption = "Off" Then Exit Sub
If Range("HelpActive").Value = 1 And HelpIcon.Caption = "On" Then Exit Sub

If Range("HelpActive").Value = 0 Then xPos = 1 Else xPos = 0
If xPos = vbNullString Then Range("HelpActive").Value = 0: xPos = 1

Call App_CLICK.HelpStatus_Clk(xPos)

End Sub

Private Sub HelpStatus_Click()

If Range("HelpActive").Value = 0 And HelpIcon.Caption = "Off" Then Exit Sub
If Range("HelpActive").Value = 1 And HelpIcon.Caption = "On" Then Exit Sub

If Range("HelpActive").Value = 0 Then xPos = 1 Else xPos = 0
If xPos = vbNullString Then Range("HelpActive").Value = 0: xPos = 1

Call App_CLICK.HelpStatus_Clk(xPos)

End Sub
Private Sub FreezeBtn_Click()

Call App_CLICK.FreezeBtn_Clk

End Sub
Private Sub PostSetupBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 3
Call App_TOOLS.NaviBarColor(xBtn)
Call App_TOOLS.NaviBarUnderline(xBtn)

xMsg = 55: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub QueueBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 4
Call App_TOOLS.NaviBarColor(xBtn)
Call App_TOOLS.NaviBarUnderline(xBtn)

xMsg = 53: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub BreakBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 5
Call App_TOOLS.NaviBarColor(xBtn)
Call App_TOOLS.NaviBarUnderline(xBtn)


xMsg = 51: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub

Private Sub StartBtn_Click()

Call App_CLICK.Start_Clk

End Sub
Private Sub BreakBtn_Click()

App_CLICK.Break_Clk

End Sub

Private Sub CtrlBoxBtn_Click()

Call App_Focus.SH_CTRLBOX

End Sub
Private Sub QueueBtn_Click()

Me.Hide
Call App_Focus.SH_ETWEETXLQUEUE

End Sub
Private Sub PostSetupBtn_Click()

Me.Hide
Call App_Focus.SH_ETWEETXLPOST

End Sub
Private Sub ProfileSetupBtn_Click()

Me.Hide
Call App_Focus.SH_ETWEETXLSETUP

End Sub
Private Sub DraftsHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

xArt = "<lib>xtwt;winform(3);add.draft(*);$" '//xlas
Call lexKey(xArt)
        
End Sub
Private Sub DraftHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        
Call App_CLICK.DraftHdr_Clk
        
End Sub
Private Sub RmvAllDrafts_Click()

Call App_CLICK.DeleteAllDraft_Clk

End Sub
Private Sub ClrSetupBtn_Click()

Call App_CLICK.ClrSetup_Clk

End Sub
Private Sub LoadLinkerBtn_Click()

Range("LoadLess").Value = 1
xLink = ""
Call App_IMPORT.MyLink(xLink)
Range("LoadLess").Value = 0

End Sub
Private Sub LastLinkBtn_Click()

Range("LoadLess").Value = 1
xLink = AppLoc & "\mtsett\lastlink.tmp"
Call App_IMPORT.MyLink(xLink)
Range("LoadLess").Value = 0

End Sub
Private Sub LinkerHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call App_CLICK.LinkerHdr_Clk

End Sub
Private Sub ReloadLinkerBtn_Click()

Range("LoadLess").Value = 1
xLink = Range("RemLink").Value
Call App_IMPORT.MyLink(xLink)
Range("LoadLess").Value = 0

End Sub
Private Sub RuntimeHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        
Call App_CLICK.RuntimeHdr_Clk
            
End Sub
Private Sub OffsetHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

ETWEETXLPOST.OffsetBox.Value = "00:00:00"

If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then _
ETWEETXLPOST.xlFlowStrip.Value = "Offset refreshed..."
        
End Sub
Private Sub PostHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'//Clear post box
Call App_CLICK.PostHdr_Clk

End Sub
Private Sub SaveLinkerBtn_Click()

If Range("ConnectTrig").Value = 1 Then
Call App_CLICK.SaveLinkerBtn_Clk
    Else
        xMsg = 6: Call App_MSG.AppMsg(xMsg)
            End If

End Sub
Private Sub SendAPI_Click()

If Range("SendAPI").Value = 0 And SendAPI.Value = False Then Exit Sub
If Range("SendAPI").Value = 1 And SendAPI.Value = True Then Exit Sub

If Range("SendAPI").Value = 0 Then xPos = 1 Else xPos = 0
If xPos = vbNullString Then Range("SendAPI").Value = 0: xPos = 1

Call App_CLICK.SendAPI_Clk(xPos)

End Sub
Private Sub UserBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call App_CLICK.RmvUserBox_EnClk

End Sub
Private Sub UserHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        
Call App_CLICK.UserHdr_Clk
        
End Sub
Private Sub AddDraft_Click()

App_CLICK.AddDraft_Clk

End Sub
'//ADD RUNTIME BUTTON
Private Sub AddRuntime_Click()

xPos = 0
Call App_CLICK.AddRuntime_Clk(xPos)

End Sub
Private Sub AddRuntime_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode.Value = 38 Then
UpHrDwnBtn.SetFocus
End If

End Sub
Private Sub MedDemoScroll_SpinDown()

On Error Resume Next

lastRw = Cells(Rows.Count, "I").End(xlUp).Row

Range("MedScrollPos").Value = Range("MedScrollPos").Value - 1

'//LEFT
If Range("MedScrollPos").Value < 0 Then Range("MedScrollPos").Value = lastRw - 1

MedLinkHldr = Range("MediaScroll").Offset(Range("MedScrollPos").Value)
MedLinkHldr = Replace(MedLinkHldr, """", "")

If MedLinkHldr <> "" Then

If Dir(MedLinkHldr) <> "" Then
    MedDemo.Picture = LoadPicture(MedLinkHldr)
    MedDemo.PictureSizeMode = fmPictureSizeModeStretch
        
        End If
            Else
                MedDemo.Picture = Nothing
                    End If
        
        MedCt.Caption = Range("MedScrollPos").Value + 1
        Range("MedScrollLink").Value = MedLinkHldr

End Sub
Private Sub MedDemoScroll_SpinUp()

On Error Resume Next

lastRw = Cells(Rows.Count, "I").End(xlUp).Row

Range("MedScrollPos").Value = Range("MedScrollPos").Value + 1

'//RIGHT
If Range("MedScrollPos").Value > 3 Then Range("MedScrollPos").Value = 0

MedLinkHldr = Range("MediaScroll").Offset(Range("MedScrollPos").Value)
MedLinkHldr = Replace(MedLinkHldr, """", "")

If MedLinkHldr <> "" Then

If Dir(MedLinkHldr) <> "" Then
    MedDemo.Picture = LoadPicture(MedLinkHldr)
    MedDemo.PictureSizeMode = fmPictureSizeModeStretch
        
        End If
            Else
                MedDemo.Picture = Nothing
                    End If

        MedCt.Caption = Range("MedScrollPos").Value + 1
        Range("MedScrollLink").Value = MedLinkHldr

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
Range("MediaScroll").Offset(X, 0).Value = MedArr(X)
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

Range("MedScrollLink").Value = MediaHldr
  
End Sub
Private Sub PostThreadScroll_SpinDown()

On Error Resume Next

lastRw = Cells(Rows.Count, "Y").End(xlUp).Row
If lastRw < 1 Then lastRw = Cells(Rows.Count, "Z").End(xlUp).Row

Call Cleanup.ClnMediaScroll

Range("ThreadScrollPos").Value = Range("ThreadScrollPos").Value - 1

'//LEFT
If Range("ThreadScrollPos").Value <= 0 Then Range("ThreadScrollPos").Value = lastRw - 1

ThreadHldr = Range("PostThread").Offset(Range("ThreadScrollPos").Value)
MediaHldr = Range("MedThread").Offset(Range("ThreadScrollPos").Value)

'//media check
If InStr(1, MediaHldr, """ """) Then

MedArr = Split(MediaHldr, """ """)
For X = 0 To UBound(MedArr)
Range("MediaScroll").Offset(X, 0).Value = MedArr(X)
    Next
        Else
            Range("MediaScroll").Offset(0, 0).Value = MediaHldr
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
        
ThreadCt.Caption = Range("ThreadScrollPos").Value
Range("MedScrollPos").Value = 0

End Sub
Private Sub PostThreadScroll_SpinUp()

On Error Resume Next

lastRw = Cells(Rows.Count, "Y").End(xlUp).Row
If lastRw < 1 Then lastRw = Cells(Rows.Count, "Z").End(xlUp).Row

Call Cleanup.ClnMediaScroll

Range("ThreadScrollPos").Value = Range("ThreadScrollPos").Value + 1

'//RIGHT
If Range("ThreadScrollPos").Value >= lastRw Then Range("ThreadScrollPos").Value = 1

ThreadHldr = Range("PostThread").Offset(Range("ThreadScrollPos").Value)
MediaHldr = Range("MedThread").Offset(Range("ThreadScrollPos").Value)

'//media check
If InStr(1, MediaHldr, """ """) Then

MedArr = Split(MediaHldr, """ """)
For X = 0 To UBound(MedArr)
Range("MediaScroll").Offset(X, 0).Value = MedArr(X)
    Next
        Else
            Range("MediaScroll").Offset(0, 0).Value = MediaHldr
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
        
ThreadCt.Caption = Range("ThreadScrollPos").Value
Range("MedScrollPos").Value = 0

End Sub
Private Sub LinkerBox_Click()

If Range("DataPullTrig").Value <> 1 Then iNum = LinkerBox.ListIndex: _
ETWEETXLPOST.ProfileListBox.Value = Range("Profilelink").Offset(iNum + 1, 0).Value

xTwt = LinkerBox.Value
'//remove numbered count
If xTwt <> vbNullString Then xTwtArr = Split(xTwt, ") "): xTwt = xTwtArr(1)

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

Call App_CHANGE.PostBox_Chg

End Sub
'//RMV RUNTIME BUTTON
Private Sub RmvRuntime_Click()

Call App_CLICK.RmvRuntime_Clk

End Sub

'//RMV RUNTIME DBL CLICK
Private Sub RuntimeBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

App_CLICK.RuntimeBox_Clk

End Sub
'//RMV RUNTIME ENTER KEY
Private Sub RuntimeBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 13 Then

Call App_CLICK.RmvRuntime_EnClk

End If

End Sub

'//RMV USER ENTER KEY
Private Sub UserBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 13 Then

Call App_CLICK.RmvUserBox_EnClk

End If

End Sub

'//REFRESH OFFSET
Private Sub OffsetBox_Change()

App_CHANGE.OffsetBox_Chg

End Sub
'//RMV DRAFT FROM DATABASE
Private Sub RmvDraft_Click()

Call App_CLICK.DeleteDraft_Clk

End Sub
Private Sub DynOffset_Click()

xPos = vbNullString
Call App_CLICK.DynOffset_Clk(xPos)

End Sub
Private Sub PostBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'//Key Enter
If KeyCode.Value = 13 Then
PostBox.EnterKeyBehavior = True
Exit Sub
End If

'//Key Tab
If KeyCode.Value = 9 Then
PostBox.TabKeyBehavior = True
Exit Sub
End If

'//Key Ctrl
If KeyCode.Value = vbKeyControl Then
Range("xlasKeyCtrl").Value = vbKeyControl
KeyCode.Value = 0
Exit Sub
End If

'//Alt key
If KeyCode.Value = 18 Then
Range("xlasKeyCtrl").Value = Range("xlasKeyCtrl").Value + 18
KeyCode.Value = 0
Exit Sub
End If

'//Key Ctrl+D
If KeyCode.Value = vbKeyD Then
If Range("xlasKeyCtrl").Value = 17 Then
PostBox.Value = ""
DraftBox.Value = ""
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
       
'//Key Ctrl+H
If KeyCode.Value = vbKeyH Then
If Range("xlasKeyCtrl").Value = 17 Then
Range("xlasWinForm").Value = 31
XLREPLACE.Show
Range("xlasWinForm").Value = 3
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+S
If KeyCode.Value = vbKeyS Then
If Range("xlasKeyCtrl").Value = 17 Then
Call App_CLICK.SavePost_Clk
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
'//Key Ctrl+T
If KeyCode.Value = vbKeyT Then
If Range("xlasKeyCtrl").Value = 17 Then
Call App_CLICK.AddThread_Clk
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
'//Key Ctrl+R
If KeyCode.Value = vbKeyR Then
If Range("xlasKeyCtrl").Value = 17 Then
Call App_CLICK.RmvThread_Clk
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Alt+D
If KeyCode.Value = vbKeyD Then
If Range("xlasKeyCtrl").Value = 35 Then
Call App_CLICK.DeleteDraft_Clk
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
'//Key Ctrl+Alt+R
If KeyCode.Value = vbKeyR Then
If Range("xlasKeyCtrl").Value = 35 Then
Call App_CLICK.RmvAllThread_Clk
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

        
        Range("xlasKeyCtrl").Value = vbNullString

End Sub
Private Sub AddThreadBtn_Click()

If ThreadCt.Caption = vbNullString Or 0 Then ThreadCt.Caption = 1

If PostBox.Value <> "" Then
Call App_CLICK.AddThread_Clk
End If

End Sub

Private Sub RmvThreadBtn_Click()

lastRw = Cells(Rows.Count, "Y").End(xlUp).Row
If lastRw < 1 Then lastRw = Cells(Rows.Count, "Z").End(xlUp).Row: If lastRw <= 1 Then Exit Sub

Call App_CLICK.RmvThread_Clk

End Sub
Private Sub RmvAllThreadBtn_Click()

Call App_CLICK.RmvAllThread_Clk

End Sub
Private Sub AddMedBtn_Click()

xMed = vbNullString
Call App_CLICK.AddPostMed_Clk(xMed)

End Sub
Private Sub RmvMedBtn_Click()

Call App_CLICK.RmvPostMed_Clk

End Sub
Private Sub DraftBox_Change()

Call App_CHANGE.DraftBox_Chg

End Sub

Private Sub UserListBox_Change()

Range("User").Value = Replace(ETWEETXLPOST.UserListBox.Value, Range("Scure").Value, "")
xUser = Range("User").Value

If xUser <> "" Then
Call App_CLICK.SetActive_Clk(xUser)
End If

If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
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
Range("xlasKeyCtrl").Value = vbKeyControl
KeyCode.Value = 0
Exit Sub
End If

'//Alt key
If KeyCode.Value = 18 Then
Range("xlasKeyCtrl").Value = Range("xlasKeyCtrl").Value + 18
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
If Range("xlasKeyCtrl").Value = 17 Then
Call App_CLICK.SavePost_Clk
Range("xlasKeyCtrl").Value = ""
KeyCode.Value = 0
End If
    Exit Sub
        End If

'//Key Ctrl+Alt+D
If KeyCode.Value = vbKeyD Then
If Range("xlasKeyCtrl").Value = 35 Then
Call App_CLICK.DeleteDraft_Clk
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

End Sub
Private Sub ProfileListBox_Click()

If InStr(1, xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
xlFlowStrip.Value = ProfileListBox.Value & "..."
End If

Range("Profile").Value = ETWEETXLPOST.ProfileListBox.Value
Range("DataPullTrig").Value = 0

xType = 0: Call App_IMPORT.MyTweetData(xType)
Call App_IMPORT.MyProfileData

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

Call App_CHANGE.TimeBox_Chg

End Sub
Private Sub TimeBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'//Add time w/ enter key
If KeyCode.Value = 13 Then
Call AddRuntime_Click
KeyCode.Value = 0
End If

End Sub
Private Sub TimeHdr_Click()

Call App_CLICK.TimerHdr_Clk

End Sub
'///////////////////////////////////////HOUR ADJUSTMENT BUTTON//////////////////////////////////////////////
Private Sub UpHrBtn_Click()

Dim xTimes As Integer

xTimes = 1

Call App_CLICK.UpHrBtn_Clk(xTimes)

End Sub
Private Sub UpHrDwnBtn_SpinUp()

Dim xTimes As Integer

xTimes = 1

Call App_CLICK.UpHrBtn_Clk(xTimes)

End Sub
Private Sub DwnHrBtn_Click()

Dim xTimes As Integer

xTimes = 1

Call App_CLICK.DwnHrBtn_Clk(xTimes)

End Sub
Private Sub UpHrDwnBtn_SpinDown()

Dim xTimes As Integer

xTimes = 1

Call App_CLICK.DwnHrBtn_Clk(xTimes)

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
Call AddRuntime_Click
Next

Exit Sub

End If

End Sub
'///////////////////////////////////////MINUTE ADJUSTMENT BUTTON//////////////////////////////////////////////
Private Sub UpMinBtn_Click()

Dim xTimes As Integer

xTimes = 1

Call App_CLICK.UpMinBtn_Clk(xTimes)

End Sub
Private Sub UpMinDwnBtn_SpinUp()

Dim xTimes As Integer

xTimes = 1

App_CLICK.UpMinBtn_Clk (xTimes)

End Sub
Private Sub DwnMinBtn_Click()

Dim xTimes As Integer

xTimes = 1

Call App_CLICK.DwnMinBtn_Clk(xTimes)

End Sub
Private Sub UpMinDwnBtn_SpinDown()

Dim xTimes As Integer

xTimes = 1
Call App_CLICK.DwnMinBtn_Clk(xTimes)

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
Call AddRuntime_Click
Next

Exit Sub

End If

End Sub
'///////////////////////////////////////SECOND ADJUSTMENT BUTTON//////////////////////////////////////////////
Private Sub UpSecBtn_Click()

Dim xTimes As Integer

xTimes = 1

Call App_CLICK.UpSecBtn_Clk(xTimes)

End Sub
Private Sub UpSecDwnBtn_SpinUp()

Dim xTimes As Integer

xTimes = 1

Call App_CLICK.UpSecBtn_Clk(xTimes)

End Sub
Private Sub DwnSecBtn_Click()

Dim xTimes As Integer

xTimes = 1

Call App_CLICK.DwnSecBtn_Clk(xTimes)

End Sub
Private Sub UpSecDwnBtn_SpinDown()

Dim xTimes As Integer

xTimes = 1

Call App_CLICK.DwnSecBtn_Clk(xTimes)

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
Call AddRuntime_Click
Next

Exit Sub

End If

End Sub
Private Sub AddLinkBtn_Click()

'//Reset connect trigger
If Range("AppActive").Value <> 1 Then
Range("ConnectTrig").Value = 0
End If
        
xPos = 0

Call App_CLICK.AddLink_Clk(xPos)

End Sub
Private Sub RmvLinkBtn_Click()

'//Reset connect trigger
If Range("AppActive").Value <> 1 Then
Range("ConnectTrig").Value = 0
End If

Call App_CLICK.RmvLink_Clk

End Sub
Private Sub AddUserBtn_Click()

'//Reset connect trigger
If Range("AppActive").Value <> 1 Then
Range("ConnectTrig").Value = 0
End If

xPos = 0
Call App_CLICK.AddUser_Clk(xPos)

End Sub
Private Sub RmvUserBtn_Click()

'//Reset connect trigger
If Range("AppActive").Value <> 1 Then
Range("ConnectTrig").Value = 0
End If

Call App_CLICK.RmvUser_Clk

End Sub
Private Sub LinkerBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call App_CLICK.RmvLinkerBox_EnClk

End Sub
Private Sub LinkerBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'//ENTER KEY
If KeyCode = 13 Then

Call App_CLICK.RmvLinkerBox_EnClk

End If

End Sub


Private Sub xlFlowStrip_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

xKey = KeyCode.Value
Call App_TOOLS.RunFlowStrip(xKey)
        
End Sub
Private Sub xlFlowStripBar_Click()

Call App_CLICK.xlFlowStripBar_Clk

End Sub
Private Sub SavePostBtn_Click()

Call App_CLICK.SavePost_Clk

End Sub
Private Sub ConnectBtn_Click()

Call App_TOOLS.xDisable
Call App_CLICK.ConnectPost_Clk

End Sub
Private Sub HomeBtn_Click()

Call App_Focus.HdForms
Call App_Focus.SH_ETWEETXLHOME

End Sub
Private Sub LogoBg_Click()

ETWEETXLPOST.Hide
ETWEETXLHOME.Show

End Sub
Private Sub UsersHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

xArt = "<lib>xtwt;winform(3);add.user(*);$" '//xlas
Call lexKey(xArt)


End Sub

Private Sub ViewMedBtn_Click()

App_CLICK.ViewMedBtn_Clk

End Sub
Private Sub DraftFilterBtn_Click()

If DraftFilterBtn.Caption = "..." Then xFil = 0
If DraftFilterBtn.Caption = "•" Then xFil = 1
        
Call App_CLICK.DraftFilterBtn_Clk(xFil)

End Sub
'//hover effects
Private Sub HelpIcon_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 14
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub


Private Sub HelpStatus_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 14
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub DraftHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 1: Call HoverHelp(xMsg)

xHov = 1
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub DraftsHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 2: Call HoverHelp(xMsg)

xHov = 2
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub OffsetHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 3: Call HoverHelp(xMsg)

xHov = 3
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub PostHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 4: Call HoverHelp(xMsg)

xHov = 4
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub RuntimeHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 5: Call HoverHelp(xMsg)

xHov = 5
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub TimeHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 6: Call HoverHelp(xMsg)

xHov = 6
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub UserHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 7: Call HoverHelp(xMsg)

xHov = 7
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub UsersHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 8: Call HoverHelp(xMsg)

xHov = 8
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub LinkerHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 9: Call HoverHelp(xMsg)

xHov = 9
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub FreezeBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 10: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault
xHov = 15: Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub LogoBg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 11: Call HoverHelp(xMsg)

xHov = 10
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub CtrlBoxBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 12: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault
xHov = 16: Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub SendAPI_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 13: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub DynOffset_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 14: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub DraftFilterBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 15: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub RmvAllDrafts_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 16: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub RmvDraft_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 17: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub AddDraft_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 18: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub xlFlowStripBar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 19: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub ActivePresetBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 20: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub AddMedBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 29: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub RmvMedBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 30: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub

Private Sub ViewMedBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 31: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub SavePostBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 32: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub AddThreadBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 33: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub RmvThreadBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 34: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub RmvAllThreadBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 35: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub ConnectBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 36: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub AddUserBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 37: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub RmvUserBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 38: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub AddLinkBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 39: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub RmvLinkBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 40: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub AddRuntime_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 41: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub RmvRuntime_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 42: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub UserBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 43: Call HoverHelp(xMsg)
xHov = 11
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub LinkerBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 44: Call HoverHelp(xMsg)
xHov = 12
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

MouseMove = 0
Button = 0

End Sub
Private Sub RuntimeBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 45: Call HoverHelp(xMsg)
xHov = 13
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub SaveLinkerBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)


xMsg = 46: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub LoadLinkerBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)


xMsg = 47: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub ReloadLinkerBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)


xMsg = 48: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub ClrSetupBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)


xMsg = 49: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub LastLinkBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)


xMsg = 50: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub PostBg1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub
Private Sub PostBg2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub

Private Sub PostBg3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub

Private Sub PostBg4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub

Private Sub PostBg5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub

Private Sub PostBg6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub


