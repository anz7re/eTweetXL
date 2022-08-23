VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ETWEETXLSETUP 
   Caption         =   "eTweetXL"
   ClientHeight    =   11220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13800
   OleObjectBlob   =   "ETWEETXLSETUP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ETWEETXLSETUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

'//Add default browser
BrowserBox.AddItem ("Firefox")

'//Welcome
If Me.xlFlowStrip.Value = "" Then Me.xlFlowStrip.Value = eTweetXL_INFO.AppWelcome

Range("DataPullTrig").Value = 0
Range("EditStatus").Value = 0

End Sub
Private Sub UserForm_Activate()

'//Set default browser
ETWEETXLSETUP.BrowserBox.Value = "Firefox"

'//Cleanup
ETWEETXLSETUP.ProfileListBox.Clear

'//Show runtime action message
Call eTweetXL_GET.getRtState

'//WinForm #
xWin = 12: Call setWindow(xWin)

'//Update application state
Call eTweetXL_TOOLS.updAppState

'//Import profile information
Call eTweetXL_GET.getProfileNames

'//Window title
If Me.xlFlowStrip.Value = vbNullString Or Range("AppState").Value2 <> 1 Then Me.xlFlowStrip.Value = "Profile Setup..."

End Sub
Private Sub FreezeBtn_Click()

Call eTweetXL_CLICK.FreezeBtn_Clk

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
Private Sub APISetupBtn_Click()

ETWEETXLAPISETUP.Show

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
Private Sub HideBtn_Click()

Call eTweetXL_CLICK.HideBtn_Clk

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
Private Sub HomeBtn_Click()

Range("DataPullTrig").Value = 0
Me.Hide
Call eTweetXL_FOCUS.shw_ETWEETXLHOME

End Sub
Private Sub ProfileListBox_Click()

If Range("Profile").Value2 = ETWEETXLSETUP.ProfileNameBox.Value Then
    If ProfileListBox.Value = "" Then
          ProfileListBox.Value = ETWEETXLSETUP.ProfileNameBox.Value
                End If
                    End If

If ProfileListBox.Value <> "" Then
Range("Profile").Value2 = ProfileListBox.Value
ProfileNameBox.Value = ProfileListBox.Value
End If

Range("DataPullTrig").Value = 0

If InStr(1, xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
xlFlowStrip.Value = ProfileListBox.Value & "..."
End If

Call eTweetXL_GET.getProfileData

End Sub
Private Sub NewUser_Click()

If InStr(1, UsernameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(12);mk.user( -list" & UsernameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If

xInfo = vbNullString
Call eTweetXL_CLICK.NewUser_Clk(xInfo)

End Sub
Private Sub RmvAllUsers_Click()

Call eTweetXL_CLICK.DelAllUsersBtn_Clk

End Sub
Private Sub RmvUser_Click()

If InStr(1, UsernameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(12);del.user( -list" & UsernameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If

xInfo = vbNullString
Call eTweetXL_CLICK.DelUserBtn_Clk(xInfo)

End Sub
Private Sub ShowPassBox_Click()

xUser = UserListBox.Value
xUser = Replace(UserListBox.Value, Range("Scure").Value, "")
Call eTweetXL_GET.getTargetData(xUser)

End Sub
Private Sub NewProfile_Click()

If InStr(1, ProfileNameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(12);mk.profile( -list" & ProfileNameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If

xInfo = vbNullString
Call eTweetXL_CLICK.NewProfile_Clk(xInfo)

End Sub
Private Sub RmvAllProfiles_Click()

Call eTweetXL_CLICK.RmvAllProfilesBtn_Clk

End Sub
Private Sub RmvProfile_Click()

If InStr(1, ProfileNameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(12);del.profile( -list" & ProfileNameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If

xInfo = vbNullString
Call eTweetXL_CLICK.RmvProfileBtn_Clk(xInfo)

End Sub
Private Sub PassBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

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

'//Key Ctrl+S
If KeyCode.Value = vbKeyS Then
If Range("xlasKeyCtrl").Value2 = 17 Then
If InStr(1, UsernameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(12);mk.user( -list" & UsernameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If
Call eTweetXL_CLICK.NewUser_Clk(xInfo)
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If


End Sub
Private Sub ProfileNameBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)


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

'//Key Ctrl+Alt+D
If KeyCode.Value = vbKeyD Then
If Range("xlasKeyCtrl").Value2 = 35 Then
If InStr(1, ProfileNameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(12);del.profile( -list" & ProfileNameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If
xInfo = vbNullString
Call eTweetXL_CLICK.RmvProfileBtn_Clk(xInfo)
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+S
If KeyCode.Value = vbKeyS Then
If Range("xlasKeyCtrl").Value2 = 17 Then
If InStr(1, ProfileNameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(12);mk.profile( -list" & ProfileNameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If
xInfo = vbNullString
Call eTweetXL_CLICK.NewProfile_Clk(xInfo)
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

Range("xlasKeyCtrl").Value2 = vbNullString

End Sub
Private Sub UsernameBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

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

'//Key Ctrl+Alt+D
If KeyCode.Value = vbKeyD Then
If Range("xlasKeyCtrl").Value2 = 35 Then
If InStr(1, UsernameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(12);del.user( -list" & UsernameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If
xInfo = vbNullString
Call eTweetXL_CLICK.DelUserBtn_Clk(xInfo)
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+S
If KeyCode.Value = vbKeyS Then
If Range("xlasKeyCtrl").Value2 = 17 Then
If InStr(1, UsernameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(12);mk.user( -list" & UsernameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If
Call eTweetXL_CLICK.NewUser_Clk(xInfo)
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

Range("xlasKeyCtrl").Value2 = vbNullString

End Sub
Private Sub xlFlowStrip_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'//WinForm #
xWin = 12: Call setWindow(xWin)

xKey = KeyCode.Value
Call eTweetXL_TOOLS.runFlowStrip(xKey)
KeyCode.Value = xKey
        
End Sub
Private Sub xlFlowStripBar_Click()

Call eTweetXL_CLICK.xlFlowStripBar_Clk

End Sub
Private Sub UserListBox_Change()

xUser = Replace(UserListBox.Value, Range("Scure").Value, "")
UsernameBox.Value = xUser
Range("User").Value = xUser
Call eTweetXL_GET.getTargetData(xUser)

End Sub
Private Sub UsernameBox_Change()

If UsernameBox.SpecialEffect <> fmSpecialEffectSunken Then
UsernameBox.SpecialEffect = fmSpecialEffectSunken
End If

End Sub
Private Sub PassBox_Change()

If PassBox.SpecialEffect <> fmSpecialEffectSunken Then
PassBox.SpecialEffect = fmSpecialEffectSunken
End If

End Sub
Private Sub ProfileNameBox_Change()

If ProfileNameBox.SpecialEffect <> fmSpecialEffectSunken Then
ProfileNameBox.SpecialEffect = fmSpecialEffectSunken
End If

End Sub
Private Sub AppTag_Click()

ETWEETXLSETUP.Hide
ETWEETXLHOME.Show

End Sub
'//hover effects
Private Sub HomeBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 0
Call eTweetXL_TOOLS.fxsNaviBar(xBtn)
Call eTweetXL_TOOLS.undNaviBar(xBtn)

End Sub
Private Sub StartBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 1
Call eTweetXL_TOOLS.fxsNaviBar(xBtn)
Call eTweetXL_TOOLS.undNaviBar(xBtn)

xMsg = 52: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub ProfileSetupBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 2
Call eTweetXL_TOOLS.fxsNaviBar(xBtn)
Call eTweetXL_TOOLS.undNaviBar(xBtn)

xMsg = 54: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub PostSetupBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 3
Call eTweetXL_TOOLS.fxsNaviBar(xBtn)
Call eTweetXL_TOOLS.undNaviBar(xBtn)

xMsg = 55: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub QueueBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 4
Call eTweetXL_TOOLS.fxsNaviBar(xBtn)
Call eTweetXL_TOOLS.undNaviBar(xBtn)

xMsg = 53: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub BreakBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

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
Private Sub xlFlowStripBar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 19: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub ActiveUser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 20: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub RmvProfile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 23: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub

Private Sub RmvUser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 24: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub NewProfile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 25: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub NewUser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 26: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub RmvAllProfiles_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 27: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub RmvAllUsers_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 28: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub SetupBg1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub SetupBg2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub SetupBg3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub SetupBg4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub HideBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover
xHov = 17: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub UserForm_Terminate()

Range("DataPullTrig").Value = 0

Range("xlasWinForm").Value2 = Range("xlasWinFormLast").Value2

End Sub


