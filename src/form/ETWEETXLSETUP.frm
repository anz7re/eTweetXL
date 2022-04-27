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

Me.Caption = AppTag

'//Add default browser
BrowserBox.AddItem ("Firefox")

'//Welcome
If Me.xlFlowStrip.Value = "" Then Me.xlFlowStrip.Value = App_INFO.AppWelcome

Range("DataPullTrig").Value = 0
Range("SetupEdit").Value = 0

End Sub
Private Sub UserForm_Activate()

'//WinForm #
Range("xlasWinForm").Value = 2

'//Set default browser
ETWEETXLSETUP.BrowserBox.Value = "Firefox"

'//Cleanup
ETWEETXLSETUP.ProfileListBox.Clear

'//Show runtime action message
Call App_TOOLS.ShowRtAction

'//Update active state
Call App_TOOLS.UpdateActive

'//Import profile information
Call App_IMPORT.MyProfileNames

'//Window title
If Me.xlFlowStrip.Value = vbNullString Or Range("AppActive").Value <> 1 Then Me.xlFlowStrip.Value = "Profile Setup..."

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
Private Sub PostSetupBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 3
Call App_TOOLS.NaviBarColor(xBtn)
Call App_TOOLS.NaviBarUnderline(xBtn)

xMsg = 55: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub FreezeBtn_Click()

Call App_CLICK.FreezeBtn_Clk

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
Private Sub APISetupBtn_Click()

ETWEETXLAPISETUP.Show

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
Private Sub HomeBtn_Click()

Range("DataPullTrig").Value = 0
Call App_Focus.HdForms
Call App_Focus.SH_ETWEETXLHOME

End Sub
Private Sub ProfileListBox_Click()

If Range("Profile").Value = ETWEETXLSETUP.ProfileNameBox.Value Then
    If ProfileListBox.Value = "" Then
          ProfileListBox.Value = ETWEETXLSETUP.ProfileNameBox.Value
                End If
                    End If

If ProfileListBox.Value <> "" Then
Range("Profile").Value = ProfileListBox.Value
ProfileNameBox.Value = ProfileListBox.Value
End If

Range("DataPullTrig").Value = 0

If InStr(1, xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
xlFlowStrip.Value = ProfileListBox.Value & "..."
End If

Call App_IMPORT.MyProfileData

End Sub
Private Sub NewUser_Click()

If InStr(1, UsernameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(2);mk.user( -list" & UsernameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If

xInfo = vbNullString
Call App_CLICK.NewUser_Clk(xInfo)

End Sub
Private Sub RmvAllUsers_Click()

Call App_CLICK.DelAllUsers_Clk

End Sub
Private Sub RmvUser_Click()

If InStr(1, UsernameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(2);del.user( -list" & UsernameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If

xInfo = vbNullString
Call App_CLICK.DelUser_Clk(xInfo)

End Sub
Private Sub ShowPassBox_Click()

xUser = UserListBox.Value
xUser = Replace(UserListBox.Value, Range("Scure").Value, "")
Call App_IMPORT.MyPassData(xUser)

End Sub
Private Sub NewProfile_Click()

If InStr(1, ProfileNameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(2);mk.profile( -list" & ProfileNameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If

xInfo = vbNullString
Call App_CLICK.NewProfile_Clk(xInfo)

End Sub
Private Sub RmvAllProfiles_Click()

Call App_CLICK.RmvAllProfiles_Clk

End Sub
Private Sub RmvProfile_Click()

If InStr(1, ProfileNameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(2);del.profile( -list" & ProfileNameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If

xInfo = vbNullString
Call App_CLICK.RmvProfile_Clk(xInfo)

End Sub
Private Sub PassBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

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

'//Key Ctrl+S
If KeyCode.Value = vbKeyS Then
If Range("xlasKeyCtrl").Value = 17 Then
If InStr(1, UsernameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(2);mk.user( -list" & UsernameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If
Call App_CLICK.NewUser_Clk(xInfo)
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If


End Sub
Private Sub ProfileNameBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)


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

'//Key Ctrl+Alt+D
If KeyCode.Value = vbKeyD Then
If Range("xlasKeyCtrl").Value = 35 Then
If InStr(1, ProfileNameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(2);del.profile( -list" & ProfileNameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If
xInfo = vbNullString
Call App_CLICK.RmvProfile_Clk(xInfo)
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+S
If KeyCode.Value = vbKeyS Then
If Range("xlasKeyCtrl").Value = 17 Then
If InStr(1, ProfileNameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(2);mk.profile( -list" & ProfileNameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If
xInfo = vbNullString
Call App_CLICK.NewProfile_Clk(xInfo)
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

Range("xlasKeyCtrl").Value = vbNullString

End Sub
Private Sub UsernameBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

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

'//Key Ctrl+Alt+D
If KeyCode.Value = vbKeyD Then
If Range("xlasKeyCtrl").Value = 35 Then
If InStr(1, UsernameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(2);del.user( -list" & UsernameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If
xInfo = vbNullString
Call App_CLICK.DelUser_Clk(xInfo)
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+S
If KeyCode.Value = vbKeyS Then
If Range("xlasKeyCtrl").Value = 17 Then
If InStr(1, UsernameBox.Value, ",") Then
xArt = "<lib>xtwt;winform(2);mk.user( -list" & UsernameBox.Value & ");$" '//xlas
Call lexKey(xArt)
Exit Sub
End If
Call App_CLICK.NewUser_Clk(xInfo)
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

Range("xlasKeyCtrl").Value = vbNullString

End Sub
Private Sub xlFlowStrip_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

xKey = KeyCode.Value
Call App_TOOLS.RunFlowStrip(xKey)
        
End Sub
Private Sub xlFlowStripBar_Click()

Call App_CLICK.xlFlowStripBar_Clk

End Sub
Private Sub UserListBox_Change()

xUser = Replace(UserListBox.Value, Range("Scure").Value, "")
UsernameBox.Value = xUser
Range("User").Value = xUser
Call App_IMPORT.MyPassData(xUser)

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
Private Sub LogoBg_Click()

ETWEETXLSETUP.Hide
ETWEETXLHOME.Show

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
Private Sub RmvProfile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

xMsg = 23: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub

Private Sub RmvUser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 24: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub NewProfile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

xMsg = 25: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub NewUser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

xMsg = 26: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub RmvAllProfiles_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

xMsg = 27: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub RmvAllUsers_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

xMsg = 28: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub SetupBg1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub
Private Sub SetupBg2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub
Private Sub SetupBg3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub
Private Sub SetupBg4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub
Private Sub UserForm_Terminate()

Range("DataPullTrig").Value = 0

End Sub


