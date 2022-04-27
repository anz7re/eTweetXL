VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ETWEETXLHOME 
   Caption         =   "eTweetXL"
   ClientHeight    =   9165.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11025
   OleObjectBlob   =   "ETWEETXLHOME.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ETWEETXLHOME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

Me.Caption = AppTag
ActivePresetBox.BackColor = -2147483633
Range("DataPullTrig").Value = 0

'//Refresh progress bar...
Call App_TOOLS.ProgBarRefresher

'//Set default web browser (Firefox)
If Dir(AppLoc & "\mtsett\webset.txt") = "" Then _
Open AppLoc & "\mtsett\webset.txt" For Output As #1: Print #1, "Firefox": Close #1

End Sub
Private Sub UserForm_Activate()

On Error Resume Next

'//WinForm #
Range("xlasWinForm").Value = 1

'//Check Linker for info
If Range("LinkerTotal").Value > 0 Then
Range("LinkTrig").Value = 1
End If

'//Update active state
Call App_TOOLS.UpdateActive

'//Refresh progress bar
Call App_TOOLS.ProgBarRefresher

'//Set navigation font colors back to black
Call App_TOOLS.NaviBarDefault

'//Window message
If Me.xlFlowStrip.Value = vbNullString Or Range("AppActive").Value <> 1 Then Me.xlFlowStrip.Value = App_INFO.AppWelcome

End Sub
Private Sub CtrlBoxBtn_Click()

Call App_Focus.SH_CTRLBOX

End Sub
Private Sub xlFlowStrip_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

xKey = KeyCode.Value
Call App_TOOLS.RunFlowStrip(xKey)

End Sub
Private Sub LogoBg_Click()

ETWEETXLHOME.Hide
ETWEETXLQUEUE.Show

End Sub
Private Sub xlFlowStripBar_Click()

Call App_CLICK.xlFlowStripBar_Clk

End Sub
Private Sub OffsetTimer_Click()

App_CLICK.OffsetTimer_Clk

End Sub
Private Sub ProfileListBox_Change()

Range("DataPullTrig").Value = 0

If ProfileListBox.Value <> "" Then
Range("Profile").Value = ProfileListBox.Value
ETWEETXLSETUP.ActivePresetBox.Caption = Range("User").Value
End If

End Sub
Private Sub ProfileListBox_Click()

Call App_IMPORT.MyProfileNames

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
'//Start button
Private Sub StartBtn_Click()

Call App_CLICK.Start_Clk

End Sub
Private Sub StartBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 1
Call App_TOOLS.NaviBarColor(xBtn)
Call App_TOOLS.NaviBarUnderline(xBtn)

xMsg = 52: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
'Profile setup button
Private Sub ProfileSetupBtn_Click()

Me.Hide
ETWEETXLSETUP.Show

End Sub
Private Sub ProfileSetupBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 2
Call App_TOOLS.NaviBarColor(xBtn)
Call App_TOOLS.NaviBarUnderline(xBtn)

xMsg = 54: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
'Post setup button
Private Sub PostSetupBtn_Click()

Me.Hide
ETWEETXLPOST.Show

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
'Queue button
Private Sub QueueBtn_Click()

Me.Hide
ETWEETXLQUEUE.Show

End Sub
Private Sub QueueBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 4
Call App_TOOLS.NaviBarColor(xBtn)
Call App_TOOLS.NaviBarUnderline(xBtn)

xMsg = 53: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
'//Break button
Private Sub BreakBtn_Click()

App_CLICK.Break_Clk

End Sub
Private Sub BreakBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 5
Call App_TOOLS.NaviBarColor(xBtn)
Call App_TOOLS.NaviBarUnderline(xBtn)


xMsg = 51: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

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

Call App_TOOLS.HoverDefault
xHov = 10: Call App_TOOLS.HoverEffect(xHov)

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
Private Sub LinkerActive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 21: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub ProgBarBg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 22: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

End Sub
Private Sub HomeBg1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub

Private Sub UserForm_Terminate()

Range("DataPullTrig").Value = 0

If Range("LinkTrig").Value <> 1 Then
Call Cleanup.ClnMainSpace
End If

End Sub
