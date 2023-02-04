VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ETWEETXLHOME 
   Caption         =   "eTweetXL"
   ClientHeight    =   9615.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960.001
   OleObjectBlob   =   "ETWEETXLHOME.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ETWEETXLHOME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

'ActiveUser.BackColor = -2147483633
Range("DataPullTrig").Value2 = 0

'//Refresh progress bar...
Call eTweetXL_TOOLS.updProgBar

'//Set default web browser (Firefox)
If Dir(AppLoc & "\mtsett\webset.txt") = "" Then _
Open AppLoc & "\mtsett\webset.txt" For Output As #1: Print #1, "Firefox": Close #1

End Sub
Private Sub UserForm_Activate()

On Error Resume Next

'//Check Linker for info
If Range("LinkerTotal").Value2 > 0 Then
Range("LinkTrig").Value = 1
End If

'//WinForm #
xWin = 11: Call setWindow(xWin)

'//Update application state
Call eTweetXL_TOOLS.updAppState

'//Refresh progress bar
Call eTweetXL_TOOLS.updProgBar

'//Set navigation font colors back to black
Call eTweetXL_TOOLS.dfsNaviBar

'//Window message
If Me.xlFlowStrip.Value = vbNullString Or Range("AppState").Value2 <> 1 Then Me.xlFlowStrip.Value = eTweetXL_INFO.AppWelcome

End Sub
Private Sub CtrlBoxBtn_Click()

Call eTweetXL_FOCUS.shw_CTRLBOX

End Sub
Private Sub xlFlowStrip_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'//WinForm #
xWin = 11: Call setWindow(xWin)

xKey = KeyCode.Value
Call eTweetXL_TOOLS.runFlowStrip(xKey)
KeyCode.Value = xKey

End Sub
Private Sub AppTag_Click()

ETWEETXLHOME.Hide
ETWEETXLQUEUE.Show

End Sub
Private Sub xlFlowStripBar_Click()

Call eTweetXL_CLICK.xlFlowStripBar_Clk

End Sub
Private Sub ProfileListBox_Change()

Range("DataPullTrig").Value2 = 0

If ProfileListBox.Value <> "" Then
Range("Profile").Value2 = ProfileListBox.Value
'ETWEETXLSETUP.ActiveUser.Caption = Range("User").Value
End If

End Sub
Private Sub ProfileListBox_Click()

Call eTweetXL_GET.getProfileNames

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
Private Sub HideBtn_Click()

Call eTweetXL_CLICK.HideBtn_Clk

End Sub
'//Start button
Private Sub StartBtn_Click()

Call eTweetXL_CLICK.StartBtn_Clk

End Sub
'Profile setup button
Private Sub ProfileSetupBtn_Click()

Me.Hide
ETWEETXLSETUP.Show

End Sub
'Post setup button
Private Sub PostSetupBtn_Click()

Me.Hide
ETWEETXLPOST.Show

End Sub
Private Sub FreezeBtn_Click()

Call eTweetXL_CLICK.FreezeBtn_Clk

End Sub
'Queue button
Private Sub QueueBtn_Click()

Me.Hide
ETWEETXLQUEUE.Show

End Sub
'//Break button
Private Sub BreakBtn_Click()

eTweetXL_CLICK.BreakBtn_Clk

End Sub
'//hover effects
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
xHov = 14
Call eTweetXL_TOOLS.dfsHover
Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub HelpStatus_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xHov = 14
Call eTweetXL_TOOLS.dfsHover
Call eTweetXL_TOOLS.fxsHover(xHov)

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

'On Error Resume Next

'Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
'xMsg = 20: Call HoverHelp(xMsg)

'Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub AppStatus_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 21: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub ProgBarBg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 22: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub HideBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover
xHov = 17: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub HomeBg1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub

Private Sub UserForm_Terminate()

Range("DataPullTrig").Value2 = 0

If Range("LinkTrig").Value2 <> 1 Then
Call clnMain
End If

'//backup current queue state
Call eTweetXL_POST.pstLastQueue

Range("xlasWinForm").Value2 = Range("xlasWinFormLast").Value2

End Sub
