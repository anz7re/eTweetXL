VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CTRLBOX 
   ClientHeight    =   9615.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   27255
   OleObjectBlob   =   "CTRLBOX.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CTRLBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

Application.EnableCancelKey = xlDisabled

Call getEnvironment(appEnv, appBlk)

'//Record previous WinForm (for switching back to an application window)
Workbooks(appEnv).Worksheets(appBlk).Range("xlasBlkAddr100").Value2 = Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2

Me.Caption = CtrlTag
Range("xlasSilent").Value2 = 0
Range("xlasInvert").Value2 = 0
Range("xlasRemember").Value = 0
Range("xlasKeyCtrl").Value2 = 0
Range("xlasSaveFile").Value2 = vbNullString

'//Add menu hover options...
Call addOptions(xType)

'//set default window size
Call dfsWindow

End Sub
Private Sub UserForm_Activate()

Application.EnableCancelKey = xlInterrupt

Call getEnvironment(appEnv, appBlk)

'//WinForm #
xWin = 100: Call setWindow(xWin)

'//Update application state
Call CtrlBox_TOOLS.updAppState

If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinFormLast").Value2 = vbNullString Then Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinFormLast").Value2 = 10

FileSel.Visible = False
EditSel.Visible = False
DebugSel.Visible = False
OptionsSel.Visible = False
RunSel.Visible = False
HelpSel.Visible = False
WindowSel.Visible = False

CTRLBOX.RemWinSizeValue.Caption = "0"

If Workbooks(appEnv).Worksheets(appBlk).Range("xlasConsoleType").Value2 <> "" Then
CtrlBoxWindow.Value = Workbooks(appEnv).Worksheets(appBlk).Range("xlasConsoleType").Value2
End If

End Sub
Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Me.Top = 0
Me.Left = 0

End Sub
Private Sub CtrlBoxWindow_Change()

Call getEnvironment(appEnv, appBlk)

'//WinForm #
xWin = 101: Call setWindow(xWin)

'//Set default screen state
Call dfsMainScreen

'//Check if remembering
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasRemember").Value = 1 Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasAMemory").Value = CtrlBoxWindow.Value
End If

'//WinForm #
xWin = 100: Call setWindow(xWin)

'//Set Window statistics
Call setWindowStats

'//Cleanup
If InStr(1, CtrlBoxWindow, "cls;") Then
CtrlBoxWindow.Value = vbNullString
End If

End Sub
Private Sub CtrlBoxWindow_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Call getEnvironment(appEnv, appBlk)

'//WinForm #
xWin = 101: Call setWindow(xWin)

Call CtrlBox_TOOLS.dfsHover

End Sub
Private Sub CtrlBoxWindow_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

Dim Art As String

Call getEnvironment(appEnv, appBlk)

Art = CTRLBOX.CtrlBoxWindow.Value

'//WinForm #
xWin = 101: Call setWindow(xWin)

'//Shift key (run script)
If KeyCode.Value = vbKeyShift Then
Workbooks(appEnv).Worksheets(appBlk).Activate
If InStr(1, Art, "$") Then '//check for run trigger
On Error GoTo ErrMsg
CTRLBOX.Hide '//hide control box
Call xlas(Art)
Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinFormLast").Value2 '//set window to previous
Workbooks(appEnv).Worksheets(appBlk).Range("xlasConsoleType").Value2 = CtrlBoxWindow.Value '//save control box text
Workbooks(appEnv).Save
Workbooks(appEnv).Worksheets(appBlk).Activate
CTRLBOX.Show '//show control box
Exit Sub
    End If
        End If

'//Enter key
If KeyCode.Value = 13 Then
CtrlBoxWindow.EnterKeyBehavior = True
Exit Sub
End If

'//Tab key
If KeyCode.Value = 9 Then
CtrlBoxWindow.TabKeyBehavior = True
Exit Sub
End If

'//Alt key
If KeyCode.Value = 18 Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 + 18
KeyCode.Value = 0
Exit Sub
End If

'//Ctrl key
If KeyCode.Value = vbKeyControl Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbKeyControl
KeyCode.Value = 0
Exit Sub
End If

'//Key Ctrl+D
If KeyCode.Value = vbKeyD Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 17 Then
'//Clear screen...
Call Ctrlbox_CLICK.ClearScreen_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+F
If KeyCode.Value = vbKeyF Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 17 Then
'//Clear screen...
XLFONTBOX.Show
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
'//Key Ctrl+I
If KeyCode.Value = vbKeyI Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 17 Then
'//Invert window hotkey...
Call Ctrlbox_CLICK.InvertScreen_Clk
Call CtrlBox_TOOLS.dfsMainScreen
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
'//Key Ctrl+H
If KeyCode.Value = vbKeyH Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 17 Then
'//Replace tool hotkey...
XLREPLACE.Show
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+N
If KeyCode.Value = vbKeyN Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 17 Then
'//Create a new project hotkey...
Call Ctrlbox_CLICK.NewFile_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
'//Key Ctrl+O
If KeyCode.Value = vbKeyO Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 17 Then
'//Open new project hotkey...
Call Ctrlbox_CLICK.OpenFile_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
'//Key Ctrl+R
If KeyCode.Value = vbKeyR Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 17 Then
'//Remember typing hotkey...
Call Ctrlbox_CLICK.Remember_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
'//Key Ctrl+S
If KeyCode.Value = vbKeyS Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 17 Then
'//Save project hotkey...
Call Ctrlbox_CLICK.SaveFile_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Q
If KeyCode.Value = vbKeyQ Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 17 Then
'//Close ctrl box hotkey w/o saving first...
ThisWorkbook.Close
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
       
'//Key Ctrl+W
If KeyCode.Value = vbKeyW Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 17 Then
'//Minimize window size
If Range("xlasBlkAddr105").Value2 = 1 Then
Range("xlasBlkAddr105").Value2 = 0 '//options trigger
Range("xlasBlkAddr107").Value2 = 0 '//taskbar trigger
Call Ctrlbox_CLICK.Maximize_Clk
WindowSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): xType = 0: Call addOptions(xType)
Exit Sub
End If
'//Maximize window size
Range("xlasBlkAddr105").Value2 = 1 '//options trigger
Range("xlasBlkAddr107").Value2 = 1 '//taskbar trigger
Call Ctrlbox_CLICK.Maximize_Clk
WindowSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): xType = 1: Call addOptions(xType)
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Up
If KeyCode.Value = vbKeyUp Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 17 Then
'//Increase zoom
Call Ctrlbox_CLICK.ZoomUp_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Down
If KeyCode.Value = vbKeyDown Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 17 Then
'//Decrease zoom
Call Ctrlbox_CLICK.ZoomDown_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Alt+Q
If KeyCode.Value = vbKeyQ Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 35 Then
'//Save and close ctrl box hotkey...
ThisWorkbook.Save: ThisWorkbook.Close
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Alt+S
If KeyCode.Value = vbKeyS Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 35 Then
'//Save as project hotkey...
Call Ctrlbox_CLICK.SaveAsFile_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Alt+R
If KeyCode.Value = vbKeyR Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 35 Then
'//Recall memory hotkey...
Call Ctrlbox_CLICK.Recall_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
     
'//Key Ctrl+Alt+W
If KeyCode.Value = vbKeyW Then
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = 35 Then
'//Hide hotkey...
Call Ctrlbox_CLICK.Hide_Clk
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

Workbooks(appEnv).Worksheets(appBlk).Range("xlasKeyCtrl").Value2 = vbNullString

Exit Sub
ErrMsg:
xMsg = 2: Call CtrlBox_MSG.AppMsg(xMsg, errLvl)
CTRLBOX.Show '//show control box

End Sub
Private Sub FileBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 1
Call CtrlBox_TOOLS.fxsHover(xHov)
Call CtrlBox_TOOLS.undHover(xHov)

End Sub
Private Sub FileBtn_Click()

xBtn = 1
Call CtrlBox_TOOLS.dfsHover
Call CtrlBox_TOOLS.selOption(xBtn)

End Sub
Private Sub FileSel_Click()

'//Start new project
If InStr(1, FileSel.Value, "New ") Then
Call Ctrlbox_CLICK.NewFile_Clk
FileSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Open project
If InStr(1, FileSel.Value, "Open ") Then
Call Ctrlbox_CLICK.OpenFile_Clk
FileSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Save As project
If InStr(1, FileSel.Value, "Save As ") Then
Call Ctrlbox_CLICK.SaveAsFile_Clk
FileSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Save project
If InStr(1, FileSel.Value, "Save  ") Then
Call Ctrlbox_CLICK.SaveFile_Clk
FileSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Save & exit
If InStr(1, FileSel.Value, "Save & Exit") Then
ThisWorkbook.Save: ThisWorkbook.Close
FileSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Exit
If InStr(1, FileSel.Value, "Exit ") Then
ThisWorkbook.Close
FileSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

FileSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)

End Sub
Private Sub EditBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 2
Call CtrlBox_TOOLS.fxsHover(xHov)
Call CtrlBox_TOOLS.undHover(xHov)

End Sub
Private Sub EditBtn_Click()

xBtn = 2
Call CtrlBox_TOOLS.dfsHover
Call CtrlBox_TOOLS.selOption(xBtn)

End Sub
Private Sub EditSel_Click()

'//Clear screen
If InStr(1, EditSel.Value, "Clear Screen ") Then
Call Ctrlbox_CLICK.ClearScreen_Clk
EditSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Undo
If InStr(1, EditSel.Value, "Undo ") Then
CTRLBOX.CtrlBoxWindow.SetFocus
Art = "<lib>xbas;app.key(^z);$"
Call xlas(Art)
EditSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Cut
If InStr(1, EditSel.Value, "Cut ") Then
CTRLBOX.CtrlBoxWindow.SetFocus
Art = "<lib>xbas;app.key(^x);$"
Call xlas(Art)
EditSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Copy
If InStr(1, EditSel.Value, "Copy ") Then
CTRLBOX.CtrlBoxWindow.SetFocus
Art = "<lib>xbas;app.key(^c);$"
Call xlas(Art)
EditSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Paste
If InStr(1, EditSel.Value, "Paste ") Then
CTRLBOX.CtrlBoxWindow.SetFocus
Art = "<lib>xbas;app.key(^v);$"
Call xlas(Art)
EditSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Replace
If InStr(1, EditSel.Value, "Replace ") Then
CTRLBOX.CtrlBoxWindow.SetFocus
Art = "<lib>xbas;app.key(^h);$"
Call xlas(Art)
EditSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Select all
If InStr(1, EditSel.Value, "Select All ") Then
CTRLBOX.CtrlBoxWindow.SetFocus
Art = "<lib>xbas;app.key(^a);$"
Call xlas(Art)
EditSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

EditSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)

End Sub
Private Sub DebugBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 3
Call CtrlBox_TOOLS.fxsHover(xHov)
Call CtrlBox_TOOLS.undHover(xHov)

End Sub
Private Sub DebugBtn_Click()

xBtn = 3
Call CtrlBox_TOOLS.dfsHover
Call CtrlBox_TOOLS.selOption(xBtn)

End Sub
Private Sub DebugSel_Click()

''//Change font/background settings
'If InStr(1, DebugSel.Value, "Comment Block ") Then
'Call Ctrlbox_CLICK.CommentBlock_Clk
'DebugSel.Clear: Art = "<lib>xbas;delayevent(10);$": call xlas(Art): Call addOptions(xType)
'End If
'
''//Change font/background settings
'If InStr(1, DebugSel.Value, "Uncomment Block ") Then
'Call Ctrlbox_CLICK.UncommentBlock_Clk
'DebugSel.Clear: Art = "<lib>xbas;delayevent(10);$": call xlas(Art): Call addOptions(xType)
'End If


End Sub
Private Sub OptionsBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 4
Call CtrlBox_TOOLS.fxsHover(xHov)
Call CtrlBox_TOOLS.undHover(xHov)

End Sub
Private Sub OptionsBtn_Click()

xBtn = 4
Call CtrlBox_TOOLS.dfsHover
Call CtrlBox_TOOLS.selOption(xBtn)

End Sub
Private Sub OptionsSel_Click()

'//Change font/background settings
If InStr(1, OptionsSel.Value, "Screen Style ") Then
XLFONTBOX.Show
OptionsSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
End If

OptionsSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)

End Sub
Private Sub RunBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 5
Call CtrlBox_TOOLS.fxsHover(xHov)
Call CtrlBox_TOOLS.undHover(xHov)

End Sub
Private Sub RunBtn_Click()

xBtn = 5
Call CtrlBox_TOOLS.dfsHover
Call CtrlBox_TOOLS.selOption(xBtn)

End Sub
Private Sub RunSel_Click()

Dim Art As String

Call getEnvironment(appEnv, appBlk)

Art = CTRLBOX.CtrlBoxWindow.Value

'//Run Script...
If InStr(1, RunSel.Value, "Run Script ") Then
Workbooks(appEnv).Worksheets(appBlk).Activate
On Error GoTo ErrMsg
Art = Art & "$"
CTRLBOX.Hide '//hide control box
Call xlas(Art)
Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinFormLast").Value2 '//set window to previous
Workbooks(appEnv).Worksheets(appBlk).Range("xlasConsoleType").Value2 = CtrlBoxWindow.Value '//save control box text
Workbooks(appEnv).Save
Workbooks(appEnv).Worksheets(appBlk).Activate
Unload CTRLBOX
CTRLBOX.Show '//show control box
Exit Sub
        End If

RunSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)

Exit Sub
ErrMsg:
xMsg = 2: Call CtrlBox_MSG.AppMsg(xMsg, errLvl)
CTRLBOX.Show '//show control box

End Sub
Private Sub WindowBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 6
Call CtrlBox_TOOLS.fxsHover(xHov)
Call CtrlBox_TOOLS.undHover(xHov)

End Sub
Private Sub WindowBtn_Click()

xBtn = 6
Call CtrlBox_TOOLS.dfsHover
Call CtrlBox_TOOLS.selOption(xBtn)

End Sub
Private Sub WindowSel_Click()

'//Hide Control Box+
If InStr(1, WindowSel.Value, "Hide ") Then
Call Ctrlbox_CLICK.Hide_Clk
WindowSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Invert window colors
If InStr(1, WindowSel.Value, "Invert Screen ") Then
Call Ctrlbox_CLICK.InvertScreen_Clk
Call CtrlBox_TOOLS.dfsMainScreen
WindowSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Remember everything being typed into the Control Box until turned off
If InStr(1, WindowSel.Value, "Remember ") Then
Call Ctrlbox_CLICK.Remember_Clk
WindowSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Recall what was remembered
If InStr(1, WindowSel.Value, "Recall ") Then
Call Ctrlbox_CLICK.Recall_Clk
WindowSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Maximize window
If InStr(1, WindowSel.Value, "Maximize ") Then
Range("xlasBlkAddr105").Value2 = 1 '//options trigger
Range("xlasBlkAddr107").Value2 = 1 '//taskbar trigger
Call Ctrlbox_CLICK.Maximize_Clk
WindowSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): xType = 1: Call addOptions(xType)
Exit Sub
End If

'//Minimize window
If InStr(1, WindowSel.Value, "Minimize ") Then
Range("xlasBlkAddr105").Value2 = 0 '//options trigger
Range("xlasBlkAddr107").Value2 = 0 '//taskbar trigger
Call Ctrlbox_CLICK.Maximize_Clk
WindowSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): xType = 0: Call addOptions(xType)
Exit Sub
End If

'//Zoom in
If InStr(1, WindowSel.Value, "Zoom In ") Then
Call Ctrlbox_CLICK.ZoomUp_Clk
WindowSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Zoom out
If InStr(1, WindowSel.Value, "Zoom Out ") Then
Call Ctrlbox_CLICK.ZoomDown_Clk
WindowSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

WindowSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)

End Sub
Private Sub HelpBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 7
Call CtrlBox_TOOLS.fxsHover(xHov)
Call CtrlBox_TOOLS.undHover(xHov)

End Sub
Private Sub HelpBtn_Click()

xBtn = 7
Call CtrlBox_TOOLS.dfsHover
Call CtrlBox_TOOLS.selOption(xBtn)

End Sub
Private Sub HelpSel_Click()

'//Basic application information
If InStr(1, HelpSel.Value, "About Control Box+ ") Then
CTRLBOXABOUT.Show
HelpSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

'//Send feedback
If InStr(1, HelpSel.Value, "Send Feedback ") Then
Call Ctrlbox_CLICK.SendFeedback_Clk
HelpSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If


'//Frequently asked questions
If InStr(1, HelpSel.Value, "FAQ ") Then
HelpSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)
Exit Sub
End If

HelpSel.Clear: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): Call addOptions(xType)

End Sub
Private Sub RemWinSize_Click()

RemWinSizeValue = "0"
CtrlBoxWindow.Font.Size = 12
Call setWindowStats

End Sub
Private Sub UserForm_Terminate()

Call getEnvironment(appEnv, appBlk)

Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = Workbooks(appEnv).Worksheets(appBlk).Range("xlasBlkAddr100").Value2
Workbooks(appEnv).Worksheets(appBlk).Range("xlasBlkAddr100").Value2 = vbNullString
Workbooks(appEnv).Worksheets(appBlk).Range("xlasConsoleType").Value2 = CtrlBoxWindow.Value

End Sub
