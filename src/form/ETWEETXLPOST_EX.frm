VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ETWEETXLPOST_EX 
   Caption         =   "eTweetXL"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12480
   OleObjectBlob   =   "ETWEETXLPOST_EX.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ETWEETXLPOST_EX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()

'//WinForm #
xWin = 16: Call setWindow(xWin)

Range("AppendTrig").Value2 = 0: AppendBtn.ForeColor = &H80000011
Range("ReflectTrig").Value2 = 0: ReflectBtn.ForeColor = &H80000011

ETWEETXLPOST_EX.PostBox.Value = ETWEETXLPOST.PostBox.Value

End Sub
Private Sub PostBox_Change()

'//WinForm #
xWin = 161: Call setWindow(xWin)

Call eTweetXL_CHANGE.PostBox_Chg

If Range("ReflectTrig").Value = 1 Then
If Range("AppendTrig").Value = 1 Then
    ETWEETXLPOST.PostBox.Value = ETWEETXLPOST.PostBox.Value & vbNewLine & ETWEETXLPOST_EX.PostBox.Value
        Else
            ETWEETXLPOST.PostBox.Value = ETWEETXLPOST_EX.PostBox.Value
                End If
                    End If
'//WinForm #
xWin = 16: Call setWindow(xWin)

End Sub
Private Sub AddSizeHBtn_Click()

Call eTweetXL_CLICK.AddSizeHBtn_Clk

End Sub
Private Sub AddSizeVBtn_Click()

Call eTweetXL_CLICK.AddSizeVBtn_Clk

End Sub
Private Sub AppendBtn_Click()

If Range("AppendTrig").Value2 = 0 Then Range("AppendTrig").Value2 = 1: AppendBtn.ForeColor = vbGreen _
Else Range("AppendTrig").Value2 = 0: AppendBtn.ForeColor = &H80000011

End Sub
Private Sub ReflectBtn_Click()

Call eTweetXL_CLICK.ReflectBtn_Clk

End Sub
Private Sub RmvSizeHBtn_Click()

Call eTweetXL_CLICK.RmvSizeHBtn_Clk

End Sub
Private Sub RmvSizeVBtn_Click()

Call eTweetXL_CLICK.RmvSizeVBtn_Clk

End Sub
Private Sub AddThreadBtn_Click()

Call eTweetXL_CLICK.AddThreadBtn_Clk

End Sub
Private Sub RmvAllThreadBtn_Click()

Call eTweetXL_CLICK.RmvAllThreadBtn_Clk

End Sub
Private Sub RmvThreadBtn_Click()

Call eTweetXL_CLICK.RmvThreadBtn_Clk

End Sub
Private Sub SavePostBtn_Click()

Call eTweetXL_CLICK.SavePostBtn_Clk

End Sub
Private Sub SplitPostBtn_Click()

eTweetXL_CLICK.SplitPostBtn_Clk

End Sub
Private Sub TrimPostBtn_Click()

eTweetXL_CLICK.TrimPostBtn_Clk

End Sub
Private Sub LoadPostBtn_Click()

Range("ReflectTrig").Value = 0
Call eTweetXL_CLICK.ReflectBtn_Clk
Call eTweetXL_CLICK.LoadPostBtn_Clk(xName, xPath)

End Sub
Private Sub ExitBtn_Click()

Me.Hide
Range("xlasWinForm").Value2 = 13

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
Range("xlasWinForm").Value2 = 161
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
Range("xlasWinForm").Value2 = 161
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

'//Key Ctrl+Shift+Up
If KeyCode.Value = vbKeyUp Then
If Range("xlasKeyCtrl").Value2 = 33 Then
Call eTweetXL_CLICK.RmvSizeVBtn_Clk
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Shift+Down
If KeyCode.Value = vbKeyDown Then
If Range("xlasKeyCtrl").Value2 = 33 Then
Call eTweetXL_CLICK.AddSizeVBtn_Clk
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Shift+Left
If KeyCode.Value = vbKeyLeft Then
If Range("xlasKeyCtrl").Value2 = 33 Then
Call eTweetXL_CLICK.RmvSizeHBtn_Clk
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+Shift+Right
If KeyCode.Value = vbKeyRight Then
If Range("xlasKeyCtrl").Value2 = 33 Then
Call eTweetXL_CLICK.AddSizeHBtn_Clk
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
Range("xlasKeyCtrl").Value2 = vbNullString
        
End Sub
Private Sub AppendBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 1: Call eTweetXL_TOOLS.undPostEx(xBtn)

End Sub
Private Sub LoadPostBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 2: Call eTweetXL_TOOLS.undPostEx(xBtn)

End Sub

Private Sub ReflectBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 3: Call eTweetXL_TOOLS.undPostEx(xBtn)

End Sub
Private Sub AddSizeHBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 4: Call eTweetXL_TOOLS.undPostEx(xBtn)

End Sub
Private Sub RmvSizeHBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 5: Call eTweetXL_TOOLS.undPostEx(xBtn)

End Sub
Private Sub AddSizeVBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 6: Call eTweetXL_TOOLS.undPostEx(xBtn)

End Sub
Private Sub RmvSizeVBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 7: Call eTweetXL_TOOLS.undPostEx(xBtn)

End Sub
Private Sub SplitPostBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 8: Call eTweetXL_TOOLS.undPostEx(xBtn)

End Sub
Private Sub TrimPostBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 9: Call eTweetXL_TOOLS.undPostEx(xBtn)

End Sub
Private Sub AddThreadBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 10: Call eTweetXL_TOOLS.undPostEx(xBtn)

End Sub
Private Sub RmvThreadBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 11: Call eTweetXL_TOOLS.undPostEx(xBtn)

End Sub
Private Sub RmvAllThreadBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 12: Call eTweetXL_TOOLS.undPostEx(xBtn)

End Sub
Private Sub SavePostBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 13: Call eTweetXL_TOOLS.undPostEx(xBtn)

End Sub
Private Sub ExitBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 14: Call eTweetXL_TOOLS.undPostEx(xBtn)

End Sub
Private Sub UserForm_Terminate()

Range("xlasWinForm").Value2 = 13

End Sub
