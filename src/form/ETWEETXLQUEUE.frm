VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ETWEETXLQUEUE 
   Caption         =   "eTweetXL"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15975
   OleObjectBlob   =   "ETWEETXLQUEUE.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ETWEETXLQUEUE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

'//Refresh media Scroll...
Range("MedScrollPos").Value = 0

'//Welcome...
If Me.xlFlowStrip.Value = "" Then Me.xlFlowStrip.Value = eTweetXL_INFO.AppWelcome

Me.Caption = AppTag

End Sub
Private Sub UserForm_Activate()

'//Cleanup
ETWEETXLQUEUE.QueueBox.Clear
ETWEETXLQUEUE.RuntimeBox.Clear
ETWEETXLQUEUE.UserBox.Clear

'//Show runtime action message
Call eTweetXL_GET.getRtState

'//WinForm #
ThisWorkbook.ActiveSheet.Range("xlasWinForm").Value2 = 4

'//Update application state
Call eTweetXL_TOOLS.updAppState

Call eTweetXL_GET.getQueueData

'//Window title
If Me.xlFlowStrip.Value = vbNullString Or Range("AppState").Value2 <> 1 Then Me.xlFlowStrip.Value = "Queue..."

End Sub
Private Sub HomeBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xBtn = 0
Call eTweetXL_TOOLS.fxsNaviBar(xBtn)
Call eTweetXL_TOOLS.undNaviBar(xBtn)

xMsg = 52: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover

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
Private Sub StartBtn_Click()

Call eTweetXL_CLICK.StartBtn_Clk

End Sub
Private Sub BreakBtn_Click()

Call eTweetXL_CLICK.BreakBtn_Clk

End Sub

Private Sub CtrlBoxBtn_Click()

Call eTweetXL_FOCUS.shw_CTRLBOX

End Sub
Private Sub FreezeBtn_Click()

Call eTweetXL_CLICK.FreezeBtn_Clk

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

Private Sub PostHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'//Clear post box
Call eTweetXL_CLICK.PostHdr_Clk

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
    MedDemo.Picture = Nothing
    MedDemo.Picture = LoadPicture(MediaHldr)
    MedDemo.PictureSizeMode = fmPictureSizeModeStretch
        End If

Range("MedScrollLink").Value = MediaHldr
  
End Sub
Private Sub RuntimeBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call eTweetXL_CLICK.RuntimeBox_Clk

End Sub
Private Sub HomeBtn_Click()

Me.Hide
Call eTweetXL_FOCUS.shw_ETWEETXLHOME

End Sub
Private Sub AppTag_Click()

ETWEETXLQUEUE.Hide
ETWEETXLHOME.Show

End Sub
Private Sub PostBox_Change()

Call eTweetXL_CHANGE.PostBox_Chg

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
Range("xlasWinForm").Value2 = 41
XLFONTBOX.Show
Range("xlasWinForm").Value2 = 3
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If

'//Key Ctrl+H
If KeyCode.Value = vbKeyH Then
If Range("xlasKeyCtrl").Value2 = 17 Then
Range("xlasWinForm").Value2 = 41
Range("xlasWinFormLast").Value2 = 4
XLREPLACE.Show
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
        
'//Key Ctrl+Alt+R
If KeyCode.Value = vbKeyR Then
If Range("xlasKeyCtrl").Value2 = 35 Then
Call eTweetXL_CLICK.RmvAllThreadBtn_Clk
Range("xlasKeyCtrl").Value2 = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
        Range("xlasKeyCtrl").Value2 = ""
        
End Sub
Private Sub xlFlowStripBar_Click()

Call eTweetXL_CLICK.xlFlowStripBar_Clk

End Sub
Private Sub xlFlowStrip_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

xKey = KeyCode.Value
Call eTweetXL_TOOLS.runFlowStrip(xKey)
        
End Sub
Private Sub QueueBox_Click()

If Range("DataPullTrig").Value <> 1 Then iNum = QueueBox.ListIndex: _
ETWEETXLPOST.ProfileListBox.Value = Range("ProfileLink").Offset(iNum + 1, 0).Value

xTwt = QueueBox.Value
'//remove numbered count
If xTwt <> vbNullString Then xTwtArr = Split(xTwt, ") "): xTwt = xTwtArr(1)
ETWEETXLPOST.DraftBox.Value = xTwt

End Sub
Private Sub QueueBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

If QueueBox.Value <> vbNullString Then MsgBox (QueueBox.Value), vbQuestion, AppTag

End Sub
Private Sub AddThreadBtn_Click()

If ThreadCt.Caption = vbNullString Or 0 Then ThreadCt.Caption = 1

If PostBox.Value <> vbNullString Then
Call eTweetXL_CLICK.AddThreadBtn_Clk
End If

End Sub

Private Sub RmvThreadBtn_Click()

lastR = Cells(Rows.Count, "Z").End(xlUp).Row: If lastR <= 1 Then Exit Sub

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

eTweetXL_CLICK.RmvPostMedBtn_Clk

End Sub
Private Sub MedDemoScroll_SpinDown()

On Error Resume Next

lastR = Cells(Rows.Count, "I").End(xlUp).Row

Range("MedScrollPos").Value = Range("MedScrollPos").Value - 1

'//LEFT
If Range("MedScrollPos").Value < 0 Then Range("MedScrollPos").Value = lastR - 1

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

lastR = Cells(Rows.Count, "I").End(xlUp).Row

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
Private Sub PostThreadScroll_SpinDown()

On Error Resume Next

lastR = Cells(Rows.Count, "Y").End(xlUp).Row
If lastR < 1 Then lastR = Cells(Rows.Count, "Z").End(xlUp).Row

Call clnMediaScroll

Range("ThreadScrollPos").Value = Range("ThreadScrollPos").Value - 1

'//LEFT
If Range("ThreadScrollPos").Value <= 0 Then Range("ThreadScrollPos").Value = lastR - 1

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

lastR = Cells(Rows.Count, "Y").End(xlUp).Row
If lastR < 1 Then lastR = Cells(Rows.Count, "Z").End(xlUp).Row

Call clnMediaScroll

Range("ThreadScrollPos").Value = Range("ThreadScrollPos").Value + 1

'//RIGHT
If Range("ThreadScrollPos").Value >= lastR Then Range("ThreadScrollPos").Value = 1

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
Private Sub SavePostBtn_Click()

eTweetXL_CLICK.SavePostBtn_Clk

End Sub
Private Sub ViewMedBtn_Click()

eTweetXL_CLICK.ViewMedBtn_Clk

End Sub
'//hover effects
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
Private Sub PostHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 4: Call HoverHelp(xMsg)

Call eTweetXL_TOOLS.dfsHover
xHov = 4: Call eTweetXL_TOOLS.fxsHover(xHov)

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
Private Sub QueueBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Range("HoverPos").Value = 44 '//temporary for loading drafts

Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub RuntimeBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
xMsg = 45: Call HoverHelp(xMsg)
Call eTweetXL_TOOLS.dfsHover
xHov = 13: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
Private Sub QueueBg1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub QueueBg2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub QueueBg3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub QueueBg4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub QueueBg5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover

End Sub
Private Sub HideBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim xWin As Object: X = 0: Y = 0: Call basPostWinFormPos(xWin, X, Y): Set xWin = Nothing
Call eTweetXL_TOOLS.dfsHover
xHov = 17: Call eTweetXL_TOOLS.fxsHover(xHov)

End Sub
