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
If Me.xlFlowStrip.Value = "" Then Me.xlFlowStrip.Value = App_INFO.AppWelcome

Me.Caption = AppTag

End Sub
Private Sub UserForm_Activate()

'//WinForm #
Range("xlasWinForm").Value = 4

'//Cleanup
ETWEETXLQUEUE.QueueBox.Clear
ETWEETXLQUEUE.RuntimeBox.Clear
ETWEETXLQUEUE.UserBox.Clear

'//Show runtime action message
Call App_TOOLS.ShowRtAction

'//Update active state
Call App_TOOLS.UpdateActive
Call App_IMPORT.MyNextQueue

'//Window title
If Me.xlFlowStrip.Value = vbNullString Or Range("AppActive").Value <> 1 Then Me.xlFlowStrip.Value = "Queue..."

End Sub
Private Sub HomeBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xBtn = 0
Call App_TOOLS.NaviBarColor(xBtn)
Call App_TOOLS.NaviBarUnderline(xBtn)

xMsg = 52: Call HoverHelp(xMsg)

Call App_TOOLS.HoverDefault

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
Private Sub StartBtn_Click()

Call App_CLICK.Start_Clk

End Sub
Private Sub BreakBtn_Click()

Call App_CLICK.Break_Clk

End Sub

Private Sub CtrlBoxBtn_Click()

Call App_Focus.SH_CTRLBOX

End Sub
Private Sub FreezeBtn_Click()

Call App_CLICK.FreezeBtn_Clk

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

Private Sub PostHdr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'//Clear post box
Call App_CLICK.PostHdr_Clk

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
Private Sub RuntimeBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call App_CLICK.RuntimeBox_Clk

End Sub
Private Sub HomeBtn_Click()

Call App_Focus.HdForms
Call App_Focus.SH_ETWEETXLHOME

End Sub
Private Sub LogoBg_Click()

ETWEETXLQUEUE.Hide
ETWEETXLHOME.Show

End Sub
Private Sub PostBox_Change()

Call App_CHANGE.PostBox_Chg

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
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
'//Key Ctrl+H
If KeyCode.Value = vbKeyH Then
If Range("xlasKeyCtrl").Value = 17 Then
Range("xlasWinForm").Value = 41
Range("xlasWinFormLast").Value = 4
XLREPLACE.Show
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
        
'//Key Ctrl+Alt+R
If KeyCode.Value = vbKeyR Then
If Range("xlasKeyCtrl").Value = 35 Then
Call App_CLICK.RmvAllThread_Clk
Range("xlasKeyCtrl").Value = vbNullString
KeyCode.Value = 0
Exit Sub
End If
End If
        
        Range("xlasKeyCtrl").Value = ""
        
End Sub
Private Sub xlFlowStripBar_Click()

Call App_CLICK.xlFlowStripBar_Clk

End Sub
Private Sub xlFlowStrip_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

xKey = KeyCode.Value
Call App_TOOLS.RunFlowStrip(xKey)
        
End Sub
Private Sub QueueBox_Click()

If Range("DataPullTrig").Value <> 1 Then iNum = QueueBox.ListIndex: _
ETWEETXLPOST.ProfileListBox.Value = Range("Profilelink").Offset(iNum + 1, 0).Value

xTwt = QueueBox.Value
'//remove numbered count
If xTwt <> vbNullString Then xTwtArr = Split(xTwt, ") "): xTwt = xTwtArr(1)
ETWEETXLPOST.DraftBox.Value = xTwt

End Sub
Private Sub AddThreadBtn_Click()

If ThreadCt.Caption = vbNullString Or 0 Then ThreadCt.Caption = 1

If PostBox.Value <> vbNullString Then
Call App_CLICK.AddThread_Clk
End If

End Sub

Private Sub RmvThreadBtn_Click()

lastRw = Cells(Rows.Count, "Z").End(xlUp).Row: If lastRw <= 1 Then Exit Sub

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

App_CLICK.RmvPostMed_Clk

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
Private Sub SavePostBtn_Click()

App_CLICK.SavePost_Clk

End Sub
Private Sub ViewMedBtn_Click()

App_CLICK.ViewMedBtn_Clk

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
Private Sub PostHdr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 4: Call HoverHelp(xMsg)

xHov = 4
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
Private Sub QueueBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next

Range("HoverPos").Value = 44 '//temporary for loading drafts

Call App_TOOLS.HoverDefault

End Sub
Private Sub RuntimeBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

On Error Resume Next


xMsg = 45: Call HoverHelp(xMsg)
xHov = 13
Call App_TOOLS.HoverDefault
Call App_TOOLS.HoverEffect(xHov)

End Sub
Private Sub QueueBg1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub
Private Sub QueueBg2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub
Private Sub QueueBg3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub
Private Sub QueueBg4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub
Private Sub QueueBg5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call App_TOOLS.HoverDefault

End Sub
