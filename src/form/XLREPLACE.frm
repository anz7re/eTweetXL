VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} XLREPLACE 
   Caption         =   "xlReplace"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5700
   OleObjectBlob   =   "XLREPLACE.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "XLREPLACE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()

Dim xWin As Object

'//Find running window
Call getWindow(xWin)
If Range("xlasWinForm").Value2 > 100 Then
        ElseIf Range("xlasWinForm").Value2 = 100 Then Set xWin = CTRLBOX.CtrlBoxWindow
            ElseIf Range("xlasInputField").Value2 <> 99 Then
                Set xWin = xWin.xlFlowStrip
                    End If
                    
XLREPLACE.FindWhatBox.SetFocus
If XLREPLACE.FindWhatBox.Value = vbNullString Then XLREPLACE.FindWhatBox.Value = xWin.SelText

Set xWin = Nothing

End Sub
Private Sub FindNextBtn_Click()

xBtn = 1
Call ButtonDown(xBtn)

xFind = FindWhatBox.Value
Call xlReplace_CLICK.FindNextBtn_Clk(xFind)
XLREPLACE.Hide: Art = "<lib>xbas;delayevent(10);$": Call xlas(Art): XLREPLACE.Show

End Sub
Private Sub ReplaceBtn_Click()

xBtn = 2
Call ButtonDown(xBtn)

xFind = XLREPLACE.FindWhatBox
Call xlReplace_CLICK.ReplaceBtn_Clk(xFind)
XLREPLACE.Hide: XLREPLACE.Show

End Sub
Private Sub ReplaceAllBtn_Click()

xBtn = 3
Call ButtonDown(xBtn)

xFind = XLREPLACE.FindWhatBox
Call xlReplace_CLICK.ReplaceAllBtn_Clk(xFind)
XLREPLACE.Hide: XLREPLACE.Show

End Sub
Private Sub CancelBtn_Click()

xBtn = 4
Call ButtonDown(xBtn)
XLREPLACE.Hide

End Sub
Private Sub FindWhatBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = vbKeyEscape Then XLREPLACE.Hide: Exit Sub

If KeyCode = vbKeyTab Then ReplaceWithBox.SelStart = 0: ReplaceWithBox.SetFocus: Exit Sub

End Sub
Private Sub ReplaceWithBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = vbKeyEscape Then XLREPLACE.Hide: Exit Sub

If KeyCode = vbKeyTab Then FindWhatBox.SelStart = 0: FindWhatBox.SetFocus: Exit Sub

End Sub
Private Sub MatchCaseBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = vbKeyEscape Then XLREPLACE.Hide

End Sub
Private Sub WrapAroundBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = vbKeyEscape Then XLREPLACE.Hide

End Sub
Private Sub UserForm_Terminate()

Range("xlasWinForm").Value2 = Range("xlasWinFormLast").Value2

End Sub
