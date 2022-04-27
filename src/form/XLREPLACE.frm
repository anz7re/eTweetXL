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
Call findWindow(xWin)
If Range("xlasWinForm").Value > 10 Then
        ElseIf Range("xlasWinForm").Value = 10 Then Set xWin = CTRLBOX.CtrlBoxWindow
            Else
                Set xWin = xWin.xlFlowStrip
                    End If
                    
XLREPLACE.FindWhatBox.SetFocus
If XLREPLACE.FindWhatBox.Value = vbNullString Then XLREPLACE.FindWhatBox.Value = xWin.SelText

Set xWin = Nothing

End Sub
Private Sub FindNextBtn_Click()

xBtn = 1
Call ButtonPress(xBtn)

xFind = FindWhatBox.Value
Call xlReplace_CLICK.FindNext_Clk(xFind)
XLREPLACE.Hide

End Sub
Private Sub ReplaceBtn_Click()

xBtn = 2
Call ButtonPress(xBtn)

xFind = XLREPLACE.FindWhatBox
Call xlReplace_CLICK.Replace_Clk(xFind)
XLREPLACE.Hide: XLREPLACE.Show

End Sub
Private Sub ReplaceAllBtn_Click()

xBtn = 3
Call ButtonPress(xBtn)

xFind = XLREPLACE.FindWhatBox
Call xlReplace_CLICK.ReplaceAll_Clk(xFind)
XLREPLACE.Hide: XLREPLACE.Show

End Sub
Private Sub CancelBtn_Click()

xBtn = 4
Call ButtonPress(xBtn)
XLREPLACE.Hide

End Sub
Public Function ButtonPress(xBtn)

Dim BtnObj, RepFormObj, RepBtnObj As Object

If xBtn = 1 Then Set RepBtnObj = XLREPLACE.FindNextBtn
If xBtn = 2 Then Set RepBtnObj = XLREPLACE.ReplaceBtn
If xBtn = 3 Then Set RepBtnObj = XLREPLACE.ReplaceAllBtn
If xBtn = 4 Then Set RepBtnObj = XLREPLACE.CancelBtn

RepBtnObj.BackColor = RGB(237, 247, 252)
RepBtnObj.BorderColor = vbBlue

Set RepFormObj = XLREPLACE

On Error Resume Next

For Each BtnObj In RepFormObj.Controls

If BtnObj <> RepBtnObj Then
    If InStr(1, BtnObj.Value, "Btn") Then
    BtnObj.BackColor = &H80000016
    BtnObj.BorderColor = &H8000000C
        End If
            End If
                Next
                
                Set BtnObj = Nothing
                Set RepFormObj = Nothing
                Set RepBtnObj = Nothing

End Function
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

Range("xlasWinForm").Value = Range("xlasWinFormLast").Value

End Sub
