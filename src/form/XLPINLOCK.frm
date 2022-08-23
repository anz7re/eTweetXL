VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} XLPINLOCK 
   Caption         =   "Pinlock"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5295
   OleObjectBlob   =   "XLPINLOCK.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "XLPINLOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub PinlockBox_Change()

Dim PBoxHldr, LastDg As Integer

If Len(PinlockBox.Value) > 5 Then
PinlockBox.Value = ""
InvisBox.Value = ""
Exit Sub
End If

If ShowPinBox.Value = False Then '//HIDE
    PBoxHldr = PinlockBox.Value
    If InStr(1, PBoxHldr, "*") Then
    PBoxArr = Split(PBoxHldr, "*")
    LastDg = UBound(PBoxArr)
        If PBoxArr(LastDg) <> "" Then
        InvisBox.Value = InvisBox.Value & PBoxArr(LastDg)
        PBoxHldr = Replace(PBoxHldr, PBoxArr(LastDg), "*")
        PinlockBox.Value = PBoxHldr
            End If
        Else
            InvisBox.Value = PBoxHldr
            PinlockBox.Value = "*"
                End If
                    End If

End Sub
Private Sub UserForm_Activate()

If XLPINLOCK.PinlockBox.Value <> "" Then XLPINLOCK.PinlockBox.Value = ""
If XLPINLOCK.InvisBox.Value <> "" Then XLPINLOCK.InvisBox.Value = ""

End Sub
Private Sub UserForm_Initialize()

Range("PassQuit").Value = 0
XLPINLOCK.PinlockBox.Value = ""
XLPINLOCK.InvisBox.Value = ""

End Sub
Private Sub EnterBtn_Click()

On Error GoTo ErrMsg

If ShowPinBox.Value = False Then

If Int(InvisBox.Value) = Int(Range("Target").Value) Then
Call UserForm_Terminate
Exit Sub
    Else
        GoTo ErrMsg
            End If
            
                Else
                
If Int(PinlockBox.Value) = Int(Range("Target").Value) Then
Call UserForm_Terminate
Exit Sub
    Else
        GoTo ErrMsg
            End If
                            
                End If

GoTo CleanMe

ErrMsg:
Range("PassQuit").Value = Range("PassQuit").Value + 1
MsgBox ("Invalid passcode"), vbExclamation, eTweetXL_INFO.AppName
If Range("PassQuit").Value >= 5 Then Range("PassQuit").Value = 0: ThisWorkbook.Close (SaveChanges = False)

CleanMe:
PinlockBox.Value = ""
InvisBox.Value = ""

End Sub
Private Sub UserForm_Terminate()
    
On Error GoTo ErrMsg
    
If ShowPinBox.Value = False Then

    If Int(InvisBox.Value) <> Int(Range("Target").Value) Then
    XLPINLOCK.Show
    GoTo ErrMsg
    Exit Sub
        End If
            pLatch = Range("Ucure").Value
            Range(pLatch).Value = 1
            XLPINLOCK.Hide
            Exit Sub
                    
                    Else
                
    If Int(PinlockBox.Value) <> Int(Range("Target").Value) Then
    XLPINLOCK.Show
    GoTo ErrMsg
    Exit Sub
        End If
            pLatch = Range("Ucure").Value
            Range(pLatch).Value = 1
            XLPINLOCK.Hide
            Exit Sub
        
        End If
ErrMsg:
MsgBox ("Invalid action"), vbExclamation, eTweetXL_INFO.AppName
End Sub


