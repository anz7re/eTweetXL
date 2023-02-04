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

Dim PinHldr As String
Dim xPos As Integer

If Len(PinlockBox.Value) > 5 Then
PinlockBox.Value = ""
InvisBox.Value = ""
Exit Sub
End If

If ShowPinBox.Value = False Then '//HIDE
    PinHldr = PinlockBox.Value
    If InStr(1, PinHldr, "*") Then
    PinBoxArr = Split(PinHldr, "*")
    xPos = UBound(PinBoxArr)
        If PinBoxArr(xPos) <> "" Then
        InvisBox.Value = InvisBox.Value & PinBoxArr(xPos)
        PinHldr = Replace(PinHldr, PinBoxArr(xPos), "*")
        PinlockBox.Value = PinHldr
            End If
        Else
            InvisBox.Value = PinHldr
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

Dim PinHldr As String

On Error GoTo ErrMsg

If ShowPinBox.Value = False Then

PinHldr = InvisBox.Value: X = PinHldr: Call basBinaryHash1(X, xVerify, xHash): PinHldr = """" & xHash & """"

If PinHldr = Range("Target").Value Then
Call UserForm_Terminate
Exit Sub
    Else
        GoTo ErrMsg
            End If
            
                Else
               
PinHldr = PinlockBox.Value: X = PinHldr: Call basBinaryHash1(X, xVerify, xHash): PinHldr = """" & xHash & """"

If PinHldr = Range("Target").Value Then
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
Private Sub CloseBtn_Click()

Me.Hide

End Sub
Private Sub UserForm_Terminate()
    
On Error GoTo ErrMsg
    
If ShowPinBox.Value = False Then

PinHldr = InvisBox.Value: X = PinHldr: Call basBinaryHash1(X, xVerify, xHash): PinHldr = """" & xHash & """"

    If PinHldr <> Range("Target").Value Then
    XLPINLOCK.Show
    GoTo ErrMsg
    Exit Sub
        End If
            pLatch = Range("Ucure").Value
            Range(pLatch).Value = 1
            XLPINLOCK.Hide
            Exit Sub
                    
                    Else
                    
PinHldr = PinlockBox.Value: X = PinHldr: Call basBinaryHash1(X, xVerify, xHash): PinHldr = """" & xHash & """"

    If PinHldr <> Range("Target").Value Then
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


