Attribute VB_Name = "xlReplace_CLICK"
Public Function ButtonDown(ByVal xBtn As Byte)

Dim BtnObj, ReplFormObj, ReplBtnObj As Object

If xBtn = 1 Then Set ReplBtnObj = XLREPLACE.FindNextBtn
If xBtn = 2 Then Set ReplBtnObj = XLREPLACE.ReplaceBtn
If xBtn = 3 Then Set ReplBtnObj = XLREPLACE.ReplaceAllBtn
If xBtn = 4 Then Set ReplBtnObj = XLREPLACE.CancelBtn

ReplBtnObj.BackColor = RGB(185, 231, 170)
ReplBtnObj.BorderColor = vbGreen
ReplBtnObj.ForeColor = vbBlack

Set ReplFormObj = XLREPLACE

On Error Resume Next

For Each BtnObj In ReplFormObj.Controls

If BtnObj <> ReplBtnObj Then
    If InStr(1, BtnObj.Value, "Btn") Then
    BtnObj.BackColor = &H80000006
    BtnObj.BorderColor = &H8000000C
    BtnObj.ForeColor = &H8000000E
        End If
            End If
                Next
                
                Set BtnObj = Nothing
                Set ReplFormObj = Nothing
                Set ReplBtnObj = Nothing

End Function
'//FIND NEXT FUNCTIONALITY
Public Function FindNextBtn_Clk(xFind)

Dim xWin As Object
Dim x, StrLen As Integer
Dim xType As String
x = 0

'//Find running window
Call getWindow(xWin)
If Range("xlasWinForm").Value2 > 100 Then
        ElseIf Range("xlasWinForm").Value2 = 100 Then Set xWin = CTRLBOX.CtrlBoxWindow
            Else
                Set xWin = xWin.xlFlowStrip
                    End If
                    
'//Check for wrap around
If XLREPLACE.WrapAroundBox.Value = False Then If xWin.SelStart = CInt(Len(xWin.Value)) - CInt(Len(xWin.SelText)) Then _
MsgBox ("""" & xFind & """" & " not found"), vbInformation, "xlReplace": Exit Function

If xWin.SelText = vbNullString Or xWin.SelText = " " Then xWin.SelStart = x Else x = xWin.SelStart + 1

If XLREPLACE.MatchCaseBox.Value = True Then GoTo MatchCase

xType = "TC"
If InStr(1, xWin.Text, xFind, vbTextCompare) = False Then MsgBox ("""" & xFind & """" & " not found"), vbInformation, "xlReplace": Exit Function

GoTo CheckLen

MatchCase:
xType = "BC"
If InStr(1, xWin.Text, xFind, vbBinaryCompare) = False Then MsgBox ("""" & xFind & """" & " not found"), vbInformation, "xlReplace": Exit Function

CheckLen:

StrLen = Len(xWin.Text) + 1

ReCheck:
xWin.SetFocus
xWin.SelStart = x
xWin.SelLength = Len(xFind)

x = x + 1
If xType = "TC" Then If InStr(1, xWin.SelText, xFind, vbTextCompare) = False And x < StrLen Then GoTo ReCheck
If xType = "BC" Then If InStr(1, xWin.SelText, xFind, vbBinaryCompare) = False And x < StrLen Then GoTo ReCheck

Range("xlasWinForm").Value2 = Range("xlasWinFormLast").Value2
Set xWin = Nothing

End Function
'//REPLACE FUNCTIONALITY
Public Function ReplaceBtn_Clk(xFind)

Dim xWin As Object
Dim x, StrLen As Integer
Dim xType As String
x = 0

'//Find running window
Call getWindow(xWin)
If Range("xlasWinForm").Value2 > 100 Then
        ElseIf Range("xlasWinForm").Value2 = 100 Then Set xWin = CTRLBOX.CtrlBoxWindow
            Else
                Set xWin = xWin.xlFlowStrip
                    End If

'//Check for wrap around
If XLREPLACE.WrapAroundBox.Value = False Then If xWin.SelStart = CInt(Len(xWin.Value)) - CInt(Len(xWin.SelText)) Then _
MsgBox ("""" & xFind & """" & " not found"), vbInformation, "xlReplace": Exit Function

If XLREPLACE.MatchCaseBox.Value = True Then GoTo MatchCase

xType = "TC"
If InStr(1, xWin.Text, xFind, vbTextCompare) = False Then MsgBox ("""" & xFind & """" & " not found"), vbInformation, "xlReplace": Exit Function

GoTo CheckLen

MatchCase:
xType = "BC"
If InStr(1, xWin.Text, xFind, vbBinaryCompare) = False Then MsgBox ("""" & xFind & """" & " not found"), vbInformation, "xlReplace": Exit Function

CheckLen:

StrLen = Len(xWin.Text) + 1

ReCheck:
xWin.SetFocus
xWin.SelStart = x
xWin.SelLength = Len(xFind)

x = x + 1
If xType = "TC" Then If InStr(1, xWin.SelText, xFind, vbTextCompare) = False And x < StrLen Then GoTo ReCheck
If xType = "BC" Then If InStr(1, xWin.SelText, xFind, vbBinaryCompare) = False And x < StrLen Then GoTo ReCheck

If xType = "TC" Then If InStr(1, xWin.SelText, xFind, vbTextCompare) Then xWin.SelText = XLREPLACE.ReplaceWithBox.Value
If xType = "BC" Then If InStr(1, xWin.SelText, xFind, vbBinaryCompare) Then xWin.SelText = XLREPLACE.ReplaceWithBox.Value

Range("xlasWinForm").Value2 = Range("xlasWinFormLast").Value2
Set xWin = Nothing

End Function
'//REPLACE ALL FUNCTIONALITY
Public Function ReplaceAllBtn_Clk(xFind)

Dim xWin As Object
Dim x, StrLen As Integer
Dim xType As String
x = 0

'//Find running window
Call getWindow(xWin)
If Range("xlasWinForm").Value2 > 100 Then
        ElseIf Range("xlasWinForm").Value2 = 100 Then Set xWin = CTRLBOX.CtrlBoxWindow
            Else
                Set xWin = xWin.xlFlowStrip
                    End If


If XLREPLACE.MatchCaseBox.Value = True Then GoTo MatchCase

xType = "TC"
If InStr(1, xWin.Text, xFind, vbTextCompare) = False Then MsgBox ("""" & xFind & """" & " not found"), vbInformation, "xlReplace": Exit Function

MatchCase:
xType = "BC"
If InStr(1, xWin.Text, xFind, vbBinaryCompare) = False Then MsgBox ("""" & xFind & """" & " not found"), vbInformation, "xlReplace": Exit Function

StrLen = Len(xWin.Text) + 1

Do Until x = StrLen
ReCheck:
xWin.SetFocus
xWin.SelStart = x
xWin.SelLength = Len(xFind)

x = x + 1
If xType = "TC" Then If InStr(1, xWin.SelText, xFind, vbTextCompare) = False And x < StrLen Then GoTo ReCheck
If xType = "BC" Then If InStr(1, xWin.SelText, xFind, vbBinaryCompare) = False And x < StrLen Then GoTo ReCheck

If xType = "TC" Then If InStr(1, xWin.SelText, xFind, vbTextCompare) Then xWin.SelText = XLREPLACE.ReplaceWithBox.Value
If xType = "BC" Then If InStr(1, xWin.SelText, xFind, vbBinaryCompare) Then xWin.SelText = XLREPLACE.ReplaceWithBox.Value
Loop

Range("xlasWinForm").Value2 = Range("xlasWinFormLast").Value2
Set xWin = Nothing

End Function
