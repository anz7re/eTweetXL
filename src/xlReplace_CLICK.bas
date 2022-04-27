Attribute VB_Name = "xlReplace_CLICK"
'//FIND NEXT FUNCTIONALITY
Public Function FindNext_Clk(xFind)

Dim xWin As Object
Dim X, StrLen As Integer
Dim xType As String
X = 0

'//Find running window
Call findWindow(xWin)
If Range("xlasWinForm").Value > 10 Then
        ElseIf Range("xlasWinForm").Value = 10 Then Set xWin = CTRLBOX.CtrlBoxWindow
            Else
                Set xWin = xWin.xlFlowStrip
                    End If
                    
If xWin.SelText = vbNullString Or xWin.SelText = " " Then xWin.SelStart = X Else X = xWin.SelStart + 1

If XLREPLACE.MatchCaseBox.Value = True Then GoTo MatchCase

xType = "TC"
If InStr(1, xWin.Text, xFind, vbTextCompare) = False Then Exit Function

GoTo CheckLen

MatchCase:
xType = "BC"
If InStr(1, xWin.Text, xFind, vbBinaryCompare) = False Then Exit Function

CheckLen:

StrLen = Len(xWin.Text) + 1

ReCheck:
xWin.SetFocus
xWin.SelStart = X
xWin.SelLength = Len(xFind)

X = X + 1
If xType = "TC" Then If InStr(1, xWin.SelText, xFind, vbTextCompare) = False And X < StrLen Then GoTo ReCheck
If xType = "BC" Then If InStr(1, xWin.SelText, xFind, vbBinaryCompare) = False And X < StrLen Then GoTo ReCheck

Range("xlasWinForm").Value = Range("xlasWinFormLast").Value
Set xWin = Nothing

End Function
'//REPLACE FUNCTIONALITY
Public Function Replace_Clk(xFind)

Dim xWin As Object
Dim X, StrLen As Integer
Dim xType As String
X = 0

'//Find running window
Call findWindow(xWin)
If Range("xlasWinForm").Value > 10 Then
        ElseIf Range("xlasWinForm").Value = 10 Then Set xWin = CTRLBOX.CtrlBoxWindow
            Else
                Set xWin = xWin.xlFlowStrip
                    End If

If XLREPLACE.MatchCaseBox.Value = True Then GoTo MatchCase

xType = "TC"
If InStr(1, xWin.Text, xFind, vbTextCompare) = False Then Exit Function

GoTo CheckLen

MatchCase:
xType = "BC"
If InStr(1, xWin.Text, xFind, vbBinaryCompare) = False Then Exit Function

CheckLen:

StrLen = Len(xWin.Text) + 1

ReCheck:
xWin.SetFocus
xWin.SelStart = X
xWin.SelLength = Len(xFind)

X = X + 1
If xType = "TC" Then If InStr(1, xWin.SelText, xFind, vbTextCompare) = False And X < StrLen Then GoTo ReCheck
If xType = "BC" Then If InStr(1, xWin.SelText, xFind, vbBinaryCompare) = False And X < StrLen Then GoTo ReCheck

If xType = "TC" Then If InStr(1, xWin.SelText, xFind, vbTextCompare) Then xWin.SelText = XLREPLACE.ReplaceWithBox.Value
If xType = "BC" Then If InStr(1, xWin.SelText, xFind, vbBinaryCompare) Then xWin.SelText = XLREPLACE.ReplaceWithBox.Value

Range("xlasWinForm").Value = Range("xlasWinFormLast").Value
Set xWin = Nothing

End Function
'//REPLACE ALL FUNCTIONALITY
Public Function ReplaceAll_Clk(xFind)

Dim xWin As Object
Dim X, StrLen As Integer
Dim xType As String
X = 0

'//Find running window
Call findWindow(xWin)
If Range("xlasWinForm").Value > 10 Then
        ElseIf Range("xlasWinForm").Value = 10 Then Set xWin = CTRLBOX.CtrlBoxWindow
            Else
                Set xWin = xWin.xlFlowStrip
                    End If


If XLREPLACE.MatchCaseBox.Value = True Then GoTo MatchCase

xType = "TC"
If InStr(1, xWin.Text, xFind, vbTextCompare) = False Then Exit Function

MatchCase:
xType = "BC"
If InStr(1, xWin.Text, xFind, vbBinaryCompare) = False Then Exit Function

StrLen = Len(xWin.Text) + 1

Do Until X = StrLen
ReCheck:
xWin.SetFocus
xWin.SelStart = X
xWin.SelLength = Len(xFind)

X = X + 1
If xType = "TC" Then If InStr(1, xWin.SelText, xFind, vbTextCompare) = False And X < StrLen Then GoTo ReCheck
If xType = "BC" Then If InStr(1, xWin.SelText, xFind, vbBinaryCompare) = False And X < StrLen Then GoTo ReCheck

If xType = "TC" Then If InStr(1, xWin.SelText, xFind, vbTextCompare) Then xWin.SelText = XLREPLACE.ReplaceWithBox.Value
If xType = "BC" Then If InStr(1, xWin.SelText, xFind, vbBinaryCompare) Then xWin.SelText = XLREPLACE.ReplaceWithBox.Value
Loop

Range("xlasWinForm").Value = Range("xlasWinFormLast").Value
Set xWin = Nothing

End Function
