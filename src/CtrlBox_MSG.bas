Attribute VB_Name = "CtrlBox_MSG"
'/############################\
'//Application Error Messages\\
'///########################\\\

Function AppMsg(xMsg, errLvl) As Integer

If Range("xlasSilent").Value2 <> 1 Then

'/1/xlFlowStrip syntax error
If xMsg = 1 Then

Dim oFlowStrip As Object

'//Check xlFlowStrip Window...
If Range("xlasWinForm").Value2 = 1 Then Set oFlowStrip = ETWEETXLHOME.xlFlowStrip
If Range("xlasWinForm").Value2 = 2 Then Set oFlowStrip = ETWEETXLSETUP.xlFlowStrip
If Range("xlasWinForm").Value2 = 3 Then Set oFlowStrip = ETWEETXLPOST.xlFlowStrip
If Range("xlasWinForm").Value2 = 4 Then Set oFlowStrip = ETWEETXLQUEUE.xlFlowStrip
If Range("xlasWinForm").Value2 = 10 Then Set oFlowStrip = CTRLBOX.CtrlBoxWindow

oFlowStrip.ForeColor = vbRed
'//msg
MsgBox ("Syntax error"), vbExclamation, CtrlTag
Exit Function
End If

'/2/Components missing
If xMsg = 2 Then
MsgBox ("Components missing"), vbCritical, CtrlTag
End If

    End If

End Function

