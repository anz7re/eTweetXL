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
If Range("xlasWinForm").Value2 = 11 Then Set oFlowStrip = ETWEETXLHOME.xlFlowStrip
If Range("xlasWinForm").Value2 = 12 Then Set oFlowStrip = ETWEETXLSETUP.xlFlowStrip
If Range("xlasWinForm").Value2 = 13 Then Set oFlowStrip = ETWEETXLPOST.xlFlowStrip
If Range("xlasWinForm").Value2 = 14 Then Set oFlowStrip = ETWEETXLQUEUE.xlFlowStrip
If Range("xlasWinForm").Value2 = 100 Then Set oFlowStrip = CTRLBOX.CtrlBoxWindow

oFlowStrip.ForeColor = vbRed
'//msg
MsgBox ("Syntax error"), vbExclamation, CtrlBox_INFO.AppTag
Exit Function
End If

'/2/Components missing
If xMsg = 2 Then
MsgBox ("Components missing"), vbCritical, CtrlBox_INFO.AppTag
End If

    End If

End Function

