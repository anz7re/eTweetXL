Attribute VB_Name = "xlColorSwatch_CLICK"
Public Sub Sw_Clk(xSw)

On Error Resume Next

Dim oControl As Object: Dim oAppTxt As Object: Dim oAppBg1 As Object

'//application object variables
Set oAppTxt = CTRLBOX.CtrlBoxWindow
Set oAppBg1 = CTRLBOX.CtrlBoxWindow

For Each oControl In XLCOLORSWATCH.Controls
If InStr(1, oControl.name, "Sw" & xSw) Then
xSwArr = Split(oControl.Caption, ",")
XLCOLORSWATCH.SwCtrl.Caption = 1
XLCOLORSWATCH.RColBox = xSwArr(0)
XLCOLORSWATCH.GColBox = xSwArr(1)
XLCOLORSWATCH.SwCtrl.Caption = 0
XLCOLORSWATCH.BColBox = xSwArr(2)
'//set window background color
If InStr(1, XLCOLORSWATCH.CurrType, "B") Then: _
oAppTxt.BackColor = oControl.BackColor: _
oAppBg1.BackColor = oControl.BackColor

'//set window font color
If InStr(1, XLCOLORSWATCH.CurrType, "F") Then oAppText.ForeColor = oControl.BackColor
End If
Next

Set oAppTxt = Nothing
Set oAppBg1 = Nothing

End Sub



