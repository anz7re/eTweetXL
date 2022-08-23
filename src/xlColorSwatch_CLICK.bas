Attribute VB_Name = "xlColorSwatch_CLICK"
Public Sub Sw_Clk(xSw)

On Error Resume Next

Dim oControl As Object

Call getWindow(xWin)

For Each oControl In XLCOLORSWATCH.Controls
If InStr(1, oControl.name, "Sw" & xSw) Then
xSwArr = Split(oControl.Caption, ",")
XLCOLORSWATCH.SwCtrl.Caption = 1
XLCOLORSWATCH.RColBox = xSwArr(0)
XLCOLORSWATCH.GColBox = xSwArr(1)
XLCOLORSWATCH.SwCtrl.Caption = 0
XLCOLORSWATCH.BColBox = xSwArr(2)

'//set window background color
If InStr(1, XLCOLORSWATCH.CurrType, "B") Then xWin.BackColor = oControl.BackColor

'//set window font color
If InStr(1, XLCOLORSWATCH.CurrType, "F") Then xWin.ForeColor = oControl.BackColor
End If
Next

Set xWin = Nothing

End Sub



