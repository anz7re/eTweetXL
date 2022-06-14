Attribute VB_Name = "xlColorSwatch_TOOLS"
Public Function setColors(xRGB)

On Error Resume Next

Dim oControl As Object
Dim X As Integer

Call fndWindow(xWin)

'//Find RGB
xRGBArr = Split(xRGB, ",")

For X = 0 To UBound(xRGBArr)
If xRGBArr(X) = vbNullString Then xRGBArr(X) = 0
Next

'//Set gradient swatch
X = 1
For Each oControl In XLCOLORSWATCH.Controls
If InStr(1, oControl.name, "Sw" & X) Then
oControl.BackColor = RGB((-1 * (xRGBArr(0) - ((2 * X) + 3)) - xR), (xRGBArr(1) - 10), (xRGBArr(2) - 20))
oControl.Caption = xRGBArr(0) - ((2 * X) + 3) & "," & (xRGBArr(1) - 10) & "," & (xRGBArr(2) - 20)
oControl.ForeColor = oControl.BackColor
X = X + 1
End If
Next

XLCOLORSWATCH.SwBaseLrg.BackColor = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
XLCOLORSWATCH.SwBaseSm.BackColor = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))

If InStr(1, XLCOLORSWATCH.CurrType.Caption, "B") Then Range("xlasBlkAddr96").Value = xRGBArr(0) & "," & xRGBArr(1) & "," & xRGBArr(2)
If InStr(1, XLCOLORSWATCH.CurrType.Caption, "F") Then Range("xlasBlkAddr97").Value = xRGBArr(0) & "," & xRGBArr(1) & "," & xRGBArr(2)

'//set window background color
If Range("xlasBlkAddr96").Value <> vbNullString Then
xRGBArr = Split(Range("xlasBlkAddr96").Value, ",")
R = xRGBArr(0): If R = vbNullString Then R = 0
G = xRGBArr(1): If G = vbNullString Then G = 0
B = xRGBArr(2): If B = vbNullString Then B = 0
XLFONTBOX.CurrBColor.BackColor = RGB(R, G, B)
xWin.BackColor = RGB(R, G, B)
End If

'//set window font color
If Range("xlasBlkAddr97").Value <> vbNullString Then
xRGBArr = Split(Range("xlasBlkAddr97").Value, ",")
R = xRGBArr(0): If R = vbNullString Then R = 0
G = xRGBArr(1): If G = vbNullString Then G = 0
B = xRGBArr(2): If B = vbNullString Then B = 0
XLFONTBOX.CurrFColor.BackColor = RGB(R, G, B)
xWin.ForeColor = RGB(R, G, B)
End If

Set xWin = Nothing

End Function






