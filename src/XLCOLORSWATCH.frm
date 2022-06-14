VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} XLCOLORSWATCH 
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5070
   OleObjectBlob   =   "XLCOLORSWATCH.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "XLCOLORSWATCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub UserForm_Activate()

On Error Resume Next

If CurrType.Caption = "B" Then xRGBArr = Split(Range("xlasBlkAddr96").Value, ","): XLCOLORSWATCH.RColBox.Value = xRGBArr(0): XLCOLORSWATCH.GColBox.Value = xRGBArr(1): XLCOLORSWATCH.BColBox.Value = xRGBArr(2): CurrType.Caption = "Background Color"
If CurrType.Caption = "F" Then xRGBArr = Split(Range("xlasBlkAddr97").Value, ","): XLCOLORSWATCH.RColBox.Value = xRGBArr(0): XLCOLORSWATCH.GColBox.Value = xRGBArr(1): XLCOLORSWATCH.BColBox.Value = xRGBArr(2): CurrType.Caption = "Font Color"

SwBaseLrg.BackColor = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
SwBaseSm.BackColor = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))

SwCtrl.Caption = 0

End Sub
Private Sub RColBox_Change()

On Error GoTo reZero

If Len(RColBox.Value) > 3 Then RColBox.Value = "0"
If CInt(RColBox.Value) > 255 Then RColBox.Value = "255"
If CInt(RColBox.Value) < 0 Then RColBox.Value = "0"

If SwCtrl.Caption = "0" Then
xRGB = RColBox.Value & "," & GColBox.Value & "," & BColBox.Value
Call xlColorSwatch_TOOLS.setColors(xRGB)
End If

Exit Sub

reZero:
RColBox.Value = 0

End Sub
Private Sub GColBox_Change()

On Error GoTo reZero

If Len(GColBox.Value) > 3 Then GColBox.Value = "0"
If CInt(GColBox.Value) > 255 Then GColBox.Value = "255"
If CInt(RColBox.Value) < 0 Then RColBox.Value = "0"

If SwCtrl.Caption = "0" Then
xRGB = RColBox.Value & "," & GColBox.Value & "," & BColBox.Value
Call xlColorSwatch_TOOLS.setColors(xRGB)
End If

Exit Sub

reZero:
GColBox.Value = 0

End Sub
Private Sub BColBox_Change()

On Error GoTo reZero

If Len(BColBox.Value) > 3 Then BColBox.Value = "0"
If CInt(BColBox.Value) > 255 Then BColBox.Value = "255"
If CInt(RColBox.Value) < 0 Then RColBox.Value = "0"

If SwCtrl.Caption = "0" Then
xRGB = RColBox.Value & "," & GColBox.Value & "," & BColBox.Value
Call xlColorSwatch_TOOLS.setColors(xRGB)
End If

Exit Sub

reZero:
BColBox.Value = 0

End Sub
Private Sub RColUpDown_SpinDown()

On Error GoTo reZero

If RColBox.Value = vbNullString Or RColBox.Value < 1 Then RColBox.Value = 1
RColBox.Value = RColBox.Value - 1

Exit Sub

reZero:
RColBox.Value = 0

End Sub
Private Sub RColUpDown_SpinUp()

On Error GoTo reZero

If RColBox.Value = vbNullString Then RColBox.Value = 0
RColBox.Value = RColBox.Value + 1

Exit Sub

reZero:
RColBox.Value = 0

End Sub
Private Sub GColUpDown_SpinDown()

On Error GoTo reZero

If GColBox.Value = vbNullString Or GColBox.Value < 1 Then GColBox.Value = 1
GColBox.Value = GColBox.Value - 1

Exit Sub

reZero:
GColBox.Value = 0

End Sub
Private Sub GColUpDown_SpinUp()

On Error GoTo reZero

If GColBox.Value = vbNullString Then GColBox.Value = 0
GColBox.Value = GColBox.Value + 1

Exit Sub

reZero:
GColBox.Value = 0

End Sub
Private Sub BColUpDown_SpinDown()

On Error GoTo reZero

If BColBox.Value = vbNullString Or BColBox.Value < 1 Then BColBox.Value = 1
BColBox.Value = BColBox.Value - 1

Exit Sub

reZero:
BColBox.Value = 0

End Sub
Private Sub BColUpDown_SpinUp()

On Error GoTo reZero

If BColBox.Value = vbNullString Then BColBox.Value = 0
BColBox.Value = BColBox.Value + 1

Exit Sub
reZero:
BColBox.Value = 0

End Sub
Private Sub SwBaseLrg_Click()

xSw = "BaseLrg": Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub

Private Sub SwBaseSm_Click()

xSw = "BaseSm": Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw1_Click()

SwBaseLrg.BackColor = Sw1.BackColor
SwBaseSm.BackColor = Sw1.BackColor

xSw = 1: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw2_Click()

SwBaseLrg.BackColor = Sw2.BackColor
SwBaseSm.BackColor = Sw2.BackColor

xSw = 2: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw3_Click()

SwBaseLrg.BackColor = Sw3.BackColor
SwBaseSm.BackColor = Sw3.BackColor

xSw = 3: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw4_Click()

SwBaseLrg.BackColor = Sw4.BackColor
SwBaseSm.BackColor = Sw4.BackColor

xSw = 4: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw5_Click()

SwBaseLrg.BackColor = Sw5.BackColor
SwBaseSm.BackColor = Sw5BackColor

xSw = 5: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw6_Click()

SwBaseLrg.BackColor = Sw6.BackColor
SwBaseSm.BackColor = Sw6.BackColor

xSw = 6: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw7_Click()

SwBaseLrg.BackColor = Sw7.BackColor
SwBaseSm.BackColor = Sw7.BackColor

xSw = 7: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw8_Click()

SwBaseLrg.BackColor = Sw8.BackColor
SwBaseSm.BackColor = Sw8.BackColor

xSw = 8: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw9_Click()

SwBaseLrg.BackColor = Sw9.BackColor
SwBaseSm.BackColor = Sw9.BackColor

xSw = 9: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw10_Click()

SwBaseLrg.BackColor = Sw10.BackColor
SwBaseSm.BackColor = Sw10.BackColor

xSw = 10: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw11_Click()

SwBaseLrg.BackColor = Sw11.BackColor
SwBaseSm.BackColor = Sw11.BackColor

xSw = 11: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw12_Click()

SwBaseLrg.BackColor = Sw12.BackColor
SwBaseSm.BackColor = Sw12.BackColor

xSw = 12: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw13_Click()

SwBaseLrg.BackColor = Sw13.BackColor
SwBaseSm.BackColor = Sw13.BackColor

xSw = 13: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw14_Click()

SwBaseLrg.BackColor = Sw14.BackColor
SwBaseSm.BackColor = Sw14.BackColor

xSw = 14: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw15_Click()

SwBaseLrg.BackColor = Sw15.BackColor
SwBaseSm.BackColor = Sw15.BackColor

xSw = 15: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw16_Click()

SwBaseLrg.BackColor = Sw16.BackColor
SwBaseSm.BackColor = Sw16.BackColor

xSw = 16: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw17_Click()

SwBaseLrg.BackColor = Sw17.BackColor
SwBaseSm.BackColor = Sw17.BackColor

xSw = 17: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw18_Click()

SwBaseLrg.BackColor = Sw18.BackColor
SwBaseSm.BackColor = Sw18.BackColor

xSw = 18: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw19_Click()

SwBaseLrg.BackColor = Sw19.BackColor
SwBaseSm.BackColor = Sw19.BackColor

xSw = 19: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw20_Click()

SwBaseLrg.BackColor = Sw20.BackColor
SwBaseSm.BackColor = Sw20.BackColor

xSw = 20: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw21_Click()

SwBaseLrg.BackColor = Sw21.BackColor
SwBaseSm.BackColor = Sw21.BackColor

xSw = 21: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw22_Click()

SwBaseLrg.BackColor = Sw22.BackColor
SwBaseSm.BackColor = Sw22.BackColor

xSw = 22: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw23_Click()

SwBaseLrg.BackColor = Sw23.BackColor
SwBaseSm.BackColor = Sw23.BackColor

xSw = 23: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw24_Click()

SwBaseLrg.BackColor = Sw24.BackColor
SwBaseSm.BackColor = Sw24.BackColor

xSw = 24: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub Sw25_Click()

SwBaseLrg.BackColor = Sw25.BackColor
SwBaseSm.BackColor = Sw25.BackColor

xSw = 25: Call xlColorSwatch_CLICK.Sw_Clk(xSw)

End Sub
Private Sub UserForm_Terminate()

xRGB = RColBox.Value & "," & GColBox.Value & "," & BColBox.Value
Call xlColorSwatch_TOOLS.setColors(xRGB)

End Sub

