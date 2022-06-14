Attribute VB_Name = "CtrlBox_TOOLS"
Public Sub dfsMainScreen()

On Error Resume Next

Call fndEnvironment(appEnv, appBlk)

CTRLBOX.CtrlBoxWindow.Font.Size = ((12 * CInt(CTRLBOX.RemWinSizeValue)) / 24) + 12

If Workbooks(appEnv).Worksheets(appBlk).Range("xlasInvert").Value2 = 1 Then defCol1 = vbBlack: _
defCol2 = vbWhite: CTRLBOX.CtrlBoxWindow.ForeColor = defCol1: CTRLBOX.CtrlBoxWindow.BackColor = defCol2: Exit Sub
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasInvert").Value2 = 2 Then defCol1 = vbWhite: _
defCol2 = vbBlack: CTRLBOX.CtrlBoxWindow.ForeColor = defCol1: CTRLBOX.CtrlBoxWindow.BackColor = defCol2: Exit Sub

If Workbooks(appEnv).Worksheets(appBlk).Range("xlasCtrlBoxFColor").Value <> vbNullString Then defCol1 = Workbooks(appEnv).Worksheets(appBlk).Range("xlasCtrlBoxFColor").Value
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasCtrlBoxBColor").Value <> vbNullString Then defCol2 = Workbooks(appEnv).Worksheets(appBlk).Range("xlasCtrlBoxBColor").Value

If defCol1 <> vbNullString Then
defColArr = Split(Range("xlasCtrlBoxBColor").Value, ",")
R = defColArr(0): If R = vbNullString Then R = 0
G = defColArr(1): If G = vbNullString Then G = 0
B = defColArr(2): If B = vbNullString Then B = 0
XLFONTBOX.CurrBColor.BackColor = RGB(R, G, B)
CTRLBOX.CtrlBoxWindow.BackColor = RGB(R, G, B)
End If

If defCol2 <> vbNullString Then
defCol1Arr = Split(Range("xlasCtrlBoxFColor").Value, ",")
R = defCol1Arr(0): If R = vbNullString Then R = 0
G = defCol1Arr(1): If G = vbNullString Then G = 0
B = defCol1Arr(2): If B = vbNullString Then B = 0
XLFONTBOX.CurrFColor.BackColor = RGB(R, G, B)
CTRLBOX.CtrlBoxWindow.ForeColor = RGB(R, G, B)
End If

End Sub
Public Sub addOptions()

CTRLBOX.FileSel.Clear
CTRLBOX.EditSel.Clear
CTRLBOX.DebugSel.Clear
CTRLBOX.OptionsSel.Clear
CTRLBOX.RunSel.Clear
CTRLBOX.WindowSel.Clear
CTRLBOX.HelpSel.Clear

'//File options...
CTRLBOX.FileSel.AddItem ("New              Ctrl+N")
CTRLBOX.FileSel.AddItem ("Open            Ctrl+O")
CTRLBOX.FileSel.AddItem ("Save             Ctrl+S")
CTRLBOX.FileSel.AddItem ("Save As        Ctrl+Alt+S")
CTRLBOX.FileSel.AddItem ("Save & Exit  Ctrl+Alt+Q")
CTRLBOX.FileSel.AddItem ("Exit               Ctrl+Q")
CTRLBOX.FileSel.AddItem ("")
CTRLBOX.FileSel.AddItem ("")

'//Edit options...
CTRLBOX.EditSel.AddItem ("Undo                 Ctrl+Z")
CTRLBOX.EditSel.AddItem ("Cut                   Ctrl+X")
CTRLBOX.EditSel.AddItem ("Copy                 Ctrl+C")
CTRLBOX.EditSel.AddItem ("Paste                Ctrl+V")
CTRLBOX.EditSel.AddItem ("Replace             Ctrl+H")
CTRLBOX.EditSel.AddItem ("Clear Screen     Ctrl+D")
CTRLBOX.EditSel.AddItem ("Select All           Ctrl+A")
CTRLBOX.EditSel.AddItem ("")
CTRLBOX.EditSel.AddItem ("")

'//Debug options...
CTRLBOX.DebugSel.AddItem ("")

'//Options options...
CTRLBOX.OptionsSel.AddItem ("Screen Style           Ctrl+F")
CTRLBOX.OptionsSel.AddItem ("")
CTRLBOX.OptionsSel.AddItem ("")

'//Run options...
CTRLBOX.RunSel.AddItem ("Run Script         Shift")
CTRLBOX.RunSel.AddItem ("")
CTRLBOX.RunSel.AddItem ("")

'//Window options...
CTRLBOX.WindowSel.AddItem ("Hide                      Ctrl+Alt+W")
CTRLBOX.WindowSel.AddItem ("Invert Screen       Ctrl+I")
CTRLBOX.WindowSel.AddItem ("Remember            Ctrl+R")
CTRLBOX.WindowSel.AddItem ("Recall                    Ctrl+Alt+R")
CTRLBOX.WindowSel.AddItem ("Maximize               Ctrl+W")
CTRLBOX.WindowSel.AddItem ("Zoom In                Ctrl+Up")
CTRLBOX.WindowSel.AddItem ("Zoom Out              Ctrl+Down")
CTRLBOX.WindowSel.AddItem ("")
CTRLBOX.WindowSel.AddItem ("")

'//Help options...
CTRLBOX.HelpSel.AddItem ("About Control Box+  ")
CTRLBOX.HelpSel.AddItem ("Send Feedback       ")
CTRLBOX.HelpSel.AddItem ("")
CTRLBOX.HelpSel.AddItem ("")

End Sub
Public Function selOption(xBtn)

'//File...
If xBtn = 1 Then
CTRLBOX.FileSel.SpecialEffect = fmSpecialEffectEtched
CTRLBOX.FileSel.Height = 85
CTRLBOX.FileSel.Left = 12
CTRLBOX.FileSel.Top = 18
CTRLBOX.FileSel.Width = 115

'//Edit...
ElseIf xBtn = 2 Then
CTRLBOX.EditSel.SpecialEffect = fmSpecialEffectEtched
CTRLBOX.EditSel.Height = 95
CTRLBOX.EditSel.Left = 42
CTRLBOX.EditSel.Top = 18
CTRLBOX.EditSel.Width = 115

'//Debug...
ElseIf xBtn = 3 Then
CTRLBOX.DebugSel.SpecialEffect = fmSpecialEffectEtched
CTRLBOX.DebugSel.Height = 80
CTRLBOX.DebugSel.Left = 72
CTRLBOX.DebugSel.Top = 18
CTRLBOX.DebugSel.Width = 115

'//Options...
ElseIf xBtn = 4 Then
CTRLBOX.OptionsSel.SpecialEffect = fmSpecialEffectEtched
CTRLBOX.OptionsSel.Height = 40
CTRLBOX.OptionsSel.Left = 114
CTRLBOX.OptionsSel.Top = 18
CTRLBOX.OptionsSel.Width = 115

'//Run...
ElseIf xBtn = 5 Then
CTRLBOX.RunSel.SpecialEffect = fmSpecialEffectEtched
CTRLBOX.RunSel.Height = 40
CTRLBOX.RunSel.Left = 162
CTRLBOX.RunSel.Top = 18
CTRLBOX.RunSel.Width = 115

'//Window...
ElseIf xBtn = 6 Then
CTRLBOX.WindowSel.SpecialEffect = fmSpecialEffectEtched
CTRLBOX.WindowSel.Height = 95
CTRLBOX.WindowSel.Left = 192
CTRLBOX.WindowSel.Top = 18
CTRLBOX.WindowSel.Width = 125

'//Help...
ElseIf xBtn = 7 Then
CTRLBOX.HelpSel.SpecialEffect = fmSpecialEffectEtched
CTRLBOX.HelpSel.Height = 45
CTRLBOX.HelpSel.Left = 240
CTRLBOX.HelpSel.Top = 18
CTRLBOX.HelpSel.Width = 115

Exit Function

End If

End Function
Public Function fxsHover(xHov)

On Error Resume Next

'//File...
If xHov = 1 Then
CTRLBOX.FileBtn.ForeColor = RGB(185, 231, 170)
CTRLBOX.EditBtn.ForeColor = &H8000000E
CTRLBOX.DebugBtn.ForeColor = &H8000000E
CTRLBOX.OptionsBtn.ForeColor = &H8000000E
CTRLBOX.RunBtn.ForeColor = &H8000000E
CTRLBOX.WindowBtn.ForeColor = &H8000000E
CTRLBOX.HelpBtn.ForeColor = &H8000000E

'//Edit...
ElseIf xHov = 2 Then
CTRLBOX.EditBtn.ForeColor = RGB(185, 231, 170)
CTRLBOX.FileBtn.ForeColor = &H8000000E
CTRLBOX.DebugBtn.ForeColor = &H8000000E
CTRLBOX.OptionsBtn.ForeColor = &H8000000E
CTRLBOX.RunBtn.ForeColor = &H8000000E
CTRLBOX.WindowBtn.ForeColor = &H8000000E
CTRLBOX.HelpBtn.ForeColor = &H8000000E

'//Debug...
ElseIf xHov = 3 Then
CTRLBOX.DebugBtn.ForeColor = RGB(185, 231, 170)
CTRLBOX.FileBtn.ForeColor = &H8000000E
CTRLBOX.RunBtn.ForeColor = &H8000000E
CTRLBOX.WindowBtn.ForeColor = &H8000000E
CTRLBOX.OptionsBtn.ForeColor = &H8000000E
CTRLBOX.EditBtn.ForeColor = &H8000000E
CTRLBOX.HelpBtn.ForeColor = &H8000000E

'//Options...
ElseIf xHov = 4 Then
CTRLBOX.OptionsBtn.ForeColor = RGB(185, 231, 170)
CTRLBOX.FileBtn.ForeColor = &H8000000E
CTRLBOX.DebugBtn.ForeColor = &H8000000E
CTRLBOX.EditBtn.ForeColor = &H8000000E
CTRLBOX.RunBtn.ForeColor = &H8000000E
CTRLBOX.WindowBtn.ForeColor = &H8000000E
CTRLBOX.HelpBtn.ForeColor = &H8000000E

'//Run...
ElseIf xHov = 5 Then
CTRLBOX.RunBtn.ForeColor = RGB(185, 231, 170)
CTRLBOX.FileBtn.ForeColor = &H8000000E
CTRLBOX.DebugBtn.ForeColor = &H8000000E
CTRLBOX.OptionsBtn.ForeColor = &H8000000E
CTRLBOX.EditBtn.ForeColor = &H8000000E
CTRLBOX.WindowBtn.ForeColor = &H8000000E
CTRLBOX.HelpBtn.ForeColor = &H8000000E

'//Window...
ElseIf xHov = 6 Then
CTRLBOX.WindowBtn.ForeColor = RGB(185, 231, 170)
CTRLBOX.FileBtn.ForeColor = &H8000000E
CTRLBOX.DebugBtn.ForeColor = &H8000000E
CTRLBOX.OptionsBtn.ForeColor = &H8000000E
CTRLBOX.EditBtn.ForeColor = &H8000000E
CTRLBOX.RunBtn.ForeColor = &H8000000E
CTRLBOX.HelpBtn.ForeColor = &H8000000E

'//Help...
ElseIf xHov = 7 Then
CTRLBOX.HelpBtn.ForeColor = RGB(185, 231, 170)
CTRLBOX.FileBtn.ForeColor = &H8000000E
CTRLBOX.DebugBtn.ForeColor = &H8000000E
CTRLBOX.OptionsBtn.ForeColor = &H8000000E
CTRLBOX.EditBtn.ForeColor = &H8000000E
CTRLBOX.RunBtn.ForeColor = &H8000000E
CTRLBOX.WindowBtn.ForeColor = &H8000000E

Exit Function

End If

End Function
Sub dfsHover()

On Error Resume Next

'//File...
CTRLBOX.FileSel.SpecialEffect = fmSpecialEffectFlat
CTRLBOX.FileSel.Height = 0
CTRLBOX.FileSel.Left = 0
CTRLBOX.FileSel.Top = 0
CTRLBOX.FileSel.Width = 0
CTRLBOX.FileSel.Visible = True

'//Edit...
CTRLBOX.EditSel.SpecialEffect = fmSpecialEffectFlat
CTRLBOX.EditSel.Height = 0
CTRLBOX.EditSel.Left = 0
CTRLBOX.EditSel.Top = 0
CTRLBOX.EditSel.Width = 0
CTRLBOX.EditSel.Visible = True

'//Debug...
CTRLBOX.DebugSel.SpecialEffect = fmSpecialEffectFlat
CTRLBOX.DebugSel.Height = 0
CTRLBOX.DebugSel.Left = 0
CTRLBOX.DebugSel.Top = 0
CTRLBOX.DebugSel.Width = 0
CTRLBOX.DebugSel.Visible = True

'//Options...
CTRLBOX.OptionsSel.SpecialEffect = fmSpecialEffectFlat
CTRLBOX.OptionsSel.Height = 0
CTRLBOX.OptionsSel.Left = 0
CTRLBOX.OptionsSel.Top = 0
CTRLBOX.OptionsSel.Width = 0
CTRLBOX.OptionsSel.Visible = True

'//Run...
CTRLBOX.RunSel.SpecialEffect = fmSpecialEffectFlat
CTRLBOX.RunSel.Height = 0
CTRLBOX.RunSel.Left = 0
CTRLBOX.RunSel.Top = 0
CTRLBOX.RunSel.Width = 0
CTRLBOX.RunSel.Visible = True

'//Window...
CTRLBOX.WindowSel.SpecialEffect = fmSpecialEffectFlat
CTRLBOX.WindowSel.Height = 0
CTRLBOX.WindowSel.Left = 0
CTRLBOX.WindowSel.Top = 0
CTRLBOX.WindowSel.Width = 0
CTRLBOX.WindowSel.Visible = True

'//Help...
CTRLBOX.HelpSel.SpecialEffect = fmSpecialEffectFlat
CTRLBOX.HelpSel.Height = 0
CTRLBOX.HelpSel.Left = 0
CTRLBOX.HelpSel.Top = 0
CTRLBOX.HelpSel.Width = 0
CTRLBOX.HelpSel.Visible = True

'//Default font colors...
CTRLBOX.FileBtn.ForeColor = &H8000000E
CTRLBOX.EditBtn.ForeColor = &H8000000E
CTRLBOX.DebugBtn.ForeColor = &H8000000E
CTRLBOX.OptionsBtn.ForeColor = &H8000000E
CTRLBOX.RunBtn.ForeColor = &H8000000E
CTRLBOX.WindowBtn.ForeColor = &H8000000E
CTRLBOX.HelpBtn.ForeColor = &H8000000E

'//Default font size...
CTRLBOX.FileBtn.Font.Size = 9
CTRLBOX.EditBtn.Font.Size = 9
CTRLBOX.DebugBtn.Font.Size = 9
CTRLBOX.OptionsBtn.Font.Size = 9
CTRLBOX.RunBtn.Font.Size = 9
CTRLBOX.WindowBtn.Font.Size = 9
CTRLBOX.HelpBtn.Font.Size = 9

End Sub
Function undHover(xHov)

On Error Resume Next

'//File...
If xHov = 1 Then
CTRLBOX.FileBtn.Font.Underline = True
CTRLBOX.EditBtn.Font.Underline = False
CTRLBOX.OptionsBtn.Font.Underline = False
CTRLBOX.DebugBtn.Font.Underline = False
CTRLBOX.RunBtn.Font.Underline = False
CTRLBOX.WindowBtn.Font.Underline = False
CTRLBOX.HelpBtn.Font.Underline = False
Exit Function
End If

'//Edit...
If xHov = 2 Then
CTRLBOX.EditBtn.Font.Underline = True
CTRLBOX.FileBtn.Font.Underline = False
CTRLBOX.OptionsBtn.Font.Underline = False
CTRLBOX.DebugBtn.Font.Underline = False
CTRLBOX.RunBtn.Font.Underline = False
CTRLBOX.WindowBtn.Font.Underline = False
CTRLBOX.HelpBtn.Font.Underline = False
Exit Function
End If

'//Debug...
If xHov = 3 Then
CTRLBOX.DebugBtn.Font.Underline = True
CTRLBOX.FileBtn.Font.Underline = False
CTRLBOX.EditBtn.Font.Underline = False
CTRLBOX.OptionsBtn.Font.Underline = False
CTRLBOX.RunBtn.Font.Underline = False
CTRLBOX.WindowBtn.Font.Underline = False
CTRLBOX.HelpBtn.Font.Underline = False
Exit Function
End If

'//Options...
If xHov = 4 Then
CTRLBOX.OptionsBtn.Font.Underline = True
CTRLBOX.FileBtn.Font.Underline = False
CTRLBOX.EditBtn.Font.Underline = False
CTRLBOX.DebugBtn.Font.Underline = False
CTRLBOX.RunBtn.Font.Underline = False
CTRLBOX.WindowBtn.Font.Underline = False
Exit Function
End If

'//Run...
If xHov = 5 Then
CTRLBOX.RunBtn.Font.Underline = True
CTRLBOX.FileBtn.Font.Underline = False
CTRLBOX.EditBtn.Font.Underline = False
CTRLBOX.OptionsBtn.Font.Underline = False
CTRLBOX.DebugBtn.Font.Underline = False
CTRLBOX.WindowBtn.Font.Underline = False
CTRLBOX.HelpBtn.Font.Underline = False
Exit Function
End If

'//Window...
If xHov = 6 Then
CTRLBOX.WindowBtn.Font.Underline = True
CTRLBOX.FileBtn.Font.Underline = False
CTRLBOX.EditBtn.Font.Underline = False
CTRLBOX.OptionsBtn.Font.Underline = False
CTRLBOX.DebugBtn.Font.Underline = False
CTRLBOX.RunBtn.Font.Underline = False
CTRLBOX.HelpBtn.Font.Underline = False
Exit Function
End If

'//Help...
If xHov = 7 Then
CTRLBOX.HelpBtn.Font.Underline = True
CTRLBOX.FileBtn.Font.Underline = False
CTRLBOX.EditBtn.Font.Underline = False
CTRLBOX.OptionsBtn.Font.Underline = False
CTRLBOX.DebugBtn.Font.Underline = False
CTRLBOX.RunBtn.Font.Underline = False
CTRLBOX.WindowBtn.Font.Underline = False
Exit Function
End If

End Function
Public Function dfsWindow()

CTRLBOX.Height = 510
CTRLBOX.Width = 510
CTRLBOX.CtrlBoxWindow.Height = 438.75
CTRLBOX.CtrlBoxWindow.Width = 480
CTRLBOX.SideBar1.Height = 438
CTRLBOX.SideBar1.Left = 480
CTRLBOX.RemCol.Top = 462
CTRLBOX.RemEnco.Left = 432
CTRLBOX.RemLen.Top = 462
CTRLBOX.RemLight.Left = 462
CTRLBOX.RemLight.Top = 456
CTRLBOX.RemLine.Top = 462
CTRLBOX.RemLines.Top = 462
CTRLBOX.RemStatus.Left = 402
CTRLBOX.RemStatus.Top = 462
CTRLBOX.RemSys.Top = 462
CTRLBOX.RemWinSize.Top = 462

End Function
Public Function setColors(xRGB)

On Error Resume Next

Dim X As Integer
Dim oControl As Object

'//Find RGB
xRGBArr = Split(xRGB, ",")

For X = 0 To UBound(xRGBArr)
If xRGBArr(X) = vbNullString Then xRGBArr(X) = 0
Next

'//Set gradient swatch
X = 1
For Each oControl In XLFONTSWATCH.Controls
If InStr(1, oControl.name, "Sw" & X) Then
oControl.BackColor = RGB((-1 * (xRGBArr(0) - ((2 * X) + 3)) - xR), (xRGBArr(1) - 10), (xRGBArr(2) - 20))
oControl.Caption = xRGBArr(0) - ((2 * X) + 3) & "," & (xRGBArr(1) - 10) & "," & (xRGBArr(2) - 20)
oControl.ForeColor = oControl.BackColor
X = X + 1
End If
Next

XLFONTSWATCH.SwBaseLrg.BackColor = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
XLFONTSWATCH.SwBaseSm.BackColor = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))

If InStr(1, XLFONTSWATCH.CurrType.Caption, "B") Then Range("xlasCtrlBoxBColor").Value = xRGBArr(0) & "," & xRGBArr(1) & "," & xRGBArr(2)
If InStr(1, XLFONTSWATCH.CurrType.Caption, "F") Then Range("xlasCtrlBoxFColor").Value = xRGBArr(0) & "," & xRGBArr(1) & "," & xRGBArr(2)

If Range("xlasCtrlBoxBColor").Value <> vbNullString Then
xRGBArr = Split(Range("xlasCtrlBoxBColor").Value, ",")
R = xRGBArr(0): If R = vbNullString Then R = 0
G = xRGBArr(1): If G = vbNullString Then G = 0
B = xRGBArr(2): If B = vbNullString Then B = 0
XLFONTBOX.CurrBColor.BackColor = RGB(R, G, B)
CTRLBOX.CtrlBoxWindow.BackColor = RGB(R, G, B)
End If

If Range("xlasCtrlBoxFColor").Value <> vbNullString Then
xRGBArr = Split(Range("xlasCtrlBoxFColor").Value, ",")
R = xRGBArr(0): If R = vbNullString Then R = 0
G = xRGBArr(1): If G = vbNullString Then G = 0
B = xRGBArr(2): If B = vbNullString Then B = 0
XLFONTBOX.CurrFColor.BackColor = RGB(R, G, B)
CTRLBOX.CtrlBoxWindow.ForeColor = RGB(R, G, B)
End If

End Function
Public Function setWindowStats()

On Error Resume Next

CTRLBOX.RemLine.Caption = "Ln " & CTRLBOX.CtrlBoxWindow.CurLine
CTRLBOX.RemCol.Caption = "Col " & CTRLBOX.CtrlBoxWindow.SelStart
CTRLBOX.RemWinSize.Caption = CStr(100 + CInt(CTRLBOX.RemWinSizeValue.Caption)) & "%"
CTRLBOX.RemLen.Caption = "Len " & Len(CTRLBOX.CtrlBoxWindow.Value)
CTRLBOX.RemLines.Caption = "Lns " & CTRLBOX.CtrlBoxWindow.LineCount

End Function
Public Sub updAppState()

If Range("xlasWinForm").Value2 = vbNullString Then Range("xlasWinForm").Value2 = 11

'//Set previous WinForm #
Range("xlasWinFormLast").Value2 = Range("xlasWinForm").Value2

Call fndWindow(xWin)
X = xWin.Left: Y = xWin.Top
Call basPostWinFormPos(xWin, X, Y)
Call basSetWinFormPos(xWin, X, Y)
Set xWin = Nothing

End Sub


