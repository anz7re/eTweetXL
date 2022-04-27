Attribute VB_Name = "Ctrlbox_CLICK"
Public Sub NewFile_Clk()

Dim xFile As String

On Error GoTo EndMacro

xFile = InputBox("Enter a name for your project:", CtrlTag)

If xFile <> "" Then

CTRLBOX.Caption = CtrlTag & " - " & xFile

CTRLBOX.CtrlBoxWindow.Value = vbNullString
Range("xlasSaveFile").Value = vbNullString

End If

EndMacro:
End Sub
Public Sub OpenFile_Clk()

Dim xData, xDataHldr, xFile As String

On Error GoTo EndMacro

xFile = Application.GetOpenFilename()

If xFile <> "" Then

Open xFile For Input As #1
Do Until EOF(1)
Line Input #1, xData
xDataHldr = xDataHldr & xData
Loop
Close #1

CTRLBOX.CtrlBoxWindow.Value = xDataHldr

End If

EndMacro:
End Sub
Public Sub SaveFile_Clk()

Dim xFile, xStr As String

On Error GoTo EndMacro

If Range("xlasSaveFile").Value <> "" Then

xFile = Range("xlasSaveFile").Value
xStr = CTRLBOX.CtrlBoxWindow.Value

Open xFile For Output As #1
Print #1, xStr
Close #1

Else

Call Ctrlbox_CLICK.SaveAsFile_Clk

End If

EndMacro:
End Sub
Public Sub SaveAsFile_Clk()

Dim xFile, xStr As String

On Error GoTo EndMacro

xFile = Application.GetSaveAsFilename(ProjTitle)

If xFile <> "" Then
If xFile <> False Then

xStr = CTRLBOX.CtrlBoxWindow.Value

Open xFile For Output As #1
Print #1, xStr
Close #1

Range("xlasSaveFile").Value = xFile

End If
    End If

EndMacro:
End Sub
Public Sub Remember_Clk()

If Range("xlasRemember").Value <> 1 Then
Range("xlasAMemory").Value = vbNullString
Range("xlasRemember").Value = 1
CTRLBOX.RemLight.Visible = True
CTRLBOX.RemStatus.Caption = "Remembering..."
Else
    Range("xlasRemember").Value = 0
    CTRLBOX.RemLight.Visible = False
    CTRLBOX.RemStatus.Caption = vbNullString
        End If

End Sub
Public Sub Recall_Clk()

CTRLBOX.CtrlBoxWindow.Value = Range("xlasAMemory").Value

End Sub
Public Sub InvertScreen_Clk()

If Range("xlasInvert").Value <> 2 Then Range("xlasInvert").Value = 2 Else Range("xlasInvert").Value = 1

End Sub
Public Sub ClearScreen_Clk()

CTRLBOX.CtrlBoxWindow.Value = vbNullString

End Sub
Public Sub Maximize_Clk()

If CTRLBOX.Width = 1375 Then Call CtrlBox_TOOLS.WindowDefault: Exit Sub

CTRLBOX.Top = 0
CTRLBOX.Left = 0
CTRLBOX.Height = 850
CTRLBOX.Width = 1375
CTRLBOX.CtrlBoxWindow.Height = 745
CTRLBOX.CtrlBoxWindow.Width = 1344
CTRLBOX.SideBar1.Height = 745
CTRLBOX.SideBar1.Left = 1344
CTRLBOX.RemLight.Left = 1320
CTRLBOX.RemStatus.Left = 1260
CTRLBOX.RemLight.Top = 764
CTRLBOX.RemStatus.Top = 770

End Sub
Public Sub Sw_Clk(xSw)

On Error Resume Next

Dim oControl As Object

For Each oControl In XLFONTSWATCH.Controls
If InStr(1, oControl.name, "Sw" & xSw) Then
xSwArr = Split(oControl.Caption, ",")
XLFONTSWATCH.SwCtrl.Caption = 1
XLFONTSWATCH.RColBox = xSwArr(0)
XLFONTSWATCH.GColBox = xSwArr(1)
XLFONTSWATCH.SwCtrl.Caption = 0
XLFONTSWATCH.BColBox = xSwArr(2)
If InStr(1, XLFONTSWATCH.CurrType, "B") Then CTRLBOX.CtrlBoxWindow.BackColor = oControl.BackColor
If InStr(1, XLFONTSWATCH.CurrType, "F") Then CTRLBOX.CtrlBoxWindow.ForeColor = oControl.BackColor
End If
Next

End Sub



