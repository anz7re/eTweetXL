Attribute VB_Name = "Ctrlbox_CLICK"
Public Sub CommentBlock_Clk()

End Sub
Public Sub UncommentBlock_Clk()

End Sub
Public Sub Hide_Clk()

Dim M As Byte

Unload CTRLBOX
M = MsgBox("Display Control Box+?", vbOKOnly, CtrlBox_INFO.AppTag)
    If M = vbOK Then
        Call CtrlBox_FOCUS.shw_CTRLBOX
            End If

End Sub
Public Sub NewFile_Clk()

Dim xFile As String

On Error GoTo EndMacro

xFile = InputBox("Enter a name for your project:", CtrlBox_INFO.AppTag)

If xFile <> "" Then

CTRLBOX.Caption = CtrlBox_INFO.AppTag & " - " & xFile

CTRLBOX.CtrlBoxWindow.Value = vbNullString
Range("xlasSaveFile").Value2 = vbNullString

End If

EndMacro:
End Sub
Public Sub OpenFile_Clk()

Dim xData, xDataHldr, xFile As String

On Error GoTo EndMacro

xFile = Application.GetOpenFilename()

If xFile <> "" Then

Range("xlasSaveFile").Value2 = xFile

Open xFile For Input As #1
Do Until EOF(1)
Line Input #1, xData
xDataHldr = xDataHldr & vbCrLf & xData
Loop
Close #1

CTRLBOX.CtrlBoxWindow.Value = xDataHldr
CTRLBOX.Caption = CtrlBox_INFO.AppTag & " - " & xFile

End If

EndMacro:
End Sub
Public Sub SaveFile_Clk()

Dim xFile, xStr As String

On Error GoTo EndMacro

If Range("xlasSaveFile").Value2 <> "" Then

xFile = Range("xlasSaveFile").Value2
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

xFile = Application.GetSaveAsFilename(CtrlBox_INFO.AppTitle)

If xFile <> "" Then
If xFile <> False Then

xStr = CTRLBOX.CtrlBoxWindow.Value

Open xFile For Output As #1
Print #1, xStr
Close #1

Range("xlasSaveFile").Value2 = xFile

End If
    End If

EndMacro:
End Sub
Public Sub SendFeedback_Clk()

Art = "<lib>xbas;sh(mailto:mail@autokit.tech);$": Call xlas(Art)

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

CTRLBOX.CtrlBoxWindow.Value = CTRLBOX.CtrlBoxWindow.Value & Range("xlasAMemory").Value

End Sub
Public Sub InvertScreen_Clk()

If Range("xlasInvert").Value2 <> 2 Then Range("xlasInvert").Value2 = 2 Else Range("xlasInvert").Value2 = 1

End Sub
Public Sub ClearScreen_Clk()

CTRLBOX.CtrlBoxWindow.Value = vbNullString

End Sub
Public Sub Maximize_Clk()

If CTRLBOX.Width > 510 Then: Call CtrlBox_TOOLS.shTskBar: Art = "<lib>xbas;wait(3s);$": Call xlas(Art): Unload CTRLBOX: Call CtrlBox_TOOLS.dfsWindow: CTRLBOX.Show: Exit Sub

Call CtrlBox_TOOLS.hdTskBar: Art = "<lib>xbas;wait(3s);$": Call xlas(Art): Call CtrlBox_TOOLS.fxsWindow

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
Public Function ZoomUp_Clk()

If CInt(CTRLBOX.RemWinSizeValue.Caption) < 400 Then
CTRLBOX.RemWinSizeValue.Caption = CTRLBOX.RemWinSizeValue.Caption + 10
End If

Call setWindowStats
Call dfsMainScreen

End Function
Public Function ZoomDown_Clk()

If CInt(CTRLBOX.RemWinSizeValue.Caption) > -50 Then
CTRLBOX.RemWinSizeValue.Caption = CTRLBOX.RemWinSizeValue.Caption - 10
End If

Call setWindowStats
Call dfsMainScreen

End Function



