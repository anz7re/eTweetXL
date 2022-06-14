VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} XLFONTBOX 
   ClientHeight    =   1530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4650
   OleObjectBlob   =   "XLFONTBOX.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "XLFONTBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()

On Error Resume Next

If Range("xlasBlkAddr96").Value <> vbNullString Then xRGBArr = Split(Range("xlasBlkAddr96").Value, ","): CurrBColor.BackColor = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
If Range("xlasBlkAddr97").Value <> vbNullString Then xRGBArr = Split(Range("xlasBlkAddr97").Value, ","): CurrFColor.BackColor = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))

End Sub

Private Sub UserForm_Initialize()

Dim oFSO, oFile, oFldr As Object
Dim X As Integer
Dim FontArr(1000) As String

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFldr = oFSO.GetFolder(drv & "\Windows\Fonts\")

X = 1
For Each oFile In oFldr.Files

If InStr(1, oFile.Path, ".ttf") Then
FontArr(X) = Replace(oFile.name, ".ttf", vbNullString)
XLFONTBOX.FontStyleBox.AddItem (FontArr(X))
XLFONTBOX.FontSizeBox.AddItem (X + 2)
X = X + 1
End If
Next

FontStyleBox.Value = "MS Reference Sans Serif"
XLFONTBOX.FontUnderlineBtn.Font.Underline = True

End Sub
Private Sub CurrBColor_Click()

XLCOLORSWATCH.CurrType.Caption = "B"
XLCOLORSWATCH.Show

End Sub
Private Sub CurrFColor_Click()

XLCOLORSWATCH.CurrType.Caption = "F"
XLCOLORSWATCH.Show

End Sub
Private Sub FontSizeBox_Change()

'//set font size
Dim oAppTxt As Object

'//application object variables
Set oAppTxt = CTRLBOX.CtrlBoxWindow

oAppTxt.Font.Size = FontSizeBox.Value

End Sub
Private Sub FontStyleBox_Change()

'//set font style
Dim oAppTxt As Object

'//application object variables
Set oAppTxt = CTRLBOX.CtrlBoxWindow

oAppTxt.Font.name = FontStyleBox.Value

Set oAppTxt = Nothing

End Sub
Private Sub FontBoldBtn_Click()

'//set font bold
Dim oAppTxt As Object

'//application object variables
Set oAppTxt = CTRLBOX.CtrlBoxWindow

If oAppTxt.Font.Bold = False Then
oAppTxt.Font.Bold = True
Else
    oAppTxt.Font.Bold = False
        End If

Set oAppTxt = Nothing

End Sub
Private Sub FontItalicBtn_Click()

'//set font bold
Dim oAppTxt As Object

'//application object variables
Set oAppTxt = CTRLBOX.CtrlBoxWindow

'//set font italic
If oAppTxt.Font.Italic = False Then
oAppTxt.Font.Italic = True
Else
    oAppTxt.Font.Italic = False
        End If

Set oAppTxt = Nothing

End Sub
Private Sub FontUnderlineBtn_Click()

'//set font bold
Dim oAppTxt As Object

'//application object variables
Set oAppTxt = CTRLBOX.CtrlBoxWindow

'//set font underline
If oAppTxt.Font.Underline = False Then
oAppTxt.Font.Underline = True
Else
    oAppTxt.Font.Underline = False
        End If
        
Set oAppTxt = Nothing

End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Call dfsHover

End Sub
Private Sub FontSizeUpBtn_Click()

If FontSizeBox.Value = vbNullString Then FontSizeBox.Value = 7
FontSizeBox.Value = FontSizeBox.Value + 1

End Sub
Private Sub FontSizeDwnBtn_Click()

If FontSizeBox.Value = vbNullString Then FontSizeBox.Value = 7
FontSizeBox.Value = FontSizeBox.Value - 1
If FontSizeBox.Value < 1 Then FontSizeBox.Value = 1

End Sub
Private Sub FontBoldBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 1
Call fxsHover(xHov)

End Sub
Private Sub FontItalicBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 2
Call fxsHover(xHov)

End Sub
Private Sub FontUnderlineBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 3
Call fxsHover(xHov)

End Sub
Private Sub FontSizeUpBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 4
Call fxsHover(xHov)

End Sub
Private Sub FontSizeDwnBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 5
Call fxsHover(xHov)

End Sub
Private Sub BgColorBtn_Click()

XLCOLORSWATCH.CurrType.Caption = "B"
XLCOLORSWATCH.Show

End Sub
Private Sub BgColorBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 6
Call fxsHover(xHov)

End Sub
Private Sub FontColorBtn_Click()

XLCOLORSWATCH.CurrType.Caption = "F"
XLCOLORSWATCH.Show

End Sub
Private Sub FontColorBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

xHov = 7
Call fxsHover(xHov)

End Sub
Private Function fxsHover(xHov)

On Error Resume Next

If xHov = 1 Then
XLFONTBOX.FontBoldBtn.BackColor = RGB(185, 231, 170)
XLFONTBOX.FontBoldBtn.BackStyle = fmBackStyleOpaque
XLFONTBOX.FontItalicBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontUnderlineBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontSizeUpBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontSizeDwnBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.BgColorBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontColorBtn.BackStyle = fmBackStyleTransparent

ElseIf xHov = 2 Then
XLFONTBOX.FontItalicBtn.BackColor = RGB(185, 231, 170)
XLFONTBOX.FontItalicBtn.BackStyle = fmBackStyleOpaque
XLFONTBOX.FontBoldBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontUnderlineBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontSizeUpBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontSizeDwnBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.BgColorBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontColorBtn.BackStyle = fmBackStyleTransparent

ElseIf xHov = 3 Then
XLFONTBOX.FontUnderlineBtn.BackColor = RGB(185, 231, 170)
XLFONTBOX.FontUnderlineBtn.BackStyle = fmBackStyleOpaque
XLFONTBOX.FontBoldBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontSizeDwnBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.BgColorBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontSizeUpBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontItalicBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontColorBtn.BackStyle = fmBackStyleTransparent

ElseIf xHov = 4 Then
XLFONTBOX.FontSizeUpBtn.BackColor = RGB(185, 231, 170)
XLFONTBOX.FontSizeUpBtn.BackStyle = fmBackStyleOpaque
XLFONTBOX.FontBoldBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontUnderlineBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontItalicBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontSizeDwnBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.BgColorBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontColorBtn.BackStyle = fmBackStyleTransparent

ElseIf xHov = 5 Then
XLFONTBOX.FontSizeDwnBtn.BackColor = RGB(185, 231, 170)
XLFONTBOX.FontSizeDwnBtn.BackStyle = fmBackStyleOpaque
XLFONTBOX.FontBoldBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontUnderlineBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontSizeUpBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontItalicBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.BgColorBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontColorBtn.BackStyle = fmBackStyleTransparent

ElseIf xHov = 6 Then
XLFONTBOX.BgColorBtn.BackColor = RGB(185, 231, 170)
XLFONTBOX.BgColorBtn.BackStyle = fmBackStyleOpaque
XLFONTBOX.FontBoldBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontUnderlineBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontSizeUpBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontItalicBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontSizeDwnBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontColorBtn.BackStyle = fmBackStyleTransparent

ElseIf xHov = 7 Then
XLFONTBOX.FontColorBtn.BackColor = RGB(185, 231, 170)
XLFONTBOX.FontColorBtn.BackStyle = fmBackStyleOpaque
XLFONTBOX.FontBoldBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontUnderlineBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontSizeUpBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontItalicBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontSizeDwnBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.BgColorBtn.BackStyle = fmBackStyleTransparent

'//(hover light blue) RGB(201, 233, 246)

Exit Function

End If

End Function
Private Function dfsHover()

XLFONTBOX.FontBoldBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontItalicBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontUnderlineBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontSizeUpBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontSizeDwnBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.BgColorBtn.BackStyle = fmBackStyleTransparent
XLFONTBOX.FontColorBtn.BackStyle = fmBackStyleTransparent

End Function
