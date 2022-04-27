Attribute VB_Name = "xlAppScript_xbas"
'//Library API Calls
'
'//user32 mouse event library API call (click() article)
Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal mFlags As Long, ByVal mX As Long, ByVal mY As Long, ByVal mButtons As Long, ByVal mInfo As Long)
'//user32 cursor position library API call (click() article)
Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'//mouse event API call functions (click() article)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10
Public Function runLib$(xArt)
'/\_____________________________________________________________________________________________________________________________
'//
'//     xbas (basic) Library
'//        Version: 1.0.8
'/\_____________________________________________________________________________________________________________________________
'//
'//     License Information:
'//
'//     Copyright (C) 2022-present, Autokit Technology.
'//
'//     Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
'//
'//     1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
'//
'//     2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
'//
'//     3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
'//
'//     THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
'//     THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
'//     (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
'//     HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'//
'/\_____________________________________________________________________________________________________________________________
'//
'//     xbas is a simple scripting library for automating basic tasks in Microsoft Excel & Windows.
'//
'//
'//
'//     Basic Lib Requirements: Windows 10, MS Excel Version 2107, PowerShell 5.1.19041.1023
'//
'//                             (previous versions not tested &/or unsupported)
'/\_____________________________________________________________________________________________________________________________
'//
'//     Latest Revision: 4/23/2022
'/\_____________________________________________________________________________________________________________________________
'//
'//     Developer(s): anz7re
'//     Contact: support@xlappscript.org | support@autokit.tech | anz7re@autokit.tech
'//     Web: xlappscript.org | autokit.tech
'/\_____________________________________________________________________________________________________________________________
        
        '//Library variable declarations
        Dim oFSO As Object: Dim oDrv As Object: Dim oFile As Object: Dim oSubFldr As Object
        Dim appEnv As String: Dim appBlk As String: Dim xCell As String: Dim FX As String: Dim HX As String
        Dim sysShell As String: Dim wbMacro As String: Dim xArtArr() As String: Dim xArtCl As String: Dim xArtArrCl() As String
        Dim xExt As String: Dim xExtArr() As String: Dim xRGBArr() As String: Dim xMod As String: Dim xWb As String: Dim xVar As String
        Dim BX As Long: Dim EX As Long: Dim CX As Long: Dim PX As Long: Dim SX As Long: Dim TX As Long: Dim x1 As Long: Dim y1 As Long: Dim x2 As Long: Dim y2 As Long
        Dim C As Byte: Dim E As Byte: Dim K As Byte: Dim S As Byte: Dim T As Byte: Dim errLvl As Byte
        Dim X As Variant
        
        '//Pre-cleanup
        x1 = 0: x2 = 0: y1 = 0: y2 = 0: BX = 0: CX = 0: PX = 0: SX = 0: TX = 0: C = 0: E = 0: K = 0: S = 0: T = 0: X = 0: X = CByte(X)
        Call modArtQ(xArt)
        
        '//Find application environment & block
        Call findEnvironment(appEnv, appBlk)
        
        '//Find flags
        If InStr(1, xArt, "--") Or InStr(1, xArt, "++") Then _
        Call libFlag(xArt, errLvl): If xArt = 1 Then Exit Function Else _
        Call libSwitch(xArt, errLvl) '//Find switches
        
        '//Set library error level
        If Range("xlasLibErrLvl").Value = 0 Then On Error GoTo ErrMsg Else _
        If Range("xlasLibErrLvl").Value = 1 Then On Error Resume Next
        
'/\_____________________________________
'//
'//          APPLICATION ARTICLES
'/\_____________________________________
'//
'//Application build...
If InStr(1, xArt, "build(", vbTextCompare) Then
xArt = Replace(xArt, "app.", vbNullString, , , vbTextCompare)
xArt = Replace(xArt, "build(", vbNullString, , , vbTextCompare)
Call modArtQ(xArt)

If InStr(1, xArt, ",") Then xArtArr = Split(xArt, ",") Else MsgBox MsgBox(Application.Build): Exit Function  '//no excerpt provided
Exit Function

'//Application printer...
ElseIf InStr(1, xArt, "printer(", vbTextCompare) Then
xArt = Replace(xArt, "app.", vbNullString, , , vbTextCompare)
xArt = Replace(xArt, "printer(", vbNullString, , , vbTextCompare)
Call modArtQ(xArt)

If InStr(1, xArt, ",") Then xArtArr = Split(xArt, ",") Else MsgBox (Application.ActivePrinter): Exit Function '//no excerpt provided
Exit Function

'//Application name...
ElseIf InStr(1, xArt, "name(", vbTextCompare) Then
xArt = Replace(xArt, "app.", vbNullString, , , vbTextCompare)
xArt = Replace(xArt, "name(", vbNullString, , , vbTextCompare)
Call modArtQ(xArt)

If InStr(1, xArt, ",") Then xArtArr = Split(xArt, ",") Else MsgBox (Application.name): Exit Function '//no excerpt provided
Exit Function

'//Application run module...
ElseIf InStr(1, xArt, "run(", vbTextCompare) Then

xArt = Replace(xArt, "app.", vbNullString, , , vbTextCompare)
xArt = Replace(xArt, "run(", vbNullString, , , vbTextCompare)
Call modArtQ(xArt)

If InStr(1, xArt, ",") Then xArtArr = Split(xArt, ",") Else X = Application.Run(xArt): Exit Function '//no arguments provided
xMod = xArtArr(0) '//extract module

X = 1
Do Until X > UBound(xArtArr) '//extract argument(s)
xArt = xArtArr(X) & ",": xArtCl = xArtCl & xArt
X = X + 1
Loop

xArt = xArtCl
If Right(xArt, Len(xArt) - Len(xArt) + 1) = "," Then xArt = Left(xArt, Len(xArt) - 1)

X = Application.Run(xMod, (xArt))
Exit Function
'//#
'//
'/\_____________________________________
'//
'//          CELL/RANGE ARTICLES
'/\_____________________________________
'//
'//Modify cell...
ElseIf InStr(1, xArt, "cell(", vbTextCompare) Then
Call modArtQ(xArt)

If InStr(1, xArt, ",") = False Then MsgBox (Application.ActiveCell.Address): Exit Function '//no excerpt provided

'//Check for modifiers...
If InStr(1, xArt, ".") Then
If InStr(1, xArt, " .") Then xArtArr = Split(xArt, " .")
If InStr(1, xArt, ").") Then xArtArr = Split(xArt, ").")

Do Until X > UBound(xArtArr)

xArt = xArtArr(X): Call modArtP(xArt): Call modArtQ(xArt): xArtArr(X) = xArt

'//Extract cell...
If InStr(1, xArtArr(X), "cell", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "cell", vbNullString, , , vbTextCompare)
If InStr(1, xArtArr(0), "=") Then xArtArrCl = Split(xArtArr(0), "="): _
xArtArrCl = Split(xArtArrCl(1), ",") Else: _
xArtArrCl = Split(xArtArr(0), ",")
xArt = xArtArrCl(0): Call modArtP(xArt): xArtArrCl(0) = xArt
xArt = xArtArrCl(1): Call modArtP(xArt): xArtArrCl(1) = xArt
x1 = CInt(xArtArrCl(0)): y1 = CInt(xArtArrCl(1))
End If
'//Select cell...
If InStr(1, xArtArr(X), "sel", vbTextCompare) Then
Cells(x1, y1).Select
End If
'//Clean cell...
If InStr(1, xArtArr(X), "cln", vbTextCompare) Then
Cells(x1, y1).ClearContents
End If
'//Clear cell...
If InStr(1, xArtArr(X), "clr", vbTextCompare) Then
Cells(x1, y1).Clear
End If
'//Copy cell...
If InStr(1, xArtArr(X), "copy") Then
If InStr(1, xArtArr(X), "copy&") Then C = 1
If InStr(1, xArtArr(X), "copy&!") Then C = 2
If InStr(1, xArtArr(X), "copy&!!") Then C = 3

    xArtArr(X) = Replace(xArtArr(X), "copy", vbNullString, vbTextCompare)
    xArtArr(X) = Replace(xArtArr(X), "!", vbNullString)
    xArtArr(X) = Replace(xArtArr(X), "&", vbNullString)
    xArt = xArtArr(X): Call modArtP(xArt): xArtArr(X) = xArt
    
    ActiveCell.Copy
    
    If C = vbNullString Then ActiveCell.Copy '//just copy
     
    If C = 1 Then '//copy paste cell contents
        ActiveWorkbook.Worksheets(appBlk).Cells(xArtArr(X)).Activate
            ActiveCell.PasteSpecial
                End If
                
    If C = 2 Then '//copy paste clean contents
        xCell = ActiveCell.Address
            ActiveWorkbook.Worksheets(appBlk).Cells(xArtArr(X)).Activate
                ActiveCell.PasteSpecial
                    ActiveWorkbook.Worksheets(appBlk).Cells(xCell).ClearContents
                        End If
                        
    If C = 3 Then '//copy paste clear cell contents
        xCell = ActiveCell.Address
            ActiveWorkbook.Worksheets(appBlk).Cells(xArtArr(X)).Activate
                ActiveCell.PasteSpecial
                    ActiveWorkbook.Worksheets(appBlk).Cells(xCell).Clear
                        End If
                        
                                End If

'//Paste cell...
If InStr(1, xArtArr(X), "paste", vbTextCompare) Then
xArt = xArtArr(X): Call modArtP(xArt)
ActiveCell.PasteSpecial
End If
'//Set cell name...
If InStr(1, xArtArr(X), "name", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "name ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "name(", vbNullString, , , vbTextCompare)
xArt = xArtArr(X): Call modArtP(xArt)
'//no name entered (clear name)
If xArtArr(X) = vbNullString Then
xArtArr(X) = Cells(x1, y1).name.name
ActiveWorkbook.Names(xArtArr(X)).Delete
    Else
        Cells(x1, y1).name = xArtArr(X)
            End If
                End If
'//Set cell value...
If InStr(1, xArtArr(X), "value", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "value ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "value(", vbNullString, , , vbTextCompare)
Cells(x1, y1).Value = xArtArr(X)
End If
'//Set cell font color...
If InStr(1, xArtArr(X), "fcolor", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "fcolor ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "fcolor(", vbNullString, , , vbTextCompare)
HX = xArtArr(X)
Call basColor(HX) '//find color
HX = Replace(HX, ")", vbNullString)
HX = Replace(HX, "(", vbNullString)
xRGBArr = Split(HX, ",")
Cells(x1, y1).Font.Color = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
End If
'//Set cell font size...
If InStr(1, xArtArr(X), "fsize", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "fsize ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "fsize(", vbNullString, , , vbTextCompare)
Cells(x1, y1).Font.Size = xArtArr(X)
End If
'//Set cell font type...
If InStr(1, xArtArr(X), "ftype", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "ftype", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "ftype(", vbNullString, , , vbTextCompare)
Cells(x1, y1).Font.FontStyle = xArtArr(X)
End If
'//Set cell pattern...
If InStr(1, xArtArr(X), "pattern", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "pattern", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "pattern(", vbNullString, , , vbTextCompare)
PX = xArtArr(X)
Call basPattern(PX) '//find pattern
Cells(x1, y1).Interior.Pattern = PX
End If
'//Set cell border direction...
If InStr(1, xArtArr(X), "border", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "border ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "border(", vbNullString, , , vbTextCompare)
BX = xArtArr(X)
Call basBorder(BX) '//find border
Cells(x1, y1).BorderAround (BX)
End If
'//Set cell border type...
If InStr(1, xArtArr(X), "btype", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "border ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "border(", vbNullString, , , vbTextCompare)
SX = xArtArr(X)
Call basBorderStyle(SX) '//find border type
Cells(x1, y1).Borders.LineStyle = SX
End If
'//Set cell color...
If InStr(1, xArtArr(X), "bgcolor", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "bgcolor ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "bgcolor(", vbNullString, , , vbTextCompare)
HX = xArtArr(X)
Call basColor(HX) '//find color
HX = Replace(HX, ")", vbNullString)
HX = Replace(HX, "(", vbNullString)
xRGBArr = Split(HX, ",")
Cells(x1, y1).Interior.Color = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
End If
'//Read cell value into variable...
If InStr(1, xArtArr(X), "read", vbTextCompare) Then
If InStr(1, xArtArr(0), "=") Then
xArtArr = Split(xArtArr(0), "=")
xArtArr(0) = Trim(xArtArr(0))
xVar = Cells(x1, y1).Value
xArt = appEnv & ",#!" & xArtArr(0) & "=" & xVar & ",#!" & X & ",#!" & 1: Call kinExpand(xArt)
End If
    End If

X = X + 1
Loop
Exit Function
End If

Exit Function
'//#
'//
'//Modify range...
ElseIf InStr(1, xArt, "rng(", vbTextCompare) Then

'//Check for modifiers...
If InStr(1, xArt, ".") Then
If InStr(1, xArt, " .") Then xArtArr = Split(xArt, " .")
If InStr(1, xArt, ").") Then xArtArr = Split(xArt, ").")

Do Until X > UBound(xArtArr)

xArt = xArtArr(X): Call modArtP(xArt): Call modArtQ(xArt): xArtArr(X) = xArt

'//Extract range...
If InStr(1, xArtArr(X), "rng", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "rng", vbNullString, , , vbTextCompare)
If InStr(1, xArtArr(0), "=") Then xArtArrCl = Split(xArtArr(0), "="): _
xArt = xArtArrCl(1): Call modArtP(xArt): xArtArr(0) = xArt Else: _
xArt = xArtArr(X): Call modArtP(xArt): xArtArr(X) = xArt
End If
'//Select range...
If InStr(1, xArtArr(X), "sel", vbTextCompare) Then
Range(xArtArr(0)).Select
End If
'//Clean range...
If InStr(1, xArtArr(X), "cln", vbTextCompare) Then
Range(xArtArr(0)).ClearContents
End If
'//Clear range...
If InStr(1, xArtArr(X), "clr", vbTextCompare) Then
Range(xArtArr(0)).Clear
End If
'//Copy range...
If InStr(1, xArtArr(X), "copy") Then
If InStr(1, xArtArr(X), "copy&") Then C = 1
If InStr(1, xArtArr(X), "copy&!") Then C = 2
If InStr(1, xArtArr(X), "copy&!!") Then C = 3

    xArtArr(X) = Replace(xArtArr(X), "copy", vbNullString, vbTextCompare)
    xArtArr(X) = Replace(xArtArr(X), "!", vbNullString)
    xArtArr(X) = Replace(xArtArr(X), "&", vbNullString)
    xArt = xArtArr(X): Call modArtP(xArt): xArtArr(X) = xArt
    
    ActiveCell.Copy
    
    If C = vbNullString Then ActiveCell.Copy '//just copy
     
    If C = 1 Then '//copy paste range contents
        ActiveWorkbook.Worksheets(appBlk).Range(xArtArr(X)).Activate
            ActiveCell.PasteSpecial
                End If
                
    If C = 2 Then '//copy paste clean contents
        xCell = ActiveCell.Address
            ActiveWorkbook.Worksheets(appBlk).Range(xArtArr(X)).Activate
                ActiveCell.PasteSpecial
                    ActiveWorkbook.Worksheets(appBlk).Range(xCell).ClearContents
                        End If
                        
    If C = 3 Then '//copy paste clear range contents
        xCell = ActiveCell.Address
            ActiveWorkbook.Worksheets(appBlk).Range(xArtArr(X)).Activate
                ActiveCell.PasteSpecial
                    ActiveWorkbook.Worksheets(appBlk).Range(xCell).Clear
                        End If
                        
                                End If
            

'//Paste range...
If InStr(1, xArtArr(X), "paste", vbTextCompare) Then
ActiveCell.PasteSpecial
End If
'//Set range name...
If InStr(1, xArtArr(X), "name", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "name ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "name(", vbNullString, , , vbTextCompare)
'//no name entered (clear name)
If xArtArr(X) = vbNullString Then
xArtArr(X) = Range(xArtArr(0)).name.name
ActiveWorkbook.Names(xArtArr(X)).Delete
    Else
        Range(xArtArr(0)).name = xArtArr(X)
            End If
                End If
'//Set range value...
If InStr(1, xArtArr(X), "value", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "value ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "value(", vbNullString, , , vbTextCompare)
Range(xArtArr(0)).Value = xArtArr(X)
End If
'//Set range font color...
If InStr(1, xArtArr(X), "fcolor", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "fcolor ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "fcolor(", vbNullString, , , vbTextCompare)
HX = xArtArr(X)
Call basColor(HX) '//find color
HX = Replace(HX, ")", vbNullString)
HX = Replace(HX, "(", vbNullString)
xRGBArr = Split(HX, ",")
Range(xArtArr(0)).Font.Color = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
End If
'//Set range font size...
If InStr(1, xArtArr(X), "fsize", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "fsize ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "fsize(", vbNullString, , , vbTextCompare)
Range(xArtArr(0)).Font.Size = xArtArr(X)
End If
'//Set range font type...
If InStr(1, xArtArr(X), "ftype", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "ftype ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "ftype(", vbNullString, , , vbTextCompare)
Range(xArtArr(0)).Font.FontStyle = xArtArr(X)
End If
'//Set range pattern...
If InStr(1, xArtArr(X), "pattern", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "pattern ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "pattern(", vbNullString, , , vbTextCompare)
PX = xArtArr(X)
Call basPattern(PX) '//find pattern
Range(xArtArr(0)).Interior.Pattern = PX
End If
'//Set range border direction...
If InStr(1, xArtArr(X), "border", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "border ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "border(", vbNullString, , , vbTextCompare)
BX = xArtArr(X)
Call basBorder(BX) '//find border
Range(xArtArr(0)).BorderAround (BX)
End If
'//Set range border type...
If InStr(1, xArtArr(X), "btype(", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "btype ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "btype(", vbNullString, , , vbTextCompare)
SX = xArtArr(X)
Call basBorderStyle(SX) '//find border type
Range(xArtArr(0)).Borders.LineStyle = SX
End If
'//Set range color...
If InStr(1, xArtArr(X), "bgcolor", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "bgcolor ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "bgcolor(", vbNullString, , , vbTextCompare)
HX = xArtArr(X)
Call basColor(HX) '//find color
HX = Replace(HX, ")", vbNullString)
HX = Replace(HX, "(", vbNullString)
xRGBArr = Split(HX, ",")
Range(xArtArr(0)).Interior.Color = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
End If
'//Read range value into variable...
If InStr(1, xArtArr(X), "read", vbTextCompare) Then
If xArtArrCl(0) <> Empty Then
xVar = Range(xArtArr(0)).Value
xArt = appEnv & ",#!" & xArtArrCl(0) & "=" & xVar & ",#!" & X & ",#!" & 1: Call kinExpand(xArt)
End If
    End If

X = X + 1
Loop
Exit Function
End If

'//no modifier
xArt = Replace(xArt, "rng(", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)
'//Activate range...
Range(xArt).Activate
Exit Function
'//#
'//
'//Select & modify cell/range...
ElseIf InStr(1, xArt, "sel(", vbTextCompare) Then

'//Check for modifiers...
If InStr(1, xArt, ".") Then
If InStr(1, xArt, " .") Then xArtArr = Split(xArt, " .")
If InStr(1, xArt, ").") Then xArtArr = Split(xArt, ").")

Do Until X > UBound(xArtArr)

xArt = xArtArr(X): Call modArtP(xArt): Call modArtQ(xArt): xArtArr(X) = xArt

'//Select cell...
If InStr(1, xArtArr(X), "sel", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "sel", vbNullString, , , vbTextCompare)
If InStr(1, xArtArr(0), "=") Then xArtArrCl = Split(xArtArr(0), "="): _
xArt = xArtArrCl(1): Call modArtP(xArt): xArtArr(0) = xArt Else: _
xArt = xArtArr(X): Call modArtP(xArt): xArtArr(X) = xArt
Range(xArtArr(X)).Select
End If
'//Clean cell...
If InStr(1, xArtArr(X), "cln", vbTextCompare) Then
Range(xArtArr(0)).ClearContents
End If
'//Clear cell...
If InStr(1, xArtArr(X), "clr", vbTextCompare) Then
Range(xArtArr(0)).Clear
End If
'//Copy cell...
If InStr(1, xArtArr(X), "copy", vbTextCompare) Then
If InStr(1, xArtArr(X), "copy&", vbTextCompare) Then C = 1
If InStr(1, xArtArr(X), "copy&!", vbTextCompare) Then C = 2
If InStr(1, xArtArr(X), "copy&!!", vbTextCompare) Then C = 3

    xArtArr(X) = Replace(xArtArr(X), "copy", vbNullString, , , vbTextCompare)
    xArtArr(X) = Replace(xArtArr(X), "!", vbNullString)
    xArtArr(X) = Replace(xArtArr(X), "&", vbNullString)
    xArt = xArtArr(X): Call modArtP(xArt): xArtArr(X) = xArt
    
    ActiveCell.Copy
    
    If C = vbNullString Then ActiveCell.Copy
     
    If C = 1 Then '//copy paste
        ActiveWorkbook.Worksheets(appBlk).Range(xArtArr(X)).Activate
            ActiveCell.PasteSpecial
                End If
                
    If C = 2 Then '//copy paste clear contents
        xCell = ActiveCell.Address
            ActiveWorkbook.Worksheets(appBlk).Range(xArtArr(X)).Activate
                ActiveCell.PasteSpecial
                    ActiveWorkbook.Worksheets(appBlk).Range(xCell).ClearContents
                        End If
                        
        If C = 3 Then '//copy paste clear cell
        xCell = ActiveCell.Address
            ActiveWorkbook.Worksheets(appBlk).Range(xArtArr(X)).Activate
                ActiveCell.PasteSpecial
                    ActiveWorkbook.Worksheets(appBlk).Range(xCell).Clear
                        End If
                        
                                End If
                                
'//Paste cell...
If InStr(1, xArtArr(X), "paste", vbTextCompare) Then
ActiveCell.PasteSpecial
End If
'//Set cell name...
If InStr(1, xArtArr(X), "name", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "name ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "name(", vbNullString, , , vbTextCompare)
'//no name entered (clear name)
If xArtArr(X) = vbNullString Then
xArtArr(X) = Range(xArtArr(0)).name.name
ActiveWorkbook.Names(xArtArr(X)).Delete
    Else
        Range(xArtArr(0)).name = xArtArr(X)
            End If
                End If
'//Set cell value...
If InStr(1, xArtArr(X), "value", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "value ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "value(", vbNullString, , , vbTextCompare)
Range(xArtArr(0)).Value = xArtArr(X)
End If
'//Set cell font color...
If InStr(1, xArtArr(X), "fcolor", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "fcolor ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "fcolor(", vbNullString, , , vbTextCompare)
HX = xArtArr(X)
Call basColor(HX) '//find color
HX = Replace(HX, ")", vbNullString)
HX = Replace(HX, "(", vbNullString)
xRGBArr = Split(HX, ",")
Range(xArtArr(0)).Font.Color = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
End If
'//Set cell font size...
If InStr(1, xArtArr(X), "fsize", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "fsize ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "fsize(", vbNullString, , , vbTextCompare)
Range(xArtArr(0)).Font.Size = xArtArr(X)
End If
'//Set cell font type...
If InStr(1, xArtArr(X), "ftype", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "ftype ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "ftype(", vbNullString, , , vbTextCompare)
Range(xArtArr(0)).Font.FontStyle = xArtArr(X)
End If
'//Set cell pattern...
If InStr(1, xArtArr(X), "pattern", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "pattern ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "pattern(", vbNullString, , , vbTextCompare)
PX = xArtArr(X)
Call basPattern(PX) '//find pattern
Range(xArtArr(0)).Interior.Pattern = PX
End If
'//Set cell border direction...
If InStr(1, xArtArr(X), "border", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "border ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "border(", vbNullString, , , vbTextCompare)
BX = xArtArr(X)
Call basBorder(BX) '//find border
Range(xArtArr(0)).BorderAround = BX
End If
'//Set cell border type...
If InStr(1, xArtArr(X), "btype", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "btype ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "btype(", vbNullString, , , vbTextCompare)
SX = xArtArr(X)
Call basBorderStyle(SX) '//find border type
Range(xArtArr(0)).Borders.LineStyle = SX
End If
'//Set cell color...
If InStr(1, xArtArr(X), "bgcolor", vbTextCompare) Then
xArtArr(X) = Replace(xArtArr(X), "bgcolor ", vbNullString, , , vbTextCompare)
xArtArr(X) = Replace(xArtArr(X), "bgcolor(", vbNullString, , , vbTextCompare)
HX = xArtArr(X)
Call basColor(HX) '//find color
HX = Replace(HX, ")", vbNullString)
HX = Replace(HX, "(", vbNullString)
xRGBArr = Split(HX, ",")
Range(xArtArr(0)).Interior.Color = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
End If
'//Read cell value into variable...
If InStr(1, xArtArr(X), "read", vbTextCompare) Then
If xArtArrCl(0) <> Empty Then
xVar = Range(xArtArr(0)).Value
xArt = appEnv & ",#!" & xArtArrCl(0) & "=" & xVar & ",#!" & X & ",#!" & 1: Call kinExpand(xArt)
End If
    End If
    
X = X + 1
Loop
Exit Function
End If
'//no modifier
xArt = Replace(xArt, "sel(", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)
'//Activate cell...
Range(xArt).Select
Exit Function
'//#
'//
'/\_____________________________________
'//
'//        WORKBOOK ARTICLES
'/\_____________________________________
'//
'//Modify Workbook...
ElseIf InStr(1, xArt, "wb(", vbTextCompare) Then
xArt = Replace(xArt, "wb(", vbNullString, , , vbTextCompare)

If InStr(1, xArt, ".active", vbTextCompare) Then If InStr(1, xArt, ").active", vbTextCompare) = False Then ActiveWorkbook.Activate  '//activate current workbook
If InStr(1, xArt, ").active", vbTextCompare) Then '//activate specific workbook
xArt = Replace(xArt, ".active", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)
Workbooks(xArt).Activate
Range("MAS2").name = "xlasEnvironment": Range("xlasEnvironment").Value = appEnv '//link environment to workbook
Range("MAS3").name = "xlasBlock": Range("xlasBlock").Value = appBlk '//link block to workbook
Exit Function
End If

If InStr(1, xArt, ").open", vbTextCompare) Then '//open workbook

xArt = Replace(xArt, ".open", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)
Workbooks.Open (xArt)
Range("MAS2").name = "xlasEnvironment": Range("xlasEnvironment").Value = appEnv '//link environment to workbook
Range("MAS3").name = "xlasBlock": Range("xlasBlock").Value = appBlk '//link block to workbook
Workbooks(appEnv).Worksheets(appBlk).Activate
Exit Function
End If

If InStr(1, xArt, ").hd", vbTextCompare) Then ActiveWorkbook.Application.Visible = False   '//hide active workbook
If InStr(1, xArt, ").sh", vbTextCompare) Then ActiveWorkbook.Application.Visible = True  '//show active workbook

If InStr(1, xArt, ").close", vbTextCompare) Then _
xArt = Replace(xArt, ".close", vbNullString, , , vbTextCompare): _
Call modArtP(xArt): Call modArtQ(xArt): _
If xArt = vbNullString Then ActiveWorkbook.Close: Exit Function Else _
Workbooks(xArt).Close: Exit Function '//close workbook

'If InStr(1, xArt, ").export", vbTextCompare) Then ActiveWorkbook.ExportAsFixedFormat = vbnullstring '//export file
If InStr(1, xArt, ").nwin", vbTextCompare) Then ActiveWorkbook.NewWindow: Exit Function '//create new window

If InStr(1, xArt, ").save", vbTextCompare) And InStr(1, xArt, ").saveas", vbTextCompare) = False Then _
xArt = Replace(xArt, ".save", vbNullString, , , vbTextCompare): _
Call modArtP(xArt): Call modArtQ(xArt): _
If xArt = vbNullString Then ActiveWorkbook.Save: Exit Function Else _
Workbooks(xArt).Save: Exit Function '//save workbook

If InStr(1, xArt, ").saveas", vbTextCompare) Then '//save as workbook

Call modArtP(xArt): Call modArtQ(xArt)

xArt = Replace(xArt, ".saveas", vbNullString, , , vbTextCompare)
xArtArr = Split(xArt, ",")
If UBound(xArtArr) = 1 Then
EX = xArtArr(1): Call basSaveFormat(EX)
If EX <> "(*Err)" Then
Range("MAS2").name = "xlasEnvironment": Range("xlasEnvironment").Value = appEnv '//link environment to workbook
Range("MAS3").name = "xlasBlock": Range("xlasBlock").Value = appBlk '//link block to workbook
ActiveWorkbook.SaveAs FileName:=xArtArr(0), FileFormat:=xExt
End If
    End If
        Exit Function
            End If
    
If InStr(1, xArt, ").name", vbTextCompare) Then MsgBox (ActiveWorkbook.name), 0, "": Exit Function '//get name of workbook
If InStr(1, xArt, ").path", vbTextCompare) Then MsgBox (ActiveWorkbook.Path), 0, "": Exit Function '//get path of workbook

If InStr(1, xArt, ").add", vbTextCompare) Then '//add worksheet

Call modArtP(xArt): Call modArtQ(xArt)

If InStr(1, xArt, ").addafter", vbTextCompare) Then C = 1: xArt = Replace(xArt, ".addafter", vbNullString, , , vbTextCompare) '//add after worksheet
If InStr(1, xArt, ").addbefore", vbTextCompare) Then C = 2: xArt = Replace(xArt, ".addbefore", vbNullString, , , vbTextCompare) '//add before worksheet

xArt = Replace(xArt, ".add", vbNullString, , , vbTextCompare)
If xArt = vbNullString Then '//default add no parameters
xArt = "Sheet" & ActiveWorkbook.Worksheets.Count + 1
Worksheets.Add.name = xArt
Exit Function
End If

If InStr(1, xArt, ",") = False Then
'//single parameter... (set count w/ default worksheet name & place before or after first/last sheet)
If C = 1 Then Worksheets.Add After:=Worksheets(Worksheets.Count), Count:=Int(xArt): Exit Function
If C = 2 Then Worksheets.Add Before:=Worksheets(Worksheets.Count), Count:=Int(xArt): Exit Function
    Else
xArtArr = Split(xArt, ",")
If UBound(xArtArr) = 1 Then
'//two parameters... (set add worksheet name & place before or after assigned sheet
If C = 1 Then Worksheets.Add(After:=Worksheets(xArtArr(0))).name = xArtArr(1): Exit Function
If C = 2 Then Worksheets.Add(Before:=Worksheets(xArtArr(0))).name = xArtArr(1): Exit Function
ElseIf UBound(xArtArr) = 2 Then
'//three parameters... (set add worksheet name & place before or after assigned  sheet w/ count)
If C = 1 Then Worksheets.Add(After:=Worksheets(xArtArr(0)), Count:=Int(xArtArr(2))).name = xArtArr(1): Exit Function
If C = 2 Then Worksheets.Add(Before:=Worksheets(xArtArr(0)), Count:=Int(xArtArr(2))).name = xArtArr(1): Exit Function
                    End If
                        End If
        
If InStr(1, xArt, ").new", vbTextCompare) Then '//add new workbook

Call modArtP(xArt): Call modArtQ(xArt)

xArt = Replace(xArt, ".new", vbNullString, , , vbTextCompare)
xArtArr = Split(xArt, ",")
If UBound(xArtArr) = 1 Then
EX = xArtArr(1): Call basSaveFormat(EX)
If EX <> "(*Err)" Then
Application.Workbooks.Add.SaveAs FileName:=xArtArr(0), FileFormat:=xExt
Workbooks(appEnv).Worksheets(appBlk).Activate
End If
    End If
        Exit Function
            End If
                End If


'//Run workbook module...
If InStr(1, xArt, ").run", vbTextCompare) Then

xArt = Replace(xArt, ".run", vbNullString, , , vbTextCompare)
Call modArtD(xArt): Call modArtQ(xArt)

If InStr(1, xArt, ",") Then xArtArr = Split(xArt, ",") Else GoTo wbRunNoArg

xArtArr(0) = Trim(xArtArr(0)): xArtArr(1) = Trim(xArtArr(1))
xWb = xArtArr(0) '//extract workbook
xMod = xArtArr(1) '//extract module

X = 2
Do Until X > UBound(xArtArr) '//extract argument(s)
xArt = xArtArr(X) & ",": xArtCl = xArtCl & xArt
X = X + 1
Loop

xArt = xArtCl
If Right(xArt, Len(xArt) - Len(xArt) + 1) = "," Then xArt = Left(xArt, Len(xArt) - 1)

X = Application.Run("'" & xWb & "'!" & xMod, (xArt))
Exit Function
    
'//no arguments provided
wbRunNoArg:
X = Application.Run(xArt)
Exit Function
End If

'//Delete workbook cell name...
If InStr(1, xArt, ").delname", vbTextCompare) Then

xArt = Replace(xArt, ".delname", vbNullString, , , vbTextCompare)
Call modArtD(xArt): Call modArtP(xArt): Call modArtQ(xArt)

If InStr(1, xArt, ",") Then xArtArr = Split(xArt, ",") Else GoTo ErrMsg

xArtArr(0) = Trim(xArtArr(0)): xArtArr(1) = Trim(xArtArr(1))

Workbooks(xArtArr(0)).Names(xArtArr(1)).Delete

Exit Function

End If

'//excerpt not supplied
MsgBox (ActiveWorkbook.name)
Exit Function
'//#
'//
'/\_____________________________________
'//
'//       FILE/DIRECTORY ARTICLES
'/\_____________________________________
'//
'//
'//Modify file
ElseIf InStr(1, xArt, "fil(", vbTextCompare) Then

If InStr(1, xArt, ".fil", vbTextCompare) Then
xArt = Replace(xArt, "fil(", vbNullString, , , vbTextCompare)

If InStr(1, xArt, "del.", vbTextCompare) Then E = 1: xArt = Replace(xArt, "del.", vbNullString, , , vbTextCompare)
If InStr(1, xArt, "mk.", vbTextCompare) Then E = 2: xArt = Replace(xArt, "mk.", vbNullString, , , vbTextCompare)
If Left(xArt, 1) = " " Then xArt = Right(xArt, Len(xArt) - 1)
If E = 0 Then errLvl = 1: GoTo ErrMsg

Call modArtP(xArt): Call modArtQ(xArt)
Set oFSO = CreateObject("Scripting.FileSystemObject")
If E = 1 Then: Set oFSO = CreateObject("Scripting.FileSystemObject"): oFSO.DeleteFile (xArt): Set oFSO = Nothing: Exit Function '//delete file
If E = 2 Then: _

If InStr(1, xArt, ",") Then
E = E & "1"
xArtArr = Split(xArt, ",")
xArt = xArtArr(0): Call modArtP(xArt): Call modArtQ(xArt): xArtArr(0) = xArt
xArt = xArtArr(1): Call modArtP(xArt): xArtArr(1) = xArt
xArtArr(1) = LTrim(xArtArr(1)) '//remove leading space
End If

If E = 2 Then Open (xArt) For Output As #1: Print #1, vbNullString: Close #1: Exit Function
If E = 21 Then Open (xArtArr(0)) For Output As #1: Print #1, xArtArr(1): Close #1: Exit Function
Exit Function

Else
errLvl = 1: GoTo ErrMsg
End If

'//Modify folder
ElseIf InStr(1, xArt, "dir(", vbTextCompare) Then

If InStr(1, xArt, ".dir", vbTextCompare) Then
xArt = Replace(xArt, "dir(", vbNullString, , , vbTextCompare)

If InStr(1, xArt, "del.", vbTextCompare) Then E = 1: xArt = Replace(xArt, "del.", vbNullString, , , vbTextCompare)
If InStr(1, xArt, "mk.", vbTextCompare) Then E = 2: xArt = Replace(xArt, "mk.", vbNullString, , , vbTextCompare)
If Left(xArt, 1) = " " Then xArt = Right(xArt, Len(xArt) - 1)
If E = 0 Then errLvl = 1: GoTo ErrMsg

Call modArtP(xArt): Call modArtQ(xArt)
If E = 1 Then: Set oFSO = CreateObject("Scripting.FileSystemObject"): oFSO.DeleteFolder (xArt): Set oFSO = Nothing: Exit Function '//create file
If E = 2 Then: _

If InStr(1, xArt, ",") Then
E = E & "1"
xArtArr = Split(xArt, ",")
xArt = xArtArr(0): Call modArtP(xArt): Call modArtQ(xArt): xArtArr(0) = xArt
xArt = xArtArr(1): Call modArtP(xArt): xArtArr(1) = xArt
xArtArr(1) = LTrim(xArtArr(1)) '//remove leading space
End If

If E = 2 Then MkDir (xArt): Exit Function
If E = 21 Then MkDir (xArtArr(0)): MkDir (xArtArr(0) & "/" & xArtArr(1)): Exit Function
Exit Function

Else
errLvl = 1: GoTo ErrMsg
End If


'//Delete empty directory
ElseIf InStr(1, xArt, "dfldr(", vbTextCompare) Then

xArt = Replace(xArt, "dfldr", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)

If Dir(xArt, vbDirectory) <> "" Then RmDir (xArt): Exit Function

'//Delete file
ElseIf InStr(1, xArt, "dfile(", vbTextCompare) Then

xArt = Replace(xArt, "dfile", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)

If Dir(xArt) <> "" Then Kill (xArt): Exit Function

'//Create empty directory
ElseIf InStr(1, xArt, "mfldr(", vbTextCompare) Then

xArt = Replace(xArt, "mfldr", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)

MkDir (xArt): Exit Function

'//Create file
ElseIf InStr(1, xArt, "mfile(", vbTextCompare) Then

xArt = Replace(xArt, "mfile", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)

Open (xArt) For Output As #1: Print #1, vbNullString: Close #1: Exit Function

'//Move file or folder
ElseIf InStr(1, xArt, "move(", vbTextCompare) Then
xArt = Replace(xArt, "move(", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)

If InStr(1, xArt, ",") Then

xArtArr = Split(xArt, ",")

xArt = "move " & xArtArr(0) & " " & xArtArr(1)
    
sysShell = Shell("cmd.exe /s /c" & xArt, 0)
sysShell = vbNullString
End If
Exit Function

'//Rename file or folder
ElseIf InStr(1, xArt, "ren(", vbTextCompare) Then
xArt = Replace(xArt, "ren(", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)

If InStr(1, xArt, ",") Then

xArtArr = Split(xArt, ",")

If UBound(xArtArr) = 2 Then GoTo renAll
If InStr(1, xArt, "app.r") Then GoTo renVBA

'//default
xArt = "ren " & xArtArr(0) & " " & xArtArr(1): sysShell = Shell("cmd.exe /s /c" & xArt, 0): sysShell = vbNullString
End If
Exit Function

renVBA:
Name xArtArr(0) As xArtArr(1)
Exit Function

renAll:
Dim xDate, xName, xTime As String
Dim xNum As Long
xNum = 1

Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oSubFldr = oFSO.GetFolder(xArtArr(0))

xExtArr = Split(xArtArr(1), "."): EX = xExtArr(1): xName = xExtArr(0)

If InStr(1, xArtArr(2), "-num", vbTextCompare) Then xNum = 1: GoTo renAllNum
If InStr(1, xArtArr(2), "-datenum", vbTextCompare) And InStr(1, xArtArr(2), "mtime", vbTextCompare) = False Then xDate = Date: xDate = Replace(xDate, "/", "-"): xNum = 1: GoTo renAllDateNum
If InStr(1, xArtArr(2), "-datenumtime", vbTextCompare) Then xDate = Date: xDate = Replace(xDate, "/", "-"): xNum = 1: xTime = Time: xTime = Replace(xTime, ":", vbNullString): xTime = Replace(xTime, " ", vbNullString): GoTo renAllDateNumTime

renAllNum:
For Each oFile In oSubFldr.Files
xArt = "ren " & oFile.Path & " " & xName & "_" & xNum & "." & EX
sysShell = Shell("cmd.exe /s /c" & xArt, 0)
sysShell = vbNullString
xNum = xNum + 1
Next
Set oFSO = Nothing
Set oFile = Nothing
Set oSubFldr = Nothing
Exit Function

renAllDateNum:
For Each oFile In oSubFldr.Files
xArt = "ren " & oFile.Path & " " & xName & "_" & xDate & "_" & xNum & "." & EX
sysShell = Shell("cmd.exe /s /c" & xArt, 0)
sysShell = vbNullString
xNum = xNum + 1
Next
Set oFSO = Nothing
Set oFile = Nothing
Set oSubFldr = Nothing
Exit Function

renAllDateNumTime:
For Each oFile In oSubFldr.Files
xNum = xNum + xArt = "ren " & oFile.Path & " " & xName & "_" & xDate & "_" & xNum & "_" & xTime & "." & EX
sysShell = Shell("cmd.exe /s /c" & xArt, 0)
sysShell = vbNullString
xNum = xNum + 1
Next
Set oFSO = Nothing
Set oFile = Nothing
Set oSubFldr = Nothing
Exit Function

'//Read file
ElseIf InStr(1, xArt, "read(", vbTextCompare) Then
xArt = Replace(xArt, "read", vbNullString, , , vbTextCompare)

'//switches
If InStr(1, xArt, "-all", vbTextCompare) Then S = 1: xArt = Replace(xArt, "-all", vbNullString, , , vbTextCompare): GoTo SetRead
If InStr(1, xArt, "-nlall", vbTextCompare) Then S = 2: xArt = Replace(xArt, "-nlall", vbNullString, , , vbTextCompare): GoTo SetRead
If InStr(1, xArt, "-string", vbTextCompare) Then S = 3: xArt = Replace(xArt, "-string", vbNullString, , , vbTextCompare): GoTo SetRead
If InStr(1, xArt, "-xstring", vbTextCompare) Then S = 4: xArt = Replace(xArt, "-xstring", vbNullString, , , vbTextCompare): GoTo SetRead
If InStr(1, xArt, "-nlstring", vbTextCompare) Then S = 5: xArt = Replace(xArt, "-nlstring", vbNullString, , , vbTextCompare): GoTo SetRead
If InStr(1, xArt, "-line", vbTextCompare) Then S = 6: xArt = Replace(xArt, "-line", vbNullString, , , vbTextCompare): GoTo SetRead
If InStr(1, xArt, "-xline", vbTextCompare) Then S = 7: xArt = Replace(xArt, "-xline", vbNullString, , , vbTextCompare): GoTo SetRead
If InStr(1, xArt, "-nlline", vbTextCompare) Then S = 8: xArt = Replace(xArt, "-nlline", vbNullString, , , vbTextCompare): GoTo SetRead

SetRead:
'//enhancers(s):
If InStr(1, xArt, "count.") Then E = 1: xArt = Replace(xArt, "count.", vbNullString, , , vbTextCompare)

xArtArr = Split(xArt, "=")
xArt = xArtArr(1): Call modArtP(xArt): Call modArtQ(xArt): xArt = Trim(xArt): xArtArr(1) = xArt
xArt = xArtArr(0)

'//read for all
If S = 1 Then
Open xArtArr(1) For Input As #1: Do Until EOF(1): Line Input #1, xArtCl: xVar = xVar & xArtCl: Loop: Close #1
GoTo EndRead
End If

'//read for newline all
If S = 2 Then
Open xArtArr(1) For Input As #1: Do Until EOF(1): Line Input #1, xArtCl: xVar = xVar & xArtCl & vbNewLine: Loop: Close #1
GoTo EndRead
End If

'//read for string
If S = 3 Then
xArtArr = Split(xArtArr(1), ","): xArtArr(1) = Trim(xArtArr(1))
Open xArtArr(1) For Input As #1
Do Until EOF(1): Line Input #1, xArtCl
If InStr(1, xArtCl, xArtArr(0)) Then xVar = xArtCl: Close #1: GoTo EndRead
Loop: Close #1
GoTo EndRead
End If

'//read for all string
If S = 4 Then
xArtArr = Split(xArtArr(1), ","): xArtArr(1) = Trim(xArtArr(1))
Open xArtArr(1) For Input As #1
Do Until EOF(1): Line Input #1, xArtCl
If InStr(1, xArtCl, xArtArr(0)) Then xVar = xVar & xArtCl
Loop: Close #1
GoTo EndRead
End If

'//read for newline string
If S = 5 Then
xArtArr = Split(xArtArr(1), ","): xArtArr(1) = Trim(xArtArr(1))
Open xArtArr(1) For Input As #1
Do Until EOF(1): Line Input #1, xArtCl
If InStr(1, xArtCl, xArtArr(0)) Then xVar = xVar & xArtCl & vbNewLine
Loop: Close #1
GoTo EndRead
End If

'//read for line
If S = 6 Then
xArtArr = Split(xArtArr(1), ","): xArtArr(1) = Trim(xArtArr(1))
Open xArtArr(1) For Input As #1
For X = 1 To xArtArr(0)
Line Input #1, xArtCl
Next: Close #1: xVar = xArtCl
GoTo EndRead
End If

'//read for all line
If S = 7 Then
xArtArr = Split(xArtArr(1), ","): xArtArr(1) = Trim(xArtArr(1))
Open xArtArr(1) For Input As #1
For X = 1 To xArtArr(0)
Line Input #1, xArtCl
xVar = xVar & xArtCl
Next: Close #1
GoTo EndRead
End If

'//read for newline line
If S = 8 Then
xArtArr = Split(xArtArr(1), ","): xArtArr(1) = Trim(xArtArr(1))
Open xArtArr(1) For Input As #1
For X = 1 To xArtArr(0)
Line Input #1, xArtCl
xVar = xVar & xArtCl & vbNewLine
Next: Close #1
GoTo EndRead
End If

EndRead:
'//count
If E = 1 Then xVar = Len(xVar)

xArt = appEnv & ",#!" & xArt & "=" & xVar & ",#!" & X & ",#!" & 1: Call kinExpand(xArt)

Exit Function
'//#
'//
'/\_____________________________________
'//
'//          HALT ARTICLES
'/\_____________________________________
'//
'//Pause...
ElseIf InStr(1, xArt, "wait(", vbTextCompare) Then
xArt = Replace(xArt, "wait(", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)

If InStr(1, xArt, "ms", vbTextCompare) Then C = C & "1" '//millisecond
If InStr(1, xArt, "s", vbTextCompare) Then C = C & "2" '//second
If InStr(1, xArt, "m", vbTextCompare) Then C = C & "3" '//minute
If InStr(1, xArt, "h", vbTextCompare) Then C = C & "4" '//hour

If C <> 0 Then

Dim xTimeArr(3) As String
Dim xMil, xSec, xMin, xHr As String
Dim AppWait As Variant

If InStr(1, xArt, "ms", vbTextCompare) Then xMilArr = Split(xArt, "ms", , vbTextCompare): xTimeArr(0) = xMilArr(0): xMil = "T"
If InStr(1, xArt, "s", vbTextCompare) Then xSecArr = Split(xArt, "s", , vbTextCompare): xTimeArr(1) = xSecArr(0): xSec = "T"
If InStr(1, xArt, "m", vbTextCompare) Then xMinArr = Split(xArt, "m", , vbTextCompare): xTimeArr(2) = xMinArr(0): xMin = "T"
If InStr(1, xArt, "h", vbTextCompare) Then xHrArr = Split(xArt, "h", , vbTextCompare): xTimeArr(3) = xHrArr(0): xHr = "T"

'//set millisecond
If xMil = "T" Then
xArt = xTimeArr(0)
Call findChar(xArt): If xArt = "(*Err)" Then GoTo ErrMsg
xArt = -1 * (xArt * -0.00000001)
Application.Wait (Now + xArt)
Exit Function
End If
        
'//set second
If xSec = "T" Then
xArt = xTimeArr(1)
Call findChar(xArt): If xArt = "(*Err)" Then GoTo ErrMsg
If Len(xTimeArr(1)) < 2 Then
xTimeArr(1) = "0" & xTimeArr(1): xSec = xTimeArr(1)
Else: xSec = xTimeArr(1)
End If
    Else: xSec = "00"
        End If
        
'//set minute
If xMin = "T" Then
xArt = xTimeArr(2)
Call findChar(xArt): If xArt = "(*Err)" Then GoTo ErrMsg
If Len(xTimeArr(2)) < 2 Then
xTimeArr(2) = "0" & xTimeArr(2): xMin = xTimeArr(2)
Else: xMin = xTimeArr(2)
End If
    Else: xMin = "00"
        End If
        
'//set hour
If xHr = "T" Then
xArt = xTimeArr(3)
Call findChar(xArt): If xArt = "(*Err)" Then GoTo ErrMsg
If Len(xTimeArr(3)) < 2 Then
xTimeArr(3) = "0" & xTimeArr(3): xHr = xTimeArr(3)
Else: xHr = xTimeArr(3)
End If
    Else: xHr = "00"
        End If

    AppWait = TimeSerial(xHr, xMin, xSec)
    Application.Wait Now + TimeValue(AppWait)
Else
    '//00:00:00 format
    If InStr(1, xArt, ":") Then
    xArtArr = Split(xArt, ":")
    AppWait = TimeSerial(xArtArr(0), xArtArr(1), xArtArr(2))
    Application.Wait Now + TimeValue(AppWait)
        Else
            GoTo ErrMsg
                End If
                    End If
Exit Function

'//Waste time...
ElseIf InStr(1, xArt, "wastetime(", vbTextCompare) Then
xArt = Replace(xArt, "wastetime(", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)
For TX = 0 To xArt * 360: T = 1: Call basWasteTime(T): Next
Exit Function
'//#
'//
'/\_____________________________________
'//
'//         INPUT-HOST ARTICLES
'/\_____________________________________

ElseIf InStr(1, xArt, "input(", vbTextCompare) Then

   xArt = Replace(xArt, "input(", vbNullString, , , vbTextCompare)
   Call modArtP(xArt)
   
   xArtArr = Split(xArt, "="): xArt = xArtArr(0)
   xArtArr = Split(xArtArr(1), ",")
    
   If UBound(xArtArr) = 1 Then xVar = InputBox(xArtArr(0), xArtArr(1))
   If UBound(xArtArr) = 2 Then xVar = InputBox(xArtArr(0), xArtArr(1), xArtArr(2))

   xArt = appEnv & ",#!" & xArt & "=" & xVar & ",#!" & X & ",#!" & 1: Call kinExpand(xArt)
   
Exit Function
'//#
'//
'/\_____________________________________
'//
'//     OUTPUT-HOST ARTICLES
'/\_____________________________________
'//
'//Output w/ CMD
ElseIf InStr(1, xArt, "echo(", vbTextCompare) Then

    If InStr(1, xArt, "o(0)", vbTextCompare) Then S = 1: xArt = Replace(xArt, "echo(0)", vbNullString): GoTo setEcho
    If InStr(1, xArt, "o(1)", vbTextCompare) Then S = 2: xArt = Replace(xArt, "echo(1)", vbNullString): GoTo setEcho
    If InStr(1, xArt, "o(2)", vbTextCompare) Then S = 3: xArt = Replace(xArt, "echo(2)", vbNullString): GoTo setEcho
    If InStr(1, xArt, "o(3)", vbTextCompare) Then S = 4: xArt = Replace(xArt, "echo(3)", vbNullString): GoTo setEcho
    If InStr(1, xArt, "o(4)", vbTextCompare) Then S = 5: xArt = Replace(xArt, "echo(4)", vbNullString): GoTo setEcho
    If InStr(1, xArt, "o(5)", vbTextCompare) Then S = 6: xArt = Replace(xArt, "echo(5)", vbNullString): GoTo setEcho
    If InStr(1, xArt, "o(6)", vbTextCompare) Then S = 7: xArt = Replace(xArt, "echo(6)", vbNullString): GoTo setEcho
    
    xArt = Replace(xArt, "echo(", vbNullString, , , vbTextCompare)
    
setEcho:
Call modArtP(xArt)
  
   sysShell = Shell("cmd.exe /k echo " & xArt, S)
   sysShell = vbNullString
   Exit Function
   
'//Output w/ default message box
ElseIf InStr(1, xArt, "host(", vbTextCompare) Then

   xArt = Replace(xArt, "host(", vbNullString, , , vbTextCompare)
   Call modArtQ(xArt)
   If Right(xArt, 1) = ")" Then xArt = Left(xArt, Len(xArt) - 1)
   MsgBox (xArt)
   Exit Function
   
   
'//Output w/ VBA message box
ElseIf InStr(1, xArt, "msg(", vbTextCompare) Then

   xArt = Replace(xArt, "msg(", vbNullString, , , vbTextCompare)
    Call modArtP(xArt)
   
   If InStr(1, xArt, "=") Then '//check for variable
   xArtArr = Split(xArt, "=")
   xArt = xArtArr(0)
   If UBound(xArtArr) = 1 Then xArtArr = Split(xArtArr(1), ","): _
   xArtArr(0) = Trim(xArtArr(0))
   
   If UBound(xArtArr) = 0 Then xVar = MsgBox(xArtArr(0)): GoTo EndMsg
   If UBound(xArtArr) = 1 Then xArtArr(1) = Trim(xArtArr(1)): xVar = MsgBox(xArtArr(0), xArtArr(1)): GoTo EndMsg
   If UBound(xArtArr) = 2 Then xArtArr(1) = Trim(xArtArr(1)): xArtArr(2) = Trim(xArtArr(2)): xVar = MsgBox(xArtArr(0), xArtArr(1), xArtArr(2)): GoTo EndMsg
   
EndMsg:
   xArt = appEnv & ",#!" & xArt & "=" & xVar & ",#!" & X & ",#!" & 1: Call kinExpand(xArt)
   Exit Function
   
   End If
   
   MsgBox (xArt) '//no arguments
   Exit Function
'//#
'//
'/\_____________________________________
'//
'//      KEYSTROKE ARTICLES
'/\_____________________________________
'//
    ElseIf InStr(1, xArt, "key(", vbTextCompare) Then
    
    If InStr(1, xArt, ").clr", vbTextCompare) Then
    Dim oKey, oTemp As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oTemp = oFSO.GetFolder(drv & envHome & "\.z7\utility\temp")
    If oFSO.FolderExists(oTemp) = True Then For Each oKey In oTemp.Files: Kill (oKey): Next: Set oTemp = Nothing: Set oKey = Nothing
    Exit Function
    End If
    
    If InStr(1, xArt, "app.k", vbTextCompare) = False Then '//check for application key (key w/ VBA)
    
    Dim sysKey0 As String: Dim sysKey1 As String: Dim sysKey2 As String
    Dim sysKey3 As String: Dim sysKey4 As String: Dim sysKey5 As String:
    Dim sysKey6 As String
    
    If InStr(1, xArt, "y(0)", vbTextCompare) Then K = 1: xArt = Replace(xArt, "key(0)", vbNullString): GoTo setKey
    If InStr(1, xArt, "y(1)", vbTextCompare) Then K = 2: xArt = Replace(xArt, "key(1)", vbNullString): GoTo setKey
    If InStr(1, xArt, "y(2)", vbTextCompare) Then K = 3: xArt = Replace(xArt, "key(2)", vbNullString): GoTo setKey
    If InStr(1, xArt, "y(3)", vbTextCompare) Then K = 4: xArt = Replace(xArt, "key(3)", vbNullString): GoTo setKey
    If InStr(1, xArt, "y(4)", vbTextCompare) Then K = 5: xArt = Replace(xArt, "key(4)", vbNullString): GoTo setKey
    If InStr(1, xArt, "y(5)", vbTextCompare) Then K = 6: xArt = Replace(xArt, "key(5)", vbNullString): GoTo setKey
    If InStr(1, xArt, "y(6)", vbTextCompare) Then K = 7: xArt = Replace(xArt, "key(6)", vbNullString): GoTo setKey
    
setKey:
    xArt = Replace(xArt, "key", vbNullString, , , vbTextCompare)
    Call modArtP(xArt)
    xArt = Right(xArt, Len(xArt) - 1) '//remove leading quotes
    xArt = Left(xArt, Len(xArt) - 1) '//remove ending quotes
    
    
    If xArt = vbNullString Then Exit Function
  
    '//this is to help avoid file & variable collisions
    If K = 1 Then sysKey0 = drv & envHome & "\.z7\utility\temp\key0.vbs": Open sysKey0 For Output As #K: GoTo shKey
    If K = 2 Then sysKey1 = drv & envHome & "\.z7\utility\temp\key1.vbs": Open sysKey1 For Output As #K: GoTo shKey
    If K = 3 Then sysKey2 = drv & envHome & "\.z7\utility\temp\key2.vbs": Open sysKey2 For Output As #K: GoTo shKey
    If K = 4 Then sysKey3 = drv & envHome & "\.z7\utility\temp\key3.vbs": Open sysKey3 For Output As #K: GoTo shKey
    If K = 5 Then sysKey4 = drv & envHome & "\.z7\utility\temp\key4.vbs": Open sysKey4 For Output As #K: GoTo shKey
    If K = 6 Then sysKey5 = drv & envHome & "\.z7\utility\temp\key5.vbs": Open sysKey5 For Output As #K: GoTo shKey
    If K = 7 Then sysKey6 = drv & envHome & "\.z7\utility\temp\key6.vbs": Open sysKey6 For Output As #K: GoTo shKey
    '//no key specified
    If K = 0 Then K = 1: sysKey0 = drv & envHome & "\.z7\utility\temp\key0.vbs": Open sysKey0 For Output As #K: GoTo shKey
    
    '//Key using VBS...
shKey:
    Print #K, "On Error Resume Next"
    Print #K, "Dim Wshell"
    Print #K, "Set Wshell = Wscript.CreateObject(" & """" & "WScript.Shell""" & ")"
    Print #K, "Wshell.SendKeys " & """" & xArt & """"
    Print #K, "Set Wshell = Nothing"
    Print #K, "Wscript.Quit"
    Close #K
    
    If K = 1 Then Shell ("wscript.exe " & sysKey0), 0: Exit Function
    If K = 2 Then Shell ("wscript.exe " & sysKey1), 1: Exit Function
    If K = 3 Then Shell ("wscript.exe " & sysKey2), 2: Exit Function
    If K = 4 Then Shell ("wscript.exe " & sysKey3), 3: Exit Function
    If K = 5 Then Shell ("wscript.exe " & sysKey4), 4: Exit Function
    If K = 6 Then Shell ("wscript.exe " & sysKey5), 5: Exit Function
    If K = 7 Then Shell ("wscript.exe " & sysKey6), 6: Exit Function
    Exit Function
    
    Else
    
    '//Key using VBA...
    xArt = Replace(xArt, "app.", vbNullString)
    
    xArt = Replace(xArt, "key", vbNullString, , , vbTextCompare)
    Call modArtP(xArt)
    
    Application.SendKeys (xArt)
    Exit Function
    
    End If
'//#
'//
'/\_____________________________________
'//
'//        MOUSE ACTION ARTICLES
'/\_____________________________________
'//
'//Mouse click...
ElseIf InStr(1, xArt, "click(", vbTextCompare) Then

xArt = Replace(xArt, "click(", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)

'//article switches
If InStr(1, xArt, "-double", vbTextCompare) Then xArt = Replace(xArt, "-double", vbNullString, , , vbTextCompare): S = 1: GoTo setClick
If InStr(1, xArt, "-leftdown", vbTextCompare) Then xArt = Replace(xArt, "-leftdown", vbNullString, , , vbTextCompare): S = 2: GoTo setClick
If InStr(1, xArt, "-leftup", vbTextCompare) Then xArt = Replace(xArt, "-leftup", vbNullString, , , vbTextCompare): S = 3: GoTo setClick
If InStr(1, xArt, "-rightdown", vbTextCompare) Then xArt = Replace(xArt, "-rightdown", vbNullString, , , vbTextCompare): S = 4: GoTo setClick
If InStr(1, xArt, "-rightup", vbTextCompare) Then xArt = Replace(xArt, "-rightup", vbNullString, , , vbTextCompare): S = 5: GoTo setClick

setClick:
If InStr(1, xArt, ",") Then
xArt = Trim(xArt)
xArtArr = Split(xArt, ",") '//parameter
xPos = xArtArr(0) & "," & xArtArr(1)
Call basClick(S, xPos): Exit Function
End If

'//no parameter
S = 5: Call basClick(S, xPos)
Exit Function
'//#
'//
'/\_____________________________________
'//
'//        MODIFY STRING ARTICLES
'/\_____________________________________
'//
'//Convert char/string...
ElseIf InStr(1, xArt, "conv(", vbTextCompare) Then

xArt = Replace(xArt, "conv(", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)

'//article switches
If InStr(1, xArt, "-upper", vbTextCompare) Then S = vbUpperCase: GoTo setConv
If InStr(1, xArt, "-lower", vbTextCompare) Then S = vbLowerCase: GoTo setConv
If InStr(1, xArt, "-proper", vbTextCompare) Then S = vbProperCase: GoTo setConv
If InStr(1, xArt, "-unicode", vbTextCompare) Then S = vbUnicode: GoTo setConv

setConv:
xArtArr = Split(xArt, ",")
xVarArr = Split(xArtArr(0), "=") '//find variable
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next
xArt = xArtArr(1): Call modArtQ(xArt): xArtArr(1) = LTrim(xArt)

If UBound(xArtArr) = 1 Then xArt = StrConv(xArtArr(1), S): xArt = xVarArr(0) & "=" & xArt: _
xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 1: Call kinExpand(xArt): Exit Function
Exit Function

'//Replace char/string...
ElseIf InStr(1, xArt, "repl(", vbTextCompare) Then
xArt = Replace(xArt, "repl(", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)

xArtArr = Split(xArt, ",")
xVarArr = Split(xArtArr(0), "=") '//find variable
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next
xArt = xArtArr(1): Call modArtQ(xArt): xArtArr(1) = LTrim(xArt)
xArt = xArtArr(2): Call modArtQ(xArt): xArtArr(2) = LTrim(xArt)

If UBound(xArtArr) = 2 Then xArt = Replace(xArtArr(0), xArtArr(1), xArtArr(2), , , vbBinaryCompare): _
xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 1: Call kinExpand(xArt): Exit Function

If UBound(xArtArr) = 3 Then xArt = Replace(xArtArr(0), xArtArr(1), xArtArr(2), , , xArtArr(3)): _
xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 1: Call kinExpand(xArt): Exit Function
Exit Function

'//Trim starting & ending string by similiar character...
ElseIf InStr(1, xArt, "ptrim(", vbTextCompare) Then

xArt = Replace(xArt, "ptrim(", vbNullString, , , vbTextCompare)
Call modArtQ(xArt)

xVarArr = Split(xArt, "=") '//find variable
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next

If Left(xVarArr(1), 1) = "(" Then xVarArr(1) = Right(xVarArr(1), Len(xVarArr(1)) - 1):
If Right(xVarArr(1), 1) = ")" Then xArt = xVarArr(0) & "=" & Left(xVarArr(1), Len(xVarArr(1)) - 2):
xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 1: Call kinExpand(xArt): Exit Function
Exit Function

'//Trim char/string by starting left facing parentheses...
ElseIf InStr(1, xArt, "lptrim(", vbTextCompare) Then

xArt = Replace(xArt, "lptrim(", vbNullString, , , vbTextCompare)
Call modArtQ(xArt)

xVarArr = Split(xArt, "=") '//find variable
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next

If Left(xVarArr(1), 1) = "(" Then xVarArr(1) = Right(xVarArr(1), Len(xVarArr(1)) - 1):
If Right(xVarArr(1), 1) = ")" Then xArt = xVarArr(0) & "=" & Left(xVarArr(1), Len(xVarArr(1)) - 1):
xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 1: Call kinExpand(xArt): Exit Function
Exit Function

'//Trim char/string by ending right facing parentheses...
ElseIf InStr(1, xArt, "rptrim(", vbTextCompare) Then

xArt = Replace(xArt, "rptrim(", vbNullString, , , vbTextCompare)
Call modArtQ(xArt)

xVarArr = Split(xArt, "=") '//find variable
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next

If Right(xVarArr(1), 1) = ")" Then xArt = xVarArr(0) & "=" & Left(xVarArr(1), Len(xVarArr(1)) - 2): _
xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 1: Call kinExpand(xArt): Exit Function
Exit Function

'//Trim char/string by quotations...
ElseIf InStr(1, xArt, "qtrim(", vbTextCompare) Then

xArt = Replace(xArt, "qtrim(", vbNullString, , , vbTextCompare)
Call modArtP(xArt)

xVarArr = Split(xArt, "=") '//find variable
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next

If Left(xVarArr(1), 1) = """" Then xVarArr(1) = Right(xVarArr(1), Len(xVarArr(1)) - 1):
If Right(xVarArr(1), 1) = """" Then xArt = xVarArr(0) & "=" & Left(xVarArr(1), Len(xVarArr(1)) - 1):
xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 1: Call kinExpand(xArt): Exit Function
Exit Function

'//Trim starting & ending string by similiar character...
ElseIf InStr(1, xArt, "xtrim(", vbTextCompare) Then

xArt = Replace(xArt, "xtrim(", vbNullString, , , vbTextCompare)

xArtArr = Split(xArt, ",")
xVarArr = Split(xArtArr(0), "=") '//find variable
If UBound(xArtArr) > 1 Then xArtArr(1) = xArtArr(UBound(xArtArr))
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next
xArt = xArtArr(1): Call modArtP(xArt): Call modArtQ(xArt): xArtArr(1) = LTrim(xArt)

If Left(xVarArr(1), 1) = xArtArr(1) Then xVarArr(1) = Right(xVarArr(1), Len(xVarArr(1)) - 1): xArt = xVarArr(0) & "=" & xVarArr(1):
If Right(xVarArr(1), 1) = xArtArr(1) Then xVarArr(1) = Left(xVarArr(1), Len(xVarArr(1)) - 1): xArt = xVarArr(0) & "=" & xVarArr(1):
xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 1: Call kinExpand(xArt): Exit Function
Exit Function

'//Find char/string...
ElseIf InStr(1, xArt, "ins(", vbTextCompare) Then
xArt = Replace(xArt, "ins(", vbNullString, , , vbTextCompare)
xArt = Replace(xArt, """", vbNullString)
xArt = Replace(xArt, ")", vbNullString)

xArtArr = Split(xArt, ",")
xVarArr = Split(xArtArr(0), "=")
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next
If Left(xVarArr(1), 1) = " " Then xVarArr(1) = Right(xVarArr(1), Len(xVarArr(1)) - 1) '//find variable
'//
If UBound(xArtArr) = 2 Then
If Left(xArtArr(1), 1) = " " Then xArtArr(1) = Right(xArtArr(1), Len(xArtArr(1)) - 1)
If Left(xArtArr(2), 1) = " " Then xArtArr(2) = Right(xArtArr(2), Len(xArtArr(2)) - 1)

If InStr(xVarArr(1), xArtArr(1), xArtArr(2), vbBinaryCompare) Then
xArt = appEnv & ",#!" & xVarArr(0) & "=" & "TRUE" & ",#!" & X & ",#!" & 1
    Else
        xArt = appEnv & ",#!" & xVarArr(0) & "=" & "FALSE" & ",#!" & X & ",#!" & 1
            End If
                Call kinExpand(xArt): Exit Function
                    End If
                    
'//
If UBound(xArtArr) = 3 Then
If Left(xArtArr(1), 1) = " " Then xArtArr(1) = Right(xArtArr(1), Len(xArtArr(1)) - 1)
If Left(xArtArr(2), 1) = " " Then xArtArr(2) = Right(xArtArr(2), Len(xArtArr(2)) - 1)
If Left(xArtArr(3), 1) = " " Then xArtArr(3) = Right(xArtArr(3), Len(xArtArr(3)) - 1)
CX = xArtArr(3): Call basCompare(CX)
If InStr(xVarArr(1), xArtArr(1), xArtArr(2), CX) Then
xArt = appEnv & ",#!" & xVarArr(0) & "=" & "TRUE" & ",#!" & X & ",#!" & 1
    Else
        xArt = appEnv & ",#!" & xVarArr(0) & "=" & "FALSE" & ",#!" & X & ",#!" & 1
                End If
                   Call kinExpand(xArt): Exit Function
                        End If
Exit Function
   
'//Reverse string characters...
ElseIf InStr(1, xArt, "revstr(", vbTextCompare) Then

xArt = Replace(xArt, "revstr(", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)

xArtArr = Split(xArt, "=") '//find variable

xArtArr(1) = StrReverse(xArtArr(1))
xArt = xArtArr(0) & "=" & xArtArr(1)

xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 1: Call kinExpand(xArt): Exit Function
Exit Function
'//#
'//
'/\_____________________________________
'//
'//      SYSTEM SHELL/PC ARTICLES
'/\_____________________________________
'//
'//System shell...
ElseIf InStr(1, xArt, "sh(", vbTextCompare) Then

If InStr(1, xArt, "h(0)", vbTextCompare) Then S = 0: xArt = Replace(xArt, "sh(0)", vbNullString, , , vbTextCompare): GoTo setSh
If InStr(1, xArt, "h(1)", vbTextCompare) Then S = 1: xArt = Replace(xArt, "sh(1)", vbNullString, , , vbTextCompare): GoTo setSh
If InStr(1, xArt, "h(2)", vbTextCompare) Then S = 2: xArt = Replace(xArt, "sh(2)", vbNullString, , , vbTextCompare): GoTo setSh
If InStr(1, xArt, "h(3)", vbTextCompare) Then S = 3: xArt = Replace(xArt, "sh(3)", vbNullString, , , vbTextCompare): GoTo setSh
If InStr(1, xArt, "h(4)", vbTextCompare) Then S = 4: xArt = Replace(xArt, "sh(4)", vbNullString, , , vbTextCompare): GoTo setSh
If InStr(1, xArt, "h(5)", vbTextCompare) Then S = 5: xArt = Replace(xArt, "sh(5)", vbNullString, , , vbTextCompare): GoTo setSh
If InStr(1, xArt, "h(6)", vbTextCompare) Then S = 6: xArt = Replace(xArt, "sh(6)", vbNullString, , , vbTextCompare): GoTo setSh

setSh:
   xArt = Replace(xArt, "sh(", vbNullString, , , vbTextCompare)
   Call modArtP(xArt): Call modArtQ(xArt)
    
   FX = xArt
   Call basWebFilter(FX) '//check for web filter switches
   If FX <> vbNullString Then xArt = FX
   
   xArt = "start " & xArt
    
   sysShell = Shell("cmd.exe /s /c" & xArt, S)
   sysShell = vbNullString
   Exit Function

'//PC articles...
ElseIf InStr(1, xArt, "pc(", vbTextCompare) Then

xArt = Replace(xArt, "pc", vbNullString, , , vbTextCompare)

'//article switches
If InStr(1, xArt, "-file", vbTextCompare) Then xArt = Replace(xArt, "-file", vbNullString, , , vbTextCompare): S = 1
If InStr(1, xArt, "-fldr", vbTextCompare) Then xArt = Replace(xArt, "-fldr", vbNullString, , , vbTextCompare): S = 2

If InStr(1, xArt, ".exist", vbTextCompare) Then S = S & 1: xArt = Replace(xArt, ".exist", vbNullString, , , vbTextCompare): GoTo SetPC
If InStr(1, xArt, ".del", vbTextCompare) Then S = S & 2: xArt = Replace(xArt, ".del", vbNullString, , , vbTextCompare): GoTo SetPC
If InStr(1, xArt, ".open", vbTextCompare) Then S = 3: xArt = Replace(xArt, ".open", vbNullString, , , vbTextCompare): GoTo SetPC
If InStr(1, xArt, ".stop", vbTextCompare) Then S = 4: xArt = Replace(xArt, ".stop", vbNullString, , , vbTextCompare): GoTo SetPC

SetPC:
Call modArtP(xArt): Call modArtQ(xArt): xArt = Trim(xArt)

'//file exists...
If S = 1 Or S = 11 Then If Dir(xArt) <> "" Then MsgBox "TRUE": Exit Function Else MsgBox ("FALSE"): Exit Function
'//directory exists...
If S = 21 Then If Dir(xArt, vbDirectory) <> "" Then MsgBox "TRUE": Exit Function Else MsgBox ("FALSE"): Exit Function
'//delete file...
If S = 2 Or S = 12 Then Kill (xArt): Exit Function
'//delete empty directory...
If S = 22 Then RmDir (xArt): Exit Function
'//open...
If S = 3 Then xArt = "start " & xArt: sysShell = Shell("cmd.exe /s /c" & xArt, 0): sysShell = vbNullString: Exit Function
'//stop (taskkill)
If S = 4 Then xArt = "taskkill /f /im " & xArt: sysShell = Shell("cmd.exe /s /c" & xArt, 0): sysShell = vbNullString: Exit Function

'//no excerpt provided
MsgBox (xArt)
Exit Function

'//PC dot-direct articles...
ElseIf InStr(1, xArt, "pc.", vbTextCompare) Then

If InStr(1, xArt, ".copy&", vbTextCompare) Then C = 1
If InStr(1, xArt, ".copy&!", vbTextCompare) Then C = 2: GoTo setPCdot
If InStr(1, xArt, ".shutdown", vbTextCompare) Then C = 3: GoTo setPCdot
If InStr(1, xArt, ".off", vbTextCompare) Then C = 4: GoTo setPCdot
If InStr(1, xArt, ".rest", vbTextCompare) Then C = 5: GoTo setPCdot
If InStr(1, xArt, ".reboot", vbTextCompare) Then C = 6: GoTo setPCdot
If InStr(1, xArt, ".clr", vbTextCompare) Then C = 7: GoTo setPCdot

setPCdot:
xArt = Replace(xArt, "pc.", vbNullString, , , vbTextCompare)
Call modArtP(xArt)
'//article switches
If InStr(1, xArt, "-e", vbTextCompare) Then xArt = Replace(xArt, "-e", vbNullString, , , vbTextCompare): C = C & "1" '//check for switch(s)
If InStr(1, xArt, "-t", vbTextCompare) Then '//check for timer switch
Dim xT As String
xArtArr = Split(xArt, "-t")
xT = "/t " & xArtArr(1)
End If

'//Copy & paste a file
If C = 1 Then xArt = Replace(xArt, "copy&", vbNullString, , , vbTextCompare): xArtArr = Split(xArt, ","): _
xArt = "copy /y " & xArtArr(0) & " " & xArtArr(1): sysShell = Shell("cmd.exe /s /c " & xArt, vbNormalFocus): sysShell = vbNullString: Exit Function
'//Copy, paste then delete copied file
If C = 2 Then xArt = Replace(xArt, "copy&!", vbNullString, , , vbTextCompare): xArtArr = Split(xArt, ","): _
xArt = "copy /y " & xArtArr(0) & " " & xArtArr(1): sysShell = Shell("cmd.exe /s /c " & xArt, vbNormalFocus): sysShell = vbNullString: If Dir(xArtArr(0)) <> "" Then Kill (xArtArr(0)): Exit Function
'//Shutdown pc
If C = 3 Then xArt = "shutdown /s " & xT: sysShell = Shell("cmd.exe /s /c " & xArt, vbNormalFocus): sysShell = vbNullString: Exit Function
'//Shutdown pc, on next boot auto-sign in if enabled. Restart apps.
If C = 31 Then xArt = "shutdown /sg " & xT: sysShell = Shell("cmd.exe /s /c " & xArt, vbNormalFocus): sysShell = vbNullString: Exit Function
'//Logoff pc
If C = 4 Then xArt = "shutdown /l ": sysShell = Shell("cmd.exe /s /c " & xArt, vbNormalFocus): sysShell = vbNullString: Exit Function
'//Hibernate pc
If C = 5 Then xArt = "shutdown /h ": sysShell = Shell("cmd.exe /s /c " & xArt, vbNormalFocus): sysShell = vbNullString: Exit Function
'//Restart pc
If C = 6 Then xArt = "shutdown /r " & xT: sysShell = Shell("cmd.exe /s /c " & xArt, vbNormalFocus): sysShell = vbNullString: Exit Function
'//Restart pc, on next boot auto-sign in if enabled. Restart apps.
If C = 61 Then xArt = "shutdown /g " & xT: sysShell = Shell("cmd.exe /s /c " & xArt, vbNormalFocus): sysShell = vbNullString: Exit Function
'//Clear logoff queue
If C = 7 Then xArt = "shutdown /a ": sysShell = Shell("cmd.exe /s /c " & xArt, vbNormalFocus): sysShell = vbNullString: Exit Function
'//#
'//
'/\_____________________________________
'//
'//        QUERY ARTICLES
'/\_____________________________________
'//
'//Query...
ElseIf InStr(1, xArt, "q(", vbTextCompare) Then

If InStr(1, xArt, ".exist", vbTextCompare) Then C = 1: xArt = Replace(xArt, ".exist", vbNullString, , , vbTextCompare): GoTo setQ
If InStr(1, xArt, ".del", vbTextCompare) Then C = 2: xArt = Replace(xArt, ".del", vbNullString, , , vbTextCompare): GoTo setQ
If InStr(1, xArt, ".move", vbTextCompare) Then C = 3: xArt = Replace(xArt, ".move", vbNullString, , , vbTextCompare): GoTo setQ
If InStr(1, xArt, ".name", vbTextCompare) Then C = 4: xArt = Replace(xArt, ".name", vbNullString, , , vbTextCompare): GoTo setQ
If InStr(1, xArt, ".open", vbTextCompare) Then C = 5: xArt = Replace(xArt, ".open", vbNullString, , , vbTextCompare): GoTo setQ
If InStr(1, xArt, ".stop", vbTextCompare) Then C = 6: xArt = Replace(xArt, ".stop", vbNullString, , , vbTextCompare): GoTo setQ
If C = 0 Then Exit Function

setQ:
'//article switches
If InStr(1, xArt, "-loose", vbTextCompare) Then xArt = Replace(xArt, "-loose", vbNullString, , , vbTextCompare): S = 1
If InStr(1, xArt, "-strict", vbTextCompare) Then xArt = Replace(xArt, "-strict", vbNullString, , , vbTextCompare): S = 2
If InStr(1, xArt, "-file", vbTextCompare) Then xArt = Replace(xArt, "-file", vbNullString, , , vbTextCompare): S = S & 3
If InStr(1, xArt, "-fldr", vbTextCompare) Then xArt = Replace(xArt, "-fldr", vbNullString, , , vbTextCompare): S = S & 4

xArtArr = Split(xArt, "q(", , vbTextCompare)
If InStr(1, xArtArr(1), ",") Then xArtArr = Split(xArtArr(1), ","): _
xArt = xArtArr(0): Call modArtP(xArt): Call modArtQ(xArt): xArtArr(0) = Trim(xArt): xArt = xArtArr(0): _
xArt = xArtArr(1): Call modArtP(xArt): Call modArtQ(xArt): xArtArr(1) = Trim(xArt): xArt = xArtArr(0) Else: _
xArt = xArtArr(1): Call modArtP(xArt): Call modArtQ(xArt): xArtArr(1) = Trim(xArt): xArt = xArtArr(1)
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oDrv = oFSO.GetFolder("C:\") '//set drive (default is C:)
For Each oSubFldr In oDrv.SubFolders
If InStr(1, xArt, oSubFldr.name, vbTextCompare) Then xSubFldr = oSubFldr.name: GoTo hQ '//check for folder match in drive
Next

hQ:
Set oFSO = Nothing
Set oDrv = Nothing
Set oSubFldr = Nothing

Call modArtP(xArt): Call modArtQ(xArt)

QX = xArt
Call basQuery(QX, S)
xQueryArr = Split(QX, ",")

'//exists...
If C = 1 Then If xQueryArr(1) = 0 Then MsgBox ("TRUE" & vbNewLine & vbNewLine & xQueryArr(0)): Exit Function Else MsgBox ("FALSE"): Exit Function
'//delete...
If C = 2 Then Kill (xQueryArr(0)): Exit Function
'//move...
If C = 3 Then xQueryArr(0) = Replace(xQueryArr(0), " ", """" & " " & """"): xArt = "move " & xQueryArr(0) & " " & xArtArr(1): sysShell = Shell("cmd.exe /s /c" & xArt, 0): sysShell = vbNullString: Exit Function
'//name...
If C = 4 Then xQueryArr(0) = Replace(xQueryArr(0), " ", """" & " " & """"): xArt = "ren " & xQueryArr(0) & " " & xArtArr(1): sysShell = Shell("cmd.exe /s /c" & xArt, 0): sysShell = vbNullString: Exit Function
'//open...
If C = 5 Then xArt = "start " & xQueryArr(0): sysShell = Shell("cmd.exe /s /c" & xArt, 0): sysShell = vbNullString: Exit Function
'//stop (taskkill)
If C = 6 Then xArtArr = Split(xArt, "\"): xArt = xArtArr(UBound(xArtArr)): xArt = "taskkill /f /im " & xArt: sysShell = Shell("cmd.exe /s /c" & xArt, 0): sysShell = vbNullString: Exit Function
Exit Function
'//#
'//
'/\_____________________________________
'//
'//         UTILITY ARTICLES
'/\_____________________________________
'//
ElseIf InStr(1, xArt, "incr(", vbTextCompare) Then
xArt = Replace(xArt, "incr(", vbNullString, , , vbTextCompare)
Call modArtP(xArt)
If InStr(1, xArt, "+") Then C = 1: xArt = Replace(xArt, "+", vbNullString)
If InStr(1, xArt, "-") Then C = 2: xArt = Replace(xArt, "-", vbNullString)
If InStr(1, xArt, "=") Then
xArtArr = Split(xArt, "=") '//find variable
xArtArr(0) = Trim(xArtArr(0)): xArtArr(1) = Trim(xArtArr(1))

If C = 1 Then xArtArr(1) = CLng(xArtArr(1)) + CLng(xArtArr(1))
If C = 2 Then xArtArr(1) = -(CLng(xArtArr(1))) + -(CLng(xArtArr(1)))

xArt = xArtArr(0) & "=" & xArtArr(1)

xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 1: Call kinExpand(xArt): Exit Function
End If

If C = 1 Then xArt = xArt + xArt
If C = 2 Then xArt = xArt - xArt

Exit Function

'//Randomize numbers...
ElseIf InStr(1, xArt, "rnd(", vbTextCompare) Then
xArt = Replace(xArt, "rnd(", vbNullString, , , vbTextCompare)
Call modArtP(xArt): Call modArtQ(xArt)

If InStr(1, xArt, ":") Then

xArtArr = Split(xArt, "=") '//find variable
xArtArr(0) = Trim(xArtArr(0)): xArtArr(1) = Trim(xArtArr(1))

Randomize
xTempArr = Split(xArtArr(1), ":")

xArtArr(1) = CLng((xTempArr(1) * Rnd) + xTempArr(0))

xArt = xArtArr(0) & "=" & xArtArr(1)

If UBound(xArtArr) = 1 Then xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 1: Call kinExpand(xArt): Exit Function
End If

Exit Function
'//#
'//
'/\_____________________________________
'//
'//         WINFORM ARTICLES
'/\_____________________________________
'//
'//Output current window number...
ElseIf InStr(1, xArt, "me()", vbTextCompare) And Len(xArt) <= 4 Then MsgBox (Range("xlasWinForm").Value): Exit Function

    '//Set window number...
        ElseIf InStr(1, xArt, "winform(", vbTextCompare) Then
        
    '//article switches
        If InStr(1, xArt, "-last", vbTextCompare) Then _
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value = _
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinFormLast").Value: Exit Function  '//set to last window
        
        xArt = Replace(xArt, "winform(", vbNullString, , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        
        Call xlAppScript_lex.findChar(xArt)
        If xArt = "(*Err)" Then Exit Function
        
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value = xArt
        
        Exit Function
        
End If '//end
        
ErrMsg:
'//Article not found...
If errLvl <> 0 Then xArt = xArt & "(*Err)"
Workbooks(appEnv).Worksheets(appBlk).Range("xlasErrRef").Value = """" & xArt & """"
End Function
Private Function libFlag$(xArt, errLvl As Byte)

'/\_____________________________________
 '//
'//         FLAGS
'/\_____________________________________
'//
On Error GoTo ErrMsg

Call findEnvironment(appEnv, appBlk)

'//Create runtime error
If InStr(1, xArt, "--err", vbTextCompare) Then xArt = "(*Err)"

'//Run script w/ environment errors enabled (default)
If InStr(1, xArt, "--enableerr", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLibErrLvl") = 0
errLvl = 0
xArt = 1: Exit Function
End If

'//Run script w/ environment errors disabled
If InStr(1, xArt, "--disableerr", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLibErrLvl") = 1
Range("xlasEnd").Value = 0
errLvl = 0
xArt = 1: Exit Function
End If

'//Run script w/ animations/updates disabled (default)
If InStr(1, xArt, "--disableupdates", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasUpdateEnable") = 0
Call disableWbUpdates
errLvl = 0
xArt = 1: Exit Function
End If

'//Run script w/ animations/updates enabled
If InStr(1, xArt, "--enableupdates", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasUpdateEnable") = 1
Call enableWbUpdates
errLvl = 0
xArt = 1: Exit Function
End If

'//Run script w/ libraries statically disabled (default)
If InStr(1, xArt, "--disablestatic", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLocalStatic") = 0
errLvl = 0
xArt = 1: Exit Function
End If

'//Run script w/ libraries statically enabled
If InStr(1, xArt, "--enablestatic", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLocalStatic") = 1
errLvl = 0
xArt = 1: Exit Function
End If

'//Run script w/ clear runtime block addresses (default)
If InStr(1, xArt, "--disablecontain", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLocalContain") = 0
errLvl = 0
xArt = 1: Exit Function
End If

'//Run script w/o clearing runtime block addresses
If InStr(1, xArt, "--enablecontain", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLocalContain") = 1
errLvl = 0
xArt = 1: Exit Function
End If

'//Run script w/o global control variables
If InStr(1, xArt, "--disableglobal", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasGlobalControl") = 0
errLvl = 0
xArt = 1: Exit Function
End If

'//Run script w global control variables (default)
If InStr(1, xArt, "--enableglobal", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasGlobalControl").Value = 1
errLvl = 0
xArt = 1: Exit Function
End If

Exit Function

ErrMsg:
'//flag not found...
xArt = "(*Err)"

End Function
Private Function libSwitch$(xArt, errLvl As Byte)

'/\_____________________________________
 '//
'//         LIBRARY SWITCHES
'/\_____________________________________
'//
Dim xArtCl As String
Dim X As Integer

On Error GoTo ErrMsg

xArtCl = xArt
xArtArr = Split(xArt, "--")

For X = 0 To UBound(xArtArr)
xArt = xArtArr(X): Call modArtP(xArt): Call modArtQ(xArt): Call modArtS(xArt): xArtArr(X) = xArt: xArt = xArtCl
If InStr(1, xArtArr(X), "date", vbTextCompare) Then xArt = Replace(xArt, "--date", Date, , , vbTextCompare): GoTo NextStep
If InStr(1, xArtArr(X), "day", vbTextCompare) Then xArt = Replace(xArt, "--day", Day(Date), , , vbTextCompare): GoTo NextStep
If InStr(1, xArtArr(X), "present", vbTextCompare) Then xArt = Replace(xArt, "--present", Date & " " & Time, , , vbTextCompare): GoTo NextStep
If InStr(1, xArtArr(X), "me", vbTextCompare) Then xArt = Replace(xArt, "--me", ActiveWorkbook.name, , , vbTextCompare): GoTo NextStep
If InStr(1, xArtArr(X), "month", vbTextCompare) Then xArt = Replace(xArt, "--month", Month(Date), , , vbTextCompare): GoTo NextStep
If InStr(1, xArtArr(X), "now", vbTextCompare) Then xArt = Replace(xArt, "--now", Time, , , vbTextCompare): GoTo NextStep
If InStr(1, xArtArr(X), "null", vbTextCompare) Then xArt = Replace(xArt, "--null", vbNullString, , , vbTextCompare): GoTo NextStep
If InStr(1, xArtArr(X), "lparen", vbTextCompare) Then xArt = Replace(xArt, "--lparen", "(", , , vbTextCompare): GoTo NextStep
If InStr(1, xArtArr(X), "rparen", vbTextCompare) Then xArt = Replace(xArt, "--rparen", ")", , , vbTextCompare): GoTo NextStep
If InStr(1, xArtArr(X), "quote", vbTextCompare) Then xArt = Replace(xArt, "--quote", """", , , vbTextCompare): GoTo NextStep
If InStr(1, xArtArr(X), "space", vbTextCompare) Then xArt = Replace(xArt, "--space", Space(0), , , vbTextCompare): GoTo NextStep
If InStr(1, xArtArr(X), "time", vbTextCompare) Then xArt = Replace(xArt, "--time", Time, , , vbTextCompare): GoTo NextStep
If InStr(1, xArtArr(X), "year", vbTextCompare) Then xArt = Replace(xArt, "--year", Year(Date), , , vbTextCompare): GoTo NextStep
NextStep:
xArtCl = xArt
Next

Exit Function

ErrMsg:
'//switch not found...
xArt = "(*Err)"

End Function
Private Function basClick(ByVal S As Byte, ByVal xPos As String)

'/#########################\
'//  Basic Click Function #\\
'///#######################\\\

If S = 1 Then Call dblClk(xPos)
If S = 2 Then Call leftClkDown(xPos)
If S = 3 Then Call leftClkUp(xPos)
If S = 4 Then Call rightClkDown(xPos)
If S = 5 Then Call rightClkUp(xPos)

End Function
Private Sub dblClk(xPos)
'//double left click
xPosArr = Split(xPos, ",")

  SetCursorPos xPosArr(0), xPosArr(1) '//x & y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub
Private Sub leftClkDown(xPos)
'//left click down
xPosArr = Split(xPos, ",")

  SetCursorPos xPosArr(0), xPosArr(1) '//x & y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
End Sub
Private Sub leftClkUp(xPos)
'//left click up
xPosArr = Split(xPos, ",")

  SetCursorPos xPosArr(0), xPosArr(1) '//x & y position
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub
Private Sub rightClkDown(xPos)
'//right click
xPosArr = Split(xPos, ",")

  SetCursorPos xPosArr(0), xPosArr(1) '//x & y position
  mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
End Sub
Private Sub rightClkUp(xPos)
'//right click
xPosArr = Split(xPos, ",")

  SetCursorPos xPosArr(0), xPosArr(1) '//x & y position
  mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub
Private Function basQuery$(QX, ByVal S As Byte)

'/#########################\
'//   Basic Query Search  #\\
'///#######################\\\

Dim oFSO, oDir, oFile, oLastDir, oSubFldr, oSubFldr1 As Object
Dim Q_MATCH As Byte
Dim X As Integer
Set oFSO = CreateObject("Scripting.FileSystemObject")
Q_MATCH = 0

'//breakdown drive
xDrvArr = Split(QX, ":")
xDrv = xDrvArr(0)
'//check if drive exists
On Error GoTo ErrMsg
Set oDir = oFSO.GetFolder(xDrv & ":\") '//will error if can't find

'//breakdown base folder
xBase = xDrvArr(1)
xBaseArr = Split(xBase, "\")

If UBound(xBaseArr) = 2 Then '//check for multiple folders listed

'//breakdown file/folder name to query
xBase = xBaseArr(1)
xFind = xBaseArr(2)

'//check if base folder exists
Set oDir = oFSO.GetFolder(xDrv & ":\" & xBase)

xLoc = oDir & "\" & xFind
'//check if file/folder name exists as current search
If Dir(xLoc) = "" Then
Err.Clear: On Error Resume Next
'//filter current query search through all available folders within the base folder to find a match
'//
'//return full path if match found & 0 or 1 based on a successful/unsuccessful search

'//check for query assignments
If S = 1 Then Q_MATCH = 1: GoTo qFldr
If S = 2 Then Q_MATCH = 2: GoTo qFldr
If S = 13 Then Q_MATCH = 1: GoTo qFile
If S = 23 Then Q_MATCH = 2: GoTo qFile
If S = 4 Then GoTo qFldr
If S = 24 Then GoTo qFldr

'//check for query type if no assignment (file or folder based on identifier for file extension)
If InStr(1, xFind, ".") Then GoTo qFile

qFldr:
'/#########################\
'//       Folder Query     #\\
'///#######################\\\
'//
'//Query for folder...
For Each oSubFldr In oDir.SubFolders
QX = xDrv & ":\" & xBase & "\" & oSubFldr.name & "\" & xFind '//set query search to next folder
If oFSO.FolderExists(QX) = True Then GoTo qFound

'//loose match
If Q_MATCH = 1 Then If oSubFldr <> Empty Then If InStr(1, oSubFldr.name, xFind, vbTextCompare) Then QX = oSubFldr: GoTo qFound
'//strict match
If Q_MATCH = 2 Then If oSubFldr <> Empty Then If InStr(1, oSubFldr.name, xFind, vbBinaryCompare) Then QX = oSubFldr: GoTo qFound
Next

Set oDir = oFSO.GetFolder(xDrv & ":\" & xBase)

'//search through local base drive directories
    For Each oSubFldr In oDir.SubFolders
    QX = xDrv & ":\" & xBase & "\" & oSubFldr.name & "\" & xFind '//set query search to next folder
    If Dir(QX) <> "" Then
    GoTo qFound
        Else: GoTo searchLocalFol
            End If
                Next

searchLocalFol:
Set oLastDir = oDir

For Each oSubFldr1 In oLastDir.SubFolders

'//search through folders in local base drive directories
Set oDir = oFSO.GetFolder(xDrv & ":\" & xBase & "\" & oSubFldr1.name)

                For Each oSubFldr In oDir.SubFolders
                QX = oSubFldr & "\" & xFind '//set query search to next folder
                If oFSO.FolderExists(QX) = True Then GoTo qFound
                
                '//loose match
                If Q_MATCH = 1 Then If oSubFldr <> Empty Then If InStr(1, oSubFldr.name, xFind, vbTextCompare) Then QX = oSubFldr: GoTo qFound
                '//strict match
                If Q_MATCH = 2 Then If oSubFldr <> Empty Then If InStr(1, oSubFldr.name, xFind, vbBinaryCompare) Then QX = oSubFldr: GoTo qFound
                Next
                    Next
                    
                    
                        Else: GoTo qFound
                        
                        
                                GoTo ErrMsg: '//nothing found
                                
                                                                
'/#########################\
'//       File Query      #\\
'///#######################\\\
'//
'//Query for file...
qFile:

For Each oSubFldr In oDir.SubFolders
QX = xDrv & ":\" & xBase & "\" & oSubFldr.name & "\" & xFind '//set query search to next folder
If oFSO.FileExists(QX) = True Then GoTo qFound

'//loose match
If Q_MATCH = 1 Then
If oSubFldr <> Empty Then
    For Each oFile In oSubFldr.Files
      If oFile <> Empty Then If InStr(1, oFile.name, xFind, vbTextCompare) Then QX = oFile.Path: GoTo qFound
            Next
                End If
                    End If
                    
'//strict match
If Q_MATCH = 2 Then
If oSubFldr <> Empty Then
    For Each oFile In oSubFldr.Files
      If oFile <> Empty Then If InStr(1, oFile.name, xFind, vbBinaryCompare) Then QX = oFile.Path: GoTo qFound
            Next
                End If
                    End If
                        Next

Set oDir = oFSO.GetFolder(xDrv & ":\" & xBase)

'//search through local base drive directories
    For Each oSubFldr In oDir.SubFolders
    QX = xDrv & ":\" & xBase & "\" & oSubFldr.name & "\" & xFind '//set query search to next folder
    If Dir(QX) <> "" Then
    GoTo qFound
        Else: GoTo searchLocalFil
            End If
                Next

searchLocalFil:
Set oLastDir = oDir

For Each oSubFldr1 In oLastDir.SubFolders

'//search through folders in local base drive directories
Set oDir = oFSO.GetFolder(xDrv & ":\" & xBase & "\" & oSubFldr1.name)

For Each oSubFldr In oDir.SubFolders
QX = oSubFldr & "\" & xFind '//set query search to next folder
If oFSO.FileExists(QX) = True Then GoTo qFound

'//loose match
If Q_MATCH = 1 Then
If oSubFldr <> Empty Then
  For Each oFile In oSubFldr.Files
    If oFile <> Empty Then If InStr(1, oFile.name, xFind, vbTextCompare) Then QX = oFile.Path: GoTo qFound
          Next
              End If
                  End If
                
'//strict match
If Q_MATCH = 2 Then
If oSubFldr <> Empty Then
    For Each oFile In oSubFldr.Files
    If oFile <> Empty Then If InStr(1, oFile.name, xFind, vbBinaryCompare) Then QX = oFile.Path: GoTo qFound
        Next
            End If
                End If
                    
                        Next
                            Next
                    
                        GoTo ErrMsg: '//nothing found
                               
                               
'//Found our query!
qFound:
Set fso = Nothing: Set oDir = Nothing: Set oLastDir = Nothing: Set SubFldr = Nothing: Set oSubFldr1 = Nothing
QX = QX & "," & 0
Exit Function

End If
    End If

ErrMsg:
Err.Clear
Set fso = Nothing: Set oDir = Nothing: Set oLastDir = Nothing: Set SubFldr = Nothing: Set oSubFldr1 = Nothing
QX = QX & "," & 1

End Function
Private Function basSaveFormat(EX) As String

EX = Replace(EX, " ", vbNullString)

Select Case EX

Case Is = "0" Or EX = "AddIn"
EX = xlAddIn: Exit Function
Case Is = "1" Or EX = "AddIn8"
EX = xlAddIn8: Exit Function
Case Is = "2" Or EX = "CSV"
EX = xlCSV: Exit Function
Case Is = "3" Or EX = "CSVMac"
EX = xlCSVMac: Exit Function
Case Is = "4" Or EX = "CSVMSDOS"
EX = xlCSVMSDOS: Exit Function
Case Is = "5" Or EX = "CSVUTF8"
EX = xlCSVUTF8: Exit Function
Case Is = "6" Or EX = "CSVWindows"
EX = xlCSVWindows: Exit Function
Case Is = "7" Or EX = "CurrentPlatformText"
EX = xlCurrentPlatformText: Exit Function
Case Is = "8" Or EX = "DBF2"
EX = xlDBF2: Exit Function
Case Is = "9" Or EX = "DBF3"
EX = xlDBF3: Exit Function
Case Is = "10" Or EX = "DBF4"
EX = xlDBF4: Exit Function
Case Is = "11" Or EX = "DIF"
EX = xlDIF: Exit Function
Case Is = "12" Or EX = "Excel12"
EX = xlExcel12: Exit Function
Case Is = "13" Or EX = "Excel2"
EX = xlExcel2: Exit Function
Case Is = "14" Or EX = "Excel2FarEast"
EX = xlExcel2FarEast: Exit Function
Case Is = "15" Or EX = "Excel3"
EX = xlExcel3: Exit Function
Case Is = "16" Or EX = "Excel4"
EX = xlExcel4: Exit Function
Case Is = "17" Or EX = "Excel4Workbook"
EX = xlExcel4Workbook: Exit Function
Case Is = "18" Or EX = "Excel5"
EX = xlExcel5: Exit Function
Case Is = "19" Or EX = "Excel7"
EX = xlExcel7: Exit Function
Case Is = "20" Or EX = "Excel8"
EX = xlExcel8: Exit Function
Case Is = "21" Or EX = "Excel9795"
EX = xlExcel9795: Exit Function
Case Is = "22" Or EX = "Html"
EX = xlHtml: Exit Function
Case Is = "23" Or EX = "IntlAddIn"
EX = xlIntlAddIn: Exit Function
Case Is = "24" Or EX = "IntlMacro"
EX = xlIntlMacro: Exit Function
Case Is = "25" Or EX = "OpenDocumentSpreadsheet"
EX = xlOpenDocumentSpreadsheet: Exit Function
Case Is = "26" Or EX = "OpenXMLAddIn"
EX = xlOpenXMLAddIn: Exit Function
Case Is = "27" Or EX = "OpenXMLStrictWorkbook"
EX = xlOpenXMLStrictWorkbook:  Exit Function
Case Is = "28" Or EX = "OpenXMLTemplate"
EX = xlOpenXMLTemplate: Exit Function
Case Is = "29" Or EX = "OpenXMLTemplateMacroEnabled"
EX = xlOpenXMLTemplateMacroEnabled: Exit Function
Case Is = "30" Or EX = "OpenXMLWorkbook"
EX = xlOpenXMLWorkbook: Exit Function
Case Is = "31" Or EX = "OpenXMLWorkbookMacroEnabled"
EX = xlOpenXMLWorkbookMacroEnabled: Exit Function
Case Is = "32" Or EX = "SYLK"
EX = xlSYLK: Exit Function
Case Is = "33" Or EX = "Template"
EX = xlTemplate: Exit Function
Case Is = "34" Or EX = "Template8"
EX = xlTemplate8: Exit Function
Case Is = "35" Or EX = "TextMac"
EX = xlTextMac: Exit Function
Case Is = "36" Or EX = "TextMSDOS"
EX = xlTextMSDOS: Exit Function
Case Is = "37" Or EX = "TextPrinter"
EX = xlTextPrinter: Exit Function
Case Is = "38" Or EX = "TextWindows"
EX = xlTextWindows: Exit Function
Case Is = "39" Or EX = "UnicodeText"
EX = xlUnicodeText: Exit Function
Case Is = "40" Or EX = "WebArchive"
EX = xlWebArchive: Exit Function
Case Is = "41" Or EX = "WJ2WD1"
EX = xlWJ2WD1: Exit Function
Case Is = "42" Or EX = "WJ3"
EX = xlWJ3: Exit Function
Case Is = "43" Or EX = "WJ3FJ3"
EX = xlWJ3FJ3: Exit Function
Case Is = "44" Or EX = "WK1"
EX = xlWK1: Exit Function
Case Is = "45" Or EX = "WK1ALL"
EX = xlWK1ALL: Exit Function
Case Is = "46" Or EX = "WK1FMT"
EX = xlWK1FMT: Exit Function
Case Is = "47" Or EX = "WK3"
EX = xlWK3: Exit Function
Case Is = "48" Or EX = "WK3FM3"
EX = xlWK3FM3: Exit Function
Case Is = "49" Or EX = "WK4"
EX = xlWK4: Exit Function
Case Is = "50" Or EX = "WKS"
EX = xlWKS: Exit Function
Case Is = "51" Or EX = "WorkbookDefault"
EX = xlWorkbookDefault: Exit Function
Case Is = "52" Or EX = "WorkbookNormal"
EX = xlWorkbookNormal: Exit Function
Case Is = "53" Or EX = "Works2FarEast"
EX = xlWorks2FarEast: Exit Function
Case Is = "54" Or EX = "WQ1"
EX = xlWQ1: Exit Function
Case Is = "55" Or EX = "XMLSpreadsheet"
EX = xlXMLSpreadsheet: Exit Function

End Select

EX = "(*Err)"

End Function
Private Function basWebFilter(FX) As String

'//Check for web filter switch...
If InStr(1, FX, "-goog") Then GoTo FilGoogle
Exit Function

'/#########################\
'// Google search switchs #\\
'///#######################\\\
FilGoogle:
xSearchArr = Split(FX, "-goog", , vbTextCompare)
FX = xSearchArr(1)
If Left(FX, 1) = " " Then FX = Right(FX, Len(FX) - 1) '//remove leading space

'//Check for multi search...
If InStr(1, FX, ",") Then GoTo MultiGoogle

'//Image filter
    If InStr(1, FX, "-i") Then
    
    FX = Replace(FX, "-i", vbNullString)
    If Left(FX, 1) = " " Then FX = Right(FX, Len(FX) - 1) '//remove leading space
    xFil = "https://www.google.com/search?q=" & FX _
    & "^&hl=en^&tbm=isch"
    
'//Video filter
    ElseIf InStr(1, FX, "-v") Then
    
    FX = Replace(FX, "-v", vbNullString)
    If Left(FX, 1) = " " Then FX = Right(FX, Len(FX) - 1) '//remove leading space
    xFil = "https://www.google.com/search?q=" & FX _
    & "^&hl=en^&tbm=vid"
    
'//Book filter
    ElseIf InStr(1, FX, "-b") Then
    
    FX = Replace(FX, "-b", vbNullString)
    If Left(FX, 1) = " " Then FX = Right(FX, Len(FX) - 1) '//remove leading space
    xFil = "https://www.google.com/search?q=" & FX _
    & "^&hl=en^&tbm=bks"
    
'//Shop filter
    ElseIf InStr(1, FX, "-s") Then
    
    FX = Replace(FX, "-s", vbNullString)
    If Left(FX, 1) = " " Then FX = Right(FX, Len(FX) - 1) '//remove leading space
    xFil = "https://www.google.com/search?q=" & FX _
    & "^&hl=en^&tbm=shop"
    
'//News filter
    ElseIf InStr(1, FX, "-n") Then
    
    FX = Replace(FX, "-n", vbNullString)
    If Left(FX, 1) = " " Then FX = Right(FX, Len(FX) - 1) '//remove leading space
    xFil = "https://www.google.com/search?q=" & FX _
    & "^&hl=en^&tbm=nws"
    
'//Flight filter
    ElseIf InStr(1, FX, "-f") Then
    
    FX = Replace(FX, "-f", vbNullString)
    If Left(FX, 1) = " " Then FX = Right(FX, Len(FX) - 1) '//remove leading space
    xFil = "https://www.google.com/search?q=" & FX _
    & "^&hl=en^&tbm=flm"
        
            Else
            
            If Left(FX, 1) = " " Then FX = Right(FX, Len(FX) - 1) '//remove leading space
            xFil = "https://www.google.com/search?q=" & FX
        
                    End If
                    
                    FX = xSearchArr(0) & " " & xFil
                    
Exit Function
  
MultiGoogle:
'//Multi google search...
Dim X As Integer
X = 0

On Error GoTo NextStep

xMultiArr = Split(FX, ",")

Do Until X > UBound(xMultiArr)

'//Image filter
    If InStr(1, xMultiArr(X), "-i") Then
    
    Do Until xMultiArr(X) = vbNullString
    xMultiArr(X) = Replace(xMultiArr(X), "-i", vbNullString)
    If Left(xMultiArr(X), 1) = " " Then xMultiArr(X) = Right(xMultiArr(X), Len(xMultiArr(X)) - 1) '//remove leading space
    xMulti = xMulti & " " & """" & "https://www.google.com/search?q=" & xMultiArr(X) _
    & "^&hl=en^&tbm=isch" & """"
    X = X + 1
    Loop: GoTo NextStep
    End If
  
'//Video filter
    If InStr(1, xMultiArr(X), "-v") Then
    
    Do Until xMultiArr(X) = vbNullString
    xMultiArr(X) = Replace(xMultiArr(X), "-v", vbNullString)
    If Left(xMultiArr(X), 1) = " " Then xMultiArr(X) = Right(xMultiArr(X), Len(xMultiArr(X)) - 1) '//remove leading space
    xMulti = xMulti & " " & """" & "https://www.google.com/search?q=" & xMultiArr(X) _
    & "^&hl=en^&tbm=vid" & """"
    X = X + 1
    Loop: GoTo NextStep
    End If
    
'//Book filter
    If InStr(1, xMultiArr(X), "-b") Then
    
    Do Until xMultiArr(X) = vbNullString
    xMultiArr(X) = Replace(xMultiArr(X), "-b", vbNullString)
    If Left(xMultiArr(X), 1) = " " Then xMultiArr(X) = Right(xMultiArr(X), Len(xMultiArr(X)) - 1) '//remove leading space
    xMulti = xMulti & " " & """" & "https://www.google.com/search?q=" & xMultiArr(X) _
    & "^&hl=en^&tbm=bks" & """"
    X = X + 1
    Loop: GoTo NextStep
    End If
    
'//Shop filter
    If InStr(1, xMultiArr(X), "-s") Then
    
    Do Until xMultiArr(X) = vbNullString
    xMultiArr(X) = Replace(xMultiArr(X), "-s", vbNullString)
    If Left(xMultiArr(X), 1) = " " Then xMultiArr(X) = Right(xMultiArr(X), Len(xMultiArr(X)) - 1) '//remove leading space
    xMulti = xMulti & " " & """" & "https://www.google.com/search?q=" & xMultiArr(X) _
    & "^&hl=en^&tbm=shop" & """"
    X = X + 1
    Loop: GoTo NextStep
    End If
    
'//News filter
    If InStr(1, xMultiArr(X), "-n") Then
    
    Do Until xMultiArr(X) = vbNullString
    xMultiArr(X) = Replace(xMultiArr(X), "-n", vbNullString)
    If Left(xMultiArr(X), 1) = " " Then xMultiArr(X) = Right(xMultiArr(X), Len(xMultiArr(X)) - 1) '//remove leading space
    xMulti = xMulti & " " & """" & "https://www.google.com/search?q=" & xMultiArr(X) _
    & "^&hl=en^&tbm=nws" & """"
    X = X + 1
    Loop: GoTo NextStep
    End If
    
'//Flight filter
    If InStr(1, xMultiArr(X), "-f") Then
    
    Do Until xMultiArr(X) = vbNullString
    xMultiArr(X) = Replace(xMultiArr(X), "-f", vbNullString)
    If Left(xMultiArr(X), 1) = " " Then xMultiArr(X) = Right(xMultiArr(X), Len(xMultiArr(X)) - 1) '//remove leading space
    xMulti = xMulti & " " & """" & "https://www.google.com/search?q=" & xMultiArr(X) _
    & "^&hl=en^&tbm=flm" & """"
    X = X + 1
    Loop: GoTo NextStep
    End If
    
        If Left(xMultiArr(X), 1) = " " Then xMultiArr(X) = Right(xMultiArr(X), Len(xMultiArr(X)) - 1) '//remove leading space
        xMulti = xMulti & " " & """" & "https://www.google.com/search?q=" & xMultiArr(X) & """"
        
        
        X = X + 1
                        
        Loop
                                
                                
NextStep:
FX = xSearchArr(0) & " " & xMulti

If Left(FX, 1) = " " Then
Do Until Left(FX, 1) <> " "
FX = Right(FX, Len(FX) - 1) '//remove remaining lead spaces
Loop
End If

End Function
Private Function basColor(HX) As String

Dim xNotColor As String: Dim xRGBC As String
Dim X As Integer: Dim xCl As Integer
Dim I As Byte '//waste
xNotColor = "/NULL"

Retry:
X = 1
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "aliceblue;#F0F8FF;240,248,255", HX, vbTextCompare) Then xRGBCl = "aliceblue;#F0F8FF;240,248,255": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "antiquewhite;#FAEBD7;250,235,215", HX, vbTextCompare) Then xRGBCl = "antiquewhite;#FAEBD7;250,235,215": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "aqua;#00FFFF;0,255,255", HX, vbTextCompare) Then xRGBCl = "aqua;#00FFFF;0,255,255": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "aquamarine;#7FFFD4;127,255,212", HX, vbTextCompare) Then xRGBCl = "aquamarine;#7FFFD4;127,255,212": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "azure;#F0FFFF;240,255,255", HX, vbTextCompare) Then xRGBCl = "azure;#F0FFFF;240,255,255": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "beige;#F5F5DC;245,245,220", HX, vbTextCompare) Then xRGBCl = "beige;#F5F5DC;245,245,220": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "bisque;#FFE4C4;255,228,196", HX, vbTextCompare) Then xRGBCl = "bisque;#FFE4C4;255,228,196": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "black;#000000;0,0,0", HX, vbTextCompare) Then xRGBCl = "black;#000000;0,0,0": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "blanchedalmond;#FFEBCD;255,235,205", HX, vbTextCompare) Then xRGBCl = "blanchedalmond;#FFEBCD;255,235,205": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "blue;#0000FF;0,0,255", HX, vbTextCompare) Then xRGBCl = "blue;#0000FF;0,0,255": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "blueviolet;#8A2BE2;138,43,226", HX, vbTextCompare) Then xRGBCl = "blueviolet;#8A2BE2;138,43,226": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "brown;#A52A2A;165,42,42", HX, vbTextCompare) Then xRGBCl = "brown;#A52A2A;165,42,42": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "burlywood;#DEB887;222,184,135", HX, vbTextCompare) Then xRGBCl = "burlywood;#DEB887;222,184,135": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "cadetblue;#5F9EA0;95,158,160", HX, vbTextCompare) Then xRGBCl = "cadetblue;#5F9EA0;95,158,160": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "chartreuse;#7FFF00;127,255,0", HX, vbTextCompare) Then xRGBCl = "chartreuse;#7FFF00;127,255,0": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "chocolate;#D2691E;210,105,30", HX, vbTextCompare) Then xRGBCl = "chocolate;#D2691E;210,105,30": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "coral;#FF7F50;255,127,80", HX, vbTextCompare) Then xRGBCl = "coral;#FF7F50;255,127,80": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "cornflowerblue;#6495ED;100,149,237", HX, vbTextCompare) Then xRGBCl = "cornflowerblue;#6495ED;100,149,237": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "cornsilk;#FFF8DC;255,248,220", HX, vbTextCompare) Then xRGBCl = "cornsilk;#FFF8DC;255,248,220": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "crimson;#DC143C;220,20,60", HX, vbTextCompare) Then xRGBCl = "crimson;#DC143C;220,20,60": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "cyan;#00FFFF;0,255,255", HX, vbTextCompare) Then xRGBCl = "cyan;#00FFFF;0,255,255": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dblue;#00008B;0,0,139", HX, vbTextCompare) Then xRGBCl = "dblue;#00008B;0,0,139": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dcyan;#008B8B;0,139,139", HX, vbTextCompare) Then xRGBCl = "dcyan;#008B8B;0,139,139": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "deeppink;#FF1493;255,20,147", HX, vbTextCompare) Then xRGBCl = "deeppink;#FF1493;255,20,147": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "deepskyblue;#00BFFF;0,191,255", HX, vbTextCompare) Then xRGBCl = "deepskyblue;#00BFFF;0,191,255": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dgoldenrod;#B8860B;184,134,11", HX, vbTextCompare) Then xRGBCl = "dgoldenrod;#B8860B;184,134,11": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dgray;#A9A9A9;169,169,169", HX, vbTextCompare) Then xRGBCl = "dgray;#A9A9A9;169,169,169": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dgreen;#006400;0,100,0", HX, vbTextCompare) Then xRGBCl = "dgreen;#006400;0,100,0": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dimgray;#696969;105,105,105", HX, vbTextCompare) Then xRGBCl = "dimgray;#696969;105,105,105": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dkhaki;#BDB76B;189,183,107", HX, vbTextCompare) Then xRGBCl = "dkhaki;#BDB76B;189,183,107": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dmagenta;#8B008B;139,0,139", HX, vbTextCompare) Then xRGBCl = "dmagenta;#8B008B;139,0,139": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dodgerblue;#1E90FF;30,144,255", HX, vbTextCompare) Then xRGBCl = "dodgerblue;#1E90FF;30,144,255": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dolivegreen;#556B2F;85,107,47", HX, vbTextCompare) Then xRGBCl = "dolivegreen;#556B2F;85,107,47": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dorange;#FF8C00;255,140,0", HX, vbTextCompare) Then xRGBCl = "dorange;#FF8C00;255,140,0": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dorchid;#9932CC;153,50,204", HX, vbTextCompare) Then xRGBCl = "dorchid;#9932CC;153,50,204": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dred;#8B0000;139,0,0", HX, vbTextCompare) Then xRGBCl = "dred;#8B0000;139,0,0": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dsalmon;#E9967A;233,150,122", HX, vbTextCompare) Then xRGBCl = "dsalmon;#E9967A;233,150,122": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dseagreen;#8FBC8F;143,188,143", HX, vbTextCompare) Then xRGBCl = "dseagreen;#8FBC8F;143,188,143": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dslateblue;#483D8B;72,61,139", HX, vbTextCompare) Then xRGBCl = "dslateblue;#483D8B;72,61,139": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dslategray;#2F4F4F;47,79,79", HX, vbTextCompare) Then xRGBCl = "dslategray;#2F4F4F;47,79,79": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dturquoise;#00CED1;0,206,209", HX, vbTextCompare) Then xRGBCl = "dturquoise;#00CED1;0,206,209": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "dviolet;#9400D3;148,0,211", HX, vbTextCompare) Then xRGBCl = "dviolet;#9400D3;148,0,211": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "firebrick;#B22222;178,34,34", HX, vbTextCompare) Then xRGBCl = "firebrick;#B22222;178,34,34": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "floralwhite;#FFFAF0;255,250,240", HX, vbTextCompare) Then xRGBCl = "floralwhite;#FFFAF0;255,250,240": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "forestgreen;#228B22;34,139,34", HX, vbTextCompare) Then xRGBCl = "forestgreen;#228B22;34,139,34": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "gainsboro;#DCDCDC;220,220,220", HX, vbTextCompare) Then xRGBCl = "gainsboro;#DCDCDC;220,220,220": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "ghostwhite;#F8F8FF;248,248,255", HX, vbTextCompare) Then xRGBCl = "ghostwhite;#F8F8FF;248,248,255": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "gold;#FFD700;255,215,0", HX, vbTextCompare) Then xRGBCl = "gold;#FFD700;255,215,0": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "goldenrod;#DAA520;218,165,32", HX, vbTextCompare) Then xRGBCl = "goldenrod;#DAA520;218,165,32": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "gray;#808080;128,128,128", HX, vbTextCompare) Then xRGBCl = "gray;#808080;128,128,128": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "green;#008000;0,128,0", HX, vbTextCompare) Then xRGBCl = "green;#008000;0,128,0": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "greenyellow;#ADFF2F;173,255,47", HX, vbTextCompare) Then xRGBCl = "greenyellow;#ADFF2F;173,255,47": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "honeydew;#F0FFF0;240,255,240", HX, vbTextCompare) Then xRGBCl = "honeydew;#F0FFF0;240,255,240": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "hotpink;#FF69B4;255,105,180", HX, vbTextCompare) Then xRGBCl = "hotpink;#FF69B4;255,105,180": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "indigo;#4B0082;75,0,130", HX, vbTextCompare) Then xRGBCl = "indigo;#4B0082;75,0,130": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "ivory;#FFFFF0;255,255,240", HX, vbTextCompare) Then xRGBCl = "ivory;#FFFFF0;255,255,240": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "khaki;#F0E68C;240,230,140", HX, vbTextCompare) Then xRGBCl = "khaki;#F0E68C;240,230,140": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lavender;#E6E6FA;230,230,250", HX, vbTextCompare) Then xRGBCl = "lavender;#E6E6FA;230,230,250": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lavenderblush;#FFF0F5;255,240,245", HX, vbTextCompare) Then xRGBCl = "lavenderblush;#FFF0F5;255,240,245": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lawngreen;#7CFC00;124,252,0", HX, vbTextCompare) Then xRGBCl = "lawngreen;#7CFC00;124,252,0": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lblue;#ADD8E6;173,216,230", HX, vbTextCompare) Then xRGBCl = "lblue;#ADD8E6;173,216,230": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lcoral;#F08080;240,128,128", HX, vbTextCompare) Then xRGBCl = "lcoral;#F08080;240,128,128": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lcyan;#E0FFFF;224,255,255", HX, vbTextCompare) Then xRGBCl = "lcyan;#E0FFFF;224,255,255": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lemonchiffon;#FFFACD;255,250,205", HX, vbTextCompare) Then xRGBCl = "lemonchiffon;#FFFACD;255,250,205": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lgray;#D3D3D3;211,211,211", HX, vbTextCompare) Then xRGBCl = "lgray;#D3D3D3;211,211,211": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lgreen;#90EE90;144,238,144", HX, vbTextCompare) Then xRGBCl = "lgreen;#90EE90;144,238,144": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lime;#00FF00;0,255,0", HX, vbTextCompare) Then xRGBCl = "lime;#00FF00;0,255,0": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "limegreen;#32CD32;50,205,50", HX, vbTextCompare) Then xRGBCl = "limegreen;#32CD32;50,205,50": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "linen;#FAF0E6;250,240,230", HX, vbTextCompare) Then xRGBCl = "linen;#FAF0E6;250,240,230": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lpink;#FFB6C1;255,182,193", HX, vbTextCompare) Then xRGBCl = "lpink;#FFB6C1;255,182,193": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lrodyellow;#FAFAD2;250,250,210", HX, vbTextCompare) Then xRGBCl = "lrodyellow;#FAFAD2;250,250,210": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lsalmon;#FFA07A;255,160,122", HX, vbTextCompare) Then xRGBCl = "lsalmon;#FFA07A;255,160,122": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lseagreen;#20B2AA;32,178,170", HX, vbTextCompare) Then xRGBCl = "lseagreen;#20B2AA;32,178,170": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lskyblue;#87CEFA;135,206,250", HX, vbTextCompare) Then xRGBCl = "lskyblue;#87CEFA;135,206,250": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lslategray;#778899;119,136,153", HX, vbTextCompare) Then xRGBCl = "lslategray;#778899;119,136,153": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lsteelblue;#B0C4DE;176,196,222", HX, vbTextCompare) Then xRGBCl = "lsteelblue;#B0C4DE;176,196,222": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "lyellow;#FFFFE0;255,255,224", HX, vbTextCompare) Then xRGBCl = "lyellow;#FFFFE0;255,255,224": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "magenta;#FF00FF;255,0,255", HX, vbTextCompare) Then xRGBCl = "magenta;#FF00FF;255,0,255": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "maquamarine;#66CDAA;102,205,170", HX, vbTextCompare) Then xRGBCl = "maquamarine;#66CDAA;102,205,170": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "mblue;#0000CD;0,0,205", HX, vbTextCompare) Then xRGBCl = "mblue;#0000CD;0,0,205": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "midnightblue;#191970;25,25,112", HX, vbTextCompare) Then xRGBCl = "midnightblue;#191970;25,25,112": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "mintcream;#F5FFFA;245,255,250", HX, vbTextCompare) Then xRGBCl = "mintcream;#F5FFFA;245,255,250": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "mistyrose;#FFE4E1;255,228,225", HX, vbTextCompare) Then xRGBCl = "mistyrose;#FFE4E1;255,228,225": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "moccasin;#FFE4B5;255,228,181", HX, vbTextCompare) Then xRGBCl = "moccasin;#FFE4B5;255,228,181": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "morchid;#BA55D3;186,85,211", HX, vbTextCompare) Then xRGBCl = "morchid;#BA55D3;186,85,211": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "mpurple;#9370DB;147,112,219", HX, vbTextCompare) Then xRGBCl = "mpurple;#9370DB;147,112,219": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "mseagreen;#3CB371;60,179,113", HX, vbTextCompare) Then xRGBCl = "mseagreen;#3CB371;60,179,113": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "mslateblue;#7B68EE;123,104,238", HX, vbTextCompare) Then xRGBCl = "mslateblue;#7B68EE;123,104,238": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "mspringgreen;#00FA9A;0,250,154", HX, vbTextCompare) Then xRGBCl = "mspringgreen;#00FA9A;0,250,154": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "mturquoise;#48D1CC;72,209,204", HX, vbTextCompare) Then xRGBCl = "mturquoise;#48D1CC;72,209,204": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "mvioletred;#C71585;199,21,133", HX, vbTextCompare) Then xRGBCl = "mvioletred;#C71585;199,21,133": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "navajowhite;#FFDEAD;255,222,173", HX, vbTextCompare) Then xRGBCl = "navajowhite;#FFDEAD;255,222,173": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "navy;#000080;0,0,128", HX, vbTextCompare) Then xRGBCl = "navy;#000080;0,0,128": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "oldlace;#FDF5E6;253,245,230", HX, vbTextCompare) Then xRGBCl = "oldlace;#FDF5E6;253,245,230": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "olive;#808000;128,128,0", HX, vbTextCompare) Then xRGBCl = "olive;#808000;128,128,0": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "olivedrab;#6B8E23;107,142,35", HX, vbTextCompare) Then xRGBCl = "olivedrab;#6B8E23;107,142,35": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "orange;#FFA500;255,165,0", HX, vbTextCompare) Then xRGBCl = "orange;#FFA500;255,165,0": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "orangered;#FF4500;255,69,0", HX, vbTextCompare) Then xRGBCl = "orangered;#FF4500;255,69,0": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "orchid;#DA70D6;218,112,214", HX, vbTextCompare) Then xRGBCl = "orchid;#DA70D6;218,112,214": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "palegoldenrod;#EEE8AA;238,232,170", HX, vbTextCompare) Then xRGBCl = "palegoldenrod;#EEE8AA;238,232,170": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "palegreen;#98FB98;152,251,152", HX, vbTextCompare) Then xRGBCl = "palegreen;#98FB98;152,251,152": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "paleturquoise;#AFEEEE;175,238,238", HX, vbTextCompare) Then xRGBCl = "paleturquoise;#AFEEEE;175,238,238": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "palevioletred;#DB7093;219,112,147", HX, vbTextCompare) Then xRGBCl = "palevioletred;#DB7093;219,112,147": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "papayawhip;#FFEFD5;255,239,213", HX, vbTextCompare) Then xRGBCl = "papayawhip;#FFEFD5;255,239,213": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "peachpuff;#FFDAB9;255,218,185", HX, vbTextCompare) Then xRGBCl = "peachpuff;#FFDAB9;255,218,185": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "peru;#CD853F;205,133,63", HX, vbTextCompare) Then xRGBCl = "peru;#CD853F;205,133,63": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "pink;#FFC0CB;255,192,203", HX, vbTextCompare) Then xRGBCl = "pink;#FFC0CB;255,192,203": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "plum;#DDA0DD;221,160,221", HX, vbTextCompare) Then xRGBCl = "plum;#DDA0DD;221,160,221": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "powderblue;#B0E0E6;176,224,230", HX, vbTextCompare) Then xRGBCl = "powderblue;#B0E0E6;176,224,230": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "purple;#800080;128,0,128", HX, vbTextCompare) Then xRGBCl = "purple;#800080;128,0,128": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "red;#FF0000;255,0,0", HX, vbTextCompare) Then xRGBCl = "red;#FF0000;255,0,0": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "rosybrown;#BC8F8F;188,143,143", HX, vbTextCompare) Then xRGBCl = "rosybrown;#BC8F8F;188,143,143": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "royalblue;#4169E1;65,105,225", HX, vbTextCompare) Then xRGBCl = "royalblue;#4169E1;65,105,225": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "saddlebrown;#8B4513;139,69,19", HX, vbTextCompare) Then xRGBCl = "saddlebrown;#8B4513;139,69,19": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "salmon;#FA8072;250,128,114", HX, vbTextCompare) Then xRGBCl = "salmon;#FA8072;250,128,114": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "sandybrown;#F4A460;244,164,96", HX, vbTextCompare) Then xRGBCl = "sandybrown;#F4A460;244,164,96": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "seagreen;#2E8B57;46,139,87", HX, vbTextCompare) Then xRGBCl = "seagreen;#2E8B57;46,139,87": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "seashell;#FFF5EE;255,245,238", HX, vbTextCompare) Then xRGBCl = "seashell;#FFF5EE;255,245,238": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "sienna;#A0522D;160,82,45", HX, vbTextCompare) Then xRGBCl = "sienna;#A0522D;160,82,45": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "silver;#C0C0C0;192,192,192", HX, vbTextCompare) Then xRGBCl = "silver;#C0C0C0;192,192,192": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "skyblue;#87CEEB;135,206,235", HX, vbTextCompare) Then xRGBCl = "skyblue;#87CEEB;135,206,235": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "slateblue;#6A5ACD;106,90,205", HX, vbTextCompare) Then xRGBCl = "slateblue;#6A5ACD;106,90,205": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "slategray;#708090;112,128,144", HX, vbTextCompare) Then xRGBCl = "slategray;#708090;112,128,144": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "snow;#FFFAFA;255,250,250", HX, vbTextCompare) Then xRGBCl = "snow;#FFFAFA;255,250,250": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "springgreen;#00FF7F;0,255,127", HX, vbTextCompare) Then xRGBCl = "springgreen;#00FF7F;0,255,127": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "steelblue;#4682B4;70,130,180", HX, vbTextCompare) Then xRGBCl = "steelblue;#4682B4;70,130,180": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "tan;#D2B48C;210,180,140", HX, vbTextCompare) Then xRGBCl = "tan;#D2B48C;210,180,140": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "teal;#008080;0,128,128", HX, vbTextCompare) Then xRGBCl = "teal;#008080;0,128,128": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "thistle;#D8BFD8;216,191,216", HX, vbTextCompare) Then xRGBCl = "thistle;#D8BFD8;216,191,216": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "tomato;#FF6347;255,99,71", HX, vbTextCompare) Then xRGBCl = "tomato;#FF6347;255,99,71": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "turquoise;#40E0D0;64,224,208", HX, vbTextCompare) Then xRGBCl = "turquoise;#40E0D0;64,224,208": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "violet;#EE82EE;238,130,238", HX, vbTextCompare) Then xRGBCl = "violet;#EE82EE;238,130,238": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "wheat;#F5DEB3;245,222,179", HX, vbTextCompare) Then xRGBCl = "wheat;#F5DEB3;245,222,179": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "white;#FFFFFF;255,255,255", HX, vbTextCompare) Then xRGBCl = "white;#FFFFFF;255,255,255": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "whitesmoke;#F5F5F5;245,245,245", HX, vbTextCompare) Then xRGBCl = "whitesmoke;#F5F5F5;245,245,245": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "yellow;#FFFF00;255,255,0", HX, vbTextCompare) Then xRGBCl = "yellow;#FFFF00;255,255,0": GoTo ColorFound
X = X + 1: If X > (xCl) Then I = I: If InStr(1, "yellowgreen;#9ACD32;154,205,50", HX, vbTextCompare) Then xRGBCl = "yellowgreen;#9ACD32;154,205,50": GoTo ColorFound

Exit Function

ColorFound:

xRGBArr = Split(xRGBCl, ";"): If HX = xRGBArr(0) Then HX = xRGBArr(2): Exit Function

'//color not found
xCl = X
xNotColor = xRGBArr(2): GoTo Retry

End Function
Private Function basBorder(BX) As Long

'//check for border type
Select Case BX

Case Is = 0
BX = xlNone: Exit Function
Case Is = 1
BX = xlDiagonalDown: Exit Function
Case Is = 2
BX = xlDiagonalUp: Exit Function
Case Is = 3
BX = xlEdgeBottom: Exit Function
Case Is = 4
BX = xlEdgeLeft: Exit Function
Case Is = 5
BX = xlEdgeRight: Exit Function
Case Is = 6
BX = xlEdgeTop: Exit Function
Case Is = 7
BX = xlInsideHorizontal: Exit Function

Case Is = "none"
BX = xlNone: Exit Function
Case Is = "ddown"
BX = xlDiagonalDown: Exit Function
Case Is = "dup"
BX = xlDiagonalUp: Exit Function
Case Is = "bottom"
BX = xlEdgeBottom: Exit Function
Case Is = "left"
BX = xlEdgeLeft: Exit Function
Case Is = "right"
BX = xlEdgeRight: Exit Function
Case Is = "top"
BX = xlEdgeTop: Exit Function
Case Is = "inside"
BX = xlInsideHorizontal: Exit Function

End Select

End Function
Private Function basBorderStyle(SX) As Long

'//check for border style
Select Case SX

Case Is = 0
SX = xlNone: Exit Function
Case Is = 1
SX = xlContinuous: Exit Function
Case Is = 2
SX = xlDash: Exit Function
Case Is = 3
SX = xlDot: Exit Function
Case Is = 4
SX = xlDashDot: Exit Function
Case Is = 5
SX = xlDashDotDot: Exit Function
Case Is = 6
SX = xlSlantDashDot: Exit Function
Case Is = 7
SX = xlDouble: Exit Function

Case Is = "none"
SX = xlNone: Exit Function
Case Is = "line"
SX = xlContinuous: Exit Function
Case Is = "dash"
SX = xlDash: Exit Function
Case Is = "dot"
SX = xlDot: Exit Function
Case Is = "ddot"
SX = xlDashDot: Exit Function
Case Is = "ddotdot"
SX = xlDashDotDot: Exit Function
Case Is = "sddot"
SX = xlSlantDashDot: Exit Function
Case Is = "double"
SX = xlDouble: Exit Function

End Select

End Function
Private Function basCompare(CX) As Long

'//check for comparison type
Select Case CX

Case Is = 0
CX = vbBinaryCompare: Exit Function
Case Is = 1
CX = vbDatabaseCompare: Exit Function
Case Is = 2
CX = vbTextCompare: Exit Function

End Select

End Function
Private Function basPattern(PX) As Long
 
'//check for pattern
Select Case PX

Case Is = 0
PX = xlNone: Exit Function
Case Is = 1
PX = xlPatternChecker: Exit Function
Case Is = 2
PX = xlPatternCrissCross: Exit Function
Case Is = 3
PX = xlPatternDown: Exit Function
Case Is = 4
PX = xlPatternHorizontal: Exit Function
Case Is = 5
PX = xlPatternLightDown: Exit Function
Case Is = 6
PX = xlPatternLightHorizontal: Exit Function
Case Is = 7
PX = xlPatternLightUp: Exit Function
Case Is = 8
PX = xlPatternLightVertical: Exit Function
Case Is = 9
PX = xlPatternUp: Exit Function

Case Is = "none"
PX = xlNone: Exit Function
Case Is = "pcheck"
PX = xlPatternChecker: Exit Function
Case Is = "pcross"
PX = xlPatternCrissCross: Exit Function
Case Is = "pdown"
PX = xlPatternDown: Exit Function
Case Is = "phori"
PX = xlPatternHorizontal: Exit Function
Case Is = "pldown"
PX = xlPatternLightDown: Exit Function
Case Is = "plhori"
PX = xlPatternLightHorizontal: Exit Function
Case Is = "plup"
PX = xlPatternLightUp: Exit Function
Case Is = "plvert"
PX = xlPatternLightVertical: Exit Function
Case Is = "pup"
PX = xlPatternUp: Exit Function

End Select

End Function
Private Function basWasteTime(ByVal T As Byte) As Byte

T = T + 1: T = T - 1
DoEvents

End Function
Public Function disableWbUpdates() As Byte

Application.DisplayStatusBar = False
Application.EnableAnimations = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

End Function
Public Function enableWbUpdates() As Byte

Application.DisplayStatusBar = True
Application.EnableAnimations = True
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic

End Function

'//=========================================================================================================================
'//
'//         CHANGE LOG
'/\_________________________________________________________________________________________________________________________
'
'
' Version 1.0.8
'
' [ Date: 4/24/2022 ]
'
'      /``.
'     []
'   _[__]_
'  [______]
' [________] Happy Birthday André! :)
'
'
'
'
' (1): Fixed "q().stop" article not parsing the application extracted to stop
'
'
' [ Date: 4/23/2022 ]
'
' (1): added "--time" switch for getting the current time as a string
'
'
'
' [ Date: 4/21/2022 ]
'
'
' (1): Added extended syntax functionality for "cell()", "rng()", & "sel()" article modifiers
'
' These article modifiers can now be typed like:
'
' cell(1,1).value @var; or cell(1,1).value(@var);
'
' rng(a1).bgcolor @var1 .fcolor @var2; or rng(a1).bgcolor(@var1).fcolor(@var2);
'
' (2): fixed issues w/ "rng().read" sequence not returning range values
'
' (3): Prefixed library functions w/ "bas"
'
' (4): Added "--enableerr" & "--disableerr" flags for setting the library error level
'
' (5): Added ".move" modifier for "q()" article to query & move files/folders
'
'
'
' Version: 1.0.7
'
'
' [ Date: 4/18/2022 ]
'
' (1): Added "--enableglobal" & "--disableglobal" flags for triggering Global Control variables active/inactive
'
'
' [ Date: 4/17/2022 ]
'
'
' (1): Added "-loose" & "-strict" switch to "q()" article for finding binary or text string match to file/folder path provided if an
' exact not found.
'
' (2): Added "-file" & "-fldr" switches to "q()" article for assiging query search as a file or folder
'
'
' [ Date: 4/15/2022 ]
'
' (1): Added "click()" article for assigning mouse clicks/positioning
'
' Example(s):
'
' click(-double 500,500); <--- this will double click at positon 500,500 on the screen
' click(-leftdown 500,500); <--- this will left click and hold at positon 500,500 on the screen
'
'
' Version: 1.0.6
'
' [ Date: 4/3/2022 ]
'
' (1): Added "read()" article for capturing file text
'
' Article switches:
'
' -all = read all file text
' -nlall = read all file text & seperate each line
' -string = read first occurence of file text w/ this string only
' -xstring = read all file text until first occurence of this string
' -nlstring = read all file text until first occurence of this string & seperate each line
' -line = read this line of file text only
' -xline = read all file text until this line
' -nlline = read all file text until this line & seperate each line
'
' ***Example:
'
' read(-all @filePath); <--- This will read all file text w/o seperating each line
' read(-string @findStr, @filePath); <--- This will read all file text up until a string is found & return that line
' read(-nlstring @findStr, @filePath); <--- This will read all file text up until a string is found & return everything up until then
' w/ each line seperated
' read(-line 5, @filePath); <--- This will read line 5 of the file text
'
'
' [ Date: 4/2/2022 ]
'
' (1): Modified "msg()" article to return input selection value to a variable
'
' (2): Added "incr()" article to allow holding an incremented number based on an assigned value & operation.
'
' ***Example: @var = incr(+1); = 2 | @var = incr(-1) = -2 (These values will stay the same if looped through)
'
'
'
' [ Date: 4/1/2022 ]
'
' (1): Fixed issue w/ "rng().read" command causing an error
'
' (2): Made adjustments to "input()" article so its possible to return user input to a variable
'
' (3): Added ".sel" parameter to "rng()" & "cell()" articles
'
'
'
'
' Version: 1.0.5
'
' [ Date: 3/17/2022 ]
'
' (1): Added "--enablecontain" & "--disablecontain" flags for setting "Local Contain" on/off
' ***Local Contain allows previously used runtime memory addresses to retain their data
'
' [ Date: 3/13/2022 ]
'
' (1): Added "--enablestatic" & "--disablestatic" flags for setting "Local Static" on/off
' ***Local Static allows libraries to stay locked to the current runtime environment session.
'
' [ Date: 3/11/2022 ]
'
' (1): Added base articles "build()", "printer()", & "name()" & respective enhanced
' "app.build()", "app.printer()", & "app.name()" articles
'
' (2): Added "cell()" article for analyzing/modifying cells
'
'
' [ Date: 3/9/2022 ]
'
' (1): Added "fil()" & "dir()" articles & reworked them to include enhancers "mk." & "del."
'
'
' [ Date: 3/8/2022 ]
'
' (1): Added "rng().name" & "sel().name" articles for setting cell names
' ***You can alternatively clear a cell name by leaving it blank
'
' (2): Added "wb().delname" article for deleting a specific cell name
'
' (3): Added "-me" switch to "wb()" article for expanding the workbook name at runtime
'
'
'
' Version: 1.0.4
'
'
' [ Date: 3/2/2022 ]
'
' (1): Added "ptrim()" article to deal w/ removing the starting & ending parentheses of a string.
'
' (2): Added "wastetime()" article to halt the parser while still allowing for user input as oppossed to the "wait()"
' article which halts the parser but also freezes the environment.
'
'***Even though the environment becomes frozen, the "wait()" article is more precise than the wastetime() article for time
'
' wastetime(100) = approx. 1 second | wastetime(1000) = approx. 10 seconds
'
' (3): Updated "winform()" article to include "-last" switch for setting the last WinForm as the current
'
'
' [ Date: 3/1/2022 ]
'
' (1): Added "echo()" & "host()" articles.
'
' echo() = output string using cmd host | host() = output string using vba host (msgbox)
'
' ***echo() supports different window focuses by supplying a value from (0-6) ---> echo(2)(@strToShow); <--- maximized view
'
' (2): Added "conv()" & "xtrim()" articles.
'
' conv(@str, -upper); = convert char or string to a desired case (uppercase, lowercase, etc.)
'
' xtrim(@str, ":"); = remove first & last characters from string by a desired character
'
' (3): Added "lptrim", "rptrim()" & "qtrim()" articles (due to trim() article removing parentheses & quotations during parsing).
'
' lptrim(@str); = remove first & last left facing parentheses from a string
'
' rptrim(@str); = remove first & last right facing parentheses from a string
'
' qtrim(@str); = remove first & last quotes from a string
'
' (4): Added "strrev()" article to reverse strings
'
' strrev(@str); <--- string will be backwards
'
'
' [ Date: 2/28/2022 ]
'
' (1): Updated library to utilize article cleaning function within lexer
'
' (2): Included additional updates made to lexer as well as addition of the "runtime block"
'
'
'
'
' Version: 1.0.3
'
' [ Date: 2/26/2022 ]
'
' (1): Fixed issue w/ "wb().hd" & "wb().sh" articles not parsing w/ correct syntax
'
' [ Date: 2/24/2022 ]
'
' (1): Added "++e"  & "--e" switches for control over enabling/disabling workbook updates during runtime
'
'
' [ Date: 2/23/2022 ]
'
' (1): Removed "colors.txt" file & instead created a function that essentially acts the same way where a color is searched
' for within a list of color name/hex/rgb values.
'
'
' [ Date: 2/10/2022 ]
'
' (1): Fixed an issue w/ key() article leaving leading & ending quotations on the supplied keystroke when parsed.
'
'
' [ Date: 2/8/2022 ]
'
' (1): Changed "WINDOW FORM ARTICLES" labeling to "WINFORM ARTICLES"
'
' (2): Changed "form()" article to "winform()" for readability.
'
'
' [ Date: 2/5/2022 ]
'
' (1): Added "form()" article to "xbas" library for manually setting an application Window
'
'
' Version: 1.0.2
'
'
' [ Date: 1/31/2022 ]
'
' (1) Fixed issue w/ key() article leaving "key" in output
'
'
'
' Version: 1.0.1
'
'
' [ Date: 1/6/2022 ]
'
' (1) Added "ins()" article to find a char/string within another string/variable
'
' ***Will return "TRUE" or "FALSE"  to the assigned variable based on if the char/string searched for was found or not
'
' Syntax: @var = ins(@startPosition, @strToSearch, @strToFind, @compType)
'
' (2) Added "app.run(") or simply the "run()" article to allow for running a module within an opened workbook
'
' ***Currently only supports a single listed (,) argument
'
' Syntax: app.run(moduleName.subName, (arg)); also run(moduleName.subName, arg);
'
' (3) Added pc power articles as well as a copy-paste articles
'
' ***pc.shutdown & pc.reboot articles accept the "-e" switch for auto logging in & bringing up the
' previous session on start-up, & "-t" for setting a timer before shutdown.
'
' pc.copy&() = copy & paste file or folder
' pc.copy&!() = copy, paste, & delete copied file or folder
' pc.shutdown() = shutdown pc
' pc.off() = logoff
' pc.rest() = set pc to rest mode
' pc.reboot() = restart pc
' pc.clr = clear shutdown queue
'
'
'
' [ Date: 1/5/2022 ]
'
' (1) Added "repl()" article to replace a value within a string
'
' Syntax: @var = repl(@strToReplace, @strToFind, @strToReplace, @compType)
'
' ***If using 3 parameters like: @var = repl(@strToReplace, @strToFind, @strToReplace) the default comparison method
' will be binary.
'
'
' (2) Added "dfil()" & "ddir()" articles to delete files/folders
'
' ***ddir() will only delete a directory if it's completely empty, so in that instance you could
' use the del.dir() article instead to remove everything.
'
'
'
' [ Date: 1/4/2022 ]
'
' (1): Changed all replace & string check commands for articles to ignore case
'
' ***User can type sh( or SH(, q(, or Q(, etc. & that will be accepted as the same article
'
'
' [ Date: 1/3/2022 ]
'
' (1): Added "q()" article which allows the user to query search either a file or folder (depending on the (.)extension.
' W/ this command you can check for the existance, open, delete, or taskkill a file/folder.
'
' ***q() command is able to search through a total of 3 sets of directories starting from a local drive & base folder.
'
' User only needs to include the drive & base folder.
'
' Examples of drive & base folder:
' C:\Users\ <----
' C:\Windows\ <----
'
' Syntax: q(C:\Users\@fileToQuery).exists (this will prompt whether a file exists or not & it's location)
'
'
' [ Date: 1/2/2022 ]
'
' (1): Added change log, license information & library requirements. Edited library description.
'
' (2): Adjusted "key()" article so it could be split into 7 locations & variables based on a numbered reference (0-6)
' This helped stop collisions w/ VBA, & VBS when the VBA parser ran quicker than the variable was released from the previous run
'
' (You'd likely only come across this issue when trying to run consecutive key() articles w/o using a wait() offset in-between).
'
' ***Numbered references will also be attributed to a corresponding VBA shell mode (0-6) when ran
'
' (3): Shell mode can now be set for sh() article (shell modes (0-6) will correspond w/ the same VBA shell modes (0-6)).
'
' (4): When opening, activating, & "saving as" a workbook, the application environment will be linked to that workbook.
' (This helped w/ navigating back to the original application environment when performing those actions due to the newly
' opened, activated, or saved workbook now being the one activated)
'
' ***Linking is simply just relaying to the same cell (memory location), name ("xlasEnvironment"), & value (current runtime environment) to the currently
' activated workbook.
'



