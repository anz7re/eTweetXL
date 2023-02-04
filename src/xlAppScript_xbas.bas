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
Public Function runLib$(Token)
'/\_____________________________________________________________________________________________________________________________
'//
'//     xbas (basic) Library
'//        Version: 1.1.5
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
'//     Latest Revision: 2/1/2023
'/\_____________________________________________________________________________________________________________________________
'//
'//     Developer(s): anz7re (André)
'//     Contact: support@xlappscript.org | support@autokit.tech | anz7re@autokit.tech
'//     Web: xlappscript.org | autokit.tech
'/\_____________________________________________________________________________________________________________________________
        
        '//Library variable declarations
        Dim oFSO As Object: Dim oDrv As Object: Dim oFile As Object: Dim oSubFldr As Object: Dim oShell As Object: Dim oShellItem As Object: Dim xWin As Object
        Dim appEnv As String: Dim appBlk As String: Dim FX As String: Dim HX As String
        Dim sysShell As String: Dim wbMacro As String: Dim ArtH As String
        Dim xArg As String: Dim xDir As String: Dim xCell As String: Dim xExt As String: Dim xMod As String: Dim xWb As String: Dim xVar As String: Dim xVar2 As String
        Dim TokenArr() As String:  Dim TokenArrH() As String: Dim xExtArr() As String: Dim xRGBArr() As String
        Dim BX As Long: Dim EX As Long: Dim CX As Long: Dim PX As Long: Dim SX As Long: Dim TX As Long: Dim x1 As Long: Dim y1 As Long: Dim x2 As Long: Dim y2 As Long
        Dim K As Byte: Dim M As Byte: Dim S As Byte: Dim P As Byte: Dim T As Byte: Dim errLvl As Byte
        Dim X As Variant: Dim Y As Variant
        
        '//Pre-cleanup
        x1 = 0: x2 = 0: y1 = 0: y2 = 0: BX = 0: CX = 0: PX = 0: SX = 0: TX = 0: K = 0: M = 0: S = 0: P = 0: T = 0: X = 0: X = CByte(X): Y = 0: Y = CByte(Y)
        Call modArtQuotes(Token)
        
        '//Find application environment & block
        Call getEnvironment(appEnv, appBlk)
        
        '//Find flags
        If InStr(1, Token, "--") Then _
        Call libFlag(Token, errLvl): If Token = 1 Then Exit Function Else _
        Call libSwitch(Token, errLvl) '//Find switches
        
        '//Set library error level
        If Range("xlasLibErrLvl").Value2 = 0 Then On Error GoTo ErrEnd
        If Range("xlasLibErrLvl").Value2 = 1 Then On Error Resume Next

'/\_____________________________________
'//
'//          APPLICATION ARTICLES
'/\_____________________________________
'//
'//build() = Application build...
If InStr(1, Token, "build(", vbTextCompare) Then
Token = Replace(Token, "app.", vbNullString, , , vbTextCompare)
Token = Replace(Token, "build(", vbNullString, , , vbTextCompare)
Call modArtQuotes(Token)

If InStr(1, Token, ",") Then TokenArr = Split(Token, ",") Else MsgBox MsgBox(Application.Build): Exit Function  '//no excerpt provided
Exit Function

'//printer() = Application printer...
ElseIf InStr(1, Token, "printer(", vbTextCompare) Then
Token = Replace(Token, "app.", vbNullString, , , vbTextCompare)
Token = Replace(Token, "printer(", vbNullString, , , vbTextCompare)
Call modArtQuotes(Token)

If InStr(1, Token, ",") Then TokenArr = Split(Token, ",") Else MsgBox (Application.ActivePrinter): Exit Function '//no excerpt provided
Exit Function

'//name() = Application name...
ElseIf InStr(1, Token, "name(", vbTextCompare) Then
Token = Replace(Token, "app.", vbNullString, , , vbTextCompare)
Token = Replace(Token, "name(", vbNullString, , , vbTextCompare)
Call modArtQuotes(Token)

If InStr(1, Token, ",") Then TokenArr = Split(Token, ",") Else MsgBox (Application.name): Exit Function '//no excerpt provided
Exit Function

'//run() = Application run...
ElseIf InStr(1, Token, "run(", vbTextCompare) Then

Token = Replace(Token, "app.", vbNullString, , , vbTextCompare)
Token = Replace(Token, "run(", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)

'//switches
If InStr(1, Token, "-xlas", vbTextCompare) Then Token = Replace(Token, "-xlas", vbNullString, , , vbTextCompare): S = 7

'//run VBA module
If S = 0 Then
If InStr(1, Token, ",") Then TokenArr = Split(Token, ",") Else X = Application.Run(Token): Exit Function '//no arguments provided
xMod = TokenArr(0) '//extract module

X = 1
Do Until X > UBound(TokenArr) '//extract argument(s)
Token = TokenArr(X) & ",": ArtH = ArtH & Token
X = X + 1
Loop

Token = ArtH
If Right(Token, Len(Token) - Len(Token) + 1) = "," Then Token = Left(Token, Len(Token) - 1)

X = Application.Run(xMod, (Token))
Exit Function
    
'//run xlas script
ElseIf S = 7 Then Token = Trim(Token): _
Open Token For Input As #7: Token = vbNullString: _
Do Until EOF(7): Line Input #7, ArtH: Token = Token & ArtH: Loop: Close #7: Token = Token & "$": Call xlas(Token): Exit Function

End If

Exit Function
'//#
'//
'/\_____________________________________
'//
'//          CELL/RANGE ARTICLES
'/\_____________________________________
'//
'//cell() = Modify cell...
ElseIf InStr(1, Token, "cell(", vbTextCompare) Then
Call modArtQuotes(Token)

If InStr(1, Token, ",") = False Then MsgBox (Application.ActiveCell.Address): Exit Function '//no excerpt provided

'//Check for parameters...
If InStr(1, Token, ".") Then
If InStr(1, Token, " .") Then TokenArr = Split(Token, " .")
If InStr(1, Token, ").") Then TokenArr = Split(Token, ").")

Do Until X > UBound(TokenArr)

Token = TokenArr(X): Call modArtParens(Token): Call modArtQuotes(Token): TokenArr(X) = Token

'//Extract cell...
If InStr(1, TokenArr(X), "cell", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "cell", vbNullString, , , vbTextCompare)
If InStr(1, TokenArr(0), "=") Then TokenArrH = Split(TokenArr(0), "="): _
TokenArrH = Split(TokenArrH(1), ",") Else: _
TokenArrH = Split(TokenArr(0), ",")
Token = TokenArrH(0): Call modArtParens(Token): TokenArrH(0) = Token
Token = TokenArrH(1): Call modArtParens(Token): TokenArrH(1) = Token
x1 = CInt(TokenArrH(0)): y1 = CInt(TokenArrH(1))
End If
'//Select cell...
If InStr(1, TokenArr(X), "sel", vbTextCompare) Then
Cells(x1, y1).Select
End If
'//Clean cell...
If InStr(1, TokenArr(X), "cln", vbTextCompare) Then
Cells(x1, y1).ClearContents
End If
'//Clear cell...
If InStr(1, TokenArr(X), "clr", vbTextCompare) Then
Cells(x1, y1).Clear
End If
'//Copy cell...
If InStr(1, TokenArr(X), "copy") Then
If InStr(1, TokenArr(X), "copy&") Then P = 1
If InStr(1, TokenArr(X), "copy&!") Then P = 2
If InStr(1, TokenArr(X), "copy&!!") Then P = 3

    TokenArr(X) = Replace(TokenArr(X), "copy", vbNullString, vbTextCompare)
    TokenArr(X) = Replace(TokenArr(X), "!", vbNullString)
    TokenArr(X) = Replace(TokenArr(X), "&", vbNullString)
    Token = TokenArr(X): Call modArtParens(Token): TokenArr(X) = Token
    
    ActiveCell.Copy
    
    If P = vbNullString Then ActiveCell.Copy '//just copy
     
    If P = 1 Then '//copy paste cell contents
        ActiveWorkbook.Worksheets(appBlk).Cells(TokenArr(X)).Activate
            ActiveCell.PasteSpecial
                End If
                
    If P = 2 Then '//copy paste clean contents
        xCell = ActiveCell.Address
            ActiveWorkbook.Worksheets(appBlk).Cells(TokenArr(X)).Activate
                ActiveCell.PasteSpecial
                    ActiveWorkbook.Worksheets(appBlk).Cells(xCell).ClearContents
                        End If
                        
    If P = 3 Then '//copy paste clear cell contents
        xCell = ActiveCell.Address
            ActiveWorkbook.Worksheets(appBlk).Cells(TokenArr(X)).Activate
                ActiveCell.PasteSpecial
                    ActiveWorkbook.Worksheets(appBlk).Cells(xCell).Clear
                        End If
                        
                                End If

'//Paste cell...
If InStr(1, TokenArr(X), "paste", vbTextCompare) Then
Token = TokenArr(X): Call modArtParens(Token)
ActiveCell.PasteSpecial
End If
'//Set cell name...
If InStr(1, TokenArr(X), "name", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "name ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "name(", vbNullString, , , vbTextCompare)
Token = TokenArr(X): Call modArtParens(Token)
'//no name entered (clear name)
If TokenArr(X) = vbNullString Then
TokenArr(X) = Cells(x1, y1).name.name
ActiveWorkbook.Names(TokenArr(X)).Delete
    Else
        Cells(x1, y1).name = TokenArr(X)
            End If
                End If
'//Set cell value2...
If InStr(1, TokenArr(X), "value2", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "value2 ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "value2(", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "value2", vbNullString, , , vbTextCompare)
Cells(x1, y1).Value2 = TokenArr(X)
End If
'//Set cell value...
If InStr(1, TokenArr(X), "value", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "value ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "value(", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "value", vbNullString, , , vbTextCompare)
Cells(x1, y1).Value = TokenArr(X)
End If
'//Set cell font color...
If InStr(1, TokenArr(X), "fcolor", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "fcolor ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "fcolor(", vbNullString, , , vbTextCompare)
HX = TokenArr(X)
Call basColor(HX) '//find color
HX = Replace(HX, ")", vbNullString)
HX = Replace(HX, "(", vbNullString)
xRGBArr = Split(HX, ",")
Cells(x1, y1).Font.Color = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
End If
'//Set cell font size...
If InStr(1, TokenArr(X), "fsize", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "fsize ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "fsize(", vbNullString, , , vbTextCompare)
Cells(x1, y1).Font.Size = TokenArr(X)
End If
'//Set cell font type...
If InStr(1, TokenArr(X), "ftype", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "ftype", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "ftype(", vbNullString, , , vbTextCompare)
Cells(x1, y1).Font.FontStyle = TokenArr(X)
End If
'//Set cell pattern...
If InStr(1, TokenArr(X), "pattern", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "pattern", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "pattern(", vbNullString, , , vbTextCompare)
PX = TokenArr(X)
Call basPattern(PX) '//find pattern
Cells(x1, y1).Interior.Pattern = PX
End If
'//Set cell border direction...
If InStr(1, TokenArr(X), "border", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "border ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "border(", vbNullString, , , vbTextCompare)
BX = TokenArr(X)
Call basBorder(BX) '//find border
Cells(x1, y1).BorderAround (BX)
End If
'//Set cell border type...
If InStr(1, TokenArr(X), "btype", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "border ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "border(", vbNullString, , , vbTextCompare)
SX = TokenArr(X)
Call basBorderStyle(SX) '//find border type
Cells(x1, y1).Borders.LineStyle = SX
End If
'//Set cell color...
If InStr(1, TokenArr(X), "bgcolor", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "bgcolor ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "bgcolor(", vbNullString, , , vbTextCompare)
HX = TokenArr(X)
Call basColor(HX) '//find color
HX = Replace(HX, ")", vbNullString)
HX = Replace(HX, "(", vbNullString)
xRGBArr = Split(HX, ",")
Cells(x1, y1).Interior.Color = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
End If
'//Read cell value into variable...
If InStr(1, TokenArr(X), "read", vbTextCompare) Then
If InStr(1, TokenArr(0), "=") Then
TokenArr = Split(TokenArr(0), "=")
TokenArr(0) = Trim(TokenArr(0))
xVar = Cells(x1, y1).Value
Token = appEnv & "[,]" & TokenArr(0) & "=" & xVar & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk)
End If
    End If

X = X + 1
Loop
Exit Function
End If

Exit Function
'//#
'//
'//rng() = Modify range...
ElseIf InStr(1, Token, "rng(", vbTextCompare) Then

'//Check for parameters...
If InStr(1, Token, ".") Then
If InStr(1, Token, " .") Then TokenArr = Split(Token, " .")
If InStr(1, Token, ").") Then TokenArr = Split(Token, ").")

Do Until X > UBound(TokenArr)

Token = TokenArr(X): Call modArtParens(Token): Call modArtQuotes(Token): TokenArr(X) = Token

'//Extract range...
If InStr(1, TokenArr(X), "rng", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "rng", vbNullString, , , vbTextCompare)
If InStr(1, TokenArr(0), "=") Then TokenArrH = Split(TokenArr(0), "="): _
Token = TokenArrH(1): Call modArtParens(Token): TokenArr(0) = Token Else: _
Token = TokenArr(X): Call modArtParens(Token): TokenArr(X) = Token
End If
'//Select range...
If InStr(1, TokenArr(X), "sel", vbTextCompare) Then
Range(TokenArr(0)).Select
End If
'//Clean range...
If InStr(1, TokenArr(X), "cln", vbTextCompare) Then
Range(TokenArr(0)).ClearContents
End If
'//Clear range...
If InStr(1, TokenArr(X), "clr", vbTextCompare) Then
Range(TokenArr(0)).Clear
End If
'//Copy range...
If InStr(1, TokenArr(X), "copy") Then
If InStr(1, TokenArr(X), "copy&") Then P = 1
If InStr(1, TokenArr(X), "copy&!") Then P = 2
If InStr(1, TokenArr(X), "copy&!!") Then P = 3

    TokenArr(X) = Replace(TokenArr(X), "copy", vbNullString, vbTextCompare)
    TokenArr(X) = Replace(TokenArr(X), "!", vbNullString)
    TokenArr(X) = Replace(TokenArr(X), "&", vbNullString)
    Token = TokenArr(X): Call modArtParens(Token): TokenArr(X) = Token
    
    ActiveCell.Copy
    
    If P = vbNullString Then ActiveCell.Copy '//just copy
     
    If P = 1 Then '//copy paste range contents
        ActiveWorkbook.Worksheets(appBlk).Range(TokenArr(X)).Activate
            ActiveCell.PasteSpecial
                End If
                
    If P = 2 Then '//copy paste clean contents
        xCell = ActiveCell.Address
            ActiveWorkbook.Worksheets(appBlk).Range(TokenArr(X)).Activate
                ActiveCell.PasteSpecial
                    ActiveWorkbook.Worksheets(appBlk).Range(xCell).ClearContents
                        End If
                        
    If P = 3 Then '//copy paste clear range contents
        xCell = ActiveCell.Address
            ActiveWorkbook.Worksheets(appBlk).Range(TokenArr(X)).Activate
                ActiveCell.PasteSpecial
                    ActiveWorkbook.Worksheets(appBlk).Range(xCell).Clear
                        End If
                        
                                End If
            

'//Paste range...
If InStr(1, TokenArr(X), "paste", vbTextCompare) Then
ActiveCell.PasteSpecial
End If
'//Set range name...
If InStr(1, TokenArr(X), "name", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "name ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "name(", vbNullString, , , vbTextCompare)
'//no name entered (clear name)
If TokenArr(X) = vbNullString Then
TokenArr(X) = Range(TokenArr(0)).name.name
ActiveWorkbook.Names(TokenArr(X)).Delete
    Else
        Range(TokenArr(0)).name = TokenArr(X)
            End If
                End If
'//Set range value2...
If InStr(1, TokenArr(X), "value2", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "value2 ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "value2(", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "value2", vbNullString, , , vbTextCompare)
Range(TokenArr(0)).Value2 = TokenArr(X)
End If
'//Set range value...
If InStr(1, TokenArr(X), "value", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "value ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "value(", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "value", vbNullString, , , vbTextCompare)
Range(TokenArr(0)).Value = TokenArr(X)
End If
'//Set range font color...
If InStr(1, TokenArr(X), "fcolor", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "fcolor ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "fcolor(", vbNullString, , , vbTextCompare)
HX = TokenArr(X)
Call basColor(HX) '//find color
HX = Replace(HX, ")", vbNullString)
HX = Replace(HX, "(", vbNullString)
xRGBArr = Split(HX, ",")
Range(TokenArr(0)).Font.Color = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
End If
'//Set range font size...
If InStr(1, TokenArr(X), "fsize", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "fsize ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "fsize(", vbNullString, , , vbTextCompare)
Range(TokenArr(0)).Font.Size = TokenArr(X)
End If
'//Set range font type...
If InStr(1, TokenArr(X), "ftype", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "ftype ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "ftype(", vbNullString, , , vbTextCompare)
Range(TokenArr(0)).Font.FontStyle = TokenArr(X)
End If
'//Set range pattern...
If InStr(1, TokenArr(X), "pattern", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "pattern ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "pattern(", vbNullString, , , vbTextCompare)
PX = TokenArr(X)
Call basPattern(PX) '//find pattern
Range(TokenArr(0)).Interior.Pattern = PX
End If
'//Set range border direction...
If InStr(1, TokenArr(X), "border", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "border ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "border(", vbNullString, , , vbTextCompare)
BX = TokenArr(X)
Call basBorder(BX) '//find border
Range(TokenArr(0)).BorderAround (BX)
End If
'//Set range border type...
If InStr(1, TokenArr(X), "btype(", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "btype ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "btype(", vbNullString, , , vbTextCompare)
SX = TokenArr(X)
Call basBorderStyle(SX) '//find border type
Range(TokenArr(0)).Borders.LineStyle = SX
End If
'//Set range color...
If InStr(1, TokenArr(X), "bgcolor", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "bgcolor ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "bgcolor(", vbNullString, , , vbTextCompare)
HX = TokenArr(X)
Call basColor(HX) '//find color
HX = Replace(HX, ")", vbNullString)
HX = Replace(HX, "(", vbNullString)
xRGBArr = Split(HX, ",")
Range(TokenArr(0)).Interior.Color = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
End If
'//Read range value into variable...
If InStr(1, TokenArr(X), "read", vbTextCompare) Then
If TokenArrH(0) <> Empty Then
xVar = Range(TokenArr(0)).Value
Token = appEnv & "[,]" & TokenArrH(0) & "=" & xVar & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk)
End If
    End If

X = X + 1
Loop
Exit Function
End If

'//no parameter
Token = Replace(Token, "rng(", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)
'//Activate range...
Range(Token).Activate
Exit Function
'//#
'//
'//Select & modify cell/range...
ElseIf InStr(1, Token, "sel(", vbTextCompare) Then

'//Check for parameters...
If InStr(1, Token, ".") Then
If InStr(1, Token, " .") Then TokenArr = Split(Token, " .")
If InStr(1, Token, ").") Then TokenArr = Split(Token, ").")

Do Until X > UBound(TokenArr)

Token = TokenArr(X): Call modArtParens(Token): Call modArtQuotes(Token): TokenArr(X) = Token

'//Select cell...
If InStr(1, TokenArr(X), "sel", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "sel", vbNullString, , , vbTextCompare)
If InStr(1, TokenArr(0), "=") Then TokenArrH = Split(TokenArr(0), "="): _
Token = TokenArrH(1): Call modArtParens(Token): TokenArr(0) = Token Else: _
Token = TokenArr(X): Call modArtParens(Token): TokenArr(X) = Token
Range(TokenArr(X)).Select
End If
'//Clean cell...
If InStr(1, TokenArr(X), "cln", vbTextCompare) Then
Range(TokenArr(0)).ClearContents
End If
'//Clear cell...
If InStr(1, TokenArr(X), "clr", vbTextCompare) Then
Range(TokenArr(0)).Clear
End If
'//Copy cell...
If InStr(1, TokenArr(X), "copy", vbTextCompare) Then
If InStr(1, TokenArr(X), "copy&", vbTextCompare) Then P = 1
If InStr(1, TokenArr(X), "copy&!", vbTextCompare) Then P = 2
If InStr(1, TokenArr(X), "copy&!!", vbTextCompare) Then P = 3

    TokenArr(X) = Replace(TokenArr(X), "copy", vbNullString, , , vbTextCompare)
    TokenArr(X) = Replace(TokenArr(X), "!", vbNullString)
    TokenArr(X) = Replace(TokenArr(X), "&", vbNullString)
    Token = TokenArr(X): Call modArtParens(Token): TokenArr(X) = Token
    
    ActiveCell.Copy
    
    If P = vbNullString Then ActiveCell.Copy
     
    If P = 1 Then '//copy paste
        ActiveWorkbook.Worksheets(appBlk).Range(TokenArr(X)).Activate
            ActiveCell.PasteSpecial
                End If
                
    If P = 2 Then '//copy paste clear contents
        xCell = ActiveCell.Address
            ActiveWorkbook.Worksheets(appBlk).Range(TokenArr(X)).Activate
                ActiveCell.PasteSpecial
                    ActiveWorkbook.Worksheets(appBlk).Range(xCell).ClearContents
                        End If
                        
        If P = 3 Then '//copy paste clear cell
        xCell = ActiveCell.Address
            ActiveWorkbook.Worksheets(appBlk).Range(TokenArr(X)).Activate
                ActiveCell.PasteSpecial
                    ActiveWorkbook.Worksheets(appBlk).Range(xCell).Clear
                        End If
                        
                                End If
                                
'//Paste cell...
If InStr(1, TokenArr(X), "paste", vbTextCompare) Then
ActiveCell.PasteSpecial
End If
'//Set cell name...
If InStr(1, TokenArr(X), "name", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "name ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "name(", vbNullString, , , vbTextCompare)
'//no name entered (clear name)
If TokenArr(X) = vbNullString Then
TokenArr(X) = Range(TokenArr(0)).name.name
ActiveWorkbook.Names(TokenArr(X)).Delete
    Else
        Range(TokenArr(0)).name = TokenArr(X)
            End If
                End If
'//Set cell value2...
If InStr(1, TokenArr(X), "value2", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "value2 ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "value2(", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "value2", vbNullString, , , vbTextCompare)
Range(TokenArr(0)).Value2 = TokenArr(X)
End If
'//Set cell value...
If InStr(1, TokenArr(X), "value", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "value ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "value(", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "value", vbNullString, , , vbTextCompare)
Range(TokenArr(0)).Value = TokenArr(X)
End If
'//Set cell font color...
If InStr(1, TokenArr(X), "fcolor", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "fcolor ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "fcolor(", vbNullString, , , vbTextCompare)
HX = TokenArr(X)
Call basColor(HX) '//find color
HX = Replace(HX, ")", vbNullString)
HX = Replace(HX, "(", vbNullString)
xRGBArr = Split(HX, ",")
Range(TokenArr(0)).Font.Color = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
End If
'//Set cell font size...
If InStr(1, TokenArr(X), "fsize", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "fsize ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "fsize(", vbNullString, , , vbTextCompare)
Range(TokenArr(0)).Font.Size = TokenArr(X)
End If
'//Set cell font type...
If InStr(1, TokenArr(X), "ftype", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "ftype ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "ftype(", vbNullString, , , vbTextCompare)
Range(TokenArr(0)).Font.FontStyle = TokenArr(X)
End If
'//Set cell pattern...
If InStr(1, TokenArr(X), "pattern", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "pattern ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "pattern(", vbNullString, , , vbTextCompare)
PX = TokenArr(X)
Call basPattern(PX) '//find pattern
Range(TokenArr(0)).Interior.Pattern = PX
End If
'//Set cell border direction...
If InStr(1, TokenArr(X), "border", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "border ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "border(", vbNullString, , , vbTextCompare)
BX = TokenArr(X)
Call basBorder(BX) '//find border
Range(TokenArr(0)).BorderAround = BX
End If
'//Set cell border type...
If InStr(1, TokenArr(X), "btype", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "btype ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "btype(", vbNullString, , , vbTextCompare)
SX = TokenArr(X)
Call basBorderStyle(SX) '//find border type
Range(TokenArr(0)).Borders.LineStyle = SX
End If
'//Set cell color...
If InStr(1, TokenArr(X), "bgcolor", vbTextCompare) Then
TokenArr(X) = Replace(TokenArr(X), "bgcolor ", vbNullString, , , vbTextCompare)
TokenArr(X) = Replace(TokenArr(X), "bgcolor(", vbNullString, , , vbTextCompare)
HX = TokenArr(X)
Call basColor(HX) '//find color
HX = Replace(HX, ")", vbNullString)
HX = Replace(HX, "(", vbNullString)
xRGBArr = Split(HX, ",")
Range(TokenArr(0)).Interior.Color = RGB(xRGBArr(0), xRGBArr(1), xRGBArr(2))
End If
'//Read cell value into variable...
If InStr(1, TokenArr(X), "read", vbTextCompare) Then
If TokenArrH(0) <> Empty Then
xVar = Range(TokenArr(0)).Value
Token = appEnv & "[,]" & TokenArrH(0) & "=" & xVar & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk)
End If
    End If
    
X = X + 1
Loop
Exit Function
End If
'//no parameter
Token = Replace(Token, "sel(", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)
'//Activate cell...
Range(Token).Select
Exit Function
'//#
'//
'/\_____________________________________
'//
'//        WORKBOOK ARTICLES
'/\_____________________________________
'//
'//wb() = Modify Workbook...
ElseIf InStr(1, Token, "wb(", vbTextCompare) Then
Token = Replace(Token, "wb(", vbNullString, , , vbTextCompare)

If InStr(1, Token, ".active", vbTextCompare) Then If InStr(1, Token, ").active", vbTextCompare) = False Then ActiveWorkbook.Activate  '//activate current workbook
If InStr(1, Token, ").active", vbTextCompare) Then '//activate specific workbook
Token = Replace(Token, ".active", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)
Workbooks(Token).Activate
Range("MAS2").name = "xlasEnvironment": Range("xlasEnvironment").Value = appEnv '//link environment to workbook
Range("MAS3").name = "xlasBlock": Range("xlasBlock").Value = appBlk '//link block to workbook
Exit Function
End If

If InStr(1, Token, ").open", vbTextCompare) Then '//open workbook

Token = Replace(Token, ".open", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)
Workbooks.Open (Token)
Range("MAS2").name = "xlasEnvironment": Range("xlasEnvironment").Value = appEnv '//link environment to workbook
Range("MAS3").name = "xlasBlock": Range("xlasBlock").Value = appBlk '//link block to workbook
Workbooks(appEnv).Worksheets(appBlk).Activate
Exit Function
End If

If InStr(1, Token, ").hd", vbTextCompare) Then ActiveWorkbook.Application.Visible = False   '//hide active workbook
If InStr(1, Token, ").sh", vbTextCompare) Then ActiveWorkbook.Application.Visible = True  '//show active workbook

If InStr(1, Token, ").close", vbTextCompare) Then _
Token = Replace(Token, ".close", vbNullString, , , vbTextCompare): _
Call modArtParens(Token): Call modArtQuotes(Token): _
If Token = vbNullString Then ActiveWorkbook.Close: Exit Function Else _
Workbooks(Token).Close: Exit Function '//close workbook

'If InStr(1, Token, ").export", vbTextCompare) Then ActiveWorkbook.ExportAsFixedFormat = vbnullstring '//export file
If InStr(1, Token, ").newwin", vbTextCompare) Then ActiveWorkbook.NewWindow: Exit Function '//create new window

If InStr(1, Token, ").save", vbTextCompare) And InStr(1, Token, ").saveas", vbTextCompare) = False Then _
Token = Replace(Token, ".save", vbNullString, , , vbTextCompare): _
Call modArtParens(Token): Call modArtQuotes(Token): _
If Token = vbNullString Then ActiveWorkbook.Save: Exit Function Else _
Workbooks(Token).Save: Exit Function '//save workbook

If InStr(1, Token, ").saveas", vbTextCompare) Then '//save as workbook

Call modArtParens(Token): Call modArtQuotes(Token)

Token = Replace(Token, ".saveas", vbNullString, , , vbTextCompare)
TokenArr = Split(Token, ",")
If UBound(TokenArr) = 1 Then
EX = TokenArr(1): Call basSaveFormat(EX)
If EX <> "*/ERR" Then
Range("MAS2").name = "xlasEnvironment": Range("xlasEnvironment").Value = appEnv '//link environment to workbook
Range("MAS3").name = "xlasBlock": Range("xlasBlock").Value = appBlk '//link block to workbook
ActiveWorkbook.SaveAs FileName:=TokenArr(0), FileFormat:=xExt
End If
    End If
        Exit Function
            End If
    
If InStr(1, Token, ").name", vbTextCompare) Then MsgBox (ActiveWorkbook.name), 0, "": Exit Function '//get name of workbook
If InStr(1, Token, ").path", vbTextCompare) Then MsgBox (ActiveWorkbook.Path), 0, "": Exit Function '//get path of workbook

If InStr(1, Token, ").addsheet", vbTextCompare) Then '//add worksheet

Call modArtParens(Token): Call modArtQuotes(Token)

If InStr(1, Token, ").addsheetafter", vbTextCompare) Then P = 1: Token = Replace(Token, ".addafter", vbNullString, , , vbTextCompare) '//add after worksheet
If InStr(1, Token, ").addsheetbefore", vbTextCompare) Then P = 2: Token = Replace(Token, ".addbefore", vbNullString, , , vbTextCompare) '//add before worksheet

Token = Replace(Token, ".add", vbNullString, , , vbTextCompare)
If Token = vbNullString Then '//default add no arguments
Token = "Sheet" & ActiveWorkbook.Worksheets.Count + 1
Worksheets.Add.name = Token
Exit Function
End If

If InStr(1, Token, ",") = False Then
'//single argument... (set count w/ default worksheet name & place before or after first/last sheet)
If P = 1 Then Worksheets.Add After:=Worksheets(Worksheets.Count), Count:=Int(Token): Exit Function
If P = 2 Then Worksheets.Add Before:=Worksheets(Worksheets.Count), Count:=Int(Token): Exit Function
    Else
TokenArr = Split(Token, ",")
If UBound(TokenArr) = 1 Then
'//two arguments... (set add worksheet name & place before or after assigned sheet
If P = 1 Then Worksheets.Add(After:=Worksheets(TokenArr(0))).name = TokenArr(1): Exit Function
If P = 2 Then Worksheets.Add(Before:=Worksheets(TokenArr(0))).name = TokenArr(1): Exit Function
ElseIf UBound(TokenArr) = 2 Then
'//three arguments... (set add worksheet name & place before or after assigned  sheet w/ count)
If P = 1 Then Worksheets.Add(After:=Worksheets(TokenArr(0)), Count:=Int(TokenArr(2))).name = TokenArr(1): Exit Function
If P = 2 Then Worksheets.Add(Before:=Worksheets(TokenArr(0)), Count:=Int(TokenArr(2))).name = TokenArr(1): Exit Function
                    End If
                        End If
        
If InStr(1, Token, ").newbook", vbTextCompare) Then '//add new workbook

Call modArtParens(Token): Call modArtQuotes(Token)

Token = Replace(Token, ".newbook", vbNullString, , , vbTextCompare)
TokenArr = Split(Token, ",")
If UBound(TokenArr) = 1 Then
EX = TokenArr(1): Call basSaveFormat(EX)
If EX <> "*/ERR" Then
Application.Workbooks.Add.SaveAs FileName:=TokenArr(0), FileFormat:=xExt
Workbooks(appEnv).Worksheets(appBlk).Activate
End If
    End If
        Exit Function
            End If
                End If


'//Run workbook module...
If InStr(1, Token, ").run", vbTextCompare) Then

Token = Replace(Token, ".run", vbNullString, , , vbTextCompare)
Call modArtBraces(Token): Call modArtQuotes(Token)

If InStr(1, Token, ",") Then TokenArr = Split(Token, ",") Else GoTo wbRunNoArg

TokenArr(0) = Trim(TokenArr(0)): TokenArr(1) = Trim(TokenArr(1))
xWb = TokenArr(0) '//extract workbook
xMod = TokenArr(1) '//extract module

X = 2
Do Until X > UBound(TokenArr) '//extract argument(s)
Token = TokenArr(X) & ",": ArtH = ArtH & Token
X = X + 1
Loop

Token = ArtH
If Right(Token, Len(Token) - Len(Token) + 1) = "," Then Token = Left(Token, Len(Token) - 1)

X = Application.Run("'" & xWb & "'!" & xMod, (Token))
Exit Function
    
'//no arguments provided
wbRunNoArg:
X = Application.Run(Token)
Exit Function
End If

'//Delete workbook cell name...
If InStr(1, Token, ").delname", vbTextCompare) Then

Token = Replace(Token, ".delname", vbNullString, , , vbTextCompare)
Call modArtBraces(Token): Call modArtParens(Token): Call modArtQuotes(Token)

If InStr(1, Token, ",") Then TokenArr = Split(Token, ",") Else GoTo ErrEnd

TokenArr(0) = Trim(TokenArr(0)): TokenArr(1) = Trim(TokenArr(1))

Workbooks(TokenArr(0)).Names(TokenArr(1)).Delete

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
'//fil() = Modify file
ElseIf InStr(1, Token, "fil(", vbTextCompare) Then

If InStr(1, Token, ".fil", vbTextCompare) Then
Token = Replace(Token, "fil(", vbNullString, , , vbTextCompare)

If InStr(1, Token, "del.", vbTextCompare) Then M = 1: Token = Replace(Token, "del.", vbNullString, , , vbTextCompare)
If InStr(1, Token, "mk.", vbTextCompare) Then M = 2: Token = Replace(Token, "mk.", vbNullString, , , vbTextCompare)
If Left(Token, 1) = " " Then Token = Right(Token, Len(Token) - 1)
If M = 0 Then errLvl = 1: GoTo ErrEnd

Call modArtParens(Token): Call modArtQuotes(Token)
Set oFSO = CreateObject("Scripting.FileSystemObject")
If M = 1 Then: Set oFSO = CreateObject("Scripting.FileSystemObject"): oFSO.DeleteFile (Token): Set oFSO = Nothing: Exit Function '//delete file
If M = 2 Then: _

If InStr(1, Token, ",") Then
M = M & "1"
TokenArr = Split(Token, ",")
Token = TokenArr(0): Call modArtParens(Token): Call modArtQuotes(Token): TokenArr(0) = Token
Token = TokenArr(1): Call modArtParens(Token): TokenArr(1) = Token
TokenArr(1) = LTrim(TokenArr(1)) '//remove leading space
End If

If M = 2 Then Open (Token) For Output As #1: Print #1, vbNullString: Close #1: Exit Function
If M = 21 Then Open (TokenArr(0)) For Output As #1: Print #1, TokenArr(1): Close #1: Exit Function
Exit Function

Else
errLvl = 1: GoTo ErrEnd
End If

'//dir() = Modify folder
ElseIf InStr(1, Token, "dir(", vbTextCompare) Then

If InStr(1, Token, ".dir", vbTextCompare) Then
Token = Replace(Token, "dir(", vbNullString, , , vbTextCompare)

If InStr(1, Token, "del.", vbTextCompare) Then M = 1: Token = Replace(Token, "del.", vbNullString, , , vbTextCompare)
If InStr(1, Token, "mk.", vbTextCompare) Then M = 2: Token = Replace(Token, "mk.", vbNullString, , , vbTextCompare)
If Left(Token, 1) = " " Then Token = Right(Token, Len(Token) - 1)
If M = 0 Then errLvl = 1: GoTo ErrEnd

Call modArtParens(Token): Call modArtQuotes(Token)
If M = 1 Then: Set oFSO = CreateObject("Scripting.FileSystemObject"): oFSO.DeleteFolder (Token): Set oFSO = Nothing: Exit Function '//create file
If M = 2 Then: _

If InStr(1, Token, ",") Then
M = M & "1"
TokenArr = Split(Token, ",")
Token = TokenArr(0): Call modArtParens(Token): Call modArtQuotes(Token): TokenArr(0) = Token
Token = TokenArr(1): Call modArtParens(Token): TokenArr(1) = Token
TokenArr(1) = LTrim(TokenArr(1)) '//remove leading space
End If

If M = 2 Then MkDir (Token): Exit Function
If M = 21 Then MkDir (TokenArr(0)): MkDir (TokenArr(0) & "/" & TokenArr(1)): Exit Function
Exit Function

Else
errLvl = 1: GoTo ErrEnd
End If


'//dfldr() = Delete empty directory
ElseIf InStr(1, Token, "dfldr(", vbTextCompare) Then

Token = Replace(Token, "dfldr", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)

If Dir(Token, vbDirectory) <> "" Then RmDir (Token): Exit Function

'//dfile() = Delete file
ElseIf InStr(1, Token, "dfile(", vbTextCompare) Then

Token = Replace(Token, "dfile", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)

If Dir(Token) <> "" Then Kill (Token): Exit Function

'//mfldr() = Create empty directory
ElseIf InStr(1, Token, "mfldr(", vbTextCompare) Then

Token = Replace(Token, "mfldr", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)

MkDir (Token): Exit Function

'//mfile() = Create file
ElseIf InStr(1, Token, "mfile(", vbTextCompare) Then

Token = Replace(Token, "mfile", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)

Open (Token) For Output As #1: Print #1, vbNullString: Close #1: Exit Function

'//move() = Move file or folder
ElseIf InStr(1, Token, "move(", vbTextCompare) Then
Token = Replace(Token, "move(", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)

If InStr(1, Token, ",") Then

TokenArr = Split(Token, ",")

Token = "move " & TokenArr(0) & " " & TokenArr(1)
    
sysShell = Shell("cmd.exe /s /c" & Token, 0)
sysShell = vbNullString
End If
Exit Function

'//ren() = Rename file or folder
ElseIf InStr(1, Token, "ren(", vbTextCompare) Then
Token = Replace(Token, "ren(", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)

If InStr(1, Token, ",") Then

TokenArr = Split(Token, ",")

If UBound(TokenArr) = 2 Then GoTo renAll
If InStr(1, Token, "app.r") Then GoTo renVBA

'//default
Token = "ren " & TokenArr(0) & " " & TokenArr(1): sysShell = Shell("cmd.exe /s /c" & Token, 0): sysShell = vbNullString
End If
Exit Function

renVBA:
Name TokenArr(0) As TokenArr(1)
Exit Function

renAll:
Dim getDate, xName, xTime As String
Dim xNum As Long
xNum = 1

Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oSubFldr = oFSO.GetFolder(TokenArr(0))

xExtArr = Split(TokenArr(1), "."): EX = xExtArr(1): xName = xExtArr(0)

If InStr(1, TokenArr(2), "-num", vbTextCompare) Then xNum = 1: GoTo renAllNum
If InStr(1, TokenArr(2), "-datenum", vbTextCompare) And InStr(1, TokenArr(2), "mtime", vbTextCompare) = False Then getDate = Date: getDate = Replace(getDate, "/", "-"): xNum = 1: GoTo renAllDateNum
If InStr(1, TokenArr(2), "-datenumtime", vbTextCompare) Then getDate = Date: getDate = Replace(getDate, "/", "-"): xNum = 1: xTime = Time: xTime = Replace(xTime, ":", vbNullString): xTime = Replace(xTime, " ", vbNullString): GoTo renAllDateNumTime

renAllNum:
For Each oFile In oSubFldr.Files
Token = "ren " & oFile.Path & " " & xName & "_" & xNum & "." & EX
sysShell = Shell("cmd.exe /s /c" & Token, 0)
sysShell = vbNullString
xNum = xNum + 1
Next
Set oFSO = Nothing
Set oFile = Nothing
Set oSubFldr = Nothing
Exit Function

renAllDateNum:
For Each oFile In oSubFldr.Files
Token = "ren " & oFile.Path & " " & xName & "_" & getDate & "_" & xNum & "." & EX
sysShell = Shell("cmd.exe /s /c" & Token, 0)
sysShell = vbNullString
xNum = xNum + 1
Next
Set oFSO = Nothing
Set oFile = Nothing
Set oSubFldr = Nothing
Exit Function

renAllDateNumTime:
For Each oFile In oSubFldr.Files
xNum = xNum + Token = "ren " & oFile.Path & " " & xName & "_" & getDate & "_" & xNum & "_" & xTime & "." & EX
sysShell = Shell("cmd.exe /s /c" & Token, 0)
sysShell = vbNullString
xNum = xNum + 1
Next
Set oFSO = Nothing
Set oFile = Nothing
Set oSubFldr = Nothing
Exit Function

'//read() = Read file
ElseIf InStr(1, Token, "read(", vbTextCompare) Then
Token = Replace(Token, "read", vbNullString, , , vbTextCompare)

'//switches
If InStr(1, Token, "-all", vbTextCompare) Then S = 1: Token = Replace(Token, "-all", vbNullString, , , vbTextCompare): GoTo SetRead
If InStr(1, Token, "-nlall", vbTextCompare) Then S = 2: Token = Replace(Token, "-nlall", vbNullString, , , vbTextCompare): GoTo SetRead
If InStr(1, Token, "-string", vbTextCompare) Then S = 3: Token = Replace(Token, "-string", vbNullString, , , vbTextCompare): GoTo SetRead
If InStr(1, Token, "-xstring", vbTextCompare) Then S = 4: Token = Replace(Token, "-xstring", vbNullString, , , vbTextCompare): GoTo SetRead
If InStr(1, Token, "-nlstring", vbTextCompare) Then S = 5: Token = Replace(Token, "-nlstring", vbNullString, , , vbTextCompare): GoTo SetRead
If InStr(1, Token, "-line", vbTextCompare) Then S = 6: Token = Replace(Token, "-line", vbNullString, , , vbTextCompare): GoTo SetRead
If InStr(1, Token, "-xline", vbTextCompare) Then S = 7: Token = Replace(Token, "-xline", vbNullString, , , vbTextCompare): GoTo SetRead
If InStr(1, Token, "-nlline", vbTextCompare) Then S = 8: Token = Replace(Token, "-nlline", vbNullString, , , vbTextCompare): GoTo SetRead

SetRead:
'//modifier(s):
If InStr(1, Token, "count.") Then M = 1: Token = Replace(Token, "count.", vbNullString, , , vbTextCompare)

TokenArr = Split(Token, "=")
Token = TokenArr(1): Call modArtParens(Token): Call modArtQuotes(Token): Token = Trim(Token): TokenArr(1) = Token
Token = TokenArr(0)

'//read for all
If S = 1 Then
Open TokenArr(1) For Input As #1: Do Until EOF(1): Line Input #1, ArtH: xVar = xVar & ArtH: Loop: Close #1
GoTo EndRead
End If

'//read for newline all
If S = 2 Then
Open TokenArr(1) For Input As #1: Do Until EOF(1): Line Input #1, ArtH: xVar = xVar & ArtH & vbNewLine: Loop: Close #1
GoTo EndRead
End If

'//read for string
If S = 3 Then
TokenArr = Split(TokenArr(1), ","): TokenArr(1) = Trim(TokenArr(1))
Open TokenArr(1) For Input As #1
Do Until EOF(1): Line Input #1, ArtH
If InStr(1, ArtH, TokenArr(0)) Then xVar = ArtH: Close #1: GoTo EndRead
Loop: Close #1
GoTo EndRead
End If

'//read for all string
If S = 4 Then
TokenArr = Split(TokenArr(1), ","): TokenArr(1) = Trim(TokenArr(1))
Open TokenArr(1) For Input As #1
Do Until EOF(1): Line Input #1, ArtH
If InStr(1, ArtH, TokenArr(0)) Then xVar = xVar & ArtH
Loop: Close #1
GoTo EndRead
End If

'//read for newline string
If S = 5 Then
TokenArr = Split(TokenArr(1), ","): TokenArr(1) = Trim(TokenArr(1))
Open TokenArr(1) For Input As #1
Do Until EOF(1): Line Input #1, ArtH
If InStr(1, ArtH, TokenArr(0)) Then xVar = xVar & ArtH & vbNewLine
Loop: Close #1
GoTo EndRead
End If

'//read for line
If S = 6 Then
TokenArr = Split(TokenArr(1), ","): TokenArr(1) = Trim(TokenArr(1))
Open TokenArr(1) For Input As #1
For X = 1 To TokenArr(0)
Line Input #1, ArtH
Next: Close #1: xVar = ArtH
GoTo EndRead
End If

'//read for all line
If S = 7 Then
TokenArr = Split(TokenArr(1), ","): TokenArr(1) = Trim(TokenArr(1))
Open TokenArr(1) For Input As #1
For X = 1 To TokenArr(0)
Line Input #1, ArtH
xVar = xVar & ArtH
Next: Close #1
GoTo EndRead
End If

'//read for newline line
If S = 8 Then
TokenArr = Split(TokenArr(1), ","): TokenArr(1) = Trim(TokenArr(1))
Open TokenArr(1) For Input As #1
For X = 1 To TokenArr(0)
Line Input #1, ArtH
xVar = xVar & ArtH & vbNewLine
Next: Close #1
GoTo EndRead
End If

EndRead:
'//count
If M = 1 Then xVar = Len(xVar)

Token = appEnv & "[,]" & Token & "=" & xVar & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk)

Exit Function
'//#
'//
'/\_____________________________________
'//
'//          HALT ARTICLES
'/\_____________________________________
'//
'//wait() = Pause script (no actions allowed)
ElseIf InStr(1, Token, "wait(", vbTextCompare) Then
Token = Replace(Token, "wait(", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)

If InStr(1, Token, "ms", vbTextCompare) Then P = P & "1" '//millisecond
If InStr(1, Token, "s", vbTextCompare) Then P = P & "2" '//second
If InStr(1, Token, "m", vbTextCompare) Then P = P & "3" '//minute
If InStr(1, Token, "h", vbTextCompare) Then P = P & "4" '//hour

If P <> 0 Then

Dim xTimeArr(3) As String
Dim xMil, xSec, xMin, xHr As String
Dim AppWait As Variant

If InStr(1, Token, "ms", vbTextCompare) Then xMilArr = Split(Token, "ms", , vbTextCompare): xTimeArr(0) = xMilArr(0): xMil = "T"
If InStr(1, Token, "s", vbTextCompare) Then xSecArr = Split(Token, "s", , vbTextCompare): xTimeArr(1) = xSecArr(0): xSec = "T"
If InStr(1, Token, "m", vbTextCompare) Then xMinArr = Split(Token, "m", , vbTextCompare): xTimeArr(2) = xMinArr(0): xMin = "T"
If InStr(1, Token, "h", vbTextCompare) Then xHrArr = Split(Token, "h", , vbTextCompare): xTimeArr(3) = xHrArr(0): xHr = "T"

'//set millisecond
If xMil = "T" Then
Token = xTimeArr(0)
Call xlAppScript_lex.getChar(Token): If Token = "*/ERR" Then GoTo ErrEnd
Token = -1 * (Token * -0.00000001)
Application.Wait (Now + Token)
Exit Function
End If
        
'//set second
If xSec = "T" Then
Token = xTimeArr(1)
Call xlAppScript_lex.getChar(Token): If Token = "*/ERR" Then GoTo ErrEnd
If Len(xTimeArr(1)) < 2 Then
xTimeArr(1) = "0" & xTimeArr(1): xSec = xTimeArr(1)
Else: xSec = xTimeArr(1)
End If
    Else: xSec = "00"
        End If
        
'//set minute
If xMin = "T" Then
Token = xTimeArr(2)
Call xlAppScript_lex.getChar(Token): If Token = "*/ERR" Then GoTo ErrEnd
If Len(xTimeArr(2)) < 2 Then
xTimeArr(2) = "0" & xTimeArr(2): xMin = xTimeArr(2)
Else: xMin = xTimeArr(2)
End If
    Else: xMin = "00"
        End If
        
'//set hour
If xHr = "T" Then
Token = xTimeArr(3)
Call xlAppScript_lex.getChar(Token): If Token = "*/ERR" Then GoTo ErrEnd
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
    If InStr(1, Token, ":") Then
    TokenArr = Split(Token, ":")
    AppWait = TimeSerial(TokenArr(0), TokenArr(1), TokenArr(2))
    Application.Wait Now + TimeValue(AppWait)
        Else
            GoTo ErrEnd
                End If
                    End If
Exit Function

'//delayevent() = Delay script (actions allowed)
ElseIf InStr(1, Token, "delayevent(", vbTextCompare) Then
Token = Replace(Token, "delayevent(", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)
For TX = 0 To Token * 360: T = 1: Call basWasteTime(T): Next
Exit Function
'//#
'//
'/\_____________________________________
'//
'//         INPUT-HOST ARTICLES
'/\_____________________________________

ElseIf InStr(1, Token, "input(", vbTextCompare) Then

   Token = Replace(Token, "input(", vbNullString, , , vbTextCompare)
   Call modArtParens(Token)
   
   TokenArr = Split(Token, "="): Token = TokenArr(0)
   TokenArr = Split(TokenArr(1), ",")
    
   If UBound(TokenArr) = 1 Then xVar = InputBox(TokenArr(0), TokenArr(1))
   If UBound(TokenArr) = 2 Then xVar = InputBox(TokenArr(0), TokenArr(1), TokenArr(2))

   Token = appEnv & "[,]" & Token & "=" & xVar & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk)
   
Exit Function
'//#
'//
'/\_____________________________________
'//
'//     OUTPUT-HOST ARTICLES
'/\_____________________________________
'//
'//echo() = Output string w/ Command Prompt
ElseIf InStr(1, Token, "echo(", vbTextCompare) Then

    If InStr(1, Token, "o(0)", vbTextCompare) Then S = 1: Token = Replace(Token, "echo(0)", vbNullString): GoTo setEcho
    If InStr(1, Token, "o(1)", vbTextCompare) Then S = 2: Token = Replace(Token, "echo(1)", vbNullString): GoTo setEcho
    If InStr(1, Token, "o(2)", vbTextCompare) Then S = 3: Token = Replace(Token, "echo(2)", vbNullString): GoTo setEcho
    If InStr(1, Token, "o(3)", vbTextCompare) Then S = 4: Token = Replace(Token, "echo(3)", vbNullString): GoTo setEcho
    If InStr(1, Token, "o(4)", vbTextCompare) Then S = 5: Token = Replace(Token, "echo(4)", vbNullString): GoTo setEcho
    If InStr(1, Token, "o(5)", vbTextCompare) Then S = 6: Token = Replace(Token, "echo(5)", vbNullString): GoTo setEcho
    If InStr(1, Token, "o(6)", vbTextCompare) Then S = 7: Token = Replace(Token, "echo(6)", vbNullString): GoTo setEcho
    
    Token = Replace(Token, "echo(", vbNullString, , , vbTextCompare)
    
setEcho:
Call modArtParens(Token)
  
   sysShell = Shell("cmd.exe /k echo " & Token, S)
   sysShell = vbNullString
   Exit Function
   
'//Output w/ default message box
ElseIf InStr(1, Token, "host(", vbTextCompare) Then

   Token = Replace(Token, "host(", vbNullString, , , vbTextCompare)
   Call modArtParens(Token): Call modArtQuotes(Token)
   If Right(Token, 1) = ")" Then Token = Left(Token, Len(Token) - 1)
   MsgBox (Token)
   Exit Function
   
   
'//Output w/ VBA message box
ElseIf InStr(1, Token, "msg(", vbTextCompare) Then

   Token = Replace(Token, "msg(", vbNullString, , , vbTextCompare)
    Call modArtParens(Token)
   
   If InStr(1, Token, "=") Then '//check for variable
   TokenArr = Split(Token, "=")
   Token = TokenArr(0)
   If UBound(TokenArr) = 1 Then TokenArr = Split(TokenArr(1), ","): _
   TokenArr(0) = Trim(TokenArr(0))
   
   If UBound(TokenArr) = 0 Then xVar = MsgBox(TokenArr(0)): GoTo EndMsg
   If UBound(TokenArr) = 1 Then TokenArr(1) = Trim(TokenArr(1)): xVar = MsgBox(TokenArr(0), TokenArr(1)): GoTo EndMsg
   If UBound(TokenArr) = 2 Then TokenArr(1) = Trim(TokenArr(1)): TokenArr(2) = Trim(TokenArr(2)): xVar = MsgBox(TokenArr(0), TokenArr(1), TokenArr(2)): GoTo EndMsg
   
EndMsg:
   Token = appEnv & "[,]" & Token & "=" & xVar & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk)
   Exit Function
   
   End If
   
   MsgBox (Token) '//no arguments
   Exit Function
'//#
'//
'/\_____________________________________
'//
'//      KEYSTROKE ARTICLES
'/\_____________________________________
'//
    ElseIf InStr(1, Token, "key(", vbTextCompare) Then
    
    If InStr(1, Token, ").clr", vbTextCompare) Then
    Dim oKey, oTemp As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oTemp = oFSO.GetFolder(drv & envHome & "\.z7\utility\temp")
    If oFSO.FolderExists(oTemp) = True Then For Each oKey In oTemp.Files: Kill (oKey): Next: Set oTemp = Nothing: Set oKey = Nothing
    Exit Function
    End If
    
    If InStr(1, Token, "app.k", vbTextCompare) = False Then '//check for application key (key w/ VBA)
    
    Dim sysKey0 As String: Dim sysKey1 As String: Dim sysKey2 As String
    Dim sysKey3 As String: Dim sysKey4 As String: Dim sysKey5 As String:
    Dim sysKey6 As String
    
    If InStr(1, Token, "y(0)", vbTextCompare) Then K = 1: Token = Replace(Token, "key(0)", vbNullString): GoTo setKey
    If InStr(1, Token, "y(1)", vbTextCompare) Then K = 2: Token = Replace(Token, "key(1)", vbNullString): GoTo setKey
    If InStr(1, Token, "y(2)", vbTextCompare) Then K = 3: Token = Replace(Token, "key(2)", vbNullString): GoTo setKey
    If InStr(1, Token, "y(3)", vbTextCompare) Then K = 4: Token = Replace(Token, "key(3)", vbNullString): GoTo setKey
    If InStr(1, Token, "y(4)", vbTextCompare) Then K = 5: Token = Replace(Token, "key(4)", vbNullString): GoTo setKey
    If InStr(1, Token, "y(5)", vbTextCompare) Then K = 6: Token = Replace(Token, "key(5)", vbNullString): GoTo setKey
    If InStr(1, Token, "y(6)", vbTextCompare) Then K = 7: Token = Replace(Token, "key(6)", vbNullString): GoTo setKey
    
setKey:
    Token = Replace(Token, "key", vbNullString, , , vbTextCompare)
    Call modArtParens(Token)
    Token = Right(Token, Len(Token) - 1) '//remove leading quotes
    Token = Left(Token, Len(Token) - 1) '//remove ending quotes
    
    
    If Token = vbNullString Then Exit Function
  
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
    Print #K, "Dim wShell"
    Print #K, "Set wShell = wscript.CreateObject(" & """" & "wscript.Shell""" & ")"
    Print #K, "wShell.SendKeys " & """" & Token & """"
    Print #K, "Set wShell = Nothing"
    Print #K, "wscript.Quit"
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
    Token = Replace(Token, "app.", vbNullString)
    
    Token = Replace(Token, "key", vbNullString, , , vbTextCompare)
    Call modArtParens(Token)
    
    Application.SendKeys (Token)
    Exit Function
    
    End If
'//#
'//
'/\_____________________________________
'//
'//        MOUSE ACTION ARTICLES
'/\_____________________________________
'//
'//click() = Assign mouse click events
ElseIf InStr(1, Token, "click(", vbTextCompare) Then

Token = Replace(Token, "click(", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)

'//switches
If InStr(1, Token, "-double", vbTextCompare) Then Token = Replace(Token, "-double", vbNullString, , , vbTextCompare): S = 1: GoTo setClick
If InStr(1, Token, "-leftdown", vbTextCompare) Then Token = Replace(Token, "-leftdown", vbNullString, , , vbTextCompare): S = 2: GoTo setClick
If InStr(1, Token, "-leftup", vbTextCompare) Then Token = Replace(Token, "-leftup", vbNullString, , , vbTextCompare): S = 3: GoTo setClick
If InStr(1, Token, "-rightdown", vbTextCompare) Then Token = Replace(Token, "-rightdown", vbNullString, , , vbTextCompare): S = 4: GoTo setClick
If InStr(1, Token, "-rightup", vbTextCompare) Then Token = Replace(Token, "-rightup", vbNullString, , , vbTextCompare): S = 5: GoTo setClick

setClick:
If InStr(1, Token, ",") Then
Token = Trim(Token)
TokenArr = Split(Token, ",") '//arguments
xPos = TokenArr(0) & "," & TokenArr(1)
Call basClick(S, xPos): Exit Function
End If

'//no arguments
S = 5: Call basClick(S, xPos)
Exit Function
'//#
'//
'/\_____________________________________
'//
'//        MODIFY STRING ARTICLES
'/\_____________________________________
'//
'//conv32() = Convert decimal to binary bit string (32-bit signed/unsigned integers)
ElseIf InStr(1, Token, "conv32(", vbTextCompare) Then

Token = Replace(Token, "conv32(", vbNullString, , , vbTextCompare)
Call modArtParens(Token)

'//switches
If InStr(1, Token, "-8bit", vbTextCompare) Then Token = Replace(Token, "-8bit", vbNullString, , , vbTextCompare): S = 1: GoTo setConv32
If InStr(1, Token, "-16bit", vbTextCompare) Then Token = Replace(Token, "-16bit", vbNullString, , , vbTextCompare): S = 0: GoTo setConv32
If InStr(1, Token, "-32bit", vbTextCompare) Then Token = Replace(Token, "-32bit", vbNullString, , , vbTextCompare): S = 2: GoTo setConv32

setConv32:

xVarArr = Split(Token, "=") '//find variable
xRtnBits = S

X = xVarArr(1): Call basDecimalToBinary32(X, xRtnBits, xBitStr): Token = xBitStr
Token = xVarArr(0) & "=" & Token
Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function

'//Convert char/string...
ElseIf InStr(1, Token, "conv(", vbTextCompare) Then

Token = Replace(Token, "conv(", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)

'//switches
If InStr(1, Token, "-upper", vbTextCompare) Then S = vbUpperCase: GoTo setConv
If InStr(1, Token, "-lower", vbTextCompare) Then S = vbLowerCase: GoTo setConv
If InStr(1, Token, "-proper", vbTextCompare) Then S = vbProperCase: GoTo setConv
If InStr(1, Token, "-unicode", vbTextCompare) Then S = vbUnicode: GoTo setConv

setConv:
TokenArr = Split(Token, ",")
xVarArr = Split(TokenArr(0), "=") '//find variable
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next
Token = TokenArr(1): Call modArtQuotes(Token): TokenArr(1) = LTrim(Token)

If UBound(TokenArr) = 1 Then Token = StrConv(TokenArr(1), S): Token = xVarArr(0) & "=" & Token: _
Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function
Exit Function

'//hash() = Hash text strings
ElseIf InStr(1, Token, "hash(", vbTextCompare) Then

Token = Replace(Token, "hash(", vbNullString, , , vbTextCompare)
Call modArtParens(Token)

'//switches
If InStr(1, Token, "-binary1", vbTextCompare) Then Token = Replace(Token, "-binary1", vbNullString, , , vbTextCompare): S = 1: GoTo setHash

setHash:
xVarArr = Split(Token, "=") '//find variable
xVarArr(1) = Trim(xVarArr(1)): Token = xVarArr(1): Call modArtQuotes(Token)

'//hash(-binary1)
If S = 1 Then
X = Token: Call basBinaryHash1(X, xVerify, xHash): Token = xHash
Token = xVarArr(0) & "=" & Token
Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function
End If

'//repl() = Replace character/string
ElseIf InStr(1, Token, "repl(", vbTextCompare) Then
Token = Replace(Token, "repl(", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)

TokenArr = Split(Token, ",")
xVarArr = Split(TokenArr(0), "=") '//find variable
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next
Token = TokenArr(1): Call modArtQuotes(Token): TokenArr(1) = LTrim(Token)
Token = TokenArr(2): Call modArtQuotes(Token): TokenArr(2) = LTrim(Token)

If UBound(TokenArr) = 2 Then Token = Replace(TokenArr(0), TokenArr(1), TokenArr(2), , , vbBinaryCompare): _
Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function

If UBound(TokenArr) = 3 Then Token = Replace(TokenArr(0), TokenArr(1), TokenArr(2), , , TokenArr(3)): _
Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function
Exit Function

'//ptrim() = Trim starting & ending parentheses of string
ElseIf InStr(1, Token, "ptrim(", vbTextCompare) Then

If InStr(1, Token, "lptrim", vbTextCompare) Then GoTo rmvLParen
If InStr(1, Token, "rptrim", vbTextCompare) Then GoTo rmvRParen

Token = Replace(Token, "ptrim(", vbNullString, , , vbTextCompare)
Call modArtQuotes(Token)

xVarArr = Split(Token, "=") '//find variable
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next

If Left(xVarArr(1), 1) = "(" Then xVarArr(1) = Right(xVarArr(1), Len(xVarArr(1)) - 1):
If Right(xVarArr(1), 1) = ")" Then Token = xVarArr(0) & "=" & Left(xVarArr(1), Len(xVarArr(1)) - 1):
Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function
Exit Function

'//lptrim() = Trim char/string by starting left facing parentheses
rmvLParen:

Token = Replace(Token, "lptrim(", vbNullString, , , vbTextCompare)
Call modArtQuotes(Token)

xVarArr = Split(Token, "=") '//find variable
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next

If Left(xVarArr(1), 1) = "(" Then Token = xVarArr(0) & "=" & Right(xVarArr(1), Len(xVarArr(1)) - 1):
Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function
Exit Function

'//rptrim() = Trim char/string by ending right facing parentheses
rmvRParen:

Token = Replace(Token, "rptrim(", vbNullString, , , vbTextCompare)
Call modArtQuotes(Token)

xVarArr = Split(Token, "=") '//find variable
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next

If Right(xVarArr(1), 1) = ")" Then Token = xVarArr(0) & "=" & Left(xVarArr(1), Len(xVarArr(1)) - 1): _
Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function
Exit Function

'//qtrim() = Trim char/string by quotations
ElseIf InStr(1, Token, "qtrim(", vbTextCompare) Then

Token = Replace(Token, "qtrim(", vbNullString, , , vbTextCompare)
Call modArtParens(Token)

xVarArr = Split(Token, "=") '//find variable
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next

If Left(xVarArr(1), 1) = """" Then xVarArr(1) = Right(xVarArr(1), Len(xVarArr(1)) - 1):
If Right(xVarArr(1), 1) = """" Then Token = xVarArr(0) & "=" & Left(xVarArr(1), Len(xVarArr(1)) - 1):
Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function
Exit Function

'//xtrim() = Trim starting & ending string by chosen character
ElseIf InStr(1, Token, "xtrim(", vbTextCompare) Then

Token = Replace(Token, "xtrim(", vbNullString, , , vbTextCompare)

TokenArr = Split(Token, ",")
xVarArr = Split(TokenArr(0), "=") '//find variable
If UBound(TokenArr) > 1 Then TokenArr(1) = TokenArr(UBound(TokenArr))
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next
Token = TokenArr(1): Call modArtParens(Token): Call modArtQuotes(Token): TokenArr(1) = LTrim(Token)

If Left(xVarArr(1), 1) = TokenArr(1) Then xVarArr(1) = Right(xVarArr(1), Len(xVarArr(1)) - 1): Token = xVarArr(0) & "=" & xVarArr(1):
If Right(xVarArr(1), 1) = TokenArr(1) Then xVarArr(1) = Left(xVarArr(1), Len(xVarArr(1)) - 1): Token = xVarArr(0) & "=" & xVarArr(1):
Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function
Exit Function

'//ins() = Find character/string within a string
ElseIf InStr(1, Token, "ins(", vbTextCompare) Then
Token = Replace(Token, "ins(", vbNullString, , , vbTextCompare)
Token = Replace(Token, """", vbNullString)
Token = Replace(Token, ")", vbNullString)

TokenArr = Split(Token, ",")
xVarArr = Split(TokenArr(0), "=")
If UBound(xVarArr) > 1 Then For X = 2 To UBound(xVarArr): xVarArr(1) = xVarArr(1) & "=" & xVarArr(X): Next
If Left(xVarArr(1), 1) = " " Then xVarArr(1) = Right(xVarArr(1), Len(xVarArr(1)) - 1) '//find variable
'//
If UBound(TokenArr) = 2 Then
If Left(TokenArr(1), 1) = " " Then TokenArr(1) = Right(TokenArr(1), Len(TokenArr(1)) - 1)
If Left(TokenArr(2), 1) = " " Then TokenArr(2) = Right(TokenArr(2), Len(TokenArr(2)) - 1)

If InStr(xVarArr(1), TokenArr(1), TokenArr(2), vbBinaryCompare) Then
Token = appEnv & "[,]" & xVarArr(0) & "=" & "TRUE" & "[,]" & X & "[,]" & 1
    Else
        Token = appEnv & "[,]" & xVarArr(0) & "=" & "FALSE" & "[,]" & X & "[,]" & 1
            End If
                Call xlasExpand(Token, appEnv, appBlk): Exit Function
                    End If
                    
'//
If UBound(TokenArr) = 3 Then
If Left(TokenArr(1), 1) = " " Then TokenArr(1) = Right(TokenArr(1), Len(TokenArr(1)) - 1)
If Left(TokenArr(2), 1) = " " Then TokenArr(2) = Right(TokenArr(2), Len(TokenArr(2)) - 1)
If Left(TokenArr(3), 1) = " " Then TokenArr(3) = Right(TokenArr(3), Len(TokenArr(3)) - 1)
CX = TokenArr(3): Call basCompare(CX)
If InStr(xVarArr(1), TokenArr(1), TokenArr(2), CX) Then
Token = appEnv & "[,]" & xVarArr(0) & "=" & "TRUE" & "[,]" & X & "[,]" & 1
    Else
        Token = appEnv & "[,]" & xVarArr(0) & "=" & "FALSE" & "[,]" & X & "[,]" & 1
                End If
                   Call xlasExpand(Token, appEnv, appBlk): Exit Function
                        End If
Exit Function
   
'//revstr() = Reverse string characters
ElseIf InStr(1, Token, "revstr(", vbTextCompare) Then

Token = Replace(Token, "revstr(", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)

TokenArr = Split(Token, "=") '//find variable

TokenArr(1) = StrReverse(TokenArr(1))
Token = TokenArr(0) & "=" & TokenArr(1)

Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function
Exit Function
'//#
'//
'/\_____________________________________
'//
'//      SYSTEM SHELL/PC ARTICLES
'/\_____________________________________
'//
'//sh() = Quick shell
ElseIf InStr(1, Token, "sh(", vbTextCompare) Then

If InStr(1, Token, "h(0)", vbTextCompare) Then P = 0: Token = Replace(Token, "sh(0)", vbNullString, , , vbTextCompare): GoTo setSh
If InStr(1, Token, "h(1)", vbTextCompare) Then P = 1: Token = Replace(Token, "sh(1)", vbNullString, , , vbTextCompare): GoTo setSh
If InStr(1, Token, "h(2)", vbTextCompare) Then P = 2: Token = Replace(Token, "sh(2)", vbNullString, , , vbTextCompare): GoTo setSh
If InStr(1, Token, "h(3)", vbTextCompare) Then P = 3: Token = Replace(Token, "sh(3)", vbNullString, , , vbTextCompare): GoTo setSh
If InStr(1, Token, "h(4)", vbTextCompare) Then P = 4: Token = Replace(Token, "sh(4)", vbNullString, , , vbTextCompare): GoTo setSh
If InStr(1, Token, "h(5)", vbTextCompare) Then P = 5: Token = Replace(Token, "sh(5)", vbNullString, , , vbTextCompare): GoTo setSh
If InStr(1, Token, "h(6)", vbTextCompare) Then P = 6: Token = Replace(Token, "sh(6)", vbNullString, , , vbTextCompare): GoTo setSh

setSh:
   Token = Replace(Token, "sh(", vbNullString, , , vbTextCompare)
   Call modArtParens(Token): Call modArtQuotes(Token)
    
   FX = Token
   Call basWebFilter(FX) '//check for web filter switches
   If FX <> vbNullString Then Token = FX
   
   Token = "start " & Token
    
   sysShell = Shell("cmd.exe /s /c" & Token, P)
   sysShell = vbNullString
   Exit Function

'//shell32() = System shell
ElseIf InStr(1, Token, "shell32(", vbTextCompare) Then

Token = Replace(Token, "shell32(", vbNullString, , , vbTextCompare)

'//parameters
If InStr(1, Token, ".execute", vbTextCompare) Then
P = 1: Token = Replace(Token, ".execute", vbNullString, , , vbTextCompare)
'//switches
If InStr(1, Token, "-hidden", vbTextCompare) Then Token = Replace(Token, "-hidden", vbNullString, , , vbTextCompare): S = 0: GoTo GetSh32
If InStr(1, Token, "-normal", vbTextCompare) Then Token = Replace(Token, "-normal", vbNullString, , , vbTextCompare): S = 1: GoTo GetSh32
If InStr(1, Token, "-minimized", vbTextCompare) Then Token = Replace(Token, "-minimized", vbNullString, , , vbTextCompare): S = 2: GoTo GetSh32
If InStr(1, Token, "-maximized", vbTextCompare) Then Token = Replace(Token, "-maximized", vbNullString, , , vbTextCompare): S = 3: GoTo GetSh32
'//parameters
ElseIf InStr(1, Token, ".namespace", vbTextCompare) Then
P = 2: Token = Replace(Token, ".namespace", vbNullString, , , vbTextCompare)
'//switches
If InStr(1, Token, "-date", vbTextCompare) Then Token = Replace(Token, "-date", vbNullString, , , vbTextCompare): S = 1: GoTo GetSh32
If InStr(1, Token, "-name", vbTextCompare) Then Token = Replace(Token, "-name", vbNullString, , , vbTextCompare): S = 2: GoTo GetSh32
If InStr(1, Token, "-path", vbTextCompare) Then Token = Replace(Token, "-path", vbNullString, , , vbTextCompare): S = 3: GoTo GetSh32
If InStr(1, Token, "-size", vbTextCompare) Then Token = Replace(Token, "-size", vbNullString, , , vbTextCompare): S = 4: GoTo GetSh32
If InStr(1, Token, "-type", vbTextCompare) Then Token = Replace(Token, "-type", vbNullString, , , vbTextCompare): S = 5: GoTo GetSh32
'//parameters
ElseIf InStr(1, Token, ".information", vbTextCompare) Then
P = 3: Token = Replace(Token, ".information", vbNullString, , , vbTextCompare)
GoTo GetSh32
End If

GetSh32:
Set oShell = CreateObject("Shell.Application")

TokenArr = Split(Token, "=") '//find variable

Token = TokenArr(1): Call modArtParens(Token): Token = Trim(Token)

'//execute
If P = 1 Then
If InStr(1, Token, ",") Then
TokenArrH = Split(Token, ",")
If UBound(TokenArrH) >= 0 Then Token = Trim(TokenArrH(0))
If UBound(TokenArrH) >= 1 Then xArg = Trim(TokenArrH(1))
If UBound(TokenArrH) >= 2 Then xDir = Trim(TokenArrH(2))
If UBound(TokenArrH) >= 3 Then xVar = Trim(TokenArrH(3))
If UBound(TokenArrH) >= 4 Then S = Trim(TokenArrH(4))
End If

oShell.ShellExecute Token, xArg, xDir, xVar, S

Set oShell = Nothing

Token = TokenArr(0) & "=" & Token
Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk)
Exit Function

'//namespace
ElseIf P = 2 Then
FX = Token: Call basShell32Namespace(FX)
If FX <> "*/PATH" Then
Set oShellItem = oShell.Namespace(((CInt(FX))))
    Else
        Set oShellItem = oShell.Namespace((Token))
            End If

If S = 1 Then Token = oShellItem.Self.ModifyDate: GoTo SetSh32
If S = 2 Then Token = oShellItem.Self.name: GoTo SetSh32
If S = 3 Then Token = oShellItem.Self.Path: GoTo SetSh32
If S = 4 Then Token = oShellItem.Self.Size: GoTo SetSh32
If S = 5 Then Token = oShellItem.Self.Type: GoTo SetSh32

Set oShell = Nothing: Set oShellItem = Nothing

SetSh32:
Token = TokenArr(0) & "=" & Token
Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk)
Exit Function

'//system information
ElseIf P = 3 Then
FX = Token: Call basShell32GetSysInfo(FX)
If FX <> "*/ERR" Then

Token = oShell.GetSystemInformation(FX)

Set oShell = Nothing

Token = TokenArr(0) & "=" & Token
Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function
End If
Exit Function
        End If

'//no parameter...
oShell.ShellExecute Token
Token = TokenArr(0) & "=" & Token
Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk)
Exit Function

'//pc() = PC file system tasks
ElseIf InStr(1, Token, "pc(", vbTextCompare) Then

Token = Replace(Token, "pc", vbNullString, , , vbTextCompare)

'//parameters
If InStr(1, Token, ".exist", vbTextCompare) Then P = 1: Token = Replace(Token, ".exist", vbNullString, , , vbTextCompare): GoTo SetPC
If InStr(1, Token, ".del", vbTextCompare) Then P = 2: Token = Replace(Token, ".del", vbNullString, , , vbTextCompare): GoTo SetPC
If InStr(1, Token, ".open", vbTextCompare) Then P = 3: Token = Replace(Token, ".open", vbNullString, , , vbTextCompare): GoTo SetPC
If InStr(1, Token, ".stop", vbTextCompare) Then P = 4: Token = Replace(Token, ".stop", vbNullString, , , vbTextCompare): GoTo SetPC

SetPC:
'//switches
If InStr(1, Token, "-file", vbTextCompare) Then Token = Replace(Token, "-file", vbNullString, , , vbTextCompare): S = 1
If InStr(1, Token, "-fldr", vbTextCompare) Then Token = Replace(Token, "-fldr", vbNullString, , , vbTextCompare): S = 2

Call modArtParens(Token): Call modArtQuotes(Token): Token = Trim(Token)

'//file exists...
If S = 1 And P = 1 Then If Dir(Token) <> "" Then MsgBox "TRUE": Exit Function Else MsgBox ("FALSE"): Exit Function
'//directory exists...
If S = 2 And P = 1 Then If Dir(Token, vbDirectory) <> "" Then MsgBox "TRUE": Exit Function Else MsgBox ("FALSE"): Exit Function
'//delete file...
If S = 1 And P = 2 Then Kill (Token): Exit Function
'//delete empty directory...
If S = 2 And P = 2 Then RmDir (Token): Exit Function
'//open...
If P = 3 Then Token = "start " & Token: sysShell = Shell("cmd.exe /s /c" & Token, 0): sysShell = vbNullString: Exit Function
'//stop (taskkill)
If P = 4 Then Token = "taskkill /f /im " & Token: sysShell = Shell("cmd.exe /s /c" & Token, 0): sysShell = vbNullString: Exit Function

'//no excerpt provided
MsgBox (Token)
Exit Function

'//pc. = PC system control
ElseIf InStr(1, Token, "pc.", vbTextCompare) Then

'//parameters
If InStr(1, Token, ".copy&", vbTextCompare) Then P = 1
If InStr(1, Token, ".copy&!", vbTextCompare) Then P = 2: GoTo setPCdot
If InStr(1, Token, ".shutdown", vbTextCompare) Then P = 3: GoTo setPCdot
If InStr(1, Token, ".off", vbTextCompare) Then P = 4: GoTo setPCdot
If InStr(1, Token, ".rest", vbTextCompare) Then P = 5: GoTo setPCdot
If InStr(1, Token, ".reboot", vbTextCompare) Then P = 6: GoTo setPCdot
If InStr(1, Token, ".clr", vbTextCompare) Then P = 7: GoTo setPCdot

setPCdot:
Token = Replace(Token, "pc.", vbNullString, , , vbTextCompare)
Call modArtParens(Token)
'//article switches
If InStr(1, Token, "-e", vbTextCompare) Then Token = Replace(Token, "-e", vbNullString, , , vbTextCompare): P = P & "1" '//check for switch(s)
If InStr(1, Token, "-t", vbTextCompare) Then '//check for timer switch
Dim xT As String
TokenArr = Split(Token, "-t")
xT = "/t " & TokenArr(1)
End If

'//Copy & paste a file
If P = 1 Then Token = Replace(Token, "copy&", vbNullString, , , vbTextCompare): TokenArr = Split(Token, ","): _
Token = "copy /y " & TokenArr(0) & " " & TokenArr(1): sysShell = Shell("cmd.exe /s /c " & Token, vbNormalFocus): sysShell = vbNullString: Exit Function
'//Copy, paste then delete copied file
If P = 2 Then Token = Replace(Token, "copy&!", vbNullString, , , vbTextCompare): TokenArr = Split(Token, ","): _
Token = "copy /y " & TokenArr(0) & " " & TokenArr(1): sysShell = Shell("cmd.exe /s /c " & Token, vbNormalFocus): sysShell = vbNullString: If Dir(TokenArr(0)) <> "" Then Kill (TokenArr(0)): Exit Function
'//Shutdown pc
If P = 3 Then Token = "shutdown /s " & xT: sysShell = Shell("cmd.exe /s /c " & Token, vbNormalFocus): sysShell = vbNullString: Exit Function
'//Shutdown pc, on next boot auto-sign in if enabled. Restart apps.
If P = 31 Then Token = "shutdown /sg " & xT: sysShell = Shell("cmd.exe /s /c " & Token, vbNormalFocus): sysShell = vbNullString: Exit Function
'//Logoff pc
If P = 4 Then Token = "shutdown /l ": sysShell = Shell("cmd.exe /s /c " & Token, vbNormalFocus): sysShell = vbNullString: Exit Function
'//Hibernate pc
If P = 5 Then Token = "shutdown /h ": sysShell = Shell("cmd.exe /s /c " & Token, vbNormalFocus): sysShell = vbNullString: Exit Function
'//Restart pc
If P = 6 Then Token = "shutdown /r " & xT: sysShell = Shell("cmd.exe /s /c " & Token, vbNormalFocus): sysShell = vbNullString: Exit Function
'//Restart pc, on next boot auto-sign in if enabled. Restart apps.
If P = 61 Then Token = "shutdown /g " & xT: sysShell = Shell("cmd.exe /s /c " & Token, vbNormalFocus): sysShell = vbNullString: Exit Function
'//Clear logoff queue
If P = 7 Then Token = "shutdown /a ": sysShell = Shell("cmd.exe /s /c " & Token, vbNormalFocus): sysShell = vbNullString: Exit Function
'//#
'//
'/\_____________________________________
'//
'//        QUERY ARTICLES
'/\_____________________________________
'//
'//q() = File system query
ElseIf InStr(1, Token, "q(", vbTextCompare) Then

'//parameters
If InStr(1, Token, ".exist", vbTextCompare) Then P = 1: Token = Replace(Token, ".exist", vbNullString, , , vbTextCompare): GoTo setQ
If InStr(1, Token, ".del", vbTextCompare) Then P = 2: Token = Replace(Token, ".del", vbNullString, , , vbTextCompare): GoTo setQ
If InStr(1, Token, ".move", vbTextCompare) Then P = 3: Token = Replace(Token, ".move", vbNullString, , , vbTextCompare): GoTo setQ
If InStr(1, Token, ".name", vbTextCompare) Then P = 4: Token = Replace(Token, ".name", vbNullString, , , vbTextCompare): GoTo setQ
If InStr(1, Token, ".open", vbTextCompare) Then P = 5: Token = Replace(Token, ".open", vbNullString, , , vbTextCompare): GoTo setQ
If InStr(1, Token, ".stop", vbTextCompare) Then P = 6: Token = Replace(Token, ".stop", vbNullString, , , vbTextCompare): GoTo setQ
If P = 0 Then Exit Function

setQ:
'//switches
If InStr(1, Token, "-loose", vbTextCompare) Then Token = Replace(Token, "-loose", vbNullString, , , vbTextCompare): S = 1
If InStr(1, Token, "-strict", vbTextCompare) Then Token = Replace(Token, "-strict", vbNullString, , , vbTextCompare): S = 2
If InStr(1, Token, "-file", vbTextCompare) Then Token = Replace(Token, "-file", vbNullString, , , vbTextCompare): S = S & 3
If InStr(1, Token, "-fldr", vbTextCompare) Then Token = Replace(Token, "-fldr", vbNullString, , , vbTextCompare): S = S & 4

TokenArr = Split(Token, "q(", , vbTextCompare)
If InStr(1, TokenArr(1), ",") Then TokenArr = Split(TokenArr(1), ","): _
Token = TokenArr(0): Call modArtParens(Token): Call modArtQuotes(Token): TokenArr(0) = Trim(Token): Token = TokenArr(0): _
Token = TokenArr(1): Call modArtParens(Token): Call modArtQuotes(Token): TokenArr(1) = Trim(Token): Token = TokenArr(0) Else: _
Token = TokenArr(1): Call modArtParens(Token): Call modArtQuotes(Token): TokenArr(1) = Trim(Token): Token = TokenArr(1)
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oDrv = oFSO.GetFolder("C:\") '//set drive (default is C:)
For Each oSubFldr In oDrv.SubFolders
If InStr(1, Token, oSubFldr.name, vbTextCompare) Then xSubFldr = oSubFldr.name: GoTo hQ '//check for folder match in drive
Next

hQ:
Set oFSO = Nothing
Set oDrv = Nothing
Set oSubFldr = Nothing

Call modArtParens(Token): Call modArtQuotes(Token)

QX = Token
Call basQuery(QX, S)
xQueryArr = Split(QX, ",")

'//exists...
If P = 1 Then If xQueryArr(1) = 0 Then MsgBox ("TRUE" & vbNewLine & vbNewLine & xQueryArr(0)): Exit Function Else MsgBox ("FALSE"): Exit Function
'//delete...
If P = 2 Then Kill (xQueryArr(0)): Exit Function
'//move...
If P = 3 Then xQueryArr(0) = Replace(xQueryArr(0), " ", """" & " " & """"): Token = "move " & xQueryArr(0) & " " & TokenArr(1): sysShell = Shell("cmd.exe /s /c" & Token, 0): sysShell = vbNullString: Exit Function
'//name...
If P = 4 Then xQueryArr(0) = Replace(xQueryArr(0), " ", """" & " " & """"): Token = "ren " & xQueryArr(0) & " " & TokenArr(1): sysShell = Shell("cmd.exe /s /c" & Token, 0): sysShell = vbNullString: Exit Function
'//open...
If P = 5 Then Token = "start " & xQueryArr(0): sysShell = Shell("cmd.exe /s /c" & Token, 0): sysShell = vbNullString: Exit Function
'//stop (taskkill)
If P = 6 Then TokenArr = Split(Token, "\"): Token = TokenArr(UBound(TokenArr)): Token = "taskkill /f /im " & Token: sysShell = Shell("cmd.exe /s /c" & Token, 0): sysShell = vbNullString: Exit Function
Exit Function
'//#
'//
'/\_____________________________________
'//
'//         UTILITY ARTICLES
'/\_____________________________________
'//
ElseIf InStr(1, Token, "incr(", vbTextCompare) Then
Token = Replace(Token, "incr(", vbNullString, , , vbTextCompare)
Call modArtParens(Token)
If InStr(1, Token, "+") Then P = 1: Token = Replace(Token, "+", vbNullString)
If InStr(1, Token, "-") Then P = 2: Token = Replace(Token, "-", vbNullString)
If InStr(1, Token, "=") Then
TokenArr = Split(Token, "=") '//find variable
TokenArr(0) = Trim(TokenArr(0)): TokenArr(1) = Trim(TokenArr(1))

If P = 1 Then TokenArr(1) = CLng(TokenArr(1)) + CLng(TokenArr(1))
If P = 2 Then TokenArr(1) = -(CLng(TokenArr(1))) + -(CLng(TokenArr(1)))

Token = TokenArr(0) & "=" & TokenArr(1)

Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function
End If

If P = 1 Then Token = Token + Token
If P = 2 Then Token = Token - Token

Exit Function

'//rnd() = Get a random number
ElseIf InStr(1, Token, "rnd(", vbTextCompare) Then
Token = Replace(Token, "rnd(", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)

If InStr(1, Token, ":") Then

TokenArr = Split(Token, "=") '//find variable
TokenArr(0) = Trim(TokenArr(0)): TokenArr(1) = Trim(TokenArr(1))

Randomize
xTempArr = Split(TokenArr(1), ":")

TokenArr(1) = CLng((xTempArr(1) * Rnd) + xTempArr(0))

Token = TokenArr(0) & "=" & TokenArr(1)

If UBound(TokenArr) = 1 Then Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function
End If

Exit Function

'//sum() = Perform basic mathematics
ElseIf InStr(1, Token, "sum(", vbTextCompare) Then

Token = Replace(Token, "sum(", vbNullString, , , vbTextCompare)
Call modArtParens(Token): Call modArtQuotes(Token)

TokenArr = Split(Token, "=") '//find variable

Workbooks(appEnv).Worksheets(appBlk).Range("xlasBlkAddr138").Value2 = "=SUM(" & TokenArr(1) & ")"
TokenArr(1) = Workbooks(appEnv).Worksheets(appBlk).Range("xlasBlkAddr138").Value2

Token = TokenArr(0) & "=" & TokenArr(1)

Token = appEnv & "[,]" & Token & "[,]" & X & "[,]" & 1: Call xlasExpand(Token, appEnv, appBlk): Exit Function
Exit Function
'//#
'//
'/\_____________________________________
'//
'//         WINFORM ARTICLES
'/\_____________________________________
'//
'//me() = Output current window number (WinForm)
ElseIf InStr(1, Token, "me(", vbTextCompare) Then
Token = Replace(Token, "me(", vbNullString, , , vbTextCompare)

'//modifiers
If InStr(1, Token, "get.", vbTextCompare) Then M = 1: Token = Replace(Token, "get.", vbNullString, , , vbTextCompare)
If InStr(1, Token, "post.", vbTextCompare) Then M = 2: Token = Replace(Token, "post.", vbNullString, , , vbTextCompare)
If InStr(1, Token, "set.", vbTextCompare) Then M = 3: Token = Replace(Token, "set.", vbNullString, , , vbTextCompare)

'//switches
If InStr(1, Token, "-x", vbTextCompare) Then S = 1: Token = Replace(Token, "-x", vbNullString, , , vbTextCompare)
If InStr(1, Token, "-y", vbTextCompare) Then S = 2: Token = Replace(Token, "-y", vbNullString, , , vbTextCompare)
If InStr(1, Token, "-pos", vbTextCompare) Then S = 3: Token = Replace(Token, "-pos", vbNullString, , , vbTextCompare)

Call modArtParens(Token): Token = Trim(Token)
If InStr(1, Token, ",") Then TokenArr = Split(Token, ",")

Select Case True
Case M = 1 And S = 1: Call basGetWinFormPos(xWin, X, Y): MsgBox (X): Exit Function
Case M = 1 And S = 2: Call basGetWinFormPos(xWin, X, Y): MsgBox (Y): Exit Function
Case M = 1 And S = 3: Call basGetWinFormPos(xWin, X, Y): MsgBox (X & ", " & Y): Exit Function
Case M = 1 And S = 0: Call basGetWinFormPos(xWin, X, Y): MsgBox (X & ", " & Y): Exit Function
Case M = 2 And S = 1: X = Token: Call basPostWinFormPos(xWin, X, Y): Exit Function
Case M = 2 And S = 2: Y = Token: Call basPostWinFormPos(xWin, X, Y): Exit Function
Case M = 2 And S = 3: X = TokenArr(0): Y = TokenArr(1): Call basPostWinFormPos(xWin, X, Y): Exit Function
Case M = 2 And S = 0: X = TokenArr(0): Y = TokenArr(1): Call basPostWinFormPos(xWin, X, Y): Exit Function
Case M = 3 And S = 1: X = Token: Call basSetWinFormPos(xWin, X, Y): Exit Function
Case M = 3 And S = 2: Y = Token: Call basSetWinFormPos(xWin, X, Y): Exit Function
Case M = 3 And S = 3: X = TokenArr(0): Y = TokenArr(1): Call basSetWinFormPos(xWin, X, Y): Exit Function
Case M = 3 And S = 0: Call basSetWinFormPos(xWin, X, Y): Exit Function
End Select
'//no excerpt
MsgBox (Range("xlasWinForm").Value2)
Exit Function

    '//winform() = Set window number
        ElseIf InStr(1, Token, "winform(", vbTextCompare) Then
        
    '//switches
        If InStr(1, Token, "-last", vbTextCompare) Then _
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = _
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinFormLast").Value2: Exit Function  '//set to last window
        
        Token = Replace(Token, "winform(", vbNullString, , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        
        Call xlAppScript_lex.getChar(Token)
        If Token = "*/ERR" Then Exit Function
        
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = Token
        
        Exit Function
        
End If '//end
        
ErrEnd:
'//Article not found...
If errLvl <> 0 Then Token = Token & "*/ERR"
Workbooks(appEnv).Worksheets(appBlk).Range("xlasErrRef").Value = """" & Token & """"
End Function
Private Function libFlag$(Token, errLvl As Byte)

'/\_____________________________________
 '//
'//         FLAGS
'/\_____________________________________
'//
On Error GoTo ErrEnd

Call getEnvironment(appEnv, appBlk)

'//Create runtime error
If InStr(1, Token, "--err", vbTextCompare) Then Token = "*/ERR"

'//Run script w/ environment errors enabled (default)
If InStr(1, Token, "--enableerr", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLibErrLvl") = 0
errLvl = 0
Token = 1: Exit Function
End If

'//Run script w/ environment errors disabled
If InStr(1, Token, "--disableerr", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLibErrLvl") = 1
Range("xlasEnd").Value = 0
errLvl = 0
Token = 1: Exit Function
End If

'//Run script w/ animations/updates disabled (default)
If InStr(1, Token, "--disableupdates", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasUpdateEnable") = 0
Call disableWbUpdates
errLvl = 0
Token = 1: Exit Function
End If

'//Run script w/ animations/updates enabled
If InStr(1, Token, "--enableupdates", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasUpdateEnable") = 1
Call enableWbUpdates
errLvl = 0
Token = 1: Exit Function
End If

'//Run script w/ libraries statically disabled (default)
If InStr(1, Token, "--disablestatic", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLocalStatic") = 0
errLvl = 0
Token = 1: Exit Function
End If

'//Run script w/ libraries statically enabled
If InStr(1, Token, "--enablestatic", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLocalStatic") = 1
errLvl = 0
Token = 1: Exit Function
End If

'//Run script w/ clear runtime block addresses (default)
If InStr(1, Token, "--disablecontain", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLocalContain") = 0
errLvl = 0
Token = 1: Exit Function
End If

'//Run script w/o clearing runtime block addresses
If InStr(1, Token, "--enablecontain", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLocalContain") = 1
errLvl = 0
Token = 1: Exit Function
End If

'//Run script w/o global control variables
If InStr(1, Token, "--disableglobal", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasGlobalControl") = 0
errLvl = 0
Token = 1: Exit Function
End If

'//Run script w global control variables (default)
If InStr(1, Token, "--enableglobal", vbTextCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasGlobalControl").Value = 1
errLvl = 0
Token = 1: Exit Function
End If

Exit Function

ErrEnd:
'//flag not found...
Token = "*/ERR"

End Function
Private Function libSwitch$(Token, errLvl As Byte)

'/\_____________________________________
 '//
'//         LIBRARY SWITCHES
'/\_____________________________________
'//
Dim ArtH As String
Dim X As Integer

On Error GoTo ErrEnd

ArtH = Token
TokenArr = Split(Token, "--")

For X = 0 To UBound(TokenArr)
Token = TokenArr(X): Call modArtParens(Token): Call modArtQuotes(Token): Call modArtSpaces(Token): TokenArr(X) = Token: Token = ArtH
If InStr(1, TokenArr(X), "date", vbTextCompare) Then Token = Replace(Token, "--date", Date, , , vbTextCompare): GoTo NextStep
If InStr(1, TokenArr(X), "day", vbTextCompare) Then Token = Replace(Token, "--day", Day(Date), , , vbTextCompare): GoTo NextStep
If InStr(1, TokenArr(X), "present", vbTextCompare) Then Token = Replace(Token, "--present", Date & " " & Time, , , vbTextCompare): GoTo NextStep
If InStr(1, TokenArr(X), "me", vbTextCompare) Then Token = Replace(Token, "--me", ActiveWorkbook.name, , , vbTextCompare): GoTo NextStep
If InStr(1, TokenArr(X), "month", vbTextCompare) Then Token = Replace(Token, "--month", Month(Date), , , vbTextCompare): GoTo NextStep
If InStr(1, TokenArr(X), "now", vbTextCompare) Then Token = Replace(Token, "--now", Time, , , vbTextCompare): GoTo NextStep
If InStr(1, TokenArr(X), "null", vbTextCompare) Then Token = Replace(Token, "--null", vbNullString, , , vbTextCompare): GoTo NextStep
If InStr(1, TokenArr(X), "lparen", vbTextCompare) Then Token = Replace(Token, "--lparen", "(", , , vbTextCompare): GoTo NextStep
If InStr(1, TokenArr(X), "rparen", vbTextCompare) Then Token = Replace(Token, "--rparen", ")", , , vbTextCompare): GoTo NextStep
If InStr(1, TokenArr(X), "quote", vbTextCompare) Then Token = Replace(Token, "--quote", """", , , vbTextCompare): GoTo NextStep
If InStr(1, TokenArr(X), "space", vbTextCompare) Then Token = Replace(Token, "--space", Space(0), , , vbTextCompare): GoTo NextStep
If InStr(1, TokenArr(X), "time", vbTextCompare) Then Token = Replace(Token, "--time", Time, , , vbTextCompare): GoTo NextStep
If InStr(1, TokenArr(X), "year", vbTextCompare) Then Token = Replace(Token, "--year", Year(Date), , , vbTextCompare): GoTo NextStep
NextStep:
ArtH = Token
Next

Exit Function

ErrEnd:
'//switch not found...
Token = "*/ERR"

End Function
Private Function basWasteTime(ByVal T As Byte) As Byte

T = T + 1: T = T - 1
DoEvents

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
Public Function basGetWinFormPos(ByVal xWin As Object, X, Y) As Integer

'/##########################\
'//   Get WinForm Position #\\
'///########################\\\
On Error Resume Next

If xWin.name = vbNullString Then Call getWindow(xWin)
If X = 0 Then X = xWin.Left
If Y = 0 Then Y = xWin.Top
Set xWin = Nothing

End Function
Public Function basPostWinFormPos(ByVal xWin As Object, ByVal X As Integer, ByVal Y As Integer)

'/#########################\
'// Post WinForm Position #\\
'///#######################\\\
Call getEnvironment(appEnv, appBlk)

On Error Resume Next

If xWin.name = vbNullString Then Call getWindow(xWin)
If X = 0 Then X = xWin.Left
If Y = 0 Then Y = xWin.Top
Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinFormX").Value2 = X
Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinFormY").Value2 = Y
Set xWin = Nothing

End Function
Public Function basSetWinFormPos(ByVal xWin As Object, ByVal X As Integer, ByVal Y As Integer)

'/#########################\
'// Set WinForm Position  #\\
'///#######################\\\
Call getEnvironment(appEnv, appBlk)

On Error Resume Next

If xWin.name = vbNullString Then Call getWindow(xWin)
If X = 0 Then xWin.Left = Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinFormX").Value2 Else xWin.Left = X
If Y = 0 Then xWin.Top = Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinFormY").Value2 Else xWin.Top = Y
Set xWin = Nothing

End Function
Private Function basDecimalToBinary32(ByVal X As Long, ByVal xRtnBits As Byte, xBitStr) As String

'/##############################\
'// Convert Decimal To Binary  #\\
'///############################\\\
'//
'//For converting 32-bit signed integer values to binary using 2's complement

Dim xBit As Byte: Dim xBitStrLen As Byte
Dim xCntr As Byte: Dim IS_NEG As Byte: Dim xRtn As Byte
Dim xBits As String: Dim xNewBitStr As String

xBitStr = vbNullString

'//Check for bit return type
If xRtnBits = 1 Then
xRtn = 8
ElseIf xRtnBits = 0 Then
xRtn = 16
ElseIf xRtnBits = 2 Then
xRtn = 32
End If

'//Check for postive/negative integer
If X < 0 Then IS_NEG = 1

'//Get absolute value of integer
X = Abs(X)

'//Use 2's complement to create bit string
Do Until X <= 0
xBit = (X Mod 2)
If xBit = 1 Then X = X - 1
xBitStr = xBit & xBitStr
X = (X / 2)
Loop

'//Get length of bit string
xBitStrLen = Len(xBitStr)

'//Extend bit string w/ leading 0's
If xBitStrLen < 32 Then
Do Until xBitStrLen = 32
xBitStr = 0 & xBitStr
xBitStrLen = xBitStrLen + 1
Loop
End If

'//Perform bitwise NOT inversion if negative integer
If IS_NEG = 1 Then

For xCntr = 1 To xBitStrLen - 3
xBits = Left(xBitStr, xCntr)
xBit = Mid(xBits, xCntr)
If xBit = 0 Then xBit = 1 Else xBit = 0
xNewBitStr = xNewBitStr & xBit
Next

'//Get 3rd to last bit
xBits = Left(xBitStr, xCntr)
xBit = Mid(xBits, xCntr)
'//Invert 3rd to last bit bit if 0
If xBit = 0 Then xBit = 1

xNewBitStr = xNewBitStr & xBit

'//Add last 2 remaining bits to bit string
xBits = Right(xBitStr, 2)
xNewBitStr = xNewBitStr & xBits

xBitStr = Right(xNewBitStr, xRtn)

End If

End Function
Private Function basBinaryHash1(ByVal X As String, xVerify, xHash) As String

'/##############################\
'//     Basic Binary Hash1     #\\
'///############################\\\
'/
'//For creating/verifying a basic 512-bit binary hash from a given string... (collision prone/unsafe!)
'//
Dim xBit As Byte
Dim xBits As String: Dim xStr As String: Dim xPos As String
Dim xBitStrHash As String
Dim xStrLen As Integer
Dim xBitStrLen As Long: Dim xCntr As Long

xStr = X
xStrLen = Len(xStr)

'//pre-hash
For xCntr = 1 To xStrLen
X = Left(xStr, xCntr)
xPos = Mid(X, xCntr)
X = Asc(xPos)
xRtnBits = 32: Call basDecimalToBinary32(X, xRtnBits, xBitStr)
X = xBitStr

xSlide = Sqr(xCntr)

Do Until xSlide <= 32
xSlide = Sqr(xSlide)
Loop

xPos = xSlide
xSlide = xSlide * 2

'//slide character bits
Call basBitSlideRight(X, xPos, xSlide, xBitStr)
xBitStrHash = xBitStrHash & xBitStr
Next

xBitStrLen = Len(xBitStrHash)

'//create 512 character bit string
If xBitStrLen > 512 Then

xStrLen = Sqr(xBitStrLen)

Do Until xStrLen <= 512
xStrLen = Sqr(xStrLen)
Loop

xCntr = 0
Do Until Len(xBitStrHash) <= 512
xBitStrHash = Left(xBitStrHash, xStrLen + -(xCntr))
xCntr = xCntr - 1
Loop

xBitStrLen = Len(xBitStrHash)

For xCntr = xBitStrLen To 512 - 4
xBitStrHash = xBitStrHash & "0"
Next

xBitStrHash = xBitStrHash & "111"

xBitStrLen = Len(xBitStrHash)

    Else
    
        For xCntr = xBitStrLen To 512 - 4
        xBitStrHash = xBitStrHash & "0"
        Next
        
        xBitStrHash = xBitStrHash & "111"
        
        xBitStrLen = Len(xBitStrHash)

        End If

xHash = xBitStrHash

End Function
Private Function basBitSlideLeft(ByVal X As String, ByVal xPos As Long, ByVal xSlide As Long, xBitStr) As String

'/##############################\
'//  Slide Bit Strings (Left)  #\\
'///############################\\\
'/
'//For sliding a selected bit to the left (will slide all bits in the same direction!)
'//
'//Input string = X
'//
'//Start position = xPos
'//
'//Slide amount = xSlide
'//
'//Return as bit string = xBitStr

Dim xBit As String: Dim xBitStrLen As Byte: Dim xCntr As Byte
Dim xBits As String: Dim xSplitBitStr As String

End Function
Private Function basBitSlideRight(ByVal X As String, ByVal xPos As Long, ByVal xSlide As Long, xBitStr) As String

'/##############################\
'//  Slide Bit Strings (Right) #\\
'///############################\\\
'/
'//For sliding a selected bit to the right (will slide all bits in the same direction!)
'//
'//Input string = X
'//
'//Start position = xPos
'//
'//Slide amount = xSlide
'//
'//Return as bit string = xBitStr

Dim xBit As Byte: Dim xBitStrLen As Byte: Dim xCntr As Byte
Dim xBits As String: Dim xSplitBitStr As String

xBitStrLen = Len(xBitStr)
If xSlide > xBitStrLen Then xSlide = xSlide Mod xBitStrLen

'//split bit string @ position
X = Left(xBitStr, xPos)
xBitStr = Right(xBitStr, Len(xBitStr) - xPos)
'//retrieve bit to slide
xBit = Right(X, 1)
'//split bit string @ retrieved bits new position
xSplitBitStr = Right(xBitStr, Len(xBitStr) - (xSlide - 1))
'//retrieve remaining bits
xBits = Left(xBitStr, xSlide)
'//remove bit to slide from variable
X = Left(X, xPos - 1)
'//seperate remaining bits
xBits = Right(xBitStr, Len(xBits))

'//Check for sliding overlap
If Len(xBitStr) - xSlide < 1 Then
xBitStr = xBit & xBits & X
    Else
    xSplitBitStr = Left(xBitStr, Len(xBitStr) - xSlide)
    xBitStr = xBits & X & xBit & xSplitBitStr
        End If

End Function
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
On Error GoTo ErrEnd
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
                        
                        
                                GoTo ErrEnd: '//nothing found
                                
                                                                
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
                    
                        GoTo ErrEnd: '//nothing found
                               
                               
'//Found our query!
qFound:
Set fso = Nothing: Set oDir = Nothing: Set oLastDir = Nothing: Set SubFldr = Nothing: Set oSubFldr1 = Nothing
QX = QX & "," & 0
Exit Function

End If
    End If

ErrEnd:
Err.Clear
Set fso = Nothing: Set oDir = Nothing: Set oLastDir = Nothing: Set SubFldr = Nothing: Set oSubFldr1 = Nothing
QX = QX & "," & 1

End Function
Private Function basSaveFormat(EX) As String

'/#########################\
'//   Excel Save Formats  #\\
'///#######################\\\

EX = Replace(EX, " ", vbNullString)

Select Case EX
Case Is = "0" Or EX = "AddIn": EX = xlAddIn: Exit Function
Case Is = "1" Or EX = "AddIn8": EX = xlAddIn8: Exit Function
Case Is = "2" Or EX = "CSV": EX = xlCSV: Exit Function
Case Is = "3" Or EX = "CSVMac": EX = xlCSVMac: Exit Function
Case Is = "4" Or EX = "CSVMSDOS": EX = xlCSVMSDOS: Exit Function
Case Is = "5" Or EX = "CSVUTF8": EX = xlCSVUTF8: Exit Function
Case Is = "6" Or EX = "CSVWindows": EX = xlCSVWindows: Exit Function
Case Is = "7" Or EX = "CurrentPlatformText": EX = xlCurrentPlatformText: Exit Function
Case Is = "8" Or EX = "DBF2": EX = xlDBF2: Exit Function
Case Is = "9" Or EX = "DBF3": EX = xlDBF3: Exit Function
Case Is = "10" Or EX = "DBF4": EX = xlDBF4: Exit Function
Case Is = "11" Or EX = "DIF": EX = xlDIF: Exit Function
Case Is = "12" Or EX = "Excel12": EX = xlExcel12: Exit Function
Case Is = "13" Or EX = "Excel2": EX = xlExcel2: Exit Function
Case Is = "14" Or EX = "Excel2FarEast": EX = xlExcel2FarEast: Exit Function
Case Is = "15" Or EX = "Excel3": EX = xlExcel3: Exit Function
Case Is = "16" Or EX = "Excel4": EX = xlExcel4: Exit Function
Case Is = "17" Or EX = "Excel4Workbook": EX = xlExcel4Workbook: Exit Function
Case Is = "18" Or EX = "Excel5": EX = xlExcel5: Exit Function
Case Is = "19" Or EX = "Excel7": EX = xlExcel7: Exit Function
Case Is = "20" Or EX = "Excel8": EX = xlExcel8: Exit Function
Case Is = "21" Or EX = "Excel9795": EX = xlExcel9795: Exit Function
Case Is = "22" Or EX = "Html": EX = xlHtml: Exit Function
Case Is = "23" Or EX = "IntlAddIn": EX = xlIntlAddIn: Exit Function
Case Is = "24" Or EX = "IntlMacro": EX = xlIntlMacro: Exit Function
Case Is = "25" Or EX = "OpenDocumentSpreadsheet": EX = xlOpenDocumentSpreadsheet: Exit Function
Case Is = "26" Or EX = "OpenXMLAddIn": EX = xlOpenXMLAddIn: Exit Function
Case Is = "27" Or EX = "OpenXMLStrictWorkbook": EX = xlOpenXMLStrictWorkbook: Exit Function
Case Is = "28" Or EX = "OpenXMLTemplate": EX = xlOpenXMLTemplate: Exit Function
Case Is = "29" Or EX = "OpenXMLTemplateMacroEnabled": EX = xlOpenXMLTemplateMacroEnabled: Exit Function
Case Is = "30" Or EX = "OpenXMLWorkbook": EX = xlOpenXMLWorkbook: Exit Function
Case Is = "31" Or EX = "OpenXMLWorkbookMacroEnabled": EX = xlOpenXMLWorkbookMacroEnabled: Exit Function
Case Is = "32" Or EX = "SYLK": EX = xlSYLK: Exit Function
Case Is = "33" Or EX = "Template": EX = xlTemplate: Exit Function
Case Is = "34" Or EX = "Template8": EX = xlTemplate8: Exit Function
Case Is = "35" Or EX = "TextMac": EX = xlTextMac: Exit Function
Case Is = "36" Or EX = "TextMSDOS": EX = xlTextMSDOS: Exit Function
Case Is = "37" Or EX = "TextPrinter": EX = xlTextPrinter: Exit Function
Case Is = "38" Or EX = "TextWindows": EX = xlTextWindows: Exit Function
Case Is = "39" Or EX = "UnicodeText": EX = xlUnicodeText: Exit Function
Case Is = "40" Or EX = "WebArchive": EX = xlWebArchive: Exit Function
Case Is = "41" Or EX = "WJ2WD1": EX = xlWJ2WD1: Exit Function
Case Is = "42" Or EX = "WJ3": EX = xlWJ3: Exit Function
Case Is = "43" Or EX = "WJ3FJ3": EX = xlWJ3FJ3: Exit Function
Case Is = "44" Or EX = "WK1": EX = xlWK1: Exit Function
Case Is = "45" Or EX = "WK1ALL": EX = xlWK1ALL: Exit Function
Case Is = "46" Or EX = "WK1FMT": EX = xlWK1FMT: Exit Function
Case Is = "47" Or EX = "WK3": EX = xlWK3: Exit Function
Case Is = "48" Or EX = "WK3FM3": EX = xlWK3FM3: Exit Function
Case Is = "49" Or EX = "WK4": EX = xlWK4: Exit Function
Case Is = "50" Or EX = "WKS": EX = xlWKS: Exit Function
Case Is = "51" Or EX = "WorkbookDefault": EX = xlWorkbookDefault: Exit Function
Case Is = "52" Or EX = "WorkbookNormal": EX = xlWorkbookNormal: Exit Function
Case Is = "53" Or EX = "Works2FarEast": EX = xlWorks2FarEast: Exit Function
Case Is = "54" Or EX = "WQ1": EX = xlWQ1: Exit Function
Case Is = "55" Or EX = "XMLSpreadsheet": EX = xlXMLSpreadsheet: Exit Function
End Select

EX = "*/ERR"

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

'/#########################\
'//      Basic Colors     #\\
'///#######################\\\

Dim xNotColor As String: Dim xRGBC As String
Dim X As Integer: Dim XH As Integer
Dim I As Byte '//waste
xNotColor = "/NULL"

Retry:
X = 1
X = X + 1: If X > (XH) Then I = I: If InStr(1, "aliceblue;#F0F8FF;240,248,255", HX, vbTextCompare) Then xRGBCl = "aliceblue;#F0F8FF;240,248,255": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "antiquewhite;#FAEBD7;250,235,215", HX, vbTextCompare) Then xRGBCl = "antiquewhite;#FAEBD7;250,235,215": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "aqua;#00FFFF;0,255,255", HX, vbTextCompare) Then xRGBCl = "aqua;#00FFFF;0,255,255": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "aquamarine;#7FFFD4;127,255,212", HX, vbTextCompare) Then xRGBCl = "aquamarine;#7FFFD4;127,255,212": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "azure;#F0FFFF;240,255,255", HX, vbTextCompare) Then xRGBCl = "azure;#F0FFFF;240,255,255": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "beige;#F5F5DC;245,245,220", HX, vbTextCompare) Then xRGBCl = "beige;#F5F5DC;245,245,220": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "bisque;#FFE4C4;255,228,196", HX, vbTextCompare) Then xRGBCl = "bisque;#FFE4C4;255,228,196": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "black;#000000;0,0,0", HX, vbTextCompare) Then xRGBCl = "black;#000000;0,0,0": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "blanchedalmond;#FFEBCD;255,235,205", HX, vbTextCompare) Then xRGBCl = "blanchedalmond;#FFEBCD;255,235,205": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "blue;#0000FF;0,0,255", HX, vbTextCompare) Then xRGBCl = "blue;#0000FF;0,0,255": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "blueviolet;#8A2BE2;138,43,226", HX, vbTextCompare) Then xRGBCl = "blueviolet;#8A2BE2;138,43,226": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "brown;#A52A2A;165,42,42", HX, vbTextCompare) Then xRGBCl = "brown;#A52A2A;165,42,42": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "burlywood;#DEB887;222,184,135", HX, vbTextCompare) Then xRGBCl = "burlywood;#DEB887;222,184,135": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "cadetblue;#5F9EA0;95,158,160", HX, vbTextCompare) Then xRGBCl = "cadetblue;#5F9EA0;95,158,160": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "chartreuse;#7FFF00;127,255,0", HX, vbTextCompare) Then xRGBCl = "chartreuse;#7FFF00;127,255,0": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "chocolate;#D2691E;210,105,30", HX, vbTextCompare) Then xRGBCl = "chocolate;#D2691E;210,105,30": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "coral;#FF7F50;255,127,80", HX, vbTextCompare) Then xRGBCl = "coral;#FF7F50;255,127,80": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "cornflowerblue;#6495ED;100,149,237", HX, vbTextCompare) Then xRGBCl = "cornflowerblue;#6495ED;100,149,237": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "cornsilk;#FFF8DC;255,248,220", HX, vbTextCompare) Then xRGBCl = "cornsilk;#FFF8DC;255,248,220": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "crimson;#DC143C;220,20,60", HX, vbTextCompare) Then xRGBCl = "crimson;#DC143C;220,20,60": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "cyan;#00FFFF;0,255,255", HX, vbTextCompare) Then xRGBCl = "cyan;#00FFFF;0,255,255": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dblue;#00008B;0,0,139", HX, vbTextCompare) Then xRGBCl = "dblue;#00008B;0,0,139": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dcyan;#008B8B;0,139,139", HX, vbTextCompare) Then xRGBCl = "dcyan;#008B8B;0,139,139": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "deeppink;#FF1493;255,20,147", HX, vbTextCompare) Then xRGBCl = "deeppink;#FF1493;255,20,147": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "deepskyblue;#00BFFF;0,191,255", HX, vbTextCompare) Then xRGBCl = "deepskyblue;#00BFFF;0,191,255": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dgoldenrod;#B8860B;184,134,11", HX, vbTextCompare) Then xRGBCl = "dgoldenrod;#B8860B;184,134,11": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dgray;#A9A9A9;169,169,169", HX, vbTextCompare) Then xRGBCl = "dgray;#A9A9A9;169,169,169": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dgreen;#006400;0,100,0", HX, vbTextCompare) Then xRGBCl = "dgreen;#006400;0,100,0": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dimgray;#696969;105,105,105", HX, vbTextCompare) Then xRGBCl = "dimgray;#696969;105,105,105": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dkhaki;#BDB76B;189,183,107", HX, vbTextCompare) Then xRGBCl = "dkhaki;#BDB76B;189,183,107": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dmagenta;#8B008B;139,0,139", HX, vbTextCompare) Then xRGBCl = "dmagenta;#8B008B;139,0,139": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dodgerblue;#1E90FF;30,144,255", HX, vbTextCompare) Then xRGBCl = "dodgerblue;#1E90FF;30,144,255": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dolivegreen;#556B2F;85,107,47", HX, vbTextCompare) Then xRGBCl = "dolivegreen;#556B2F;85,107,47": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dorange;#FF8C00;255,140,0", HX, vbTextCompare) Then xRGBCl = "dorange;#FF8C00;255,140,0": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dorchid;#9932CC;153,50,204", HX, vbTextCompare) Then xRGBCl = "dorchid;#9932CC;153,50,204": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dred;#8B0000;139,0,0", HX, vbTextCompare) Then xRGBCl = "dred;#8B0000;139,0,0": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dsalmon;#E9967A;233,150,122", HX, vbTextCompare) Then xRGBCl = "dsalmon;#E9967A;233,150,122": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dseagreen;#8FBC8F;143,188,143", HX, vbTextCompare) Then xRGBCl = "dseagreen;#8FBC8F;143,188,143": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dslateblue;#483D8B;72,61,139", HX, vbTextCompare) Then xRGBCl = "dslateblue;#483D8B;72,61,139": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dslategray;#2F4F4F;47,79,79", HX, vbTextCompare) Then xRGBCl = "dslategray;#2F4F4F;47,79,79": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dturquoise;#00CED1;0,206,209", HX, vbTextCompare) Then xRGBCl = "dturquoise;#00CED1;0,206,209": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "dviolet;#9400D3;148,0,211", HX, vbTextCompare) Then xRGBCl = "dviolet;#9400D3;148,0,211": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "firebrick;#B22222;178,34,34", HX, vbTextCompare) Then xRGBCl = "firebrick;#B22222;178,34,34": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "floralwhite;#FFFAF0;255,250,240", HX, vbTextCompare) Then xRGBCl = "floralwhite;#FFFAF0;255,250,240": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "forestgreen;#228B22;34,139,34", HX, vbTextCompare) Then xRGBCl = "forestgreen;#228B22;34,139,34": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "gainsboro;#DCDCDC;220,220,220", HX, vbTextCompare) Then xRGBCl = "gainsboro;#DCDCDC;220,220,220": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "ghostwhite;#F8F8FF;248,248,255", HX, vbTextCompare) Then xRGBCl = "ghostwhite;#F8F8FF;248,248,255": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "gold;#FFD700;255,215,0", HX, vbTextCompare) Then xRGBCl = "gold;#FFD700;255,215,0": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "goldenrod;#DAA520;218,165,32", HX, vbTextCompare) Then xRGBCl = "goldenrod;#DAA520;218,165,32": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "gray;#808080;128,128,128", HX, vbTextCompare) Then xRGBCl = "gray;#808080;128,128,128": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "green;#008000;0,128,0", HX, vbTextCompare) Then xRGBCl = "green;#008000;0,128,0": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "greenyellow;#ADFF2F;173,255,47", HX, vbTextCompare) Then xRGBCl = "greenyellow;#ADFF2F;173,255,47": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "honeydew;#F0FFF0;240,255,240", HX, vbTextCompare) Then xRGBCl = "honeydew;#F0FFF0;240,255,240": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "hotpink;#FF69B4;255,105,180", HX, vbTextCompare) Then xRGBCl = "hotpink;#FF69B4;255,105,180": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "indigo;#4B0082;75,0,130", HX, vbTextCompare) Then xRGBCl = "indigo;#4B0082;75,0,130": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "ivory;#FFFFF0;255,255,240", HX, vbTextCompare) Then xRGBCl = "ivory;#FFFFF0;255,255,240": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "khaki;#F0E68C;240,230,140", HX, vbTextCompare) Then xRGBCl = "khaki;#F0E68C;240,230,140": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lavender;#E6E6FA;230,230,250", HX, vbTextCompare) Then xRGBCl = "lavender;#E6E6FA;230,230,250": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lavenderblush;#FFF0F5;255,240,245", HX, vbTextCompare) Then xRGBCl = "lavenderblush;#FFF0F5;255,240,245": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lawngreen;#7CFC00;124,252,0", HX, vbTextCompare) Then xRGBCl = "lawngreen;#7CFC00;124,252,0": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lblue;#ADD8E6;173,216,230", HX, vbTextCompare) Then xRGBCl = "lblue;#ADD8E6;173,216,230": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lcoral;#F08080;240,128,128", HX, vbTextCompare) Then xRGBCl = "lcoral;#F08080;240,128,128": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lcyan;#E0FFFF;224,255,255", HX, vbTextCompare) Then xRGBCl = "lcyan;#E0FFFF;224,255,255": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lemonchiffon;#FFFACD;255,250,205", HX, vbTextCompare) Then xRGBCl = "lemonchiffon;#FFFACD;255,250,205": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lgray;#D3D3D3;211,211,211", HX, vbTextCompare) Then xRGBCl = "lgray;#D3D3D3;211,211,211": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lgreen;#90EE90;144,238,144", HX, vbTextCompare) Then xRGBCl = "lgreen;#90EE90;144,238,144": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lime;#00FF00;0,255,0", HX, vbTextCompare) Then xRGBCl = "lime;#00FF00;0,255,0": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "limegreen;#32CD32;50,205,50", HX, vbTextCompare) Then xRGBCl = "limegreen;#32CD32;50,205,50": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "linen;#FAF0E6;250,240,230", HX, vbTextCompare) Then xRGBCl = "linen;#FAF0E6;250,240,230": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lpink;#FFB6C1;255,182,193", HX, vbTextCompare) Then xRGBCl = "lpink;#FFB6C1;255,182,193": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lrodyellow;#FAFAD2;250,250,210", HX, vbTextCompare) Then xRGBCl = "lrodyellow;#FAFAD2;250,250,210": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lsalmon;#FFA07A;255,160,122", HX, vbTextCompare) Then xRGBCl = "lsalmon;#FFA07A;255,160,122": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lseagreen;#20B2AA;32,178,170", HX, vbTextCompare) Then xRGBCl = "lseagreen;#20B2AA;32,178,170": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lskyblue;#87CEFA;135,206,250", HX, vbTextCompare) Then xRGBCl = "lskyblue;#87CEFA;135,206,250": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lslategray;#778899;119,136,153", HX, vbTextCompare) Then xRGBCl = "lslategray;#778899;119,136,153": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lsteelblue;#B0C4DE;176,196,222", HX, vbTextCompare) Then xRGBCl = "lsteelblue;#B0C4DE;176,196,222": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "lyellow;#FFFFE0;255,255,224", HX, vbTextCompare) Then xRGBCl = "lyellow;#FFFFE0;255,255,224": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "magenta;#FF00FF;255,0,255", HX, vbTextCompare) Then xRGBCl = "magenta;#FF00FF;255,0,255": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "maquamarine;#66CDAA;102,205,170", HX, vbTextCompare) Then xRGBCl = "maquamarine;#66CDAA;102,205,170": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "mblue;#0000CD;0,0,205", HX, vbTextCompare) Then xRGBCl = "mblue;#0000CD;0,0,205": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "midnightblue;#191970;25,25,112", HX, vbTextCompare) Then xRGBCl = "midnightblue;#191970;25,25,112": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "mintcream;#F5FFFA;245,255,250", HX, vbTextCompare) Then xRGBCl = "mintcream;#F5FFFA;245,255,250": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "mistyrose;#FFE4E1;255,228,225", HX, vbTextCompare) Then xRGBCl = "mistyrose;#FFE4E1;255,228,225": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "moccasin;#FFE4B5;255,228,181", HX, vbTextCompare) Then xRGBCl = "moccasin;#FFE4B5;255,228,181": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "morchid;#BA55D3;186,85,211", HX, vbTextCompare) Then xRGBCl = "morchid;#BA55D3;186,85,211": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "mpurple;#9370DB;147,112,219", HX, vbTextCompare) Then xRGBCl = "mpurple;#9370DB;147,112,219": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "mseagreen;#3CB371;60,179,113", HX, vbTextCompare) Then xRGBCl = "mseagreen;#3CB371;60,179,113": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "mslateblue;#7B68EE;123,104,238", HX, vbTextCompare) Then xRGBCl = "mslateblue;#7B68EE;123,104,238": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "mspringgreen;#00FA9A;0,250,154", HX, vbTextCompare) Then xRGBCl = "mspringgreen;#00FA9A;0,250,154": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "mturquoise;#48D1CC;72,209,204", HX, vbTextCompare) Then xRGBCl = "mturquoise;#48D1CC;72,209,204": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "mvioletred;#C71585;199,21,133", HX, vbTextCompare) Then xRGBCl = "mvioletred;#C71585;199,21,133": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "navajowhite;#FFDEAD;255,222,173", HX, vbTextCompare) Then xRGBCl = "navajowhite;#FFDEAD;255,222,173": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "navy;#000080;0,0,128", HX, vbTextCompare) Then xRGBCl = "navy;#000080;0,0,128": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "oldlace;#FDF5E6;253,245,230", HX, vbTextCompare) Then xRGBCl = "oldlace;#FDF5E6;253,245,230": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "olive;#808000;128,128,0", HX, vbTextCompare) Then xRGBCl = "olive;#808000;128,128,0": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "olivedrab;#6B8E23;107,142,35", HX, vbTextCompare) Then xRGBCl = "olivedrab;#6B8E23;107,142,35": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "orange;#FFA500;255,165,0", HX, vbTextCompare) Then xRGBCl = "orange;#FFA500;255,165,0": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "orangered;#FF4500;255,69,0", HX, vbTextCompare) Then xRGBCl = "orangered;#FF4500;255,69,0": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "orchid;#DA70D6;218,112,214", HX, vbTextCompare) Then xRGBCl = "orchid;#DA70D6;218,112,214": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "palegoldenrod;#EEE8AA;238,232,170", HX, vbTextCompare) Then xRGBCl = "palegoldenrod;#EEE8AA;238,232,170": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "palegreen;#98FB98;152,251,152", HX, vbTextCompare) Then xRGBCl = "palegreen;#98FB98;152,251,152": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "paleturquoise;#AFEEEE;175,238,238", HX, vbTextCompare) Then xRGBCl = "paleturquoise;#AFEEEE;175,238,238": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "palevioletred;#DB7093;219,112,147", HX, vbTextCompare) Then xRGBCl = "palevioletred;#DB7093;219,112,147": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "papayawhip;#FFEFD5;255,239,213", HX, vbTextCompare) Then xRGBCl = "papayawhip;#FFEFD5;255,239,213": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "peachpuff;#FFDAB9;255,218,185", HX, vbTextCompare) Then xRGBCl = "peachpuff;#FFDAB9;255,218,185": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "peru;#CD853F;205,133,63", HX, vbTextCompare) Then xRGBCl = "peru;#CD853F;205,133,63": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "pink;#FFC0CB;255,192,203", HX, vbTextCompare) Then xRGBCl = "pink;#FFC0CB;255,192,203": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "plum;#DDA0DD;221,160,221", HX, vbTextCompare) Then xRGBCl = "plum;#DDA0DD;221,160,221": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "powderblue;#B0E0E6;176,224,230", HX, vbTextCompare) Then xRGBCl = "powderblue;#B0E0E6;176,224,230": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "purple;#800080;128,0,128", HX, vbTextCompare) Then xRGBCl = "purple;#800080;128,0,128": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "red;#FF0000;255,0,0", HX, vbTextCompare) Then xRGBCl = "red;#FF0000;255,0,0": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "rosybrown;#BC8F8F;188,143,143", HX, vbTextCompare) Then xRGBCl = "rosybrown;#BC8F8F;188,143,143": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "royalblue;#4169E1;65,105,225", HX, vbTextCompare) Then xRGBCl = "royalblue;#4169E1;65,105,225": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "saddlebrown;#8B4513;139,69,19", HX, vbTextCompare) Then xRGBCl = "saddlebrown;#8B4513;139,69,19": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "salmon;#FA8072;250,128,114", HX, vbTextCompare) Then xRGBCl = "salmon;#FA8072;250,128,114": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "sandybrown;#F4A460;244,164,96", HX, vbTextCompare) Then xRGBCl = "sandybrown;#F4A460;244,164,96": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "seagreen;#2E8B57;46,139,87", HX, vbTextCompare) Then xRGBCl = "seagreen;#2E8B57;46,139,87": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "seashell;#FFF5EE;255,245,238", HX, vbTextCompare) Then xRGBCl = "seashell;#FFF5EE;255,245,238": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "sienna;#A0522D;160,82,45", HX, vbTextCompare) Then xRGBCl = "sienna;#A0522D;160,82,45": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "silver;#C0C0C0;192,192,192", HX, vbTextCompare) Then xRGBCl = "silver;#C0C0C0;192,192,192": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "skyblue;#87CEEB;135,206,235", HX, vbTextCompare) Then xRGBCl = "skyblue;#87CEEB;135,206,235": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "slateblue;#6A5ACD;106,90,205", HX, vbTextCompare) Then xRGBCl = "slateblue;#6A5ACD;106,90,205": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "slategray;#708090;112,128,144", HX, vbTextCompare) Then xRGBCl = "slategray;#708090;112,128,144": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "snow;#FFFAFA;255,250,250", HX, vbTextCompare) Then xRGBCl = "snow;#FFFAFA;255,250,250": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "springgreen;#00FF7F;0,255,127", HX, vbTextCompare) Then xRGBCl = "springgreen;#00FF7F;0,255,127": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "steelblue;#4682B4;70,130,180", HX, vbTextCompare) Then xRGBCl = "steelblue;#4682B4;70,130,180": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "tan;#D2B48C;210,180,140", HX, vbTextCompare) Then xRGBCl = "tan;#D2B48C;210,180,140": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "teal;#008080;0,128,128", HX, vbTextCompare) Then xRGBCl = "teal;#008080;0,128,128": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "thistle;#D8BFD8;216,191,216", HX, vbTextCompare) Then xRGBCl = "thistle;#D8BFD8;216,191,216": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "tomato;#FF6347;255,99,71", HX, vbTextCompare) Then xRGBCl = "tomato;#FF6347;255,99,71": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "turquoise;#40E0D0;64,224,208", HX, vbTextCompare) Then xRGBCl = "turquoise;#40E0D0;64,224,208": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "violet;#EE82EE;238,130,238", HX, vbTextCompare) Then xRGBCl = "violet;#EE82EE;238,130,238": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "wheat;#F5DEB3;245,222,179", HX, vbTextCompare) Then xRGBCl = "wheat;#F5DEB3;245,222,179": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "white;#FFFFFF;255,255,255", HX, vbTextCompare) Then xRGBCl = "white;#FFFFFF;255,255,255": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "whitesmoke;#F5F5F5;245,245,245", HX, vbTextCompare) Then xRGBCl = "whitesmoke;#F5F5F5;245,245,245": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "yellow;#FFFF00;255,255,0", HX, vbTextCompare) Then xRGBCl = "yellow;#FFFF00;255,255,0": GoTo ColorFound
X = X + 1: If X > (XH) Then I = I: If InStr(1, "yellowgreen;#9ACD32;154,205,50", HX, vbTextCompare) Then xRGBCl = "yellowgreen;#9ACD32;154,205,50": GoTo ColorFound

Exit Function

ColorFound:

xRGBArr = Split(xRGBCl, ";"): If HX = xRGBArr(0) Then HX = xRGBArr(2): Exit Function

'//color not found
XH = X
xNotColor = xRGBArr(2): GoTo Retry

End Function
Private Function basBorder(BX) As Long

'//get border type
Select Case BX
Case Is = 0: BX = xlNone: Exit Function
Case Is = 1: BX = xlDiagonalDown: Exit Function
Case Is = 2: BX = xlDiagonalUp: Exit Function
Case Is = 3: BX = xlEdgeBottom: Exit Function
Case Is = 4: BX = xlEdgeLeft: Exit Function
Case Is = 5: BX = xlEdgeRight: Exit Function
Case Is = 6: BX = xlEdgeTop: Exit Function
Case Is = 7: BX = xlInsideHorizontal: Exit Function

Case Is = "none": BX = xlNone: Exit Function
Case Is = "ddown": BX = xlDiagonalDown: Exit Function
Case Is = "dup": BX = xlDiagonalUp: Exit Function
Case Is = "bottom": BX = xlEdgeBottom: Exit Function
Case Is = "left": BX = xlEdgeLeft: Exit Function
Case Is = "right": BX = xlEdgeRight: Exit Function
Case Is = "top": BX = xlEdgeTop: Exit Function
Case Is = "inside": BX = xlInsideHorizontal: Exit Function
End Select

End Function
Private Function basBorderStyle(SX) As Long

'//get border style
Select Case SX
Case Is = 0: SX = xlNone: Exit Function
Case Is = 1: SX = xlContinuous: Exit Function
Case Is = 2: SX = xlDash: Exit Function
Case Is = 3: SX = xlDot: Exit Function
Case Is = 4: SX = xlDashDot: Exit Function
Case Is = 5: SX = xlDashDotDot: Exit Function
Case Is = 6: SX = xlSlantDashDot: Exit Function
Case Is = 7: SX = xlDouble: Exit Function

Case Is = "none": SX = xlNone: Exit Function
Case Is = "line": SX = xlContinuous: Exit Function
Case Is = "dash": SX = xlDash: Exit Function
Case Is = "dot": SX = xlDot: Exit Function
Case Is = "ddot": SX = xlDashDot: Exit Function
Case Is = "ddotdot": SX = xlDashDotDot: Exit Function
Case Is = "sddot": SX = xlSlantDashDot: Exit Function
Case Is = "double": SX = xlDouble: Exit Function
End Select

End Function
Private Function basCompare(CX) As Long

'//get comparison type
Select Case CX
Case Is = 0: CX = vbBinaryCompare: Exit Function
Case Is = 1: CX = vbDatabaseCompare: Exit Function
Case Is = 2: CX = vbTextCompare: Exit Function
End Select

End Function
Private Function basPattern(PX) As Long
 
'//get pattern
Select Case PX
Case Is = 0: PX = xlNone: Exit Function
Case Is = 1: PX = xlPatternChecker: Exit Function
Case Is = 2: PX = xlPatternCrissCross: Exit Function
Case Is = 3: PX = xlPatternDown: Exit Function
Case Is = 4: PX = xlPatternHorizontal: Exit Function
Case Is = 5: PX = xlPatternLightDown: Exit Function
Case Is = 6: PX = xlPatternLightHorizontal: Exit Function
Case Is = 7: PX = xlPatternLightUp: Exit Function
Case Is = 8: PX = xlPatternLightVertical: Exit Function
Case Is = 9: PX = xlPatternUp: Exit Function

Case Is = "none": PX = xlNone: Exit Function
Case Is = "pcheck": PX = xlPatternChecker: Exit Function
Case Is = "pcross": PX = xlPatternCrissCross: Exit Function
Case Is = "pdown": PX = xlPatternDown: Exit Function
Case Is = "phori": PX = xlPatternHorizontal: Exit Function
Case Is = "pldown": PX = xlPatternLightDown: Exit Function
Case Is = "plhori": PX = xlPatternLightHorizontal: Exit Function
Case Is = "plup": PX = xlPatternLightUp: Exit Function
Case Is = "plvert": PX = xlPatternLightVertical: Exit Function
Case Is = "pup": PX = xlPatternUp: Exit Function
End Select

End Function
Private Function basShell32Namespace(FX) As Integer

'//get ShellSpecialFolderConstants
  If InStr(1, FX, "ssfDESKTOP", vbTextCompare) Then FX = 0: Exit Function
  If InStr(1, FX, "ssfPROGRAMS", vbTextCompare) Then FX = 2: Exit Function
  If InStr(1, FX, "ssfCONTROLS", vbTextCompare) Then FX = 3: Exit Function
  If InStr(1, FX, "ssfPRINTERS", vbTextCompare) Then FX = 4: Exit Function
  If InStr(1, FX, "ssfPERSONAL", vbTextCompare) Then FX = 5: Exit Function
  If InStr(1, FX, "ssfFAVORITES", vbTextCompare) Then FX = 6: Exit Function
  If InStr(1, FX, "ssfSTARTUP", vbTextCompare) Then FX = 7: Exit Function
  If InStr(1, FX, "ssfRECENT", vbTextCompare) Then FX = 8: Exit Function
  If InStr(1, FX, "ssfSENDTO", vbTextCompare) Then FX = 9: Exit Function
  If InStr(1, FX, "ssfBITBUCKET", vbTextCompare) Then FX = 10: Exit Function
  If InStr(1, FX, "ssfSTARTMENU", vbTextCompare) Then FX = 11: Exit Function
  If InStr(1, FX, "ssfDESKTOPDIRECTORY", vbTextCompare) Then FX = 16: Exit Function
  If InStr(1, FX, "ssfDRIVES", vbTextCompare) Then FX = 17: Exit Function
  If InStr(1, FX, "ssfNETWORK", vbTextCompare) Then FX = 18: Exit Function
  If InStr(1, FX, "ssfNETHOOD", vbTextCompare) Then FX = 19: Exit Function
  If InStr(1, FX, "ssfFONTS", vbTextCompare) Then FX = 20: Exit Function
  If InStr(1, FX, "ssfTEMPLATES", vbTextCompare) Then FX = 21: Exit Function
  If InStr(1, FX, "ssfCOMMONSTARTMENU", vbTextCompare) Then FX = 22: Exit Function
  If InStr(1, FX, "ssfCOMMONPROGRAMS", vbTextCompare) Then FX = 23: Exit Function
  If InStr(1, FX, "ssfCOMMONSTARTUP", vbTextCompare) Then FX = 24: Exit Function
  If InStr(1, FX, "ssfCOMMONDESKTOPDIR", vbTextCompare) Then FX = 25: Exit Function
  If InStr(1, FX, "ssfAPPDATA", vbTextCompare) Then FX = 26: Exit Function
  If InStr(1, FX, "ssfPRINTHOOD", vbTextCompare) Then FX = 27: Exit Function
  If InStr(1, FX, "ssfLOCALAPPDATA", vbTextCompare) Then FX = 28: Exit Function
  If InStr(1, FX, "ssfALTSTARTUP", vbTextCompare) Then FX = 29: Exit Function
  If InStr(1, FX, "ssfCOMMONALTSTARTUP", vbTextCompare) Then FX = 30: Exit Function
  If InStr(1, FX, "ssfCOMMONFAVORITES", vbTextCompare) Then FX = 31: Exit Function
  If InStr(1, FX, "ssfINTERNETCACHE", vbTextCompare) Then FX = 32: Exit Function
  If InStr(1, FX, "ssfCOOKIES", vbTextCompare) Then FX = 33: Exit Function
  If InStr(1, FX, "ssfHISTORY", vbTextCompare) Then FX = 34: Exit Function
  If InStr(1, FX, "ssfCOMMONAPPDATA", vbTextCompare) Then FX = 35: Exit Function
  If InStr(1, FX, "ssfWINDOWS", vbTextCompare) Then FX = 36: Exit Function
  If InStr(1, FX, "ssfSYSTEM", vbTextCompare) Then FX = 37: Exit Function
  If InStr(1, FX, "ssfPROGRAMFILES", vbTextCompare) Then FX = 38: Exit Function
  If InStr(1, FX, "ssfMYPICTURES", vbTextCompare) Then FX = 39: Exit Function
  If InStr(1, FX, "ssfPROFILE", vbTextCompare) Then FX = 40: Exit Function
  If InStr(1, FX, "ssfSYSTEMx86", vbTextCompare) Then FX = 41: Exit Function
  If InStr(1, FX, "ssfPROGRAMFILEFX86", vbTextCompare) Then FX = 42: Exit Function
  
  FX = "*/PATH"
  
End Function
Private Function basShell32GetSysInfo(FX) As String

If FX = "1" Then FX = "DirectoryServiceAvailable": Exit Function
If FX = "2" Then FX = "DoubleClickTime": Exit Function
If FX = "3" Then FX = "ProcessorLevel": Exit Function
If FX = "4" Then FX = "ProcessorSpeed": Exit Function
If FX = "5" Then FX = "ProcessorArchitecture": Exit Function
If FX = "6" Then FX = "PhysicalMemoryInstalled": Exit Function
If FX = "7" Then FX = "IsOS_Professional": Exit Function
If FX = "8" Then FX = "IsOS_Personal": Exit Function
If FX = "9" Then FX = "IsOS_DomainMember": Exit Function

FX = "*/ERR"

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

