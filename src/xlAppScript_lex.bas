Attribute VB_Name = "xlAppScript_lex"
Public Function lexKey(ByVal xArt As String) As Byte
'/\______________________________________________________________________________________________________________________
'//
'//     xlAppScript Lexer
'//        Version: 1.1.2
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
'//     Latest Revision: 6/6/2022
'/\_____________________________________________________________________________________________________________________________
'//
'//     Developer(s): anz7re
'//     Contact: support@xlappscript.org | support@autokit.tech | anz7re@autokit.tech
'//     Web: xlappscript.org | autokit.tech
'/\_____________________________________________________________________________________________________________________________
'//
'//
'//Lexer variables...
Dim mPtr As Long: Dim cPtr As Long: Dim xPtr As Long: Dim xPtrH As Long: Dim R As Long: Dim X As Long
Dim appEnv As String: Dim appBlk As String: Dim xArtArr() As String: Dim xArtArrH() As String: Dim xArtH As String
Dim xKinArr() As String: Dim xPArr() As String: Dim xPtrArr() As String: Dim tArr() As String: Dim FocusInstruct(4) As String
Dim lRow As Integer
Dim C As Byte: Dim E As Byte
C = 0: mPtr = 0: cPtr = 0: xPtr = 0: xPtrH = 0: R = 0: X = 0
'//=============================================================
'//Set runtime environment...
Call fndEnvironment(appEnv, appBlk)
'//=============================================================
'//Check for run tool...
Dim xTool As Object: Call fndRunTool(xTool)
'//=============================================================
'//Set Article from run tool code if found...
If Not xTool Is Nothing Then If xArt = vbNullString Then xArt = xTool.Value
'//If not sending a script from the list of application windows
'//check to make sure an Articles being sent through...
'//
If xArt <> vbNullString Then
'//=============================================================
'//Check for run initializer...
If InStr(1, xArt, "$") Then
'//=============================================================
'//Modify runtime block...
Call modBlk(xArt)
'//=============================================================
'//Remove run identifiers, seperaters, esc...
Call escSpecial(xArt)
xArt = Replace(xArt, vbNewLine, vbNullString)
xArt = Replace(xArt, vbTab, vbNullString)
xArtArr = Split(xArt, ";"): xArtArrH = Split(xArt, ";")
'//=============================================================
Do Until xPtr = UBound(xArtArr)

If InStr(1, xArtArr(xPtr), "~") = False Then
If InStr(1, xArtArr(xPtr), "*/") Then xArt = xArtArr(xPtr): _
xArtH = xArt: Call rtnSpecial(xArt): xArtArr(xPtr) = xArt:
'//=============================================================
'//Check for application runtime environment...
If InStr(1, xArtArr(xPtr), "<env>", vbTextCompare) Then
xArtArr(xPtr) = Replace(xArtArr(xPtr), "<env>", "", , , vbTextCompare)
xArt = xArtArr(xPtr): Call modArtP(xArt): Call modArtQ(xArt): xArt = Trim(xArt)
Workbooks(appEnv).Worksheets(appBlk).Range("xlasEnvironment").Value = xArt
xArtArr(xPtr) = "#"
End If
'//Check for application runtime block...
If InStr(1, xArtArr(xPtr), "<blk>", vbTextCompare) Then
xArtArr(xPtr) = Replace(xArtArr(xPtr), "<blk>", "", , , vbTextCompare)
xArt = xArtArr(xPtr): Call modArtP(xArt): Call modArtQ(xArt): Call modArtS(xArt)
Workbooks(appEnv).Worksheets(appBlk).Range("xlasBlock").Value = xArt
xArtArr(xPtr) = "#"
End If
'//=============================================================
'//Check for xlas library...
If InStr(1, xArtArr(xPtr), "<lib>", vbTextCompare) Then
xArtArr(xPtr) = Replace(xArtArr(xPtr), "<lib>", "", , , vbTextCompare)
xArt = xArtArr(xPtr): Call modArtP(xArt): Call modArtQ(xArt): xArt = Trim(xArt)
lRow = Cells(Rows.Count, "MAL").End(xlUp).Row
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLib").Offset(lRow, 0).Value = xArt
xArtArr(xPtr) = "#"
End If
'//=============================================================
'//Check for variables...
If InStr(1, xArtArr(xPtr), "kin", vbTextCompare) Then
xArtArr(xPtr) = Replace(xArtArr(xPtr), "kin", "", , , vbTextCompare)
If Right(xArtArr(xPtr), Len(xArtArr(xPtr)) - Len(xArtArr(xPtr)) + 1) = ")" Then _
xArtArr(xPtr) = Left(xArtArr(xPtr), Len(xArtArr(xPtr)) - 1)
If Left(xArtArr(xPtr), Len(xArtArr(xPtr)) - Len(xArtArr(xPtr)) + 1) = "(" Then _
xArtArr(xPtr) = Right(xArtArr(xPtr), Len(xArtArr(xPtr)) - 1)
xKinArr = Split(xArtArr(xPtr), "=")

lRow = Cells(Rows.Count, "MAA").End(xlUp).Row
    
    For R = 0 To lRow
    
    If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(R, 0).Value2 = xKinArr(0) Then
    xKinArr(0) = Replace(xKinArr(0), " ", vbNullString)
    Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabelMod").Offset(R, 0).Value2 = "@" & xKinArr(0)
    Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(R, 0).Value2 = xKinArr(1)
    Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(R, 0).Value2 = "@" & xKinArr(0)
    Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValue").Offset(R, 0).Value2 = xKinArr(1)
    xArtArr(xPtr) = "#" & xArtArr(xPtr)
    GoTo NextVar
    Else
        End If
            Next R
            R = 0
            If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(R, 0).Value2 <> "" Then
            Do Until R = lRow: R = R + 1: Loop: End If
            xKinArr(0) = Replace(xKinArr(0), " ", vbNullString)
            Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabelMod").Offset(R, 0).Value2 = "@" & xKinArr(0)
            If xKinArr(1) = vbNullString Then xKinArr(1) = "@" & xKinArr(0)
            Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(R, 0).Value2 = xKinArr(1)
            Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(R, 0).Value2 = "@" & xKinArr(0)
            Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValue").Offset(R, 0).Value2 = xKinArr(1)
            xArtArr(xPtr) = "#" & xArtArr(xPtr)
            xArtH = "#"
            
NextVar:
                End If
                    End If
                    If xArtH <> vbNullString Then xArtArr(xPtr) = xArtH: xArtH = vbNullString
                    xArtArr(xPtr) = Replace(xArtArr(xPtr), "~", ";")
                    xPtr = xPtr + 1
                    Loop
'//=============================================================
'//Record...
xPtr = 0
Do Until xPtr = UBound(xArtArr)
xArt = xArtArr(xPtr)
    If xArt <> "" Then '//Check for empty...
    If InStr(1, xArt, "#") = False Then '//Check for comment...
    '//Record state(s) to memory location...
    Workbooks(appEnv).Worksheets(appBlk).Range("xlasState").Offset(xPtr, 0).Value2 = xPtr
    '//Record Article(s) to memory location...
    Workbooks(appEnv).Worksheets(appBlk).Range("xlasArticle").Offset(xPtr, 0).Value2 = xArt
    End If
        End If
            xPtr = xPtr + 1
            Loop
'//=============================================================
'//Run check...
xPtr = 0
xPtrH = 0
RunCheck:
Do Until xPtr >= UBound(xArtArr)
If mPtr > xPtr Then xPtr = mPtr
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasEnd").Value2 = 1 Then _
Workbooks(appEnv).Worksheets(appBlk).Range("xlasEnd").Value2 = 0: GoTo EndLex '//Check for end...
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasGoto").Value2 <> vbNullString Then xPtr = Range("xlasGoto").Value2 + 1: _
Workbooks(appEnv).Worksheets(appBlk).Range("xlasGoto").Value2 = vbNullString: mPtr = xPtr '//Check for redirection...

xArt = xArtArr(xPtr)

    If xArt <> vbNullString Then '//Check for empty...
    If InStr(1, xArt, "#") = False Then '//Check for comment...
'//=================================================================
ParseCheck:
'//Parse Article/Instruction...
If InStr(1, xArt, "goto ", vbTextCompare) Then C = C + 1
If InStr(1, xArt, "if(", vbTextCompare) Then C = C + 3
If InStr(1, xArt, "do{", vbTextCompare) Then C = C + 5
If InStr(1, xArt, "libcall(", vbTextCompare) Then C = C + 9
If InStr(1, xArt, "let ", vbTextCompare) Then C = C + 11
If xArt = "end" Or xArt = "END" Then C = 99

If C >= 3 Then
xPArr = Split(xArt, "{")
For X = 0 To UBound(xPArr)
If InStr(1, xPArr(X), "goto ", vbTextCompare) Then FocusInstruct(X) = 1
If InStr(1, xPArr(X), "if(", vbTextCompare) Then FocusInstruct(X) = 3
If LCase(xPArr(X)) = "do" Then FocusInstruct(X) = 5
If InStr(1, xPArr(X), "libcall(", vbTextCompare) Then FocusInstruct(X) = 9
If InStr(1, xPArr(X), "let ", vbTextCompare) Then FocusInstruct(X) = 11
If xArt = "end" Or xArt = "END" Then FocusInstruct(X) = 99
Next
    C = FocusInstruct(0)
                        End If
'//=============================================================
    '//goto check...
    If C = 1 Then xArtArr(xPtr) = xArt: xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
    xArt = xArt & "[,]" & xPtr & "[,]" & xPtrH & "[,]" & mPtr & "[,]" & cPtr: _
    E = 0: Call gotoSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
    xPtrArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrArr(3): cPtr = xPtrArr(4): _
    C = 0: xArtArr(xPtr) = xArtArrH(xPtr): GoTo RunCheck
    '//=============================================================
    '//if check...
    If C = 3 Then xArtArr(xPtr) = xArt: xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
    xArt = xArt & "[,]" & xPtr & "[,]" & xPtrH & "[,]" & mPtr & "[,]" & cPtr: _
    E = 0: Call ifSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
    xPtrArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrArr(3): cPtr = xPtrArr(4): _
    C = 0: xArtArr(xPtr) = xArtArrH(xPtr): GoTo ParseCheck
    '//=============================================================
    '//do check...
    If C = 5 Then xArtArr(xPtr) = xArt: xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
    xArt = xArt & "[,]" & xPtr & "[,]" & xPtrH & "[,]" & mPtr & "[,]" & cPtr: _
    E = 0: Call doSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
    xPtrArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrArr(3): cPtr = xPtrArr(4): _
    C = 0: xArtArr(xPtr) = xArtArrH(xPtr): GoTo ParseCheck
    '//=============================================================
    '//libcall check...
    If C = 9 Then xArtArr(xPtr) = xArt: xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
    xArt = xArt & "[,]" & xPtr & "[,]" & xPtr & "[,]" & mPtr & "[,]" & cPtr: _
    E = 0: Call libcallSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
    xPtrArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrArr(3): cPtr = xPtrArr(4): _
    C = 0: xArtArr(xPtr) = xArtArrH(xPtr): GoTo ParseCheck
    '//=============================================================
    '//let check...
    If C = 11 Then xArtArr(xPtr) = xArt: xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
    xArt = xArt & "[,]" & xPtr & "[,]" & mPtr & "[,]" & mPtr & "[,]" & cPtr: _
    E = 0: Call letSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
    xPtrArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrArr(3): cPtr = xPtrArr(4): _
    C = 0: xArtArr(xPtr) = xArtArrH(xPtr): GoTo ParseCheck
    '//=============================================================
    '//end check...
    If C = 99 Then xArtArr(xPtr) = xArt: Call endSet(xArt): GoTo RunCheck
    '//=============================================================
    '//variable check...
    If InStr(1, xArt, "@") Then Call kinSet(xArt) '//set variables...
    '//=============================================================
    '//Run Article...
    Call runScript(xArt)
    If xArt = "(*Err)" Then GoTo ErrRef
    '//=============================================================
    End If
        End If
            mPtr = mPtr + 1
            xPtr = xPtr + 1
            xPtrH = xPtr
                    Loop
    
EndLex:
    If Not xTool Is Nothing Then xTool.Value = Replace(xTool.Value, "$", vbNullString) '//remove run initializer...
    Exit Function
'//=============================================================
'//Error...
ErrRef:
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 <> vbNullString Then Workbooks(appEnv).Worksheets(appBlk).Range("xlasEnd").Value2 = 1
If Not xTool Is Nothing Then xTool.Value = Replace(xTool.Value, "$", vbNullString) '//remove run initializer...
                                End If
                                    End If
                
End Function
Public Function kinSet$(xArt)

'/\____________________________________________________________________________________
'//
'//     A function for identifying a variable & setting it's value before runtime
'/\____________________________________________________________________________________


Dim X As Long
'//Find application environment & block
Dim appEnv, appBlk As String
Call fndEnvironment(appEnv, appBlk)

        lRow = Workbooks(appEnv).Worksheets(appBlk).Cells(Rows.Count, "MAA").End(xlUp).Row
        
        For X = 0 To lRow
ReKin:
        If InStr(1, xArt, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(X, 0).Value2, vbTextCompare) Then
        If InStr(1, xArt, "@env", vbTextCompare) Then Call fndEnvironmentVars(xArt) '//environment variable check
    
                        If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(X, 0).Value2 <> vbNullString Or _
                        InStr(1, xArt, "=") Then
                        
                        '//null variable
                        If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(X, 0).Value2 = vbNullString Then _
                        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(X, 0).Value2 = _
                        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(X, 0).Value2: _
                        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValue").Offset(X, 0).Value2 = _
                        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(X, 0).Value2
                        End If
                        
                        If InStr(1, xArt, "=") = False Then
                        xArt = Replace(xArt, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(X, 0).Value2, _
                        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(X, 0).Value2, vbTextCompare)
                        xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 0: Call kinExpand(xArt)
                        Else
                        xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 0: Call kinExpand(xArt)
                        If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValue").Offset(X, 0).Value2 <> _
                        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(X, 0).Value2 Or _
                        InStr(1, xArt, "(@") Or InStr(1, xArt, " @") Then
                        If InStr(1, xArt, "(@") Or InStr(1, xArt, " @") Then
                        Call modArtP(xArt)
                            xArtArr = Split(xArt, "(@")
                        If InStr(1, "@" & xArtArr(UBound(xArtArr)), Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(X, 0).Value2, vbTextCompare) Then
                        xArt = Replace(xArt, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(X, 0).Value2, _
                        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(X, 0).Value2, vbTextCompare)
                                End If
                                Else
                            End If
                        Else
                        If InStr(1, xArt, "=") = False Or InStr(1, xArt, "[@") Then
                        xArt = Replace(xArt, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(X, 0).Value2, _
                        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(X, 0).Value2, vbTextCompare)
                        End If
                            End If
                                End If
                                
                    Else
                    
                        xArt = Replace(xArt, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(X, 0).Value2, _
                        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(X, 0).Value2, vbTextCompare)
                        
                        If InStr(1, xArt, "@") Then X = X + 1: GoTo ReKin
                        xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 0: Call kinExpand(xArt)
                                        
                                        End If
                                        
                                            If InStr(1, xArt, "@") = False Then Exit Function '//finished
                                            
                                                Next

End Function
Public Function kinExpand$(xArt)

'/\__________________________________________________________________________
'//
'//     A function for expanding the value of a variable before runtime
'/\__________________________________________________________________________

        
        Dim xTest As String
        Dim AX, EX, X, xBlk, xPos As Long
        Dim E As Byte
        xTest = xArt
        Dim appEnv, appBlk As String
        Call fndEnvironment(appEnv, appBlk)
        
        '//extract...
        xTestArr = Split(xTest, ",#!")
        appEnv = xTestArr(0) '//environment
        xTest = xTestArr(1) '//article to test
        X = xTestArr(2) '//position
        E = xTestArr(3) '//environment
        
        If InStr(1, xTest, "=") Then xArtArr = Split(xTest, "="): xArtArr(0) = Replace(xArtArr(0), " ", vbNullString)
        
        '//expanding from library...
        If E = 1 Then GoTo ExAlt
        
        On Error GoTo ExStr
        
        '//Numerical variable expansion,
        '//
        '//Increment
        If InStr(1, xArt, "++") Then x1 = Left(xTest, Len(xTest) - 2): xArt = CDbl(x1) + 1: _
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(X, 0) = xArt: Exit Function
        '//Decrement
        If InStr(1, xArt, "--") Then x1 = Left(xTest, Len(xTest) - 2): xArt = CDbl(x1) - 1: _
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(X, 0) = xArt: Exit Function
        '//Equal
        If InStr(1, xArt, "==") Then
        If InStr(1, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(X, 0).Value2, xArtArr(0)) Then
        If InStr(1, xArtArr(2), "@") Then
        xBlk = Cells(Rows.Count, "MAA").End(xlUp).Row
        For xPos = 0 To xBlk
        If InStr(1, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(xPos, 0).Value2, xArtArr(0)) Then
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(xPos, 0).Value2 = _
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(Range("xlasBlkAddr279").Value2, 0).Value
        ElseIf InStr(1, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(xPos, 0).Value2, xArtArr(2)) Then
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasBlkAddr279").Value2 = xPos
            End If
        Next
        xArt = "#": Exit Function
            End If
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(X, 0).Value2 = xArtArr(2)
        xArt = "#": Exit Function
                    End If
                            End If
        
        xArt = xTest

        Exit Function
        
ExStr:
Err.Clear
If InStr(1, xTest, "=") Then xArt = xTest: Exit Function
If xArtArr = Empty Then Exit Function
If UBound(xArtArr) <= 0 Then Exit Function

        '//String variable expansion
        '//
        For EX = 0 To X
        If InStr(1, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(EX, 0).Value2, xArtArr(0), vbTextCompare) Then '//check for variable in memory
    
        If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(EX, 0).Value2 = vbNullString Or _
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(EX, 0).Value2 = xArtArr(0) Then
        End If
        
        Else
        
        If Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(EX, 0).Value2 = xArtArr(0) Then
        E = 2
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(EX, 0).Value2 = xTest '//replace value in memory
        xArtArr(0) = Replace(xTest, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(EX, 0).Value2, _
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(EX, 0).Value2, vbTextCompare)
        End If
            End If
            
            If InStr(1, xArtArr(0), Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(EX, 0).Value2, vbTextCompare) Then
            xArtArr(0) = Replace(xArtArr(0), Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(EX, 0).Value2, _
            Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(EX, 0).Value2, vbTextCompare)
            End If
            
            Next
        
        If E = 0 Then xArt = xTest: Exit Function
        If E = 1 Then xArt = xArtArr(0): Exit Function
        If E = 2 Then xArt = xTest: Exit Function
        
        Exit Function
        

ExAlt:
X = Cells(Rows.Count, "MAA").End(xlUp).Row

        '//Alternative runtime expansion
        '//
        If InStr(1, xTest, "@") Then
        For EX = 0 To X
        If InStr(1, xArtArr(1), Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(EX, 0).Value2, vbTextCompare) Then xPos = EX
        If InStr(1, xArtArr(0), Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(EX, 0).Value2, vbTextCompare) Then
    
            If InStr(1, xTest, "=") = False Then
            
            xTest = Replace(xTest, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(EX, 0).Value2, _
            Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(EX, 0).Value2, vbTextCompare)
            End If
                
            If InStr(1, xArtArr(1), "@") = False Or InStr(1, xArtArr(0), "%@") Then
            xArt = xArtArr(0): Call modArtM(xArt): xArtArr(0) = xArt
            If UBound(xArtArr) > 1 Then For AX = 2 To UBound(xArtArr): xArtArr(1) = xArtArr(1) & "=" & xArtArr(AX): Next
                Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(EX, 0).Value2 = xArtArr(1)
                Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValue").Offset(EX, 0).Value2 = Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(EX, 0).Value2
                    Else
                    Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(EX, 0).Value2 = _
                    Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(xPos, 0).Value2
                        End If
                            End If
                                Next
                                    End If
        
        '//
        xArt = xTest

End Function
Public Function ifSet$(xArt, E As Byte)

'/\____________________________________________________________________________________
'//
'//     A function for parsing an if statement
'/\____________________________________________________________________________________
'//
'//
'//=============================================================
'//If-Else...
Dim mPtr, cPtr, xPtr, xPtrH, X As Long
Dim appEnv, appBlk As String
'//=============================================================
xArtArr = Split(xArt, "[:]"): xPtrArr = Split(xArt, "[,]")
If UBound(xPtrArr) = 4 Then xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrArr(3): cPtr = xPtrArr(4)
If UBound(xPtrArr) = 8 Then xPtr = xPtrArr(5): xPtrH = xPtrArr(6): mPtr = xPtrArr(7): cPtr = xPtrArr(8)
If UBound(xPtrArr) = 12 Then xPtr = xPtrArr(9): xPtrH = xPtrArr(10): mPtr = xPtrArr(11): cPtr = xPtrArr(12)
If UBound(xPtrArr) = 16 Then xPtr = xPtrArr(13): xPtrH = xPtrArr(14): mPtr = xPtrArr(15): cPtr = xPtrArr(16)
If cPtr < 1 Then cPtr = 1
'//=============================================================
BArr = Split(xArtArr(mPtr), "){"): xArt = BArr(0) '//find boolean placemarker
If InStr(1, xArt, "@") Then Call kinSet(xArt)
If UBound(BArr) = 1 Then xArtH = xArt & "){" & BArr(1): xArtArr(xPtrH) = xArtH: xArt = xArtH '//set variables...
If UBound(BArr) = 2 Then xArtH = xArt & "){" & BArr(1) & "){" & BArr(2): xArtArr(xPtrH) = xArtH: xArt = xArtH '//set variables...

Dim B, C, Opr, ECntr, FCntr, S As String
Dim A As Byte
S = vbNullString

mPtr = xPtrH
Do Until mPtr = UBound(xArtArr)
If InStr(1, xArtArr(mPtr), "else", vbTextCompare) Then ECntr = mPtr
If InStr(1, xArtArr(mPtr), "endif", vbTextCompare) Then FCntr = mPtr: GoTo FindBool
mPtr = mPtr + 1
Loop

GoTo RtnArt
'//=============================================================
'//Boolean check...
FindBool:
xArt = xArtArr(xPtrH)

If InStr(1, xArt, "-and", vbTextCompare) Then S = "1"
If InStr(1, xArt, "-or", vbTextCompare) Then S = "2"
If InStr(1, xArt, "-nor", vbTextCompare) Then S = "3"
If InStr(1, xArt, "-xor", vbTextCompare) Then S = "4"

'//...and...
If S = "1" Then
BArr = Split(xArt, "-and", , vbTextCompare)
xBArr = Split(BArr(1), "){")
BArr(0) = Replace(BArr(0), "if", vbNullString, , , vbTextCompare)
xArt = xBArr(0): Call modArtP(xArt): Call modArtQ(xArt): BArr(1) = xArt
xArt = BArr(0): Call modArtP(xArt): Call modArtQ(xArt): BArr(0) = xArt
For X = 0 To UBound(BArr)
If InStr(1, BArr(X), "-gt", vbTextCompare) Then Opr = "-gt": C = 1: GoTo CompAnd
If InStr(1, BArr(X), "-lt", vbTextCompare) Then Opr = "-lt": C = 2: GoTo CompAnd
If InStr(1, BArr(X), "-ge", vbTextCompare) Then Opr = "-ge": C = 3: GoTo CompAnd
If InStr(1, BArr(X), "-le", vbTextCompare) Then Opr = "-le": C = 4: GoTo CompAnd
If InStr(1, BArr(X), "-eq", vbTextCompare) Then Opr = "-eq": C = 5: GoTo CompAnd

CompAnd:
CArr = Split(BArr(X), Opr)
CArr(0) = Trim(CArr(0)): CArr(1) = Trim(CArr(1))
If C = 1 Then If CDbl(CArr(0)) > CDbl(CArr(1)) Then A = A + 1
If C = 2 Then If CDbl(CArr(0)) < CDbl(CArr(1)) Then A = A + 1
If C = 3 Then If CDbl(CArr(0)) >= CDbl(CArr(1)) Then A = A + 1
If C = 4 Then If CDbl(CArr(0)) <= CDbl(CArr(1)) Then A = A + 1
If C = 5 Then If (CArr(0)) = (CArr(1)) Then A = A + 1
Next
If A = 2 Then B = "T": BArr(1) = xBArr(1)
GoTo SetBool
End If

'//...or...
If S = "2" Then
BArr = Split(xArt, "-or", , vbTextCompare)
xBArr = Split(BArr(1), "){")
BArr(0) = Replace(BArr(0), "if", vbNullString, , , vbTextCompare)
xArt = xBArr(0): Call modArtP(xArt): Call modArtQ(xArt): BArr(1) = xArt
xArt = BArr(0): Call modArtP(xArt): Call modArtQ(xArt): BArr(0) = xArt
For X = 0 To UBound(BArr)
If InStr(1, BArr(X), "-gt", vbTextCompare) Then Opr = "-gt": C = 1: GoTo CompOr
If InStr(1, BArr(X), "-lt", vbTextCompare) Then Opr = "-lt": C = 2: GoTo CompOr
If InStr(1, BArr(X), "-ge", vbTextCompare) Then Opr = "-ge": C = 3: GoTo CompOr
If InStr(1, BArr(X), "-le", vbTextCompare) Then Opr = "-le": C = 4: GoTo CompOr
If InStr(1, BArr(X), "-eq", vbTextCompare) Then Opr = "-eq": C = 5: GoTo CompOr

CompOr:
CArr = Split(BArr(X), Opr)
CArr(0) = Trim(CArr(0)): CArr(1) = Trim(CArr(1))
If C = 1 Then If CDbl(CArr(0)) > CDbl(CArr(1)) Then B = "T": BArr(1) = xBArr(1): GoTo SetBool
If C = 2 Then If CDbl(CArr(0)) < CDbl(CArr(1)) Then B = "T": BArr(1) = xBArr(1): GoTo SetBool
If C = 3 Then If CDbl(CArr(0)) >= CDbl(CArr(1)) Then B = "T": BArr(1) = xBArr(1): GoTo SetBool
If C = 4 Then If CDbl(CArr(0)) <= CDbl(CArr(1)) Then B = "T": BArr(1) = xBArr(1): GoTo SetBool
If C = 5 Then If (CArr(0)) = (CArr(1)) Then B = "T": BArr(1) = xBArr(1): GoTo SetBool
Next
End If

'//...nor...
If S = "3" Then
BArr = Split(xArt, "-nor", , vbTextCompare)
xBArr = Split(BArr(1), "){")
BArr(0) = Replace(BArr(0), "if", vbNullString, , , vbTextCompare)
xArt = xBArr(0): Call modArtP(xArt): Call modArtQ(xArt): BArr(1) = xArt
xArt = BArr(0): Call modArtP(xArt): Call modArtQ(xArt): BArr(0) = xArt
For X = 0 To UBound(BArr)
If InStr(1, BArr(X), "-gt", vbTextCompare) Then Opr = "-gt": C = 1: GoTo CompNor
If InStr(1, BArr(X), "-lt", vbTextCompare) Then Opr = "-lt": C = 2: GoTo CompNor
If InStr(1, BArr(X), "-ge", vbTextCompare) Then Opr = "-ge": C = 3: GoTo CompNor
If InStr(1, BArr(X), "-le", vbTextCompare) Then Opr = "-le": C = 4: GoTo CompNor
If InStr(1, BArr(X), "-eq", vbTextCompare) Then Opr = "-eq": C = 5: GoTo CompNor

CompNor:
CArr = Split(BArr(X), Opr)
CArr(0) = Trim(CArr(0)): CArr(1) = Trim(CArr(1))
If C = 1 Then If CDbl(CArr(0)) > CDbl(CArr(1)) Then A = A + 1
If C = 2 Then If CDbl(CArr(0)) < CDbl(CArr(1)) Then A = A + 1
If C = 3 Then If CDbl(CArr(0)) >= CDbl(CArr(1)) Then A = A + 1
If C = 4 Then If CDbl(CArr(0)) <= CDbl(CArr(1)) Then A = A + 1
If C = 5 Then If (CArr(0)) = (CArr(1)) Then A = A + 1
Next
If A = 0 Then B = "T": BArr(1) = xBArr(1)
GoTo SetBool
End If

'//...xor...
If S = "4" Then
BArr = Split(xArt, "-xor", , vbTextCompare)
xBArr = Split(BArr(1), "){")
BArr(0) = Replace(BArr(0), "if", vbNullString, , , vbTextCompare)
xArt = xBArr(0): Call modArtP(xArt): Call modArtQ(xArt): BArr(1) = xArt
xArt = BArr(0): Call modArtP(xArt): Call modArtQ(xArt): BArr(0) = xArt
For X = 0 To UBound(BArr)
If InStr(1, BArr(X), "-gt", vbTextCompare) Then Opr = "-gt": C = 1: GoTo CompXor
If InStr(1, BArr(X), "-lt", vbTextCompare) Then Opr = "-lt": C = 2: GoTo CompXor
If InStr(1, BArr(X), "-ge", vbTextCompare) Then Opr = "-ge": C = 3: GoTo CompXor
If InStr(1, BArr(X), "-le", vbTextCompare) Then Opr = "-le": C = 4: GoTo CompXor
If InStr(1, BArr(X), "-eq", vbTextCompare) Then Opr = "-eq": C = 5: GoTo CompXor

CompXor:
CArr = Split(BArr(X), Opr)
CArr(0) = Trim(CArr(0)): CArr(1) = Trim(CArr(1))
If C = 1 Then If CDbl(CArr(0)) > CDbl(CArr(1)) Then A = A + 1
If C = 2 Then If CDbl(CArr(0)) < CDbl(CArr(1)) Then A = A + 1
If C = 3 Then If CDbl(CArr(0)) >= CDbl(CArr(1)) Then A = A + 1
If C = 4 Then If CDbl(CArr(0)) <= CDbl(CArr(1)) Then A = A + 1
If C = 5 Then If (CArr(0)) = (CArr(1)) Then A = A + 1
Next
If A = 1 Then B = "T": BArr(1) = xBArr(1)
GoTo SetBool
End If

'//...no operator...
BArr = Split(xArt, "){")
BArr(0) = Replace(BArr(0), "if", vbNullString, , , vbTextCompare)
xArt = BArr(0): Call modArtP(xArt): Call modArtQ(xArt): BArr(0) = xArt
If InStr(1, BArr(0), "-gt", vbTextCompare) Then Opr = "-gt": C = 1: GoTo CompDef
If InStr(1, BArr(0), "-lt", vbTextCompare) Then Opr = "-lt": C = 2: GoTo CompDef
If InStr(1, BArr(0), "-ge", vbTextCompare) Then Opr = "-ge": C = 3: GoTo CompDef
If InStr(1, BArr(0), "-le", vbTextCompare) Then Opr = "-le": C = 4: GoTo CompDef
If InStr(1, BArr(0), "-eq", vbTextCompare) Then Opr = "-eq": C = 5: GoTo CompDef

CompDef:
CArr = Split(BArr(0), Opr)
CArr(0) = Trim(CArr(0)): CArr(1) = Trim(CArr(1))
If C = 1 Then If CDbl(CArr(0)) > CDbl(CArr(1)) Then B = "T": GoTo SetBool
If C = 2 Then If CDbl(CArr(0)) < CDbl(CArr(1)) Then B = "T": GoTo SetBool
If C = 3 Then If CDbl(CArr(0)) >= CDbl(CArr(1)) Then B = "T": GoTo SetBool
If C = 4 Then If CDbl(CArr(0)) <= CDbl(CArr(1)) Then B = "T": GoTo SetBool
If C = 5 Then If (CArr(0)) = (CArr(1)) Then B = "T": GoTo SetBool

SetBool:

xArtArr(xPtrH) = xArtH '//put original article back in it's place
'//=============================================================
'//True...
If B = "T" Then
If UBound(BArr) = 1 Then xArt = BArr(1)
If UBound(BArr) = 2 Then xArt = BArr(1) & "){" & BArr(2)
If E = 0 Then mPtr = xPtr
If E = 1 Then mPtr = xPtr
GoTo RtnArt
    Else
'//=============================================================
'//False...
If ECntr <> "" Then mPtr = ECntr: GoTo FindElse
xArt = "#"
End If
If E = 0 Then xPtr = mPtr: xPtrH = xPtr
If E = 1 Then xPtrH = FCntr
GoTo RtnArt
'//=============================================================
'//Else...
FindElse:

xElseArr = Split(xArtArr(mPtr), "else")
xElseArr(1) = Replace(xElseArr(1), "{", vbNullString)
xArt = xElseArr(1)
If E = 0 Then xPtr = mPtr
If E = 1 Then xPtr = xPtrH: xPtrH = mPtr
GoTo RtnArt

'//return article
RtnArt:

xArt = xArt & "[:]" & "[,]" & xPtr & "[,]" & xPtrH & "[,]" & mPtr & "[,]" & cPtr

End Function
Public Function doSet$(xArt, E As Byte)

'/\____________________________________________________________________________________
'//
'//     A function for parsing a do loop
'/\____________________________________________________________________________________
'//
'//
'//=============================================================
'//Do Loop...
Dim mPtr, cPtr, xPtr, xPtrH, X As Long
Dim FocusInstruct(4) As String
Dim appEnv, appBlk As String
Dim C As Byte
'//=============================================================
xArtArr = Split(xArt, "[:]"): xPtrArr = Split(xArt, "[,]")
If UBound(xPtrArr) = 4 Then xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrArr(3): cPtr = xPtrArr(4)
If UBound(xPtrArr) = 8 Then xPtr = xPtrArr(5): xPtrH = xPtrArr(6): mPtr = xPtrArr(7): cPtr = xPtrArr(8)
If UBound(xPtrArr) = 12 Then xPtr = xPtrArr(9): xPtrH = xPtrArr(10): mPtr = xPtrArr(11): cPtr = xPtrArr(12)
If UBound(xPtrArr) = 16 Then xPtr = xPtrArr(13): xPtrH = xPtrArr(14): mPtr = xPtrArr(15): cPtr = xPtrArr(16)
'//=============================================================
If mPtr < 0 Then mPtr = 0
If cPtr < 1 Then cPtr = 1

Do While mPtr < UBound(xArtArr)
If cPtr <= 0 Then GoTo RtnArt '//End of looping..

   '//Read & run Article loop...
    If InStr(1, xArtArr(mPtr), "}loop", vbTextCompare) Then '//find loop iterator
    
    cPtr = Replace(xArtArr(mPtr), "}loop", "", , , vbTextCompare)
    cPtr = CLng(cPtr) * -1
    
    If CInt(xPtrH) < CInt(xPtr) Then xPtrH = xPtr '//set to first article after do statement
    
    Do Until cPtr <= 0
    
    Do Until InStr(1, xArtArr(xPtrH), "}loop", vbTextCompare) Or cPtr <= 0
    xArtArr(xPtrH) = Replace(xArtArr(xPtrH), "do{", "", , , vbTextCompare)
    xArt = xArtArr(xPtrH)
'//=================================================================
ParseCheck:
'//Parse Article/Instruction...
If InStr(1, xArt, "goto ", vbTextCompare) Then C = C + 1
If InStr(1, xArt, "if(", vbTextCompare) Then C = C + 3
If InStr(1, xArt, "do{", vbTextCompare) Then C = C + 5
If InStr(1, xArt, "libcall(", vbTextCompare) Then C = C + 9
If InStr(1, xArt, "let ", vbTextCompare) Then C = C + 11
If xArt = "end" Or xArt = "END" Then C = 99

If C >= 3 Then
xPArr = Split(xArt, "{")
For X = 0 To UBound(xPArr)
If InStr(1, xPArr(X), "goto ", vbTextCompare) Then FocusInstruct(X) = 1
If InStr(1, xPArr(X), "if(", vbTextCompare) Then FocusInstruct(X) = 3
If LCase(xPArr(X)) = "do" Then FocusInstruct(X) = 5
If InStr(1, xPArr(X), "libcall(", vbTextCompare) Then FocusInstruct(X) = 9
If InStr(1, xPArr(X), "let ", vbTextCompare) Then FocusInstruct(X) = 11
If xArt = "end" Or xArt = "END" Then FocusInstruct(X) = 99
Next
    C = FocusInstruct(0)
                        End If
'//=============================================================
'//goto check...
If C = 1 Then xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
xArt = xArt & "[,]" & xPtr & "[,]" & xPtrH & "[,]" & mPtr & "[,]" & cPtr: _
E = 1: Call gotoSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
xPtrArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = mPtr: cPtr = xPtrArr(4): _
C = 0: xArt = vbNullString: GoTo ParseCheck
'//=============================================================
'//if check...
If C = 3 Then xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
xArt = xArt & "[,]" & xPtr & "[,]" & xPtrH & "[,]" & xPtrH & "[,]" & cPtr: _
E = 1: Call ifSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
xPtrArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = mPtr: cPtr = xPtrArr(4): _
 C = 0: GoTo ParseCheck
'//=============================================================
'//do check...
If C = 5 Then xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
xArt = xArt & "[,]" & xPtr & "[,]" & xPtrH & "[,]" & mPtr & "[,]" & cPtr: _
E = 1: Call doSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
xPtrArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrArr(3): cPtr = xPtrArr(4): _
C = 0: GoTo ParseCheck
'//=============================================================
'//libcall check...
If C = 9 Then xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
xArt = xArt & "[,]" & xPtr & "[,]" & xPtr & "[,]" & mPtr & "[,]" & cPtr: _
E = 1: Call libcallSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
xPtrArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrArr(3): cPtr = xPtrArr(4): _
C = 0: GoTo ParseCheck
'//=============================================================
'//let check...
If C = 11 Then xArtArr(xPtr) = xArt: xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
xArt = xArt & "[,]" & xPtr & "[,]" & xPtrH & "[,]" & mPtr & "[,]" & cPtr: _
E = 1: Call letSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
xPtrArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrArr(3): cPtr = xPtrArr(4): _
C = 0: xArt = vbNullString: GoTo ParseCheck
'//=============================================================
'//end check...
If C = 99 Then xArtArr(xPtr) = xArt: Call endSet(xArt): GoTo RtnArt
'//=============================================================

    If InStr(1, xArt, "@") Then Call kinSet(xArt) '//set variables...
    Call runScript(xArt) '//run article...
    If xArt = "(*Err)" Then GoTo RtnArt
    
    xPtrH = xPtrH + 1
    
        Loop
            
            xPtrH = xPtr
            cPtr = cPtr - 1
            
                Loop
        
    Else
                    End If
                    
                        mPtr = mPtr + 1
                 
                            Loop
                            
                                GoTo RtnArt
                    
                                    
'//return article
RtnArt:

If CInt(xPtrH) > CInt(mPtr) Then mPtr = xPtrH: If cPtr = 0 Then mPtr = mPtr + 1

xArt = "#"
xArt = xArt & "[:]" & "[,]" & mPtr - 1 & "[,]" & mPtr - 1 & "[,]" & mPtr - 1 & "[,]" & cPtr

End Function
Public Function gotoSet$(xArt, E As Byte)

'/\____________________________________________________________________________________
'//
'//     A function for parsing a goto redirect
'/\____________________________________________________________________________________
'//
'//
Dim mPtr, cPtr, xPtr, xPtrH, X As Long
Dim appEnv, appBlk As String
'//=============================================================
xArtArr = Split(xArt, "[:]"): xPtrArr = Split(xArt, "[,]")
If UBound(xPtrArr) = 4 Then xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrArr(3): cPtr = xPtrArr(4)
If UBound(xPtrArr) = 8 Then xPtr = xPtrArr(5): xPtrH = xPtrArr(6): mPtr = xPtrArr(7): cPtr = xPtrArr(8)
If UBound(xPtrArr) = 12 Then xPtr = xPtrArr(9): xPtrH = xPtrArr(10): mPtr = xPtrArr(11): cPtr = xPtrArr(12)
If UBound(xPtrArr) = 16 Then xPtr = xPtrArr(13): xPtrH = xPtrArr(14): mPtr = xPtrArr(15): cPtr = xPtrArr(16)
'//=============================================================

Call fndEnvironment(appEnv, appBlk)

xGArr = Split(xArtArr(xPtrH), "goto ", , vbTextCompare)
xArt = xGArr(1): Call modArtP(xArt): Call modArtQ(xArt): Call modArtS(xArt)
xArt = Trim(xArt)

X = 0
Do While X < UBound(xArtArr)
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasArticle").Offset(X, 0) = ":" & xArt Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasGoto").Value22 = _
Workbooks(appEnv).Worksheets(appBlk).Range("xlasState").Offset(X, 0).Value2 '//find state...
mPtr = Workbooks(appEnv).Worksheets(appBlk).Range("xlasState").Offset(X, 0).Value2
E = 2
xArt = "#"
GoTo RtnArt
        End If
            X = X + 1
                    Loop

'//return article
RtnArt:

xArt = xArt & "[:]" & "[,]" & mPtr & "[,]" & mPtr & "[,]" & mPtr & "[,]" & cPtr

End Function
Public Function libcallSet$(xArt, E As Byte)

'/\____________________________________________________________________________________
'//
'//     A function for parsing & library call
'/\____________________________________________________________________________________
'//
'//
'//=============================================================
'//Library Call...
Dim mPtr, cPtr, xPtr, xPtrH, X As Long
Dim FocusInstruct(4), xLArr(100) As String
Dim appEnv, appBlk, xLib As String
Dim C As Byte
'//=============================================================
xArtArr = Split(xArt, "[:]"): xArtArrH = Split(xArt, "[:]")
xPtrArr = Split(xArt, "[,]")
If UBound(xPtrArr) = 4 Then xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrArr(3): cPtr = xPtrArr(4)
If UBound(xPtrArr) = 8 Then xPtr = xPtrArr(5): xPtrH = xPtrArr(6): mPtr = xPtrArr(7): cPtr = xPtrArr(8)
If UBound(xPtrArr) = 12 Then xPtr = xPtrArr(9): xPtrH = xPtrArr(10): mPtr = xPtrArr(11): cPtr = xPtrArr(12)
If UBound(xPtrArr) = 16 Then xPtr = xPtrArr(13): xPtrH = xPtrArr(14): mPtr = xPtrArr(15): cPtr = xPtrArr(16)
'//=============================================================
If mPtr < 0 Then mPtr = 0
If CInt(mPtr) > CInt(xPtrH) Then mPtr = xPtrH

Do While mPtr < UBound(xArtArr)
 
    If InStr(1, xArtArr(mPtr), "}endcall", vbTextCompare) Then '//find end of call marker
   
    lRow = Cells(Rows.Count, "MAL").End(xlUp).Row
    
    '//extract library...
    xArtArr(xPtrH) = Replace(xArtArr(xPtrH), "libcall(", "", , , vbTextCompare)
    xTempArr = Split(xArtArr(xPtrH), "libcall"): xTempArr = Split(xTempArr(0), "){")
    
    For X = 0 To UBound(xTempArr)
    xTempArr(X) = Replace(xTempArr(X), "(", "*/LPAREN")
    xTempArr(X) = Replace(xTempArr(X), ")", "*/RPAREN")
    If InStr(1, "*/LPAREN", xTempArr(X)) = False Then If InStr(1, "*/RPAREN", xTempArr(X)) = False Then xLib = xTempArr(X):
    If InStr(1, xLib, "*/LPAREN") = False Then If InStr(1, xLib, "*/RPAREN") = False Then GoTo LibFound
    Next
   
LibFound:
    xArtArr(xPtrH) = Replace(xArtArr(xPtrH), xLib & "){", vbNullString)
    xArt = xLib: Call modArtQ(xArt): Call modArtS(xArt)
    
    '//keep track of previous librarie(s)...
    For X = 1 To lRow
    xLArr(X) = Range("xlasLib").Offset(lRow, 0).Value2
    Next
    
    '//set called library to top of stack
    Range("xlasLib").Offset(1, 0).Value2 = xArt
'//=================================================================
    Do Until InStr(1, xArtArr(xPtrH), "}endcall", vbTextCompare)
   
    xArt = xArtArr(xPtrH)
'//=================================================================
ParseCheck:

If E = 2 Then GoTo RtnArt '//redirect

'//Parse Article/Instruction...
If InStr(1, xArt, "goto ", vbTextCompare) Then C = C + 1
If InStr(1, xArt, "if(", vbTextCompare) Then C = C + 3
If InStr(1, xArt, "do{", vbTextCompare) Then C = C + 5
If InStr(1, xArt, "libcall(", vbTextCompare) Then C = C + 9
If InStr(1, xArt, "let ", vbTextCompare) Then C = C + 11
If xArt = "end" Or xArt = "END" Then C = 99

If C >= 3 Then
xPArr = Split(xArt, "{")
For X = 0 To UBound(xPArr)
If InStr(1, xPArr(X), "goto ", vbTextCompare) Then FocusInstruct(X) = 1
If InStr(1, xPArr(X), "if(", vbTextCompare) Then FocusInstruct(X) = 3
If LCase(xPArr(X)) = "do" Then FocusInstruct(X) = 5
If InStr(1, xPArr(X), "libcall(", vbTextCompare) Then FocusInstruct(X) = 9
If InStr(1, xPArr(X), "let ", vbTextCompare) Then FocusInstruct(X) = 11
If xArt = "end" Or xArt = "END" Then FocusInstruct(X) = 99
Next
    C = FocusInstruct(0)
                        End If
'//=============================================================
'//goto check...
If C = 1 Then xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
xArt = xArt & "[,]" & xPtr & "[,]" & xPtrH & "[,]" & mPtr & "[,]" & cPtr: _
E = 1: Call gotoSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
xPtrArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = mPtr: cPtr = xPtrArr(4): _
C = 0: xArt = vbNullString: GoTo ParseCheck
'//=============================================================
'//if check...
If C = 3 Then xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
xArt = xArt & "[,]" & xPtr & "[,]" & xPtrH & "[,]" & xPtrH & "[,]" & cPtr: _
E = 1: Call ifSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
xPtrArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = mPtr: cPtr = xPtrArr(4): _
 C = 0: GoTo ParseCheck
'//=============================================================
'//do check...
If C = 5 Then xArtArr(xPtr) = xArt: xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
xArt = xArt & "[,]" & xPtr & "[,]" & xPtrH & "[,]" & xPtrH & "[,]" & xPtr: _
E = 1: Call doSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
tArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = tArr(2): mPtr = tArr(3): cPtr = xPtrArr(4): _
C = 0: xArtArr(xPtr) = xArtArrH(xPtr): GoTo ParseCheck
'//=============================================================
'//libcall check...
If C = 9 Then xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
xArt = xArt & "[,]" & xPtr & "[,]" & xPtr & "[,]" & xPtrH & "[,]" & cPtr: _
E = 1: Call libcallSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
xPtrArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrArr(3): cPtr = xPtrArr(4): _
C = 0: GoTo ParseCheck
'//=============================================================
'//let check...
If C = 11 Then xArtArr(xPtr) = xArt: xArt = vbNullString: For X = 0 To UBound(xArtArr): xArt = xArt & xArtArr(X) & "[:]": Next: _
xArt = xArt & "[,]" & xPtr & "[,]" & xPtrH & "[,]" & mPtr & "[,]" & cPtr: _
E = 1: Call letSet(xArt, E): tArr = Split(xArt, "[:]"): xArt = tArr(0): _
xPtrArr = Split(tArr(1), "[,]"): xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrH: cPtr = xPtrArr(4): _
C = 0: xArt = vbNullString: GoTo ParseCheck
'//=============================================================
'//end check...
If C = 99 Then xArtArr(xPtr) = xArt: Call endSet(xArt): GoTo RtnArt
'//=============================================================
    If InStr(1, xArt, "@") Then Call kinSet(xArt) '//set variables...
    Call runScript(xArt) '//run article...
    If xArt = "(*Err)" Then GoTo RtnArt
    
    xPtrH = xPtrH + 1
    
    
        Loop
            
            
            
    Else
                    End If
                    
                    '//end of library call...
                    If InStr(1, xArtArr(mPtr), "}endcall", vbTextCompare) And xLib <> vbNullString Then GoTo RtnArt
                    mPtr = mPtr + 1
                        xArt = "#"
                 
                            Loop
                    
                                GoTo RtnArt
                    
                                    
'//return article
RtnArt:
If CInt(xPtr) > CInt(mPtr) Then mPtr = xPtr - 1

xArt = "#": X = 1
Do Until xLArr(X) = vbNullString
Range("xlasLib").Offset(lRow, 0).Value2 = xLArr(X) '//return previous librarie(s) to stack
Loop

xArt = xArt & "[:]" & "[,]" & xPtr & "[,]" & mPtr & "[,]" & mPtr & "[,]" & cPtr

End Function
Public Function letSet$(xArt, E As Byte)

'/\____________________________________________________________________________________
'//
'//     A function for parsing a let statement
'/\____________________________________________________________________________________
'//
'//
Dim mPtr, cPtr, xPtr, xPtrH, X As Long
Dim appEnv, appBlk As String
'//=============================================================
xArtArr = Split(xArt, "[:]"): xPtrArr = Split(xArt, "[,]")
If UBound(xPtrArr) = 4 Then xPtr = xPtrArr(1): xPtrH = xPtrArr(2): mPtr = xPtrArr(3): cPtr = xPtrArr(4)
If UBound(xPtrArr) = 8 Then xPtr = xPtrArr(5): xPtrH = xPtrArr(6): mPtr = xPtrArr(7): cPtr = xPtrArr(8)
If UBound(xPtrArr) = 12 Then xPtr = xPtrArr(9): xPtrH = xPtrArr(10): mPtr = xPtrArr(11): cPtr = xPtrArr(12)
If UBound(xPtrArr) = 16 Then xPtr = xPtrArr(13): xPtrH = xPtrArr(14): mPtr = xPtrArr(15): cPtr = xPtrArr(16)
'//=============================================================

Call fndEnvironment(appEnv, appBlk)

xLetArr = Split(xArtArr(xPtrH), "let ", , vbTextCompare)

xSetArr = Split(xLetArr(1), "=")
xArt = xSetArr(0): Call modArtP(xArt): Call modArtQ(xArt): xArt = Trim(xArt): xSetArr(0) = xArt
xArt = xSetArr(1): Call modArtP(xArt): Call modArtQ(xArt): xArt = Trim(xArt): xSetArr(1) = xArt

xArt = xSetArr(0) & "=" & xSetArr(1)

xArt = appEnv & ",#!" & xArt & ",#!" & X & ",#!" & 1: Call kinExpand(xArt)

'//return article
RtnArt:

xArt = "#"

xArt = xArt & "[:]" & "[,]" & xPtr & "[,]" & xPtrH & "[,]" & mPtr & "[,]" & cPtr

End Function
Public Function endSet$(xArt)

'/\____________________________________________________________________________________
'//
'//     A function for parsing an end statement
'/\____________________________________________________________________________________
'//
'//
'//End script...
If xArt = "end" Or xArt = "END" Then Range("xlasEnd").Value = 1: Exit Function

End Function
Public Function escSpecial$(xArt)

'/\__________________________________________________________________________________
'//
'//     A function for escaping ([#]) special characters or strings before runtime
'/\__________________________________________________________________________________


Dim xArtH$

xArtH = xArt

If InStr(1, xArtH, "[doll]") Then
'//escape dollar symbol
xArtH = Replace(xArtH, "[doll]", "*/DOLLAR")
End If

If InStr(1, xArtH, "[hash]") Then
'//escape hash symbol
xArtH = Replace(xArtH, "[hash]", "*/HASH")
End If

If InStr(1, xArtH, "[semi]") Then
'//escape semicolon symbol
xArtH = Replace(xArtH, "[semi]", "*/SEMICOLON")
End If

If InStr(1, xArtH, "[tab]", vbTextCompare) Then
'//escape tab symbol
xArtH = Replace(xArtH, "[tab]", "*/TAB")
End If

If InStr(1, xArtH, "[feed]", vbTextCompare) Then
'//escape linefeed symbol
xArtH = Replace(xArtH, "[feed]", "*/FEED")
End If

If InStr(1, xArtH, "[nl]", vbTextCompare) Then
'//escape newline symbol
xArtH = Replace(xArtH, "[nl]", "*/NEWLINE")
End If

If InStr(1, xArtH, "[null]", vbTextCompare) Then
'//escape null symbol
xArtH = Replace(xArtH, "[null]", "*/NULL")
End If

If InStr(1, xArtH, "[space]", vbTextCompare) Then
'//escape space symbol
xArtH = Replace(xArtH, "[space]", "*/SPACE")
End If

If InStr(1, xArtH, "[lp]", vbTextCompare) Then
'//escape left parenthese symbol
xArtH = Replace(xArtH, "[lp]", "*/LPAREN")
End If

If InStr(1, xArtH, "[rp]", vbTextCompare) Then
'//escape right parentheses symbol
xArtH = Replace(xArtH, "[rp]", "*/RPAREN")
End If

xArt = xArtH

End Function
Public Function rtnSpecial$(xArt)

'/\__________________________________________________________________________________
'//
'//     A function for returning ([#]) special characters or strings before runtime
'/\__________________________________________________________________________________

Dim xArtH$

xArtH = xArt

If InStr(1, xArtH, "*/DOLLAR") Then
'//return dollar symbol
xArtH = Replace(xArtH, "*/DOLLAR", "$")
End If

If InStr(1, xArtH, "*/HASH") Then
'//return hash symbol
xArtH = Replace(xArtH, "*/HASH", "#")
End If

If InStr(1, xArtH, "*/SEMICOLON") Then
'//return semicolon symbol
xArtH = Replace(xArtH, "*/SEMICOLON", ";")
End If

If InStr(1, xArtH, "*/TAB") Then
'//return tab
xArtH = Replace(xArtH, "*/TAB", vbTab)
End If

If InStr(1, xArtH, "*/FEED") Then
'//return line feed
xArtH = Replace(xArtH, "*/FEED", vbLf)
End If

If InStr(1, xArtH, "*/NEWLINE") Then
'//return newline
xArtH = Replace(xArtH, "*/NEWLINE", vbNewLine)
End If

If InStr(1, xArtH, "*/NULL") Then
'//return null
xArtH = Replace(xArtH, "*/NULL", vbNullString)
End If

If InStr(1, xArtH, "*/SPACE") Then
'//return space
xArtH = Replace(xArtH, "*/SPACE", " ")
End If

If InStr(1, xArtH, "*/LPAREN") Then
'//return left parentheses
xArtH = Replace(xArtH, "*/LPAREN", "(")
End If

If InStr(1, xArtH, "*/RPAREN") Then
'//return right parentheses
xArtH = Replace(xArtH, "*/RPAREN", ")")
End If

xArt = xArtH

End Function
Public Function fndChar$(xArt)

'/\__________________________________________________________________________
'//
'//     A function for finding if a non-numerical characters been entered
'/\__________________________________________________________________________


On Error Resume Next

Dim xChar$
Dim xLetter%

'//Check for character(s)
xChar = "a,A,b,B,c,C,d,D,e,E,f,F,g,G,h,H,i,I,j,J,k,K,l,L,m,M,n,N,o,O,p,P,q,Q,r,R,s,S,t,T,u,U,v,V,w,W,x,X,y,Y,z,Z,`,~,!,@,#,$,%,^,&,*,(,),_,-,+,=,[,],{,},\,|,;,',<,>,?,/,."
xChar = xChar & ","","

X = 1
xLetters = Split(xChar, ",")
xLast = UBound(xLetters) - LBound(xLetters)

Do Until X = xLast
If InStr(1, xArt, xLetters(X)) Then xArt = "(*Err)": Exit Function
X = X + 1
Loop

xChar = xArt

End Function
Public Function fndEnvironment$(appEnv, appBlk)

'/\__________________________________________________________________________
'//
'//     A function for finding the current runtime environment & block
'/\__________________________________________________________________________

Dim B, N As Object

        On Error GoTo useThisEnv
        
        '//Set application runtime environment (Workbook)...
        If Range("xlasEnvironment").Value2 <> vbNullString Then appEnv = Range("xlasEnvironment").Value2
        
        '//Set application runtime block (Worksheet)...
        If Range("xlasBlock").Value2 <> vbNullString Then appBlk = Range("xlasBlock").Value2
        
        If appEnv <> vbNullString And appBlk <> vbNullString Then Exit Function '//exit if both set, otherwise find
        
        GoTo findBlock

'//Environment not found or specified (Default)
useThisEnv:
Err.Clear
If Range("xlasEnvironment").Value2 <> vbNullString Then appEnv = Range("xlasEnvironment").Value2 Else appEnv = ThisWorkbook.name

'//Find application runtime block...
findBlock:
For Each B In Workbooks(appEnv).Sheets
    For Each N In Workbooks(appEnv).Names
        If N.name = "xlasEnvironment" Then appBlk = B.name: Range("xlasBlock").Value2 = appBlk: _
        Range("xlasEnvironment").Value2 = appEnv: Exit Function
            Next
                Next

'//Block not found or specified (Default)
appBlk = appBlk: Range("xlasBlock").Value2 = appBlk: Range("xlasEnvironment").Value2 = appEnv: Exit Function

End Function
Public Function fndEnvironmentVars$(xArt)

'/\__________________________________________________________________________
'//
'//     A function for identifying & setting an environment variable
'/\__________________________________________________________________________

If InStr(1, xArt, "@envallusersprofile", vbTextCompare) Then xArt = Replace(xArt, "@envallusersprofile", Environ("ALLUSERSPROFILE"), , , vbTextCompare)
If InStr(1, xArt, "@envappdata", vbTextCompare) Then xArt = Replace(xArt, "@envappdata", Environ("APPDATA"), , , vbTextCompare)
If InStr(1, xArt, "@envcommonprogramfiles", vbTextCompare) Then xArt = Replace(xArt, "@envcommonprogramfiles", Environ("CommonProgramFiles"), , , vbTextCompare)
If InStr(1, xArt, "@envcommonprogramfilesx86", vbTextCompare) Then xArt = Replace(xArt, "@envcommonprogramfilesx86", Environ("CommonProgramFiles(x86)"), , , vbTextCompare)
If InStr(1, xArt, "@envcom", vbTextCompare) Then xArt = Replace(xArt, "@envcom", Environ("COMPUTERNAME"), , , vbTextCompare)
If InStr(1, xArt, "@envcomspec", vbTextCompare) Then xArt = Replace(xArt, "@envcomspec", Environ("ComSpec"), , , vbTextCompare)
If InStr(1, xArt, "@envdriverdata", vbTextCompare) Then xArt = Replace(xArt, "@envdriverdata", Environ("DriverData"), , , vbTextCompare)
If InStr(1, xArt, "@envhomedrive", vbTextCompare) Then xArt = Replace(xArt, "@envhomedrive", Environ("HOMEDRIVE"), , , vbTextCompare)
If InStr(1, xArt, "@envhome", vbTextCompare) Then xArt = Replace(xArt, "@envhome", Environ("HOMEPATH"), , , vbTextCompare)
If InStr(1, xArt, "@envlocalappdata", vbTextCompare) Then xArt = Replace(xArt, "@envlocalappdata", Environ("LOCALAPPDATA"), , , vbTextCompare)
If InStr(1, xArt, "@envlogonserver", vbTextCompare) Then xArt = Replace(xArt, "@envlogonserver", Environ("LOGONSERVER"), , , vbTextCompare)
If InStr(1, xArt, "@envnumofprocessors", vbTextCompare) Then xArt = Replace(xArt, "@envnumofprocessors", Environ("NUMBER_OF_PROCESSORS"), , , vbTextCompare)
If InStr(1, xArt, "@envonedrive", vbTextCompare) Then xArt = Replace(xArt, "@envonedrive", Environ("OneDrive"), , , vbTextCompare)
If InStr(1, xArt, "@envonlineservice", vbTextCompare) Then xArt = Replace(xArt, "@envonlineservices", Environ("OnlineServices"), , , vbTextCompare)
If InStr(1, xArt, "@envos", vbTextCompare) Then xArt = Replace(xArt, "@envos", Environ("OS"), , , vbTextCompare)
If InStr(1, xArt, "@envpath", vbTextCompare) Then xArt = Replace(xArt, "@envpath", Environ("PATH"), , , vbTextCompare)
If InStr(1, xArt, "@envpathext", vbTextCompare) Then xArt = Replace(xArt, "@envpathext", Environ("PATHEXT"), , , vbTextCompare)
If InStr(1, xArt, "@envplatformcode", vbTextCompare) Then xArt = Replace(xArt, "@envplatformcode", Environ("PlatformCode"), , , vbTextCompare)
If InStr(1, xArt, "@envprocessorarch", vbTextCompare) Then xArt = Replace(xArt, "@envprocessorarch", Environ("PROCESSOR_ARCHITECTURE"), , , vbTextCompare)
If InStr(1, xArt, "@envprocessorid", vbTextCompare) Then xArt = Replace(xArt, "@envprocessorid", Environ("PROCESSOR_IDENTIFIER"), , , vbTextCompare)
If InStr(1, xArt, "@envprocessorlvl", vbTextCompare) Then xArt = Replace(xArt, "@envprocessorlvl", Environ("PROCESSOR_LEVEL"), , , vbTextCompare)
If InStr(1, xArt, "@envprocessorrev", vbTextCompare) Then xArt = Replace(xArt, "@envprocessorrev", Environ("PROCESSOR_REVISION"), , , vbTextCompare)
If InStr(1, xArt, "@envprogramdata", vbTextCompare) Then xArt = Replace(xArt, "@envprogramdata", Environ("ProgramData"), , , vbTextCompare)
If InStr(1, xArt, "@envprogramfiles", vbTextCompare) Then xArt = Replace(xArt, "@envprogramfiles", Environ("ProgramFiles"), , , vbTextCompare)
If InStr(1, xArt, "@envprogramfilesx86", vbTextCompare) Then xArt = Replace(xArt, "@envprogramfilesx86", Environ("ProgramFiles(x86)"), , , vbTextCompare)
If InStr(1, xArt, "@envpsmodulepath", vbTextCompare) Then xArt = Replace(xArt, "@envpsmodulepath", Environ("PSModulePath"), , , vbTextCompare)
If InStr(1, xArt, "@envpublic", vbTextCompare) Then xArt = Replace(xArt, "@envpublic", Environ("PUBLIC"), , , vbTextCompare)
If InStr(1, xArt, "@envregioncode", vbTextCompare) Then xArt = Replace(xArt, "@envregioncode", Environ("RegionCode"), , , vbTextCompare)
If InStr(1, xArt, "@envsessionname", vbTextCompare) Then xArt = Replace(xArt, "@envsessionname", Environ("SESSIONNAME"), , , vbTextCompare)
If InStr(1, xArt, "@envsysdrive", vbTextCompare) Then xArt = Replace(xArt, "@envsysdrive", Environ("SystemDrive"), , , vbTextCompare)
If InStr(1, xArt, "@envsysroot", vbTextCompare) Then xArt = Replace(xArt, "@envsysroot", Environ("SystemRoot"), , , vbTextCompare)
If InStr(1, xArt, "@envtemp", vbTextCompare) Then xArt = Replace(xArt, "@envtemp", Environ("TEMP"), , , vbTextCompare)
If InStr(1, xArt, "@envtmp", vbTextCompare) Then xArt = Replace(xArt, "@envtmp", Environ("TMP"), , , vbTextCompare)
If InStr(1, xArt, "@envuserdomain", vbTextCompare) Then xArt = Replace(xArt, "@envuserdomain", Environ("USERDOMAIN"), , , vbTextCompare)
If InStr(1, xArt, "@envuserdomainrp", vbTextCompare) Then xArt = Replace(xArt, "@envuserdomainrp", Environ("USERDOMAIN_ROAMINGPROFILE"), , , vbTextCompare)
If InStr(1, xArt, "@envuserprofile", vbTextCompare) Then xArt = Replace(xArt, "@envuserprofile", Environ("USERPROFILE"), , , vbTextCompare)
If InStr(1, xArt, "@envuser", vbTextCompare) Then xArt = Replace(xArt, "@envuser", Environ("USERNAME"), , , vbTextCompare)
If InStr(1, xArt, "@envwindir", vbTextCompare) Then xArt = Replace(xArt, "@envwindir", Environ("windir"), , , vbTextCompare)

End Function
Public Function fndJunk%(errLvl)

'/\________________________________________________________________________________
'//
'//     A function for getting rid of junk before it's parsed
'/\________________________________________________________________________________
'//
'//Reset article from ErrLvl
xArt = errLvl
'//Reset error level to 0
errLvl = 0
'/\_____________________________________
'//
'//     JUNK FOUND
'/\_____________________________________
'//
If InStr(1, xArt, "#") Then errLvl = 1: Exit Function
If InStr(1, xArt, "[:]") Then errLvl = 1: Exit Function
If InStr(1, xArt, "[,]") Then errLvl = 1: Exit Function

If Len(xArt) <= 2 Then errLvl = 1: Exit Function

'//Check for number...
On Error GoTo NotJunk
If CDbl(xArt) * 1 = CDbl(xArt) Then
errLvl = 1: Exit Function
    Else
        End If
       
NotJunk:
Err.Clear
'/\_____________________________________
'//
'//     LABEL FOUND
'/\_____________________________________
'//
If Left(xArt, Len(xArt) - Len(xArt) + 1) = ":" Then errLvl = 1: Exit Function
'//
'/\_____________________________________
'//
'//     LOOP/CONDITIONAL FOUND
'/\_____________________________________
'//
If Left(xArt, Len(xArt) - Len(xArt) + 1) = "}" Then errLvl = 1: Exit Function
If Left(xArt, Len(xArt) - Len(xArt) + 2) = "if" Then errLvl = 1: Exit Function

End Function
Private Function fndLibs$(xLib)

'/\______________________________________________________________________________________
'//
'//     A function for finding an xlAppScript library based on it's column position
'/\______________________________________________________________________________________

Dim appEnv, appBlk As String
Call fndEnvironment(appEnv, appBlk)

xLib = "xlAppScript_" & Workbooks(appEnv).Worksheets(appBlk).Range("xlasLib").Offset(xLib, 0).Value2 & ".runLib"

End Function
Public Function fndWindow(xWin) As Object

'/\__________________________________________________________________________
'//
'//     A function for finding the current active Window/UserForm object
'/\__________________________________________________________________________

Dim appEnv, appBlk As String
Call fndEnvironment(appEnv, appBlk)

'//Check for current running window
'//
'//eTweetXL: WinForms
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 1 Then Set xWin = ETWEETXLHOME: Exit Function
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 2 Then Set xWin = ETWEETXLSETUP: Exit Function
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 3 Then Set xWin = ETWEETXLPOST: Exit Function
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 4 Then Set xWin = ETWEETXLQUEUE: Exit Function
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 5 Then Set xWin = ETWEETXLAPISETUP: Exit Function
'//eTweetXL: Input Fields
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 31 Then Set xWin = ETWEETXLPOST.PostBox: Exit Function
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 41 Then Set xWin = ETWEETXLQUEUE.PostBox: Exit Function
'//Control Box: WinForms
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 10 Then Set xWin = CTRLBOX: Exit Function
'//Control Box: Input Fields
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 11 Then Set xWin = CTRLBOX.CtrlBoxWindow: Exit Function '5/24/2022
'//AutomateXL: WinForms
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 18 Then Set xWin = AUTOMATEXLHOME: Exit Function '6/7/2022
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 19 Then Set xWin = XLMAPPER: Exit Function '6/7/2022
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 20 Then Set xWin = XLMAPPERCTRLR: Exit Function '6/7/2022

End Function
Public Function fndRunTool(xTool) As Object

'/\__________________________________________________________________________
'//
'//     A function for finding the current running tool
'/\__________________________________________________________________________

Dim appEnv, appBlk As String
Call fndEnvironment(appEnv, appBlk)

'//Check for current running tool
'//
'//eTweetXL: xlFlowStrip
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 1 Then Set xTool = ETWEETXLHOME.xlFlowStrip
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 2 Then Set xTool = ETWEETXLSETUP.xlFlowStrip
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 3 Then Set xTool = ETWEETXLPOST.xlFlowStrip
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 4 Then Set xTool = ETWEETXLQUEUE.xlFlowStrip
'//Control Box: Input Fields
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 11 Then Set xTool = CTRLBOX.CtrlBoxWindow
'//AutomateXL: xlFlowStrip
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 18 Then Set xTool = XLMAPPER.xlFlowStrip

End Function
Private Function runScript$(xArt)

'/\__________________________________________________________________________
'//
'//     A function for running xlAppScript
'/\__________________________________________________________________________

    Dim SEARCHLIB As Integer
    
    '//Check first library location...
    On Error GoTo OutsideWb
    xLib = 1: errLvl = xArt: Call fndJunk(errLvl): If errLvl = 1 Then GoTo ErrRef Else Call fndLibs(xLib): Call rtnSpecial(xArt)
    CHECKLIB = Application.Run(xLib, (xArt)): If Range("xlasErrRef").Value2 <> "" Then xArt = Range("xlasErrRef").Value2
    
OutsideWb:
If Range("MAS18").Value2 = vbNullString Then Range("MAS18").name = "xlasErrRef"
    If InStr(1, xArt, "(*Err)") Then '//Check for not found error...
    SEARCHLIB = 2
    Do Until InStr(1, xArt, "(*Err)") = False '//Search for library if error found...
    If SEARCHLIB > 100 Then GoTo ErrRef
    xArt = Replace(xArt, "(*Err)", vbNullString)
    xLib = SEARCHLIB
    Call fndLibs(xLib)
    On Error GoTo ReSearchLib
    Range("xlasErrRef").Value2 = vbNullString '//reset error reference
    FOUNDLIB = Application.Run(xLib, (xArt)): xArt = Range("xlasErrRef").Value2
    If InStr(1, xArt, "(*Err)") = False Then Exit Function
ReSearchLib:
Err.Clear
    xArt = xArt & "(*Err)"
        SEARCHLIB = SEARCHLIB + 1
            Loop
                End If
                    Exit Function
                       
ErrRef:
xArt = 1
End Function
Public Function drv() As Variant

'/\__________________________________________________________________________
'//
'//     A function for finding the current drive (default is set as C:)
'/\__________________________________________________________________________

Set drv = CreateObject("Scripting.FileSystemObject")

'//check if C: drive exists
If drv.DriveExists("C") = True Then
Set drv = Nothing
drv = "C:": Exit Function
    Else
        '//search for next useable drive if C: unavailable
        For Each drv In drv.Drives
        If drv <> "" Then drv = drv.DriveLetter & ":": Exit Function
        Next
            End If

End Function
Public Function envHome$()

'/\__________________________________________________________________________
'//
'//     A function for finding the current running user
'/\__________________________________________________________________________

'//User
envHome = Environ("HOMEPATH")

End Function
Public Function modBlk(ByVal xArt As String) As Byte

'/\__________________________________________________________________________
'//
'//     A function for modifying the xlAppScript runtime block space
'/\__________________________________________________________________________

If Range("xlasGlobalControl").Value = 1 Then Call modStack(xArt): Exit Function '//modify runtime stack

'//Clear runtime block memory addresses
Range("MAA1:MAK5000").ClearContents
If Range("xlasLocalStatic").Value2 <> 1 Then Range("MAL1:MAL5000").ClearContents
Range("MAM1:MAR5000").ClearContents
If Range("xlasLocalContain").Value2 = 1 Then Range("MAS1:MAS5000").ClearContents
Range("MAT1:MAZ5000").ClearContents

End Function
Public Function modStack(ByVal xArt As String) As Byte
'/\__________________________________________________________________________
'//
'//     A function for modifying the xlAppScript runtime stack space
'/\__________________________________________________________________________

Dim lRow As Long

lRow = Cells(Rows.Count, "MAF").End(xlUp).Row

'//Set article to top of runtime stack if empty
If Range("MAF1").Value2 = vbNullString Then
Else
'//If block space isn't empty push stack down 1 block
Range("MAF1:MAF" & lRow).Copy Destination:=Range("MAF2:MAF" & lRow + 1)
End If

Range("MAF1").Value2 = xArt: Range("MAE1").Value2 = 0

End Function
Public Function modArtD$(xArt)

'/\__________________________________________________________________________
'//
'//     A function for replacing alternative expansion characters
'/\__________________________________________________________________________


xArt = Replace(xArt, "}", ")")
xArt = Replace(xArt, "{", "(")

End Function
Public Function modArtM$(xArt)

'/\__________________________________________________________________________
'//
'//     A function for replacing alternative expansion characters
'/\__________________________________________________________________________


xArt = Replace(xArt, "%", vbNullString)

End Function
Public Function modArtP$(xArt)

'/\__________________________________________________________________________
'//
'//     A function for modifying an article/line of xlAppScript (parentheses)
'/\__________________________________________________________________________


If Left(xArt, 1) = "(" Then xArt = Right(xArt, Len(xArt) - 1)
If Right(xArt, 1) = ")" Then xArt = Left(xArt, Len(xArt) - 1)
If Left(xArt, 1) = ")" Then xArt = Right(xArt, Len(xArt) - 1)
If Right(xArt, 1) = "(" Then xArt = Left(xArt, Len(xArt) - 1)

End Function
Public Function modArtQ$(xArt)

'/\__________________________________________________________________________
'//
'//     A function for modifying an article/line of xlAppScript (quotations)
'/\__________________________________________________________________________


If Left(xArt, 1) = """" Then xArt = Right(xArt, Len(xArt) - 1)
If Right(xArt, 1) = """" Then xArt = Left(xArt, Len(xArt) - 1)
If Left(xArt, 1) = "'" Then xArt = Right(xArt, Len(xArt) - 1)
If Right(xArt, 1) = "'" Then xArt = Left(xArt, Len(xArt) - 1)

End Function
Public Function modArtS$(xArt)

'/\__________________________________________________________________________
'//
'//     A function for modifying an article/line of xlAppScript (spaces)
'/\__________________________________________________________________________

xArt = Replace(xArt, " ", vbNullString)

End Function

'//=========================================================================================================================
'//
'//         CHANGE LOG
'/\_________________________________________________________________________________________________________________________
'
'
' Version 1.1.2
'
' [ Date: 6/7/2022 ]
'
' (1): Included WinForm $'s for "AutomateXL"
'
'
' [ Date: 6/6/2022 ]
'
' (1): Removed black & red text changes to current Run Tool from the main "lex" module
'
' (2): Minor label name changes
'
' (3): Changed special character identifiers for ($), (;), ((), ()), & (#):
'
' $ --> [doll] | ; --> [semi] | ( --> [lp] | ) --> [rp] | # --> [hash]
'
'
' Version 1.1.1
'
' [ Date: 5/24/2022 ]
'
' (1): Made minor adjustments to WinForm #'s in "fndWindow" & "fndRunTool" functions
'
'
'
' Version 1.1.0
'
' [ Date: 5/11/2022 ]
'
' (1): Code optimization, function renaming, etc.
'
'
' Version 1.0.9
'
'
' [ Date: 4/23/2022 ]
'
' (1): Removed "enableWbUpdates" & "disableWbUpdates" functions from running automatically, and moved them to the "xbas" lib
' as a library dependent switch.
'
'
' (1): Adjusted "kinSet" & "kinExpand" functions to accept alternative cases w/ variables such as parsing a variable call identifier
' (@) or (=) through a file or when needing to set a variable value from another variables value.
'
' Variable Preservation Sequence = %@var
'
' Example: %@var=read(-all @filePath); <--- if the text in this file contains an "=" this will still return all text to the assigned variable
'
'
' Version 1.0.8
'
'
' [ Date: 4/20/2022 ]
'
' (1): Added function for removing the "%" alternative variable escape character.
'
' When "%" is used in front of a call to a variable like "%@var" if the variable contains "@" within a string
' it will parse through the "@" identifier as if it's a normal string instead of interpreting the string containing @
' as a variable.
'
'
' [ Date: 4/18/2022 ]
'
'
' (1): Added "xlasGlobalControl" block for triggering a control that allows Local variable usage in external environments
'
'
'
'
' [ Date: 4/17/2022 ]
'
'
' (1): "libcall" bug fixes when using after another instruction
'
' (2): Changed lexer counter variable names to contain "Ptr" (as in "Pointer")
'
' (3): Added "[(]" & "[)]" special character escape sequences for left and right parentheses
'
'
' [ Date: 4/16/2022 ]
'
'
' (1): Added "[space]" special escape character
'
' (2): Reworked "<env>" & "<blk>" commands to accept spaces
'
'
' Version 1.0.7
'
'
' [ Date: 4/2/2022 ]
'
' (1): Fixed recursive loop caused by "end" instruction when used w/ certain variations of "if-else" instruction
'
' (2): Fixed "if-else" instruction not parsing parentheses in certain instances
'
'
' [ Date: 3/31/2022 ]
'
' (1): Made adjustments to escape character function & added return function for retrieving values before searching through
' a library.
' ***New sequence is [$] <--- where the special character is enclosed in brackets
'
'
'
' Version 1.0.6
'
'
' [ Date: 3/17/2022 ]
'
' (1): Added "xlasLocalContain"
'
'
' [ Date 3/16/2022 ]
'
'
' (1): Added Runtime block addresses "xlasBlkAddr79" - "xlasBlkAddr99" (21 total)
'
'
' [ Date: 3/13/2022 ]
'
'
' (1): Added "xlasLocalStatic" address for locking libraries to the current runtime session.
'
' Version 1.0.5
'
' [ Date: 3/9/2022 ]
'
' (1): Added "let" instruction for assigning variable values on the fly after declaration.
'
'
'
' [ Date: 3/8/2022 ]
'
' (1): Added "libcall" instruction to allow directing an article string/block to a desired library. This will place the
' desired library at the top of the library stack, which will help cut down on article querying time & avoid collisions
' where articles from different libraries share similar names.
' (in instances where many libraries are used, this will allow those lines/blocks of code to parse quicker).
'
'
' [ Date: 3/7/2022 ]
'
' (1): Added an escape key function & sequence for retaining special character(s) during parsing
'
' Example: --->  [^$] (for escaping the '$' symbol)
'
'
'
'
' Version 1.0.4
'
'
'
' [ Date 3/2/2022 ]
'
' (1): Added "Alternative Expansion" functionality & characters "{", "}", & "~" which allow articles to ignore initial parsing.
'
' Replace enclosers/enders (() & ()) w/ ({) & (}), & (;) w/ (~)
'
' w/ Alernative Expansion = rng{A1}.value{100}.bgcolor{gainsboro}.fcolor{cornflowerblue}~$
' w/o Alternative Expansion = rng(A1).value(100).bgcolor(gainsboro).fcolor(cornflowerblue);$
'
' (2): Added check for "<blk>" article.
'
' The "<blk>" article's purpose is for manually setting the runtime block (worksheet)
'
'
' (3): Seperated instructional components from lexer (goto, if, do, etc.) into respective functions.
'
'
' (4): Fixed issue's w/ placing an if-statement within a do-loop & vice versa
'
'
' [ Date 3/1/2022 ]
'
'
' (1): Added "-and", "-or", "-nor", & "-xor" to accepted if-statement operators
'
'
' [ Date 2/28/2022 ]
'
'
' (1): Adjusted the "fndEnvironment" function to now find the applications "runtime block" (Worksheet hosting xlas memory/states).
' ***Changed it in a way to allow editing the name of the worksheet w/o it affecting the current runtime block
'
' (2): Created 3 functions to deal w/ cleaning up an article during parsing (modArtP, modArtQ, modArtS)
' ***Removes parentheses, quotations, & spaces
'
' (3): Seperated loops & conditionals from "lexKey" function & instead put them into their own respective functions
'
' (4): Adjusted do-loops to support nesting if statements
'
'
' [ Date 2/27/2022 ]
'
' (1): Revised "if" statement lexing to check for boolean & seperate initially before setting variable values due to retention
' of the first article following the conditional check during variable parsing.
'
' ***Revised "-eq" boolean operator to except strings
'
' (2): Moved "findbasSaveFormat" function from "lex" module to "xbas" library module (now titled "basSaveFormat")
'
'
' Version 1.0.3
'
' [ Date 2/26/2022 ]
'
' (1): Changed environment variable prefix back to "env" from "e"
'
'
' [ Date 2/24/2022 ]
'
' (1): Added "xlasUpdateEnable" range/memory location for allowing control of the "enableWbUpdates" & "disableWbUpdates" functions
' (Coincides w/ new "--e" & "++e" switches for "xbas v1.0.3" library)
'
'
' [ Date 2/7/2022 ]
'
' (1): Added additional window assignment compatibility for eTweetXL post boxes.
'
' form(31) = Tweet Setup Post Box | form(41) = Tweet Queue Post Box
'
'
'
' [ Date 2/4/2022 ]
'
' (1): Edited "kinSet" function to double check if variable was found and recheck if not before expanding.
'
'  ***Would cause issues when trying to increment a variable such as @x++
'
'
' Version 1.0.2
'
'
' [ Date: 1/31/2022 ]
'
' (1): Added "xlasLink" to "connectWb" function to help confirm the application's connected when opening/starting
'
' (2): Shortened environment variable prefixes from "env" to "e"
'
' (3): Moved connect & disconnect functions to seperate setup module
'
'
'
' Version 1.0.1
'
'
' [ Date: 1/6/2022 ]
'
' (1): Adjusted lexKey to only remove the leading & ending "()" from a declared "kin" variable. Initially would remove all parentheses
' from the article, but this would cause problems in a scenerio where we still needed the parentheses to identify a command.
'
'
'
'
'
' [ Date: 1/5/2022 ]
'
' (1): Created kinExpand function to help w/ expanding (setting the value of a variable) before, & @ runtime.
' ***Working to solve an issue where if a variable is declared w/ a value & needed again for an article such as repl() or input(),
' the variable will still retain the value it was declared as instead of expanding (will work fine when set to nothing, or
' when incrementing the var, or changing it's value to another).
'
'
'
' [ Date: 1/4/2022 ]
'
' (1): Changed name of memory locations to prefix "xlas" as to create lesser ambiquity when connecting to alternative workbooks
'
' (2): Changed "setupWb" function to "connectWb" & created a paired "disconnectWb" to help w/ adding & removing a runtime environment
'
'
' [ Date: 1/3/2022 ]
'
' (1): Added enableWbUpdates function to pair w/ disableWbUpdates to help speed up runtime
' (previously disableWbUpdates left sheet calculations on)
'
' (2): Removed call to syntax error message
'
' (3): Removed enable & disable FlowStrip functions from lex module
'
'
'
' [ Date: 1/2/2022 ]
'
' (1): Added change log, & license information & version #
'
' (2): Created "fndEnvironment" function to find the runtime environment while running xlAppScript
'
' (3): Created "fndJunk" function to remove unwanted data before getting parsed further
'
' (4): Added labeling for most functions & their purpose
'
' (5): Moved script memory locations farther over in the workbook (from MA1-MAA1, MB1-MAB1, etc.)
'
'



