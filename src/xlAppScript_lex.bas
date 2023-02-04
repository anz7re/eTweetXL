Attribute VB_Name = "xlAppScript_lex"

Public Function xlas(Art) As Byte
'/\______________________________________________________________________________________________________________________
'//
'//     xlAppScript Lexer
'//        Version: 1.2.0
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
'//     Latest Revision: 2/1/2023
'/\_____________________________________________________________________________________________________________________________
'//
'//     Developer(s): anz7re (André)
'//     Contact: support@xlappscript.org | support@autokit.tech | anz7re@autokit.tech
'//     Web: xlappscript.org | autokit.tech
'/\_____________________________________________________________________________________________________________________________
'//
'//
'/\__________________________________________________________________________________________
'//
'//     A function for screening script & environment prior to lexing
'/\__________________________________________________________________________________________
'//
'//
''//=============================================================
''//Set runtime environment...
Call getEnvironment(appEnv, appBlk):
'//=============================================================
'//Set runtime tool...
Dim Tool As Object: Call getTool(Tool, appEnv, appBlk)
'//=============================================================
'//Set article from runtime tool if found...
If Not Tool Is Nothing Then If Art = vbNullString Then Art = Tool.Value
'//=============================================================
'//Check for run initializer...
If InStr(1, Art, "$") Then
'//=============================================================
'//Escape special chracter sequences...
Call setSpecial(Art)
'//=============================================================
'//Remove line formatting characters...
Art = Replace(Art, vbNewLine, vbNullString)
Art = Replace(Art, vbTab, vbNullString)
'//=============================================================
'//Modify runtime block...
Call modBlk(Art, appEnv, appBlk)
'//=============================================================
Call xlasLex(Art, appEnv, appBlk)
'//=============================================================
End If
    
End Function
Public Function xlasLex(Token, appEnv, appBlk) As Collection

'/\__________________________________________________________________________
'//
'//     A function for lexing (tokenizing) xlAppScript code
'/\__________________________________________________________________________
'//

    Dim Tokens As New Collection
    Dim currentChar As String, currentToken As String
    Dim IS_NUMBER As Boolean, IS_STRING As Boolean
    Dim IS_COMMENT As Boolean, IS_DELIMITER As Boolean
    Dim IS_VAR As Boolean
    Dim IS_CALL As Boolean
    Dim IS_DO As Boolean
    Dim IS_FOR As Boolean
    Dim IS_GROUP As Boolean
    Dim IS_IF As Boolean
    Dim IS_LIBCALL As Boolean
    Dim IS_WHILE As Boolean
    
    For I = 1 To Len(Token)
        currentChar = Mid(Token, I, 1)
        
        '//=============================================================
        '//Check if current character is an xlas article delimiter (;)
        '//=============================================================
        If currentChar = ";" And Not IS_STRING Then
            IS_DELIMITER = True
        If Not IS_COMMENT Then Tokens.Add currentToken: IS_DELIMITER = False
            currentChar = ""
            currentToken = ""
        '//=============================================================
        '//Check if current character is a string delimiter ('') or ("")
        '//=============================================================
        ElseIf currentChar = "'" And Not IS_STRING Then
            IS_STRING = True
        ElseIf currentChar = "'" And IS_STRING Then
            IS_STRING = False
        ElseIf currentChar = """" And Not IS_STRING Then
            IS_STRING = True
        ElseIf currentChar = """" And IS_STRING Then
            IS_STRING = False
        End If
        '//=============================================================
        '//Check if current character is a comment
        '//=============================================================
        If currentChar = "#" And Not IS_STRING Then
            IS_COMMENT = True
        ElseIf IS_COMMENT And IS_DELIMITER Then
            IS_COMMENT = False
            IS_DELIMITER = False
        End If
        '//=============================================================
        '//Check if current character is a variable declaration
        '//=============================================================
        If InStr(1, currentToken, "KIN", vbTextCompare) And Not IS_STRING Then
        If InStr(1, currentToken, "KIN ", vbTextCompare) And Not IS_VAR Then
            IS_VAR = True
            Tokens.Add "KIN"
            currentToken = ""
        ElseIf InStr(1, currentToken, "KIN(", vbTextCompare) And Not IS_VAR Then
            IS_VAR = True
            Tokens.Add "KIN"
            currentToken = ""
        Else
            IS_VAR = False
        End If
        End If
        '//=============================================================
        '//Check if current token is an instruction or redirect
        '//=============================================================
        '//
        '//=============================================================
        '//call
        '//=============================================================
        If InStr(1, currentToken, "CALL(", vbTextCompare) And Not IS_STRING Then
            Tokens.Add "CALL"
            currentToken = ""
            IS_CALL = True
        ElseIf InStr(1, currentToken, ")", vbTextCompare) And IS_CALL Then
            Tokens.Add "END_CALL"
            currentToken = ""
            IS_CALL = False
        '//=============================================================
        '//do-loop
        '//=============================================================
        ElseIf InStr(1, currentToken, "DO{", vbTextCompare) And Not IS_STRING Then
            Tokens.Add "DO_LOOP"
            currentToken = ""
            IS_DO = True
        ElseIf InStr(1, currentToken, "}LOOP(", vbTextCompare) And IS_DO Then
            Tokens.Add "END_DO_LOOP"
            currentToken = ""
            IS_DO = False
        '//=============================================================
        '//goto
        '//=============================================================
        ElseIf InStr(1, currentToken, "GOTO ", vbTextCompare) Then
            Tokens.Add "GOTO"
            currentToken = ""
        '//=============================================================
        '//if-else
        '//=============================================================
        ElseIf InStr(1, currentToken, "IF(", vbTextCompare) And Not IS_STRING Then
            Tokens.Add "IF"
            currentToken = ""
            IS_IF = True
        ElseIf InStr(1, currentToken, "}ELSE{", vbTextCompare) And Not IS_STRING Then
            Tokens.Add "ELSE"
            currentToken = ""
            IS_IF = True
        ElseIf InStr(1, currentToken, "}ENDIF", vbTextCompare) And IS_IF Then
            Tokens.Add "END_IF"
            currentToken = ""
            IS_IF = False
        ElseIf InStr(1, currentToken, "){", vbTextCompare) And IS_IF Then
            Tokens.Add currentToken
            currentToken = ""
        '//=============================================================
        '//for-next
        '//=============================================================
        ElseIf InStr(1, currentToken, "FOR(", vbTextCompare) And Not IS_STRING Then
            Tokens.Add "FOR_NEXT"
            currentToken = ""
            IS_FOR = True
        ElseIf InStr(1, currentToken, "}NEXT", vbTextCompare) And IS_FOR Then
            Tokens.Add "END_NEXT"
            currentToken = ""
            IS_FOR = False
        ElseIf InStr(1, currentToken, "){", vbTextCompare) And IS_FOR Then
            Tokens.Add currentToken
            currentToken = ""
        '//=============================================================
        '//libcall
        '//=============================================================
        ElseIf InStr(1, currentToken, "LIBCALL", vbTextCompare) And Not IS_STRING Then
            Tokens.Add "LIBCALL"
            currentToken = ""
            IS_LIBCALL = True
        ElseIf InStr(1, currentToken, "}ENDLIB", vbTextCompare) And IS_LIBCALL Then
            Tokens.Add "END_LIBCALL"
            currentToken = ""
            IS_LIBCALL = False
        '//=============================================================
        '//while
        '//=============================================================
        ElseIf InStr(1, currentToken, "WHILE(", vbTextCompare) And Not IS_STRING Then
            Tokens.Add "WHILE"
            currentToken = ""
            IS_WHILE = True
        ElseIf InStr(1, currentToken, "}CONTINUE", vbTextCompare) And IS_WHILE Then
            Tokens.Add "END_WHILE"
            currentToken = ""
            IS_WHILE = False
        ElseIf InStr(1, currentToken, "){", vbTextCompare) And IS_WHILE Then
            Tokens.Add currentToken
            currentToken = ""
        End If
        If currentChar = ")" And IS_LIBCALL And Not IS_STRING Then
            Tokens.Add currentToken
            currentChar = ""
            currentToken = ""
        End If
        '//=============================================================
        '//Check if current token is a statement
        '//=============================================================
        '//
       '//=============================================================
        '//clear
        '//=============================================================
        If InStr(1, currentToken, "CLEAR ", vbTextCompare) And Not IS_STRING Then
            Tokens.Add "CLEAR"
            currentToken = ""
        '//=============================================================
        '//get
        '//=============================================================
        ElseIf InStr(1, currentToken, "GET ", vbTextCompare) And Not IS_STRING Then
            Tokens.Add "GET"
            currentToken = ""
        '//=============================================================
        '//set
        '//=============================================================
        ElseIf InStr(1, currentToken, "SET ", vbTextCompare) And Not IS_STRING Then
            Tokens.Add "SET"
            currentToken = ""
        End If
        '//=============================================================
        '//Check if current token is a structure
        '//=============================================================
        '//
        '//=============================================================
        '//group
        '//=============================================================
        If InStr(1, currentToken, "GROUP(", vbTextCompare) And Not IS_STRING Then
            Tokens.Add "GROUP"
            currentToken = ""
            IS_GROUP = True
        ElseIf InStr(1, currentToken, "}ENDGROUP", vbTextCompare) And IS_GROUP Then
            Tokens.Add "END_GROUP"
            currentToken = ""
            IS_GROUP = False
        ElseIf InStr(1, currentToken, "){", vbTextCompare) And IS_GROUP Then
            Tokens.Add currentToken
            currentToken = ""
        End If
        '//=============================================================
        '//Append current character to the current token
        '//=============================================================
        currentToken = currentToken & currentChar
    Next
    '//=============================================================
    '//Add last token to the collection
    '//=============================================================
    If Len(currentToken) > 0 Then
        Tokens.Add currentToken
    End If
    
    Call xlasParse(Tokens, appEnv, appBlk)
    
End Function
Public Function xlasParse(Tokens As Collection, appEnv, appBlk) As Collection

'/\__________________________________________________________________________
'//
'//     A function for parsing xlAppScript tokens
'/\__________________________________________________________________________
'//

    Dim xlasPtr As Long
    
    For xlasPtr = 1 To Tokens.Count
    Token = Tokens.Item(xlasPtr)
    '//=============================================================
    '//Junk check
    '//=============================================================
    Call getJunk(Token)
    If Token <> vbNullString Then
    '//=============================================================
    '//Set runtime application environment (Workbook)
    '//=============================================================
    If InStr(1, Token, "<ENV>", vbTextCompare) Then
    Call setEnv(Token, appEnv, appBlk)
    '//=============================================================
    '//Set runtime block space (Worksheet)
    '//=============================================================
    ElseIf InStr(1, Token, "<BLK>", vbTextCompare) Then
    Call setBlk(Token, appEnv, appBlk)
    '//=============================================================
    '//Set runtime libraries
    '//=============================================================
    ElseIf InStr(1, Token, "<LIB>", vbTextCompare) Then
    Call setLib(Token, appEnv, appBlk)
    '//=============================================================
    '//Set variable
    '//=============================================================
    ElseIf Token = "KIN" Then
    Call setKin(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse call instruction
    '//=============================================================
    ElseIf Token = "CALL" Then
    Call setCall(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse do-loop instruction
    '//=============================================================
    ElseIf Token = "DO_LOOP" Then
    Call setDoLoop(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse for-next instruction
    '//=============================================================
    ElseIf Token = "FOR_NEXT" Then
    Call setForNext(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse goto redirect
    '//=============================================================
    ElseIf Token = "GOTO" Then
    Call setGoto(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse group structure
    '//=============================================================
    ElseIf Token = "GROUP" Then
    Call setGroup(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse if-then instruction
    '//=============================================================
    ElseIf Token = "IF" Then
    Call setIfElse(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse libcall instruction
    '//=============================================================
    ElseIf Token = "LIBCALL" Then
    Call setLibCall(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse while instruction
    '//=============================================================
    ElseIf Token = "WHILE" Then
    Call setWhile(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse clear statement
    '//=============================================================
    ElseIf Token = "CLEAR" Then
    Call setClear(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse get statement
    '//=============================================================
    ElseIf Token = "GET" Then
    Call setGet(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse set statement
    '//=============================================================
    ElseIf Token = "SET" Then
    Call setKinSet(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Execute xlAppScript
    '//=============================================================
    Else
    Call xlasExecute(Token, appEnv, appBlk)
    End If
        End If
            Next

End Function
Public Function xlasFocusInstruction(Tokens As Collection, Token, xlasPtr, appEnv, appBlk) As Collection

'/\__________________________________________________________________________
'//
'//     A function for sorting instructions/redirects during parsing
'/\__________________________________________________________________________
'//


    '//=============================================================
    '//Parse call instruction
    '//=============================================================
    If Token = "CALL" Then
    Call setCall(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse do-loop instruction
    '//=============================================================
    ElseIf Token = "DO_LOOP" Then
    Call setDoLoop(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse for-next instruction
    '//=============================================================
    ElseIf Token = "FOR_NEXT" Then
    Call setForNext(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse goto redirect
    '//=============================================================
    ElseIf Token = "GOTO" Then
    Call setGoto(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse group structure
    '//=============================================================
    ElseIf Token = "GROUP" Then
    Call setGroup(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse if-then instruction
    '//=============================================================
    ElseIf Token = "IF" Then
    Call setIfElse(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse libcall instruction
    '//=============================================================
    ElseIf Token = "LIBCALL" Then
    Call setLibCall(Tokens, Token, xlasPtr, appEnv, appBlk)
    '//=============================================================
    '//Parse while instruction
    '//=============================================================
    ElseIf Token = "WHILE" Then
    Call setWhile(Tokens, Token, xlasPtr, appEnv, appBlk)
    End If
    

End Function
Private Function xlasExecute$(Token, appEnv, appBlk)

'/\__________________________________________________________________________
'//
'//     A function for executing xlAppScript
'/\__________________________________________________________________________

    Dim LIB_INDEX As Integer, LIB_STACK_COUNT As Integer
    Dim Lib As String
    
    '//Check for end execution
    If Token = "END" Or Token = "End" Or Token = "end" Then End
    
    '//Retrieve variable
    If InStr(1, Token, "@") Then Call getKin(Token, KIN_LABEL, KIN_VALUE, appEnv, appBlk)
    '//Retrieve environment variable
    If InStr(1, Token, "@ENV", vbTextCompare) Then Call getEnvironmentVars(Token)
    '//Retrieve libary stack count
    LIB_STACK_COUNT = Workbooks(appEnv).Worksheets(appBlk).Range("xlasLibCount").Value2
    
    '//Find a library in stack
    On Error GoTo GetNextLib
    Call getJunk(Token): If Token = vbNullString Then GoTo ErrEnd Else Lib = 1: Call getLibs(Lib, appEnv, appBlk): Call getSpecial(Token)
    CHECKLIB = Application.Run(Lib, (Token))
    If Workbooks(appEnv).Worksheets(appBlk).Range("xlasErrRef").Value2 <> "" Then Token = Workbooks(appEnv).Worksheets(appBlk).Range("xlasErrRef").Value2

GetNextLib:
If Workbooks(appEnv).Worksheets(appBlk).Range("MAS25").Value2 = vbNullString Then Workbooks(appEnv).Worksheets(appBlk).Range("MAS25").name = "xlasErrRef"
    '//Library not found
    If InStr(1, Token, "*/ERR") Then
    LIB_INDEX = 2
    '//Search for next library in stack
    Do Until InStr(1, Token, "*/ERR") = False
    If LIB_INDEX > LIB_STACK_COUNT Then GoTo ErrEnd
    Token = Replace(Token, "*/ERR", vbNullString)
    Lib = LIB_INDEX
    Call getLibs(Lib, appEnv, appBlk)
    On Error GoTo NextLib
    Workbooks(appEnv).Worksheets(appBlk).Range("xlasErrRef").Value2 = vbNullString
    '//Library found. Run script.
    FOUNDLIB = Application.Run(Lib, (Token)): Token = Workbooks(appEnv).Worksheets(appBlk).Range("xlasErrRef").Value2
    If InStr(1, Token, "*/ERR") = False Then Exit Function
NextLib:
Err.Clear
    Token = Token & "*/ERR"
        LIB_INDEX = LIB_INDEX + 1
            Loop
                End If
                Token = vbNullString
                Workbooks(appEnv).Worksheets(appBlk).Range("xlasErrRef").Value2 = vbNullString
                    Exit Function

ErrEnd:
End Function
Public Function xlasExpand$(Token, appEnv, appBlk)

'/\__________________________________________________________________________
'//
'//     A function for expanding the value of a variable at runtime
'/\__________________________________________________________________________

        
        Dim ARR_INDEX As Long, Row As Long, LastRow As Long, xPos As Long, X As Long
        Dim ENV As Byte
        Dim TOKEN_COPY As String

        TOKEN_COPY = Token
        
        '//extract article...
        TempTokenArr = Split(TOKEN_COPY, "[,]")
        appEnv = TempTokenArr(0) '//application environment
        TOKEN_COPY = TempTokenArr(1) '//article to parse
        X = TempTokenArr(2) '//position
        ENV = TempTokenArr(3) '//xlas environment
        
        If InStr(1, TOKEN_COPY, "=") Then TokenArr = Split(TOKEN_COPY, "="): TokenArr(0) = Replace(TokenArr(0), " ", vbNullString)
        
        '//expanding from library...
        If ENV = 1 Then
        
        X = Cells(Rows.Count, "MAA").End(xlUp).Row

        '//Alternative runtime expansion
        '//
        If InStr(1, TOKEN_COPY, "@") Then
        For Row = 1 To X
        If InStr(1, TokenArr(1), Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(Row, 0).Value2, vbTextCompare) Then xPos = Row
        If InStr(1, TokenArr(0), Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(Row, 0).Value2, vbTextCompare) Then
    
            If InStr(1, TOKEN_COPY, "=") = False Then
            
            TOKEN_COPY = Replace(TOKEN_COPY, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(Row, 0).Value2, _
            Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(Row, 0).Value2, vbTextCompare)
            End If
                
            If InStr(1, TokenArr(1), "@") = False Or InStr(1, TokenArr(0), "%@") Then
            Token = TokenArr(0): Call modArtPercent(Token): TokenArr(0) = Token
            If UBound(TokenArr) > 1 Then For ARR_INDEX = 2 To UBound(TokenArr): TokenArr(1) = TokenArr(1) & "=" & TokenArr(ARR_INDEX): Next
                Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(Row, 0).Value2 = TokenArr(1)
                Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValue").Offset(Row, 0).Value2 = Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(Row, 0).Value2
                    Else
                    Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(Row, 0).Value2 = _
                    Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(xPos, 0).Value2
                        End If
                            End If
                                Next
                                    End If
                                        End If
        
        '//return token value
        Token = TOKEN_COPY
        
End Function
Public Function setEnv(Token, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for setting the application runtime environment
'/\____________________________________________________________________________________

Token = Replace(Token, "<env>", vbNullString, , , vbTextCompare)
Token = Trim(Token)

Workbooks(appEnv).Worksheets(appBlk).Range("xlasEnvironment").Value2 = Token

End Function
Public Function setBlk(Token, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for setting the application runtime block
'/\____________________________________________________________________________________

Token = Replace(Token, "<blk>", vbNullString, , , vbTextCompare)
Token = Trim(Token)

Workbooks(appEnv).Worksheets(appBlk).Range("xlasBlock").Value2 = Token

End Function
Public Function setLib(Token, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for setting runtime session libraries
'/\____________________________________________________________________________________

Dim LastRow As Long

Token = Replace(Token, "<lib>", vbNullString, , , vbTextCompare)
Token = Trim(Token)

LastRow = Cells(Rows.Count, "MAL").End(xlUp).Row
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLib").Offset(LastRow, 0).Value = Token
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLibCount").Value2 = LastRow

End Function
Public Function setKin(Tokens As Collection, Token, xlasPtr, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for identifying a variable & setting it's value before runtime
'/\____________________________________________________________________________________

Dim LastRow As Long, xlasPtrCopy As Long
Dim KIN_LABEL As String, KIN_VALUE As String
Dim TokenArr() As String
xlasPtr = xlasPtr + 1
xlasPtrCopy = xlasPtr

Token = Tokens.Item(xlasPtr)
Call modArtParens(Token)

LastRow = Workbooks(appEnv).Worksheets(appBlk).Cells(Rows.Count, "MAA").End(xlUp).Row

TokenArr = Split(Token, "=")
KIN_LABEL = Trim(TokenArr(0))
KIN_VALUE = Trim(TokenArr(1))

Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabel").Offset(LastRow, 0).Value2 = "@" & KIN_LABEL
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabelMod").Offset(LastRow, 0).Value2 = "@" & KIN_LABEL
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValue").Offset(LastRow, 0).Value2 = KIN_VALUE
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(LastRow, 0).Value2 = KIN_VALUE

End Function
Public Function setKinSet(Tokens As Collection, Token, xlasPtr, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for parsing a set statement
'/\____________________________________________________________________________________

Dim LastRow As Long, Row As Long, xlasPtrCopy As Long
Dim KIN_LABEL As String, KIN_VALUE As String
Dim TokenArr() As String
xlasPtr = xlasPtr + 1
xlasPtrCopy = xlasPtr

Token = Tokens.Item(xlasPtr)

LastRow = Workbooks(appEnv).Worksheets(appBlk).Cells(Rows.Count, "MAA").End(xlUp).Row

Call modArtParens(Token)
TokenArr = Split(Token, "=")
KIN_LABEL = Trim(TokenArr(0))
KIN_VALUE = Trim(TokenArr(1))

For Row = 1 To LastRow - 1
If InStr(1, KIN_LABEL, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabelMod").Offset(Row, 0).Value2, vbBinaryCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(Row, 0).Value2 = KIN_VALUE
End If
Next

End Function
Public Function setCall(Tokens As Collection, Token, xlasPtr, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for parsing a call instruction
'/\____________________________________________________________________________________

Dim END_GROUP_PTR As Long, xlasPtrCopy As Long
Dim CALL_LABEL As String, CALL_INDEX As String
Dim TokenArr() As String
xlasPtr = xlasPtr + 1
xlasPtrCopy = xlasPtr

CALL_LABEL = Tokens.Item(xlasPtr)

Call modArtParens(CALL_LABEL)

xlasPtr = 1

'//Find call group label
Do Until xlasPtr = Tokens.Count Or InStr(1, Token, CALL_LABEL, vbTextCompare)
Token = Tokens.Item(xlasPtr)
xlasPtr = xlasPtr + 1
Loop
CALL_INDEX = xlasPtr - 1
'//Find call end pointer
Do Until xlasPtr = Tokens.Count Or InStr(1, Token, "ENDGROUP", vbTextCompare)
Token = Tokens.Item(xlasPtr)
xlasPtr = xlasPtr + 1
Loop
END_GROUP_PTR = xlasPtr

xlasPtr = CALL_INDEX + 1

Token = Tokens.Item(xlasPtr)

Do Until xlasPtr = Tokens.Count

'//Check for instruction
Call xlasFocusInstruction(Tokens, Token, xlasPtr, appEnv, appBlk)

'//Check for variable
If InStr(1, Token, "@") Then Call getKin(Token, KIN_LABEL, KIN_VALUE, appEnv, appBlk)

If InStr(1, Token, "}ENDGROUP", vbTextCompare) Then xlasPtr = xlasPtrCopy: Exit Function

'//Execute article
If Token <> vbNullString Then Call xlasExecute(Token, appEnv, appBlk)

xlasPtr = xlasPtr + 1

Token = Tokens.Item(xlasPtr)

Loop

End Function

Public Function setDoLoop(Tokens As Collection, Token, xlasPtr, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for parsing a do-loop instruction
'/\____________________________________________________________________________________
'//
'//
'//=======================================================
Dim LOOP_COUNT As Long, END_DO_LOOP_PTR As Long, xlasPtrCopy As Long
Dim LOOPED As String
LOOPED = "1"
xlasPtr = xlasPtr + 1
xlasPtrCopy = xlasPtr

Do Until xlasPtr = Tokens.Count

Token = Tokens.Item(xlasPtr)

'//Check for instruction
Call xlasFocusInstruction(Tokens, Token, xlasPtr, appEnv, appBlk)

'//Check for variable
If InStr(1, Token, "@") Then Call getKin(Token, KIN_LABEL, KIN_VALUE, appEnv, appBlk)

If InStr(1, Token, "END_DO_LOOP") Then
END_DO_LOOP_PTR = xlasPtr
LOOP_COUNT = LOOP_COUNT + 1
xlasPtr = xlasPtr + 1
LOOPED = Tokens.Item(xlasPtr): Call modArtParens(LOOPED): LOOPED = LOOPED - LOOP_COUNT: xlasPtr = xlasPtrCopy - 1: Token = vbNullString
End If

If LOOPED = "0" Then xlasPtr = END_DO_LOOP_PTR + 1: Exit Function

'//Execute article
If Token <> vbNullString Then Call xlasExecute(Token, appEnv, appBlk)

KIN_LABEL = vbNullString
KIN_VALUE = vbNullString

xlasPtr = xlasPtr + 1

Loop


End Function
Public Function setClear(Tokens As Collection, Token, xlasPtr, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for parsing a clear statement
'/\____________________________________________________________________________________

Dim LastRow As Long, Row As Long, xlasPtrCopy As Long
Dim CLEAR_LABEL As String, CLEAR_INDEX As String
Dim TokenArr() As String
xlasPtr = xlasPtr + 1
xlasPtrCopy = xlasPtr

Token = Tokens.Item(xlasPtr)

Call modArtParens(Token)
TokenArr = Split(Token, " ")
CLEAR_LABEL = Trim(TokenArr(0))
CLEAR_INDEX = Trim(TokenArr(1))

'//Clear single index
If InStr(1, CLEAR_INDEX, ":") = False Then
Workbooks(appEnv).Worksheets(appBlk).Range(CLEAR_LABEL).Offset(CLEAR_INDEX, 0).Clear
    Else
'//Clear multi index
        TokenArr = Split(CLEAR_INDEX, ":")
        
        Row = TokenArr(0)
        LastRow = TokenArr(1)
        
        Do Until Row >= LastRow
        Workbooks(appEnv).Worksheets(appBlk).Range(CLEAR_LABEL).Offset(Row, 0).Clear
        Row = Row + 1
        Loop
        
        End If

End Function
Public Function setGet(Tokens As Collection, Token, xlasPtr, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for parsing a get statement
'/\____________________________________________________________________________________

Dim LastRow As Long, Row As Long, xlasPtrCopy As Long
Dim GET_LABEL As String, GET_INDEX As String, KIN_LABEL As String, KIN_VALUE As String
Dim TokenArr() As String
xlasPtr = xlasPtr + 1
xlasPtrCopy = xlasPtr

Token = Tokens.Item(xlasPtr)

Call modArtParens(Token)
TokenArr = Split(Token, " ")
KIN_LABEL = Trim(TokenArr(0))
GET_LABEL = Trim(TokenArr(1))
GET_INDEX = Trim(TokenArr(2))

'//Clear single index
If InStr(1, GET_INDEX, ":") = False Then
KIN_VALUE = Workbooks(appEnv).Worksheets(appBlk).Range(GET_LABEL).Offset(GET_INDEX, 0).Value2
    Else
'//Clear multi index
        TokenArr = Split(GET_INDEX, ":")
        
        Row = TokenArr(0)
        LastRow = TokenArr(1)
        
        Do Until Row >= LastRow
        KIN_VALUE = KIN_VALUE & "[,]" & Workbooks(appEnv).Worksheets(appBlk).Range(GET_LABEL).Offset(Row, 0).Value2
        Row = Row + 1
        Loop
        
        KIN_VALUE = Right(KIN_VALUE, Len(KIN_VALUE) - 3)
        
        End If

Call modKin(KIN_LABEL, KIN_VALUE, appEnv, appBlk)

End Function
Public Function setGoto(Tokens As Collection, Token, xlasPtr, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for parsing a goto instruction
'/\____________________________________________________________________________________

Dim GOTO_LABEL As String
xlasPtr = xlasPtr + 1

GOTO_LABEL = Tokens.Item(xlasPtr)
xlasPtr = 1

Do Until xlasPtr = Tokens.Count

Token = Tokens.Item(xlasPtr)
If Token = ":" & GOTO_LABEL Then Exit Function
xlasPtr = xlasPtr + 1

Loop

End Function
Public Function setGroup(Tokens As Collection, Token, xlasPtr, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for parsing a group structure
'/\____________________________________________________________________________________

Do Until xlasPtr = Tokens.Count Or InStr(1, Token, "}ENDGROUP", vbTextCompare)
Token = Tokens.Item(xlasPtr)
xlasPtr = xlasPtr + 1
Loop

xlasPtr = xlasPtr - 1

End Function
Public Function setForNext(Tokens As Collection, Token, xlasPtr, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for parsing a for-next instruction
'/\____________________________________________________________________________________
'//
'//
'//=======================================================
Dim LOOP_COUNT As Long, START_FOR_NEXT As Long, END_FOR_NEXT As Long, END_FOR_NEXT_PTR As Long, xlasPtrCopy As Long
Dim KIN_LABEL As String
Dim ExpressionArr() As String, VariableArr() As String
xlasPtr = xlasPtr + 1
xlasPtrCopy = xlasPtr

Token = Tokens.Item(xlasPtr)
Call modArtBraces(Token)
Call modArtParens(Token)

If InStr(1, Token, "-EQ", vbTextCompare) Then
VariableArr = Split(Token, "-EQ", , vbTextCompare)
ElseIf InStr(1, Token, "AS", vbTextCompare) Then
VariableArr = Split(Token, "AS", , vbTextCompare)
End If

KIN_LABEL = Trim(VariableArr(0))

If InStr(1, VariableArr(1), ":") Then
ExpressionArr = Split(VariableArr(1), ":", , vbTextCompare)
ElseIf InStr(1, VariableArr(1), "TO", vbTextCompare) Then
ExpressionArr = Split(VariableArr(1), "TO", , vbTextCompare)
End If

START_FOR_NEXT = CLng(Trim(ExpressionArr(0)))
END_FOR_NEXT = CLng(Trim(ExpressionArr(1)))

KIN_VALUE = START_FOR_NEXT

Token = Tokens.Item(xlasPtr + 1)
xlasPtr = xlasPtr + 1

Do Until xlasPtr = Tokens.Count

Call modArtBraces(Token)

'//Check for instruction
Call xlasFocusInstruction(Tokens, Token, xlasPtr, appEnv, appBlk)

'//Check for variable
If InStr(1, Token, "@") Then Call getKin(Token, KIN_LABEL, KIN_VALUE, appEnv, appBlk)

'//Check for next loop
If Token = "NEXT" Or Token = "Next" Or Token = "next" Then
END_FOR_NEXT_PTR = xlasPtr
xlasPtr = xlasPtrCopy
'//Check for end value before modifying variable
If KIN_VALUE = END_FOR_NEXT Then xlasPtr = END_FOR_NEXT_PTR: Exit Function
Token = KIN_LABEL
'//Modify variable value
Call getKin(Token, KIN_LABEL, KIN_VALUE, appEnv, appBlk)
KIN_VALUE = Token
Call modKin(KIN_LABEL, KIN_VALUE, appEnv, appBlk)
KIN_VALUE = KIN_VALUE + 1
Token = vbNullString
'//Check for end value after modifying variable
If KIN_VALUE >= END_FOR_NEXT Then xlasPtr = END_FOR_NEXT_PTR: Exit Function
End If

'//Execute article
If Token <> vbNullString Then Call xlasExecute(Token, appEnv, appBlk)

xlasPtr = xlasPtr + 1

Token = Tokens.Item(xlasPtr)

Loop

End Function
Public Function setIfElse(Tokens As Collection, Token, xlasPtr, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for parsing an if-else instruction
'/\____________________________________________________________________________________

Dim ELSE_PTR As Long, END_IF_PTR As Long, xlasPtrCopy As Long
Dim BoolArr() As String, TokenArr() As String
Dim IF_CONDITION As String
Dim IS_BOOL As Boolean, IS_ELSE As Boolean, IS_TRUE As Boolean
Dim IS_LOGIC As Byte, IS_OPERATOR As Byte, IS_TRUE_COUNT As Byte, BOOL_INDEX As Byte, BOOL_COUNT As Byte
xlasPtr = xlasPtr + 1
xlasPtrCopy = xlasPtr

Token = Tokens.Item(xlasPtr)
Call modArtBraces(Token)
Call modArtParens(Token)

'//Check for instruction
Call xlasFocusInstruction(Tokens, Token, xlasPtr, appEnv, appBlk)

If xlasPtr > xlasPtrCopy Then Exit Function

'//Check for variable
If InStr(1, Token, "@") Then Call getKin(Token, KIN_LABEL, KIN_VALUE, appEnv, appBlk)

IF_CONDITION = Token

'//Check for extended logic
If InStr(1, IF_CONDITION, "-and", vbTextCompare) Then IS_BOOL = True: BoolArr = Split(IF_CONDITION, "-and", vbTextCompare): IS_LOGIC = 1
If InStr(1, IF_CONDITION, "-or", vbTextCompare) Then IS_BOOL = True: BoolArr = Split(IF_CONDITION, "-or", vbTextCompare): IS_LOGIC = 2
If InStr(1, IF_CONDITION, "-nor", vbTextCompare) Then IS_BOOL = True: BoolArr = Split(IF_CONDITION, "-nor", vbTextCompare): IS_LOGIC = 3
If InStr(1, IF_CONDITION, "-xor", vbTextCompare) Then IS_BOOL = True: BoolArr = Split(IF_CONDITION, "-xor", vbTextCompare): IS_LOGIC = 4

'//Standard if-else
If Not IS_BOOL Then
'//Check for comparison switch
If InStr(1, IF_CONDITION, "-eq", vbTextCompare) Then BoolArr = Split(IF_CONDITION, "-eq", vbTextCompare): BoolArr = Split(BoolArr(0), "-eq"): IS_OPERATOR = 1
If InStr(1, IF_CONDITION, "-gt", vbTextCompare) Then BoolArr = Split(IF_CONDITION, "-gt", vbTextCompare): BoolArr = Split(BoolArr(0), "-gt"): IS_OPERATOR = 2
If InStr(1, IF_CONDITION, "-ge", vbTextCompare) Then BoolArr = Split(IF_CONDITION, "-ge", vbTextCompare): BoolArr = Split(BoolArr(0), "-ge"): IS_OPERATOR = 3
If InStr(1, IF_CONDITION, "-le", vbTextCompare) Then BoolArr = Split(IF_CONDITION, "-le", vbTextCompare): BoolArr = Split(BoolArr(0), "-le"): IS_OPERATOR = 4
If InStr(1, IF_CONDITION, "-lt", vbTextCompare) Then BoolArr = Split(IF_CONDITION, "-lt", vbTextCompare): BoolArr = Split(BoolArr(0), "-lt"): IS_OPERATOR = 5
If InStr(1, IF_CONDITION, "-ne", vbTextCompare) Then BoolArr = Split(IF_CONDITION, "-ne", vbTextCompare): BoolArr = Split(BoolArr(0), "-ne"): IS_OPERATOR = 6

BoolArr(0) = Trim(BoolArr(0)): BoolArr(1) = Trim(BoolArr(1))

'//Perform comparisons based on supplied switch
Select Case IS_OPERATOR

Case Is = 1
If BoolArr(0) = BoolArr(1) Then IS_TRUE = True
Case Is = 2
If BoolArr(0) > BoolArr(1) Then IS_TRUE = True
Case Is = 3
If BoolArr(0) >= BoolArr(1) Then IS_TRUE = True
Case Is = 4
If BoolArr(0) <= BoolArr(1) Then IS_TRUE = True
Case Is = 5
If BoolArr(0) < BoolArr(1) Then IS_TRUE = True
Case Is = 6
If BoolArr(0) <> BoolArr(1) Then IS_TRUE = True
End Select

Else
'//Logical if-else
If IS_LOGIC = 1 Then
TokenArr = Split(BoolArr(0), "-and", , vbTextCompare)
ElseIf IS_LOGIC = 2 Then
TokenArr = Split(BoolArr(0), "-or", , vbTextCompare)
ElseIf IS_LOGIC = 3 Then
TokenArr = Split(BoolArr(0), "-nor", , vbTextCompare)
ElseIf IS_LOGIC = 4 Then
TokenArr = Split(BoolArr(0), "-xor", , vbTextCompare)
End If

IF_CONDITION = Trim(TokenArr(0)): TokenArr(1) = Trim(TokenArr(1))

For BOOL_INDEX = 0 To 1

If InStr(1, TokenArr(BOOL_INDEX), "-eq", vbTextCompare) Then BoolArr = Split(TokenArr(BOOL_INDEX), "-eq", vbTextCompare): BoolArr = Split(BoolArr(0), "-eq"): IS_OPERATOR = 1
If InStr(1, TokenArr(BOOL_INDEX), "-gt", vbTextCompare) Then BoolArr = Split(TokenArr(BOOL_INDEX), "-gt", vbTextCompare): BoolArr = Split(BoolArr(0), "-gt"): IS_OPERATOR = 2
If InStr(1, TokenArr(BOOL_INDEX), "-ge", vbTextCompare) Then BoolArr = Split(TokenArr(BOOL_INDEX), "-ge", vbTextCompare): BoolArr = Split(BoolArr(0), "-ge"): IS_OPERATOR = 3
If InStr(1, TokenArr(BOOL_INDEX), "-le", vbTextCompare) Then BoolArr = Split(TokenArr(BOOL_INDEX), "-le", vbTextCompare): BoolArr = Split(BoolArr(0), "-le"): IS_OPERATOR = 4
If InStr(1, TokenArr(BOOL_INDEX), "-lt", vbTextCompare) Then BoolArr = Split(TokenArr(BOOL_INDEX), "-lt", vbTextCompare): BoolArr = Split(BoolArr(0), "-lt"): IS_OPERATOR = 5
If InStr(1, TokenArr(BOOL_INDEX), "-ne", vbTextCompare) Then BoolArr = Split(TokenArr(BOOL_INDEX), "-ne", vbTextCompare): BoolArr = Split(BoolArr(0), "-ne"): IS_OPERATOR = 6

BoolArr(0) = Trim(BoolArr(0)): BoolArr(1) = Trim(BoolArr(1))

'//Perform comparisons based on supplied switch
Select Case IS_OPERATOR

Case Is = 1
If BoolArr(0) = BoolArr(1) Then IS_TRUE = True
Case Is = 2
If BoolArr(0) > BoolArr(1) Then IS_TRUE = True
Case Is = 3
If BoolArr(0) >= BoolArr(1) Then IS_TRUE = True
Case Is = 4
If BoolArr(0) <= BoolArr(1) Then IS_TRUE = True
Case Is = 5
If BoolArr(0) < BoolArr(1) Then IS_TRUE = True
Case Is = 6
If BoolArr(0) <> BoolArr(1) Then IS_TRUE = True
End Select
'//Track true comparisons
If IS_TRUE Then IS_TRUE_COUNT = IS_TRUE_COUNT + 1
Next

'//Check if comparison results match logic
Select Case IS_LOGIC

'//and
Case Is = 1
If IS_TRUE_COUNT = BOOL_INDEX Then IS_TRUE = True Else IS_TRUE = False
'//or
Case Is = 2
If IS_TRUE_COUNT > 0 Then IS_TRUE = True Else IS_TRUE = False
'//nor
Case Is = 3
If IS_TRUE_COUNT = 0 Then IS_TRUE = True Else IS_TRUE = False
'//xor
Case Is = 4
If IS_TRUE_COUNT = 1 Then IS_TRUE = True Else IS_TRUE = False
End Select

End If

'//Parse until instruction end to find statement pointers
Do Until xlasPtr = Tokens.Count Or InStr(1, Token, "}ENDIF", vbTextCompare)
Token = Tokens.Item(xlasPtr)
If Token = "ELSE" Then IS_ELSE = True: ELSE_PTR = xlasPtr
If InStr(1, Token, "ENDIF", vbTextCompare) Then END_IF_PTR = xlasPtr
xlasPtr = xlasPtr + 1
Loop

'//True condition
If IS_TRUE Then
xlasPtr = xlasPtrCopy + 1
Token = Tokens.Item(xlasPtr)
'//Check for instruction
Call xlasFocusInstruction(Tokens, Token, xlasPtr, appEnv, appBlk)
'//Execute first line after condition
Call xlasExecute(Token, appEnv, appBlk)
'//False condition
ElseIf Not IS_TRUE And IS_ELSE Then
xlasPtr = ELSE_PTR
Token = Tokens.Item(xlasPtr)
Else
xlasPtr = END_IF_PTR
Token = Tokens.Item(xlasPtr)
End If

End Function
Public Function setLibCall(Tokens As Collection, Token, xlasPtr, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for parsing a libcall instruction
'/\____________________________________________________________________________________
'//
'//
'//=======================================================
Dim xlasPtrCopy As Long
Dim Lib As String
xlasPtr = xlasPtr + 1
xlasPtrCopy = xlasPtr

If Tokens.Item(xlasPtr) <> "LIBCALL" Then xlasPtr = xlasPtr - 1: xlasPtrCopy = xlasPtrCopy + 1

Lib = Tokens.Item(xlasPtr + 1)
Call modArtParens(Lib)
Lib = Trim(Lib)
Call modLibStack(Lib, appEnv, appBlk)

xlasPtr = xlasPtrCopy

Do Until xlasPtr = Tokens.Count

Token = Tokens.Item(xlasPtr)
Call modArtBraces(Token)

'//Check for instruction
Call xlasFocusInstruction(Tokens, Token, xlasPtr, appEnv, appBlk)

'//Check for variable
If InStr(1, Token, "@") Then Call getKin(Token, KIN_LABEL, KIN_VALUE, appEnv, appBlk)

'//Check for end libcall
If InStr(1, Token, "ENDLIB", vbTextCompare) Then Exit Function

'//Execute article
If Token <> vbNullString Then Call xlasExecute(Token, appEnv, appBlk)

xlasPtr = xlasPtr + 1

Loop

End Function
Public Function setWhile(Tokens As Collection, Token, xlasPtr, appEnv, appBlk) As String

'/\____________________________________________________________________________________
'//
'//     A function for parsing a while instruction
'/\____________________________________________________________________________________
'//
'//
'//=======================================================
Dim LOOP_COUNT As Long, END_WHILE As Long, END_WHILE_PTR As Long, xlasPtrCopy As Long
Dim ExpressionArr() As String, VariableArr() As String
Dim KIN_LABEL As String, KIN_VALUE As String, OPERATOR As String, TOKEN_COPY As String
Dim IS_OPERATOR As Byte
xlasPtr = xlasPtr + 1
xlasPtrCopy = xlasPtr

Token = Tokens.Item(xlasPtr)
Call modArtBraces(Token)
Call modArtParens(Token)

VariableArr = Split(Token, " ")

Token = Tokens.Item(xlasPtr + 1)
TOKEN_COPY = Token
KIN_LABEL = Trim(VariableArr(0))
OPERATOR = Trim(VariableArr(1))
END_WHILE = Trim(VariableArr(2))

If InStr(1, OPERATOR, "-eq", vbTextCompare) Then
IS_OPERATOR = 1
ElseIf InStr(1, OPERATOR, "-gt", vbTextCompare) Then
IS_OPERATOR = 2
ElseIf InStr(1, OPERATOR, "-ge", vbTextCompare) Then
IS_OPERATOR = 3
ElseIf InStr(1, OPERATOR, "-le", vbTextCompare) Then
IS_OPERATOR = 4
ElseIf InStr(1, OPERATOR, "-lt", vbTextCompare) Then
IS_OPERATOR = 5
ElseIf InStr(1, OPERATOR, "-ne", vbTextCompare) Then
IS_OPERATOR = 6
End If

'//Find end while pointer
Do Until xlasPtr = Tokens.Count Or InStr(1, Token, "}CONTINUE", vbTextCompare): Token = Tokens.Item(xlasPtr): xlasPtr = xlasPtr + 1: Loop

END_WHILE_PTR = xlasPtr - 1
xlasPtr = xlasPtrCopy

Token = KIN_LABEL: Call getKin(Token, KIN_LABEL, KIN_VALUE, appEnv, appBlk)

Do Until xlasPtr = Tokens.Count

'//Check for instruction
Call xlasFocusInstruction(Tokens, Token, xlasPtr, appEnv, appBlk)

'//Check for variable
If InStr(1, Token, "@") Then Call getKin(Token, KIN_LABEL, KIN_VALUE, appEnv, appBlk)

'//end equal
If IS_OPERATOR = 1 Then
If END_WHILE <> CLng(KIN_VALUE) Then xlasPtr = END_WHILE_PTR: Exit Function
'//end greater than
ElseIf IS_OPERATOR = 2 Then
If END_WHILE >= CLng(KIN_VALUE) Then xlasPtr = END_WHILE_PTR: Exit Function
'//end greater than equal
ElseIf IS_OPERATOR = 3 Then
If END_WHILE > CLng(KIN_VALUE) Then xlasPtr = END_WHILE_PTR: Exit Function
'//end less than equal
ElseIf IS_OPERATOR = 4 Then
If END_WHILE < CLng(KIN_VALUE) Then xlasPtr = END_WHILE_PTR: Exit Function
'//end less than
ElseIf IS_OPERATOR = 5 Then
If END_WHILE <= CLng(KIN_VALUE) Then xlasPtr = END_WHILE_PTR: Exit Function
'//end not equal
ElseIf IS_OPERATOR = 6 Then
If END_WHILE = CLng(KIN_VALUE) Then xlasPtr = END_WHILE_PTR: Exit Function
End If

'//Check for end while
If InStr(1, Token, "}CONTINUE", vbTextCompare) Then
END_WHILE_PTR = xlasPtr
xlasPtr = xlasPtrCopy - 1: Token = vbNullString
End If

'//Execute article
If Token <> vbNullString Then Call xlasExecute(Token, appEnv, appBlk)

Token = KIN_LABEL: Call getKin(Token, KIN_LABEL, KIN_VALUE, appEnv, appBlk)

xlasPtr = xlasPtr + 1

If xlasPtr = xlasPtrCopy Then Token = TOKEN_COPY Else Token = Tokens.Item(xlasPtr)

Loop

End Function
Public Function getChar$(Token)

'/\__________________________________________________________________________
'//
'//     A function for flagging if a non-numerical characters been entered
'/\__________________________________________________________________________

On Error GoTo ErrEnd

Dim IS_CHAR As String

IS_CHAR = Token

IS_CHAR = IS_CHAR + 1

Exit Function

ErrEnd:
Token = "*/ERR"

End Function
Public Function getEnvironment$(appEnv, appBlk)

'/\__________________________________________________________________________
'//
'//     A function for setting the current runtime environment & block
'/\__________________________________________________________________________

Dim Blk, CellName As Object

        On Error GoTo useThisEnv

        '//Set application runtime environment (Workbook)...
        If Range("xlasEnvironment").Value2 <> vbNullString Then appEnv = Range("xlasEnvironment").Value2

        '//Set application runtime block (Worksheet)...
        If Range("xlasBlock").Value2 <> vbNullString Then appBlk = Range("xlasBlock").Value2
        
        '//Exit if both set, otherwise find...
        If appEnv <> vbNullString And appBlk <> vbNullString Then Exit Function

        GoTo findBlock

'//Environment not found or specified (Default)
useThisEnv:
Err.Clear
If Range("xlasEnvironment").Value2 <> vbNullString Then appEnv = Range("xlasEnvironment").Value2 Else appEnv = ThisWorkbook.name

'//Find application runtime block...
findBlock:
For Each Blk In Workbooks(appEnv).Sheets
    For Each CellName In Workbooks(appEnv).Names
        If CellName.name = "xlasEnvironment" Then appBlk = Blk.name: Range("xlasBlock").Value2 = appBlk: _
        Range("xlasEnvironment").Value2 = appEnv: Exit Function
            Next
                Next

'//Block not found or specified (Default)
appBlk = appBlk: Range("xlasBlock").Value2 = appBlk: Range("xlasEnvironment").Value2 = appEnv: Exit Function

End Function
Public Function getEnvironmentVars$(Token)

'/\______________________________________________________________________________
'//
'//     A function for identifying & retrieving an environment variable string
'/\______________________________________________________________________________

If InStr(1, Token, "@envallusersprofile", vbTextCompare) Then Token = Replace(Token, "@envallusersprofile", Environ("ALLUSERSPROFILE"), , , vbTextCompare)
If InStr(1, Token, "@envappdata", vbTextCompare) Then Token = Replace(Token, "@envappdata", Environ("APPDATA"), , , vbTextCompare)
If InStr(1, Token, "@envcommonprogramfiles", vbTextCompare) Then Token = Replace(Token, "@envcommonprogramfiles", Environ("CommonProgramFiles"), , , vbTextCompare)
If InStr(1, Token, "@envcommonprogramfilesx86", vbTextCompare) Then Token = Replace(Token, "@envcommonprogramfilesx86", Environ("CommonProgramFiles(x86)"), , , vbTextCompare)
If InStr(1, Token, "@envcom", vbTextCompare) Then Token = Replace(Token, "@envcom", Environ("COMPUTERNAME"), , , vbTextCompare)
If InStr(1, Token, "@envcomspec", vbTextCompare) Then Token = Replace(Token, "@envcomspec", Environ("ComSpec"), , , vbTextCompare)
If InStr(1, Token, "@envdriverdata", vbTextCompare) Then Token = Replace(Token, "@envdriverdata", Environ("DriverData"), , , vbTextCompare)
If InStr(1, Token, "@envhomedrive", vbTextCompare) Then Token = Replace(Token, "@envhomedrive", Environ("HOMEDRIVE"), , , vbTextCompare)
If InStr(1, Token, "@envhome", vbTextCompare) Then Token = Replace(Token, "@envhome", Environ("HOMEPATH"), , , vbTextCompare)
If InStr(1, Token, "@envlocalappdata", vbTextCompare) Then Token = Replace(Token, "@envlocalappdata", Environ("LOCALAPPDATA"), , , vbTextCompare)
If InStr(1, Token, "@envlogonserver", vbTextCompare) Then Token = Replace(Token, "@envlogonserver", Environ("LOGONSERVER"), , , vbTextCompare)
If InStr(1, Token, "@envnumofprocessors", vbTextCompare) Then Token = Replace(Token, "@envnumofprocessors", Environ("NUMBER_OF_PROCESSORS"), , , vbTextCompare)
If InStr(1, Token, "@envonedrive", vbTextCompare) Then Token = Replace(Token, "@envonedrive", Environ("OneDrive"), , , vbTextCompare)
If InStr(1, Token, "@envonlineservice", vbTextCompare) Then Token = Replace(Token, "@envonlineservices", Environ("OnlineServices"), , , vbTextCompare)
If InStr(1, Token, "@envos", vbTextCompare) Then Token = Replace(Token, "@envos", Environ("OS"), , , vbTextCompare)
If InStr(1, Token, "@envpath", vbTextCompare) Then Token = Replace(Token, "@envpath", Environ("PATH"), , , vbTextCompare)
If InStr(1, Token, "@envpathext", vbTextCompare) Then Token = Replace(Token, "@envpathext", Environ("PATHEXT"), , , vbTextCompare)
If InStr(1, Token, "@envplatformcode", vbTextCompare) Then Token = Replace(Token, "@envplatformcode", Environ("PlatformCode"), , , vbTextCompare)
If InStr(1, Token, "@envprocessorarch", vbTextCompare) Then Token = Replace(Token, "@envprocessorarch", Environ("PROCESSOR_ARCHITECTURE"), , , vbTextCompare)
If InStr(1, Token, "@envprocessorid", vbTextCompare) Then Token = Replace(Token, "@envprocessorid", Environ("PROCESSOR_IDENTIFIER"), , , vbTextCompare)
If InStr(1, Token, "@envprocessorlvl", vbTextCompare) Then Token = Replace(Token, "@envprocessorlvl", Environ("PROCESSOR_LEVEL"), , , vbTextCompare)
If InStr(1, Token, "@envprocessorrev", vbTextCompare) Then Token = Replace(Token, "@envprocessorrev", Environ("PROCESSOR_REVISION"), , , vbTextCompare)
If InStr(1, Token, "@envprogramdata", vbTextCompare) Then Token = Replace(Token, "@envprogramdata", Environ("ProgramData"), , , vbTextCompare)
If InStr(1, Token, "@envprogramfiles", vbTextCompare) Then Token = Replace(Token, "@envprogramfiles", Environ("ProgramFiles"), , , vbTextCompare)
If InStr(1, Token, "@envprogramfilesx86", vbTextCompare) Then Token = Replace(Token, "@envprogramfilesx86", Environ("ProgramFiles(x86)"), , , vbTextCompare)
If InStr(1, Token, "@envpsmodulepath", vbTextCompare) Then Token = Replace(Token, "@envpsmodulepath", Environ("PSModulePath"), , , vbTextCompare)
If InStr(1, Token, "@envpublic", vbTextCompare) Then Token = Replace(Token, "@envpublic", Environ("PUBLIC"), , , vbTextCompare)
If InStr(1, Token, "@envregioncode", vbTextCompare) Then Token = Replace(Token, "@envregioncode", Environ("RegionCode"), , , vbTextCompare)
If InStr(1, Token, "@envsessionname", vbTextCompare) Then Token = Replace(Token, "@envsessionname", Environ("SESSIONNAME"), , , vbTextCompare)
If InStr(1, Token, "@envsysdrive", vbTextCompare) Then Token = Replace(Token, "@envsysdrive", Environ("SystemDrive"), , , vbTextCompare)
If InStr(1, Token, "@envsysroot", vbTextCompare) Then Token = Replace(Token, "@envsysroot", Environ("SystemRoot"), , , vbTextCompare)
If InStr(1, Token, "@envtemp", vbTextCompare) Then Token = Replace(Token, "@envtemp", Environ("TEMP"), , , vbTextCompare)
If InStr(1, Token, "@envtmp", vbTextCompare) Then Token = Replace(Token, "@envtmp", Environ("TMP"), , , vbTextCompare)
If InStr(1, Token, "@envuserdomain", vbTextCompare) Then Token = Replace(Token, "@envuserdomain", Environ("USERDOMAIN"), , , vbTextCompare)
If InStr(1, Token, "@envuserdomainrp", vbTextCompare) Then Token = Replace(Token, "@envuserdomainrp", Environ("USERDOMAIN_ROAMINGPROFILE"), , , vbTextCompare)
If InStr(1, Token, "@envuserprofile", vbTextCompare) Then Token = Replace(Token, "@envuserprofile", Environ("USERPROFILE"), , , vbTextCompare)
If InStr(1, Token, "@envuser", vbTextCompare) Then Token = Replace(Token, "@envuser", Environ("USERNAME"), , , vbTextCompare)
If InStr(1, Token, "@envwindir", vbTextCompare) Then Token = Replace(Token, "@envwindir", Environ("windir"), , , vbTextCompare)

End Function
Public Function getJunk$(Token)

'/\________________________________________________________________________________
'//
'//     A function for getting rid of junk before it's parsed
'/\________________________________________________________________________________
'//
'//
'/\_____________________________________
'//
'//     JUNK FOUND
'/\_____________________________________

If Len(Token) <= 1 Then Token = vbNullString: Exit Function
If Left(Token, Len(Token) - Len(Token) + 1) = ":" Then Token = vbNullString: Exit Function
If Left(Token, Len(Token) - Len(Token) + 1) = "}" Then Token = vbNullString: Exit Function
If InStr(1, Token, "}end") Then Token = vbNullString: Exit Function

End Function
Public Function getKin$(Token, KIN_LABEL, KIN_VALUE, appEnv, appBlk)

'/\____________________________________________________________________________________
'//
'//     A function for retrieving a variable value at runtime
'/\____________________________________________________________________________________

Dim TokenArr() As String
Dim LastRow As Long, Row As Long

LastRow = Workbooks(appEnv).Worksheets(appBlk).Cells(Rows.Count, "MAA").End(xlUp).Row

'//Check for modifying variable
'//
'//Set variable value
If Left(Token, 1) = "@" And InStr(1, Token, "=") Then

TokenArr = Split(Token, "=")

KIN_LABEL = TokenArr(0)
KIN_VALUE = TokenArr(1)
Call modKin(KIN_LABEL, KIN_VALUE, appEnv, appBlk)
Exit Function

'//Increment variable value (++)
ElseIf InStr(1, Token, "++") Then

TokenArr = Split(Token, "++")

If KIN_LABEL = vbNullString Then
KIN_LABEL = Trim(TokenArr(0))

'//Retrieve variable from block space
For Row = 1 To LastRow - 1
If KIN_LABEL = Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabelMod").Offset(Row, 0).Value2 Then
KIN_VALUE = Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(Row, 0).Value2
End If
Next

End If

KIN_VALUE = KIN_VALUE + 1
Call modKin(KIN_LABEL, KIN_VALUE, appEnv, appBlk)
Token = vbNullString
Exit Function

'//Decrement variable value (--)
ElseIf InStr(1, Token, "--") Then

TokenArr = Split(Token, "--")

If KIN_LABEL = vbNullString Then
KIN_LABEL = Trim(TokenArr(0))

'//Retrieve variable from block space
For Row = 1 To LastRow - 1
If KIN_LABEL = Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabelMod").Offset(Row, 0).Value2 Then
KIN_VALUE = Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(Row, 0).Value2
End If
Next

End If

KIN_VALUE = KIN_VALUE - 1
Call modKin(KIN_LABEL, KIN_VALUE, appEnv, appBlk)
Token = vbNullString
Exit Function

End If

'//Retrieve variable from block space
For Row = 1 To LastRow - 1

If InStr(1, Token, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabelMod").Offset(Row, 0).Value2, vbBinaryCompare) Then
Token = Replace(Token, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabelMod").Offset(Row, 0).Value2, _
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(Row, 0).Value2, , , vbBinaryCompare)
KIN_VALUE = Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(Row, 0).Value2
End If

Next

End Function
Private Function getLibs$(Lib, appEnv, appBlk)

'/\______________________________________________________________________________________
'//
'//     A function for finding an xlAppScript library based on it's column position
'/\______________________________________________________________________________________

Lib = "xlAppScript_" & Workbooks(appEnv).Worksheets(appBlk).Range("xlasLib").Offset(Lib, 0).Value2 & ".runLib"

End Function
Public Function getTool(Tool, appEnv, appBlk) As Object

'/\__________________________________________________________________________
'//
'//     A function for finding the current running tool
'/\__________________________________________________________________________

'//eTweetXL: xlFlowStrip
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 11 Then Set Tool = ETWEETXLHOME.xlFlowStrip
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 12 Then Set Tool = ETWEETXLSETUP.xlFlowStrip
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 13 Then Set Tool = ETWEETXLPOST.xlFlowStrip
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 14 Then Set Tool = ETWEETXLQUEUE.xlFlowStrip
'//Control Box: Input Fields
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 101 Then Set Tool = CTRLBOX.CtrlBoxWindow
'//AutomateXL: xlFlowStrip
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 19 Then Set Tool = XLMAPPER.xlFlowStrip

End Function
Public Function getWindow(xWin) As Object

'/\__________________________________________________________________________
'//
'//     A function for finding the current active Window/UserForm object
'/\__________________________________________________________________________

Dim appEnv, appBlk As String
Call getEnvironment(appEnv, appBlk)

'//Check for current running window
'//
'//eTweetXL: Windows
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 11 Then Set xWin = ETWEETXLHOME: Exit Function
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 12 Then Set xWin = ETWEETXLSETUP: Exit Function
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 13 Then Set xWin = ETWEETXLPOST: Exit Function
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 14 Then Set xWin = ETWEETXLQUEUE: Exit Function
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 15 Then Set xWin = ETWEETXLAPISETUP: Exit Function
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 16 Then Set xWin = ETWEETXLPOST_EX: Exit Function
'//eTweetXL: Input Fields
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 131 Then Set xWin = ETWEETXLPOST.PostBox: Exit Function
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 161 Then Set xWin = ETWEETXLPOST_EX.PostBox: Exit Function
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 141 Then Set xWin = ETWEETXLQUEUE.PostBox: Exit Function
'//Control Box: Windows
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 100 Then Set xWin = CTRLBOX: Exit Function
'//Control Box: Input Fields
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 101 Then Set xWin = CTRLBOX.CtrlBoxWindow: Exit Function
'//AutomateXL: Windows
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 18 Then Set xWin = AUTOMATEXLHOME: Exit Function
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 19 Then Set xWin = XLMAPPER: Exit Function
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = 20 Then Set xWin = XLMAPPERCTRLR: Exit Function

End Function
Public Function setWindow%(xWin)

'/\__________________________________________________________________________
'//
'//     A function for setting the current & last active Window/UserForm #
'/\__________________________________________________________________________
'
Dim appEnv, appBlk As String
Call getEnvironment(appEnv, appBlk)

Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinFormLast").Value2 = Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2
Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = xWin

End Function
Public Function setSpecial$(Token)

'/\__________________________________________________________________________________
'//
'//     A function for escaping ([#]) special characters or strings before runtime
'/\__________________________________________________________________________________

Dim TokenH$

TokenH = Token

If InStr(1, TokenH, "[dblquote]", vbTextCompare) Then
'//escape double quote symbol
TokenH = Replace(TokenH, "[dblquote]", "*/DOUBLEQUOTE")
End If

If InStr(1, TokenH, "[quote]", vbTextCompare) Then
'//escape single quote symbol
TokenH = Replace(TokenH, "[quote]", "*/SINGLEQUOTE")
End If

If InStr(1, TokenH, "[doll]", vbTextCompare) Then
'//escape dollar symbol
TokenH = Replace(TokenH, "[doll]", "*/DOLLAR")
End If

If InStr(1, TokenH, "[eq]", vbTextCompare) Then
'//escape equal symbol
TokenH = Replace(TokenH, "[eq]", "*/EQUAL")
End If

If InStr(1, TokenH, "[hash]", vbTextCompare) Then
'//escape hash symbol
TokenH = Replace(TokenH, "[hash]", "*/HASH")
End If

If InStr(1, TokenH, "[semi]", vbTextCompare) Then
'//escape semicolon symbol
TokenH = Replace(TokenH, "[semi]", "*/SEMICOLON")
End If

If InStr(1, TokenH, "[tab]", vbTextCompare) Then
'//escape tab symbol
TokenH = Replace(TokenH, "[tab]", "*/TAB")
End If

If InStr(1, TokenH, "[lf]", vbTextCompare) Then
'//escape linefeed symbol
TokenH = Replace(TokenH, "[lf]", "*/LINEFEED")
End If

If InStr(1, TokenH, "[nl]", vbTextCompare) Then
'//escape newline symbol
TokenH = Replace(TokenH, "[nl]", "*/NEWLINE")
End If

If InStr(1, TokenH, "[null]", vbTextCompare) Then
'//escape null symbol
TokenH = Replace(TokenH, "[null]", "*/NULL")
End If

If InStr(1, TokenH, "[space]", vbTextCompare) Then
'//escape space symbol
TokenH = Replace(TokenH, "[space]", "*/SPACE")
End If

If InStr(1, TokenH, "[lbrk]", vbTextCompare) Then
'//escape left brackets symbol
TokenH = Replace(TokenH, "[lbrk]", "*/LBRACKETS")
End If

If InStr(1, TokenH, "[rbrk]", vbTextCompare) Then
'//escape right brackets symbol
TokenH = Replace(TokenH, "[rbrk]", "*/RBRACKETS")
End If

If InStr(1, TokenH, "[lbr]", vbTextCompare) Then
'//escape left braces symbol
TokenH = Replace(TokenH, "[lbr]", "*/LBRACE")
End If

If InStr(1, TokenH, "[rbr]", vbTextCompare) Then
'//escape right braces symbol
TokenH = Replace(TokenH, "[rbr]", "*/RBRACE")
End If

If InStr(1, TokenH, "[lp]", vbTextCompare) Then
'//escape left parenthese symbol
TokenH = Replace(TokenH, "[lp]", "*/LPAREN")
End If

If InStr(1, TokenH, "[rp]", vbTextCompare) Then
'//escape right parentheses symbol
TokenH = Replace(TokenH, "[rp]", "*/RPAREN")
End If

Token = TokenH

End Function
Public Function getSpecial$(Token)

'/\__________________________________________________________________________________
'//
'//     A function for returning ([#]) special characters or strings before runtime
'/\__________________________________________________________________________________

Dim TokenH$

TokenH = Token

If InStr(1, TokenH, "*/DOUBLEQUOTE") Then
'//return doble quote symbol
TokenH = Replace(TokenH, "*/DOUBLEQUOTE", """")
End If

If InStr(1, TokenH, "*/SINGLEQUOTE") Then
'//return single quote symbol
TokenH = Replace(TokenH, "*/SINGLEQUOTE", "'")
End If

If InStr(1, TokenH, "*/DOLLAR") Then
'//return dollar symbol
TokenH = Replace(TokenH, "*/DOLLAR", "$")
End If

If InStr(1, TokenH, "*/EQUAL") Then
'//return dollar symbol
TokenH = Replace(TokenH, "*/EQUAL", "=")
End If

If InStr(1, TokenH, "*/HASH") Then
'//return hash symbol
TokenH = Replace(TokenH, "*/HASH", "#")
End If

If InStr(1, TokenH, "*/SEMICOLON") Then
'//return semicolon symbol
TokenH = Replace(TokenH, "*/SEMICOLON", ";")
End If

If InStr(1, TokenH, "*/TAB") Then
'//return tab
TokenH = Replace(TokenH, "*/TAB", vbTab)
End If

If InStr(1, TokenH, "*/LINEFEED") Then
'//return line feed
TokenH = Replace(TokenH, "*/LINEFEED", vbLf)
End If

If InStr(1, TokenH, "*/NEWLINE") Then
'//return newline
TokenH = Replace(TokenH, "*/NEWLINE", vbNewLine)
End If

If InStr(1, TokenH, "*/NULL") Then
'//return null
TokenH = Replace(TokenH, "*/NULL", vbNullString)
End If

If InStr(1, TokenH, "*/SPACE") Then
'//return space
TokenH = Replace(TokenH, "*/SPACE", " ")
End If

If InStr(1, TokenH, "*/LBRACKETS") Then
'//return left brackets
TokenH = Replace(TokenH, "*/LBRACKETS", "[")
End If

If InStr(1, TokenH, "*/RBRACKETS") Then
'//return right brackets
TokenH = Replace(TokenH, "*/RBRACKETS", "]")
End If

If InStr(1, TokenH, "*/LBRACE") Then
'//return left braces
TokenH = Replace(TokenH, "*/LBRACE", "{")
End If

If InStr(1, TokenH, "*/RBRACE") Then
'//return right braces
TokenH = Replace(TokenH, "*/RBRACE", "}")
End If

If InStr(1, TokenH, "*/LPAREN") Then
'//return left parentheses
TokenH = Replace(TokenH, "*/LPAREN", "(")
End If

If InStr(1, TokenH, "*/RPAREN") Then
'//return right parentheses
TokenH = Replace(TokenH, "*/RPAREN", ")")
End If

Token = TokenH

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
Public Function modBlk(ByVal Token As String, appEnv, appBlk) As Byte

'/\__________________________________________________________________________
'//
'//     A function for modifying the xlAppScript runtime block space
'/\__________________________________________________________________________

If Workbooks(appEnv).Worksheets(appBlk).Range("xlasGlobalControl").Value = 1 Then Call modCallStack(Token, appEnv, appBlk): Exit Function '//modify runtime stack

Dim TokenArr() As String
Dim ART_COUNT As Long, PRE_PTR_STATE As Long

'//Clear runtime block memory addresses
Workbooks(appEnv).Worksheets(appBlk).Range("MAA1:MAK5000").ClearContents
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasLocalStatic").Value2 <> 1 Then Workbooks(appEnv).Worksheets(appBlk).Range("MAL1:MAL5000").ClearContents
Workbooks(appEnv).Worksheets(appBlk).Range("MAM1:MAR5000").ClearContents
If Workbooks(appEnv).Worksheets(appBlk).Range("xlasLocalContain").Value2 = 1 Then Workbooks(appEnv).Worksheets(appBlk).Range("MAS1:MAS5000").ClearContents
Workbooks(appEnv).Worksheets(appBlk).Range("MAT1:MAZ5000").ClearContents

'//Print article script to host worksheet
TokenArr = Split(Token, ";")
For ART_COUNT = 0 To UBound(TokenArr) - 1
If InStr(1, TokenArr(ART_COUNT), "#") = False Then
PRE_PTR_STATE = PRE_PTR_STATE + 1
Workbooks(appEnv).Worksheets(appBlk).Range("xlasArticle").Offset(PRE_PTR_STATE, 0).Value2 = TokenArr(PRE_PTR_STATE)
Workbooks(appEnv).Worksheets(appBlk).Range("xlasState").Offset(PRE_PTR_STATE, 0).Value2 = PRE_PTR_STATE
End If
Next

End Function
Public Function modCallStack(ByVal Token As String, appEnv, appBlk) As Byte

'/\__________________________________________________________________________
'//
'//     A function for modifying the xlAppScript runtime stack space
'/\__________________________________________________________________________

Dim LastRow As Long

LastRow = Cells(Rows.Count, "MAF").End(xlUp).Row

'//Set article to top of runtime stack if empty
If Workbooks(appEnv).Worksheets(appBlk).Range("MAF2").Value2 = vbNullString Then
Else
'//If block space isn't empty push stack down 1 block
Workbooks(appEnv).Worksheets(appBlk).Range("MAF2:MAF" & LastRow).Copy Destination:=Workbooks(appEnv).Worksheets(appBlk).Range("MAF3:MAF" & LastRow + 1)
End If

Workbooks(appEnv).Worksheets(appBlk).Range("MAF2").Value2 = Token: Workbooks(appEnv).Worksheets(appBlk).Range("MAE2").Value2 = 0

End Function
Public Function modKin(KIN_LABEL, KIN_VALUE, appEnv, appBlk) As Byte

'/\__________________________________________________________________________
'//
'//     A function for modifying a variable before runtime
'/\__________________________________________________________________________

Dim LastRow As Long, Row As Long

LastRow = Workbooks(appEnv).Worksheets(appBlk).Cells(Rows.Count, "MAA").End(xlUp).Row

For Row = 1 To LastRow - 1
If InStr(1, KIN_LABEL, Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinLabelMod").Offset(Row, 0).Value2, vbBinaryCompare) Then
Workbooks(appEnv).Worksheets(appBlk).Range("xlasKinValueMod").Offset(Row, 0).Value2 = KIN_VALUE
End If
Next

End Function
Public Function modLibStack(ByVal Token As String, appEnv, appBlk) As Byte

'/\__________________________________________________________________________
'//
'//     A function for modifying the xlAppScript library stack space
'/\__________________________________________________________________________

Dim LastRow As Long

LastRow = Cells(Rows.Count, "MAL").End(xlUp).Row

'//Set library to top of library stack if empty
If Workbooks(appEnv).Worksheets(appBlk).Range("MAL2").Value2 = vbNullString Then
Else
'//If block space isn't empty push stack down 1 block
Workbooks(appEnv).Worksheets(appBlk).Range("MAL2:MAL" & LastRow).Copy Destination:=Workbooks(appEnv).Worksheets(appBlk).Range("MAL3:MAL" & LastRow + 1)
End If

Workbooks(appEnv).Worksheets(appBlk).Range("MAL2").Value2 = Token
Workbooks(appEnv).Worksheets(appBlk).Range("xlasLibCount").Value2 = LastRow

End Function
Public Function modArtBraces$(Token)

'/\__________________________________________________________________________
'//
'//     A function for modifying an article/line of xlAppScript (braces)
'/\__________________________________________________________________________


If Left(Token, 1) = "{" Then Token = Right(Token, Len(Token) - 1)
If Right(Token, 1) = "}" Then Token = Left(Token, Len(Token) - 1)
If Left(Token, 1) = "}" Then Token = Right(Token, Len(Token) - 1)
If Right(Token, 1) = "{" Then Token = Left(Token, Len(Token) - 1)

End Function
Public Function modArtPercent$(Token)

'/\__________________________________________________________________________
'//
'//     A function for modifying an article/line of xlAppScript (percent)
'/\__________________________________________________________________________


If Left(Token, 1) = "%" Then Token = Right(Token, Len(Token) - 1)
If Right(Token, 1) = "%" Then Token = Left(Token, Len(Token) - 1)
If Left(Token, 1) = "%" Then Token = Right(Token, Len(Token) - 1)
If Right(Token, 1) = "%" Then Token = Left(Token, Len(Token) - 1)

End Function
Public Function modArtParens$(Token)

'/\__________________________________________________________________________
'//
'//     A function for modifying an article/line of xlAppScript (parentheses)
'/\__________________________________________________________________________


If Left(Token, 1) = "(" Then Token = Right(Token, Len(Token) - 1)
If Right(Token, 1) = ")" Then Token = Left(Token, Len(Token) - 1)
If Left(Token, 1) = ")" Then Token = Right(Token, Len(Token) - 1)
If Right(Token, 1) = "(" Then Token = Left(Token, Len(Token) - 1)

End Function
Public Function modArtQuotes$(Token)

'/\__________________________________________________________________________
'//
'//     A function for modifying an article/line of xlAppScript (quotations)
'/\__________________________________________________________________________


If Left(Token, 1) = """" Then Token = Right(Token, Len(Token) - 1)
If Right(Token, 1) = """" Then Token = Left(Token, Len(Token) - 1)
If Left(Token, 1) = "'" Then Token = Right(Token, Len(Token) - 1)
If Right(Token, 1) = "'" Then Token = Left(Token, Len(Token) - 1)

End Function
Public Function modArtSpaces$(Token)

'/\__________________________________________________________________________
'//
'//     A function for modifying an article/line of xlAppScript (spaces)
'/\__________________________________________________________________________

Token = Replace(Token, " ", vbNullString)

End Function
