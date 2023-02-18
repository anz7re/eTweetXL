Attribute VB_Name = "xlAppScript_xtwt"

Public Function runLib$(Token)
'/\____________________________________________________________________________________________________________________________
'//
'//       xtwt Library
'//         Version: 1.0.8
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
'//     xtwt is an xlAppScript library built for automating eTweetXL.
'//
'//
'//
'//     Basic Lib Requirements: Windows 10, MS Excel Version 2107, PowerShell 5.1.19041.1023, eTweetXL v1.9.0
'//
'//                             (previous versions not tested &/or unsupported)
'/\____________________________________________________________________________________________________________________________
'//
'//     Latest Revision: 2/1/2023
'/\____________________________________________________________________________________________________________________________
'//
'//     Developer(s): anz7re (André)
'//     Contact: support@xlappscript.org | support@autokit.tech | anz7re@autokit.tech
'//     Web: xlappscript.org | autokit.tech
'/\____________________________________________________________________________________________________________________________

        '//Library variable declarations
        Dim appEnv As String: Dim appBlk As String: Dim errLvl As Byte: Dim wbMacro As String
        Dim M As String: Dim S As String: Dim P As String
        Dim X As Variant: Dim x1 As Variant: Dim x2 As Variant
        Dim oBox As Object: Dim oItem As Object
        
        '//Pre-cleanup
        X = 0: X = CByte(X): x1 = 0: x1 = CByte(x1): x2 = 0: x2 = CByte(x2)
        Set oBox = Nothing: Set oItem = Nothing
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
        
        '//Check for ADA Article
        If InStr(1, Token, "app.") Then GoTo ADALink
                     
                     
'/\_____________________________________
'//
'//     DRAFT ARTICLES
'/\_____________________________________
'//


'//Modify drafts
        If InStr(1, Token, "draft(", vbTextCompare) Then
        Token = Replace(Token, "draft(", "", , , vbTextCompare)
        If InStr(1, Token, "del.", vbTextCompare) Then M = "D": _
        Token = Replace(Token, "del.", "", , , vbTextCompare) '//check for delete draft(s)...
        If InStr(1, Token, "add.", vbTextCompare) Then M = "A": _
        Token = Replace(Token, "add.", "", , , vbTextCompare) '//check for add draft(s)...
        If InStr(1, Token, "rmv.", vbTextCompare) Then M = "R": _
        Token = Replace(Token, "rmv.", "", , , vbTextCompare) '//check for remove draft(s)...
        '//special characters/wildcards
        If InStr(1, Token, "*") Then M = M & "0": _
        Token = Replace(Token, "*", vbNullString) '//Check for all...
        If InStr(1, Token, ",") Then M = M & "1" '//Check for and...
        If InStr(1, Token, ":") Then M = M & "2" '//Check for through...
        
        Call modArtParens(Token)
        
        If M = "A" Then GoTo AddDraft
        If M = "A0" Then GoTo AddAllDraft
        If M = "A1" Then GoTo AddDraft
        If M = "A2" Then GoTo AddDraft
        If M = "A01" Then GoTo AddDraft
        If M = "A02" Then GoTo AddDraft
        If M = "R" Then GoTo RmvDraft
        If M = "R0" Then GoTo RmvAllDraft
        If M = "D" Then GoTo DeleteDraft
        If M = "D0" Then GoTo DeleteAllDraft
        
        '//Set draft name (no modifier)
        ETWEETXLPOST.DraftBox.Value = Token
        Exit Function
        
        
AddDraft:
        '//Add draft to linker position
        xPos = CDbl(Token) '//position

        Call eTweetXL_CLICK.AddLinkBtn_Clk(xPos) '//Add runtime
        
        Exit Function
        
AddAllDraft:
        
        Set oBox = ETWEETXLPOST.DraftBox
        
        For X = 0 To oBox.ListCount - 1
        oBox.Value = oBox.List(X)
        Call eTweetXL_CLICK.AddLinkBtn_Clk(xPos)
        Next
        
        Set oBox = Nothing
        Exit Function
        
RmvDraft:
        '//Remove draft from linker position
        xPos = CDbl(Token) '//position
        ETWEETXLPOST.LinkerBox.RemoveItem (xPos)
        Exit Function
        
RmvAllDraft:
        '//Remove all drafts from linker
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.DraftHdr_Clk")
        Exit Function
       
DeleteDraft:
        '//Delete draft from archive
        If Token = vbNullString Then
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.DelDraftBtn_Clk")
            Else
                ETWEETXLPOST.DraftBox.Value = Token: wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.DelDraftBtn_Clk")
                    End If
        Exit Function
       
DeleteAllDraft:
        '//Delete all drafts from archive
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.DelAllDraftBtn_Clk")
        Exit Function
       
    
'//Modify thread
        ElseIf InStr(1, Token, "thread(", vbTextCompare) Then
        Token = Replace(Token, "thread(", "", , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        
        If InStr(1, Token, "add.", vbTextCompare) Then M = "A" '//Check for add...
        If InStr(1, Token, "rmv.", vbTextCompare) Then M = "R" '//Check for remove...
        '//special characters/wildcards
        If InStr(1, Token, "*") Then M = M & "0" '//Check for all...
        If InStr(1, Token, ",") Then M = M & "1" '//Check for and...
        
        Token = Replace(Token, "add.", "", , , vbTextCompare)
        Token = Replace(Token, "rmv.", "", , , vbTextCompare)
        Token = Replace(Token, "*", vbNullString)
        Call modArtParens(Token)
        
        '//check for arguments...
        If InStr(1, Token, ",") Then
        TokenArr = Split(Token, ",")
        xLast = UBound(TokenArr) - LBound(TokenArr)
        
        '//add string to single thread
        If M = "1" Then
        Range("ThreadScrollPos").Value = TokenArr(0)
        Range("PostThread").Offset(TokenArr(0), 0).Value = TokenArr(1) '//2 arguments
        If xLast > 1 Then Range("MedThread").Offset(TokenArr(0), 0).Value = TokenArr(2) '//3 arguments
        Exit Function
        End If

        '//add string to all threads
        If M = "01" Then
        lRow = Cells(Rows.Count, "Y").End(xlUp).Row
        If lRow < 1 Then lRow = Cells(Rows.Count, "Z").End(xlUp).Row
        
        For X = 1 To lRow
        Range("PostThread").Offset(X, 0).Value2 = TokenArr(1) '//2 arguments
        If xLast > 1 Then Range("MedThread").Offset(X, 0).Value2 = TokenArr(2) '//3 arguments
        Next
        Exit Function
            End If
                End If
        
        If M = vbNullString And Token = vbNullString Then GoTo AddThread
        If M = vbNullString And Token <> vbNullString Then GoTo FindThread
        If M = "A" Then GoTo AddThread
        If M = "A1" Then GoTo AddMultiThread
        If M = "A01" Then GoTo AddMultiThread
        If M = "R" Then GoTo RmvThread
        If M = "R0" Then GoTo RmvAllThread
        If M = "R1" Then GoTo RmvMultiThread
        If M = "R01" Then GoTo RmvAllThread
        
        GoTo ErrEnd
        Exit Function
        
AddThread:
        '//add single thread
        If Token <> vbNullString Then xLast = CInt(Token): GoTo AddMultiThread
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.AddThreadBtn_Clk")
        Exit Function
        
AddMultiThread:
        '//add multi thread
        Call getWindow(xWin)
        For X = 0 To xLast
        If xWin.PostBox.Value = vbNullString Then xWin.PostBox.Value = "thread" & X
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.AddThreadBtn_Clk")
        Next
        Exit Function

RmvThread:
        '//remove single focused thread
        If Token <> vbNullString Then xLast = CInt(Token): GoTo RmvMultiThread
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.RmvThreadBtn_Clk")
        Exit Function

RmvMultiThread:
        '//remove multi media
        For X = 0 To xLast
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.RmvThreadBtn_Clk")
        Next
        Exit Function
        
RmvAllThread:
        '//remove all threads
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.RmvAllThreadBtn_Clk")
        Exit Function
        
      
FindThread:
        '//find & focus thread
        Call getWindow(xWin)
        X = CInt(Token)
        Range("ThreadScrollPos").Value = X
        xWin.ThreadCt.Caption = X
        xWin.PostBox.Value = Range("PostThread").Offset(X, 0).Value2  '//2 arguments
        If xLast > 1 Then xWin.MedLinkBox.Value = Range("MedThread").Offset(X, 0).Value2
        Exit Function
       
'//Modify media
        ElseIf InStr(1, Token, "med(", vbTextCompare) Then
        Token = Replace(Token, "med(", "", , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        
        If InStr(1, Token, "add.", vbTextCompare) Then M = "A" '//Check for add...
        If InStr(1, Token, "rmv.", vbTextCompare) Then M = "R" '//Check for remove...
        '//special characters/wildcards
        If InStr(1, Token, "*") Then M = M & "0" '//Check for all...
        If InStr(1, Token, ",") Then M = M & "1" '//Check for and...
        
        Token = Replace(Token, "add.", "", , , vbTextCompare)
        Token = Replace(Token, "rmv.", "", , , vbTextCompare)
        Token = Replace(Token, "*", vbNullString)
         Call modArtParens(Token)
         
        '//check for multiple items
        If InStr(1, Token, ",") Then
        TokenArr = Split(Token, ",")
        xLast = UBound(TokenArr) - LBound(TokenArr)
        End If
        
        xMed = Token
        
        If M = "A" Then GoTo AddMedia
        If M = "A0" Then GoTo AddMultiMedia
        If M = "A1" Then GoTo AddMultiMedia
        If M = "A01" Then GoTo AddMultiMedia
        If M = "R" Then GoTo RmvMedia
        If M = "R0" Then GoTo RmvAllMedia
        If M = "R1" Then GoTo RmvMultiMedia
        If M = "R01" Then GoTo RmvAllMedia
    
        Exit Function
        
AddMedia:
        '//add single media
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.AddPostMedBtn_Clk", (xMed))
        Exit Function
        
AddMultiMedia:
        '//add multi media
        For X = 0 To xLast
        xMed = TokenArr(X)
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.AddPostMedBtn_Clk", (xMed))
        Next
        Exit Function

    
RmvMedia:
        '//rmv single media
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.RmvPostMedBtn_Clk")
        Exit Function

RmvMultiMedia:
        '//rmv multi media
        For X = 0 To xLast
        xMed = TokenArr(X)
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.RmvPostMedBtn_Clk")
        Next
        Exit Function
        
RmvAllMedia:
        Do Until ETWEETXLPOST.MedLinkBox.Value = vbNullString
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.RmvPostMedBtn_Clk")
        Loop
        Exit Function
        
'//Modify post
        ElseIf InStr(1, Token, "post(", vbTextCompare) And InStr(1, Token, ".post(", vbTextCompare) = False Then
        Token = Replace(Token, "post(", vbNullString, , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        
        If Token = vbNullString Then wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_Focus.shw_ETWEETXLPOST"): Exit Function      '//nothing entered show post setup
        
        '//special characters/wildcards
        If InStr(1, Token, ",") Then
        TokenArr = Split(Token, ",")
        '//switches
        If InStr(1, TokenArr(1), "-true", vbTextCompare) Or InStr(1, TokenArr(1), "1") Then S = "T"
        If InStr(1, TokenArr(1), "-false", vbTextCompare) Or InStr(1, TokenArr(1), "0") Then S = "F"
        Call getWindow(xWin)
        xWin.PostBox.Value = xWin.PostBox.Value & TokenArr(0)
        If S = "T" Then Call eTweetXL_CLICK.SavePostBtn_Clk
        Exit Function
        End If
        
        Call getWindow(xWin)
        xWin.PostBox.Value = xWin.PostBox.Value & Token
        Exit Function
        
'/\_____________________________________
'//
'//         PROFILE ARTICLES
'/\_____________________________________
'//
        
'//Modify profile
        ElseIf InStr(1, Token, "profile(", vbTextCompare) Then
        Token = Replace(Token, "profile(", vbNullString, , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        
        If InStr(1, Token, "del.") Then GoTo DelProfile '//check for delete profile...
        If InStr(1, Token, "mk.") Then GoTo NewProfile '//check for make new profile...
        
        Call getWindow(xWin)
        xWin.ProfileListBox.Value = Token
        Workbooks(appEnv).Worksheets(appBlk).Range("Profile").Value2 = Token
        Exit Function
        
NewProfile:
        Token = Replace(Token, "mk.", vbNullString)
        Call modArtParens(Token)
        
        '//add multiple users (no arguments)
        '//switches
        If InStr(1, Token, "-list", vbTextCompare) Then
        Token = Replace(Token, "-list", vbNullString, , , vbTextCompare)
        TokenArr = Split(Token, ",")
        For X = 0 To UBound(TokenArr)
        xInfo = TokenArr(X)
        Call eTweetXL_CLICK.NewProfile_Clk(xInfo)
        Next
        Exit Function
        End If
        
        xInfo = Token
        Call eTweetXL_CLICK.NewProfile_Clk(xInfo)
        Exit Function
        
DelProfile:
        Token = Replace(Token, "del.", vbNullString)
        Call modArtParens(Token)
        If InStr(1, Token, "-f", vbTextCompare) Then Token = Replace(Token, "-f", vbNullString, , , vbTextCompare): Range("xlasSilent").Value2 = 1 '//force deletion no prompt
        If InStr(1, Token, "*") Then GoTo DelAllProfile '//all switch
        
        '//add multiple users (no arguments)
        '//switches
        If InStr(1, Token, "-list", vbTextCompare) Then
        Token = Replace(Token, "-list", vbNullString, , , vbTextCompare)
        TokenArr = Split(Token, ",")
        For X = 0 To UBound(TokenArr)
        xInfo = TokenArr(X)
        Call eTweetXL_CLICK.RmvProfileBtn_Clk(xInfo)
        Next
        Exit Function
        End If
        
        xInfo = Token
        Call eTweetXL_CLICK.RmvProfileBtn_Clk(xInfo)
        Exit Function
        
DelAllProfile:
        Call eTweetXL_CLICK.RmvAllProfilesBtn_Clk
        Exit Function
        
'/\_____________________________________
'//
'//     USER ARTICLES
'/\_____________________________________
'//
    
'//Modify user
        ElseIf InStr(1, Token, "user(", vbTextCompare) Then
        
        Token = Replace(Token, "user(", vbNullString, , , vbTextCompare)
        If InStr(1, Token, "mk.", vbTextCompare) Then GoTo NewUser '//check for make new user(s)...
        If InStr(1, Token, "del.", vbTextCompare) Then GoTo DelUser '//check for delete user(s)...
        If InStr(1, Token, "add.", vbTextCompare) Then M = "A": _
        Token = Replace(Token, "add.", vbNullString, , , vbTextCompare) '//check for add user(s)...
        If InStr(1, Token, "rmv.", vbTextCompare) Then M = "R": _
        Token = Replace(Token, "rmv.", vbNullString, , , vbTextCompare) '//check for remove user(s)...
        '//special characters/wildcards
        If InStr(1, Token, "*") Then M = M & "0": _
        Token = Replace(Token, "*", vbNullString) '//check for all...
        If InStr(1, Token, ",") Then M = M & "1" '//check for and...
        If InStr(1, Token, ":") Then M = M & "2" '//check for through...
        
        Call modArtParens(Token)
        
        If M = "A" Then GoTo AddUser
        If M = "A0" Then GoTo AddAllUser
        If M = "A1" Then GoTo AddUser
        If M = "A2" Then GoTo AddUser
        If M = "A01" Then GoTo AddUser
        If M = "A02" Then GoTo AddUser
        If M = "R" Then GoTo RmvUser
        If M = "R0" Then GoTo RmvAllUser

        '//Set username (no modifier)
        xUser = Token
        Workbooks(appEnv).Worksheets(appBlk).Range("User").Value2 = xUser
        Call getWindow(xWin)
        xWin.UserListBox.Value = xUser
        Call eTweetXL_CLICK.SetActive_Clk(xUser)
        Exit Function
        
AddUser:
'//Add draft to Linker position
If Token = vbNullString Then Token = 0
xPos = CDbl(Token) '//position

        Call eTweetXL_CLICK.AddUserBtn_Clk(xPos) '//Add runtime
        
        Exit Function
        
AddAllUser:
xPos = 0
        '//Add all users to Linker
        Set oBox = ETWEETXLPOST.LinkerBox
        
        x1 = ETWEETXLPOST.LinkerBox.ListCount - ETWEETXLPOST.UserBox.ListCount
        x1 = x1 - 1
        
        For X = 0 To x1
        oBox.Value = oBox.List(X)
        Call eTweetXL_CLICK.AddUserBtn_Clk(xPos)
        Next
        
        Set oBox = Nothing
        Exit Function
        
RmvUser:
        '//Remove user from Linker position
        xPos = CDbl(Token) '//position
        ETWEETXLPOST.UserBox.RemoveItem (xPos)
        Exit Function
        
RmvAllUser:
        '//Remove all users from linker
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.UserHdr_Clk")
        Exit Function
    
DelUser:
Token = Replace(Token, "del.", vbNullString)
Call modArtParens(Token)
'//special characters/wildcards
If InStr(1, Token, "*") Then GoTo DelAllUser
'//switches
If InStr(1, Token, "-f", vbTextCompare) Then Token = Replace(Token, "-f", vbNullString, , , vbTextCompare): Range("xlasSilent").Value2 = 1 '//force deletion no prompt
'//add multiple users (no arguments)
If InStr(1, Token, "-list", vbTextCompare) Then
Token = Replace(Token, "-list", vbNullString, , , vbTextCompare)
TokenArr = Split(Token, ",")
For X = 0 To UBound(TokenArr)
xInfo = TokenArr(X)
Call eTweetXL_CLICK.DelUserBtn_Clk(xInfo)
Next
Exit Function
End If

xInfo = Token
Call eTweetXL_CLICK.DelUserBtn_Clk(xInfo)
Exit Function

DelAllUser:
Call eTweetXL_CLICK.DelAllUsersBtn_Clk
Exit Function

    
NewUser:
Token = Replace(Token, "mk.", "", , , vbTextCompare)
Call modArtParens(Token)

'//add multiple users (no arguments)
'//switches
If InStr(1, Token, "-list", vbTextCompare) Then
Token = Replace(Token, "-list", vbNullString, , , vbTextCompare)
TokenArr = Split(Token, ",")
For X = 0 To UBound(TokenArr)
xInfo = TokenArr(X)
Call eTweetXL_CLICK.NewUser_Clk(xInfo)
Next
Exit Function
End If

'//add single user w/ or w/o parameter(s)
    xInfo = Token
    Call eTweetXL_CLICK.NewUser_Clk(xInfo)
    Exit Function

    
        
'/\_____________________________________
'//
'//     TIME ARTICLES
'/\_____________________________________
'//
        
'//Modify time
        ElseIf InStr(1, Token, "time(", vbTextCompare) Then
            
            If InStr(1, Token, "add.", vbTextCompare) Then M = "A" '//check for add time...
            If InStr(1, Token, "rmv.", vbTextCompare) Then M = "R" '//check for remove time...
            '//special characters/wildcards
            If InStr(1, Token, "*") Then M = M & "0" '//check for all...
            If InStr(1, Token, ",") Then M = M & "1" '//check for and...
            If InStr(1, Token, ":") Then M = M & "2" '//check for through...
            
            Token = Replace(Token, "add.time", "", , , vbTextCompare)
            Token = Replace(Token, "rmv.time", "", , , vbTextCompare)
            Token = Replace(Token, "*", vbNullString)
            Call modArtParens(Token)
        
            If InStr(1, M, "A") Or InStr(1, M, "R") = False Then GoTo AddTimeDirect
            
            If InStr(1, Token, "h", vbTextCompare) Then
            TokenArr = Split(Token, "h", , vbTextCompare)
            Offset = "H"
            xCount = TokenArr(0)
                End If
            If InStr(1, Token, "m", vbTextCompare) Then
            TokenArr = Split(Token, "m", , vbTextCompare)
            Offset = "M"
            xCount = TokenArr(0)
                End If
            If InStr(1, Token, "s", vbTextCompare) Then
            TokenArr = Split(Token, "s", , vbTextCompare)
            Offset = "S"
            xCount = TokenArr(0)
                End If
            
            xCount = Replace(xCount, """", vbullstring)
            
            If M = "A" & Offset = vbNullString Then '//add time once w/o offset
            GoTo AddTimeOnce
                ElseIf M = "A" Then '//add time once w/ offset
                    GoTo AddTime
                        End If
               
            If M = "A0" Then GoTo AddTime
            If M = "A1" Then GoTo AddTime
            If M = "A2" Then GoTo AddTime
            If M = "A01" Then GoTo AddTime
            If M = "A02" Then GoTo AddTime
            If M = "R" Then GoTo RmvTime
            If M = "R0" Then GoTo RmvAllTime
                         
            Exit Function
            
AddTimeDirect:
Call modArtQuotes(Token)
xPos = 0
ETWEETXLPOST.TimeBox.Value = Token
Call eTweetXL_CLICK.AddRuntimeBtn_Clk(xPos)
Exit Function

AddTimeOnce:
Call modArtQuotes(Token)
xPos = CDbl(Token) '//position

        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.TimerHdr_Clk") '//refresh current time
        
        Call eTweetXL_CLICK.AddRuntimeBtn_Clk(xPos) '//add runtime
        
        Exit Function


AddTime:
Call modArtQuotes(Token)
If TokenArr(0) = vbNullString Then xPos = CDbl(Token) Else xPos = TokenArr(0) '//position
 
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.TimerHdr_Clk") '//refresh current time
        
        x1 = ETWEETXLPOST.LinkerBox.ListCount - ETWEETXLPOST.RuntimeBox.ListCount
        x1 = x1 - 1
        If x1 <= 0 Then x1 = 0
        
        For X = 0 To x1

        If Offset = "S" Then '//seconds
            If xCount > 0 Then
            Call eTweetXL_CLICK.UpSecBtn_Clk(xCount): xPos = 0
                Else
                    Call eTweetXL_CLICK.DwnSecBtn_Clk(xCount): xPos = 0
                    End If
                        End If
            
       If Offset = "M" Then '//minutes
            If xCount > 0 Then
            Call eTweetXL_CLICK.UpMinBtn_Clk(xCount): xPos = 0
                Else
                    Call eTweetXL_CLICK.DwnMinBtn_Clk(xCount): xPos = 0
                    End If
                        End If
         
       
       If Offset = "H" Then '//hours
            If xCount > 0 Then
            Call eTweetXL_CLICK.UpHrBtn_Clk(xCount): xPos = 0
                Else
                    Call eTweetXL_CLICK.DwnHrBtn_Clk(xCount): xPos = 0
                    End If
                        End If
                              
            
        Call eTweetXL_CLICK.AddRuntimeBtn_Clk(xPos) '//ADD RUNTIME
        
        If xPos = 0 Then xPos = vbNullString
        
        Next X
        
        Exit Function
        
RmvTime:
        '//Remove specific time from runtime
        Call modArtQuotes(Token)
        xPos = CDbl(Token) '//position
        ETWEETXLPOST.RuntimeBox.RemoveItem (xPos)
        Exit Function
                        
RmvAllTime:
        '//Remove all time from runtime
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.LinkerHdr_Clk")
        Exit Function
       
        
'/\_____________________________________
'//
'//         WINFORM ARTICLES
'/\_____________________________________
'//
'//Output current window number...
ElseIf InStr(1, Token, "me()", vbTextCompare) And Len(Token) <= 4 Then MsgBox (Range("xlasWinForm").Value2): Exit Function

'//Set window number...
        ElseIf InStr(1, Token, "winform(", vbTextCompare) Then
        
    '//switches
        If InStr(1, Token, "-last", vbTextCompare) Then _
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = _
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinFormLast").Value2: Exit Function       '//set to last window
        
        Token = Replace(Token, "winform(", "", , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        
        Call xlAppScript_lex.getChar(Token)
        If Token = "(*/ERR)" Then Exit Function
        
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = Token
        
        Exit Function
        
'//Show application windows...
ElseIf InStr(1, Token, "show.home", vbTextCompare) Or InStr(1, Token, "home()", vbTextCompare) Then ETWEETXLHOME.Show: Exit Function
ElseIf InStr(1, Token, "show.setup", vbTextCompare) Or Token = "setup()" Then ETWEETXLSETUP.Show: Exit Function
ElseIf InStr(1, Token, "show.post", vbTextCompare) Or Token = "post()" Then ETWEETXLPOST.Show: Exit Function
ElseIf InStr(1, Token, "show.queue", vbTextCompare) Or Token = "queue()" Then ETWEETXLQUEUE.Show: Exit Function
ElseIf InStr(1, Token, "show.apisetup", vbTextCompare) Or Token = "apisetup()" Then ETWEETXLAPISETUP.Show: Exit Function
ElseIf InStr(1, Token, "show.me", vbTextCompare) Then Call getWindow(xWin): xWin.Show: Exit Function
'//Hide application windows...
ElseIf InStr(1, Token, "hide.home", vbTextCompare) Then ETWEETXLHOME.Hide: Exit Function
ElseIf InStr(1, Token, "hide.setup", vbTextCompare) Then ETWEETXLSETUP.Hide: Exit Function
ElseIf InStr(1, Token, "hide.post", vbTextCompare) Then ETWEETXLPOST.Hide: Exit Function
ElseIf InStr(1, Token, "hide.queue", vbTextCompare) Then ETWEETXLQUEUE.Hide: Exit Function
ElseIf InStr(1, Token, "hide.apisetup", vbTextCompare) Then ETWEETXLAPISETUP.Hide: Exit Function
ElseIf InStr(1, Token, "hide.me", vbTextCompare) Then Call getWindow(xWin): xWin.Hide: Exit Function
Exit Function

End If '//end

'/\_____________________________________________
'//
'//     ADA ARTICLES (APPLICATION DIRECT ACTION)
'/\_____________________________________________
'//

ADALink:

'//Start application
        If InStr(1, Token, "start(", vbTextCompare) Then
        If InStr(1, Token, "app.", vbTextCompare) Or InStr(1, Token, "start(", vbTextCompare) Then
        wbMacro = "eTweetXL_CLICK.StartBtn_Clk"
        Workbooks(appEnv).Application.Run "'" & appEnv & "'!" & wbMacro: End If
        Exit Function
        
'//Break application
        ElseIf InStr(1, Token, "break(", vbTextCompare) Then
        If InStr(1, Token, "app.", vbTextCompare) Or InStr(1, Token, "break(", vbTextCompare) Then
        wbMacro = "eTweetXL_CLICK.BreakBtn_Clk"
        Workbooks(appEnv).Application.Run "'" & appEnv & "'!" & wbMacro: End If
        Exit Function
        
'//Connect Linker
        ElseIf InStr(1, Token, "app.connect(", vbTextCompare) Or InStr(1, Token, "connect(", vbTextCompare) Then
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.ConnectBtn_Clk")
        Exit Function
        
'//Set draft filter (single/threaded)
        ElseIf InStr(1, Token, "app.dfilter(", vbTextCompare) Or InStr(1, Token, "dfilter(", vbTextCompare) Then
        Token = Replace(Token, "app.", vbNullString, , , vbTextCompare)
        Token = Replace(Token, "dfilter(", vbNullString, , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        xFil = CByte(Token)
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.DraftFilterBtn_Clk", (xFil))
        Exit Function
        
'//Set dynamic offset
        ElseIf InStr(1, Token, "app.dynoffset(", vbTextCompare) Or InStr(1, Token, "dynoffset(", vbTextCompare) Then
        Token = Replace(Token, "app.", vbNullString, , , vbTextCompare)
        Token = Replace(Token, "dynoffset(", vbNullString, , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        xPos = Token
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.DynOffset_Clk", (xPos))
        Exit Function
        
'//Freeze/unfreeze application
        ElseIf InStr(1, Token, "app.freeze(", vbTextCompare) Or InStr(1, Token, "freeze(", vbTextCompare) Then
        Token = Replace(Token, "app.", vbNullString, , , vbTextCompare)
        Token = Replace(Token, "freeze(", vbNullString, , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        If Token = 0 Then Range("AppState").Value2 = 1
        If Token = 1 Then Range("AppState").Value2 = 2
        wbMacro = "eTweetXL_CLICK.FreezeBtn_Clk"
        Workbooks(appEnv).Application.Run "'" & appEnv & "'!" & wbMacro
        Exit Function
        
'//Set application help wizard on/off
        ElseIf InStr(1, Token, "app.help(", vbTextCompare) Or InStr(1, Token, "help(", vbTextCompare) Then
        Token = Replace(Token, "app.", vbNullString, , , vbTextCompare)
        Token = Replace(Token, "help(", vbNullString, , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        If Token = 0 Then Range("HelpStatus").Value2 = 0
        If Token = 1 Then Range("HelpStatus").Value2 = 1
        xPos = Token
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.HelpStatusBtn_Clk", (xPos))
        Exit Function
    
'//Hide application
        ElseIf InStr(1, Token, "app.hide(", vbTextCompare) Or InStr(1, Token, "hide(", vbTextCompare) Then
        Call eTweetXL_CLICK.HideBtn_Clk
        Exit Function
    
'//Create a post from a loaded text file
        ElseIf InStr(1, Token, "app.load.post(", vbTextCompare) Or InStr(1, Token, "load.post(", vbTextCompare) Then
        Token = Replace(Token, "app.load.post(", vbNullString, , , vbTextCompare)
        Token = Replace(Token, "app.", vbNullString, , , vbTextCompare)
        Token = Replace(Token, "load.post(", vbNullString, , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        If InStr(1, Token, ",") Then
        TokenArr = Split(Token, ",")
        '//2 arguments
        If UBound(TokenArr) = 1 Then xName = TokenArr(0): xPath = TokenArr(1): xPath = Trim(xPath): Call eTweetXL_CLICK.LoadPostBtn_Clk(xName, xPath): Exit Function
        End If
        '//1 argument
        xPath = Token: xPath = Trim(xPath): Call eTweetXL_CLICK.LoadPostBtn_Clk(xName, xPath)
        Exit Function
        
'//Load a designated link
        ElseIf InStr(1, Token, "app.load.linker(", vbTextCompare) Or InStr(1, Token, "load.linker(", vbTextCompare) Then
        
        '//switches
        If InStr(1, Token, "-last", vbTextCompare) Then
        '//Reload last connected link
        xLink = AppLoc & "\mtsett\lastlink.link"
        Call eTweetXL_GET.getLink(xLink)
        Exit Function
        End If
        
        Token = Replace(Token, "app.load.linker(", vbNullString, , , vbTextCompare)
        Token = Replace(Token, "app.", vbNullString, , , vbTextCompare)
        Token = Replace(Token, "load.linker(", vbNullString, , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        If InStr(1, Token, ",") Then GoTo LoadMultiLink
        xLink = Token
        
        If xLink <> vbNullString Then
        Call eTweetXL_GET.getLink(xLink)
        Else
        '//Reload last imported link
        xLink = Range("RemLink").Value
        Call eTweetXL_GET.getLink(xLink)
        End If
        Exit Function

LoadMultiLink:
        TokenArr = Split(Token, ",")
        
        For X = 0 To UBound(TokenArr)
        xLink = TokenArr(X)
        Call eTweetXL_GET.getLink(xLink)
        Next
        Exit Function

'//Set offset
        ElseIf InStr(1, Token, "app.offset(", vbTextCompare) Or InStr(1, Token, "offset(", vbTextCompare) Then
        Token = Replace(Token, "app.(", vbNullString, , , vbTextCompare)
        Token = Replace(Token, "offset(", vbNullString, , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        ETWEETXLPOST.OffsetBox.Value = Token
        Exit Function
      
'//Set post for API send
        ElseIf InStr(1, Token, "app.sendapi(", vbTextCompare) Or InStr(1, Token, "sendapi(", vbTextCompare) Then
        Token = Replace(Token, "app.", vbNullString, , , vbTextCompare)
        Token = Replace(Token, "sendapi(", vbNullString, , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        xPos = Token
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.SendAPI_Clk", (xPos))
        Exit Function
        
'//Clear application window
        ElseIf InStr(1, Token, "app.clr(", vbTextCompare) Or InStr(1, Token, "clr(", vbTextCompare) Then
        Token = Replace(Token, "app.", vbNullString, , , vbTextCompare)
        Token = Replace(Token, "clr.setup(", vbNullString, , , vbTextCompare)
        Call modArtParens(Token): Call modArtQuotes(Token)
        
        '//special characters/wildcards
        If Token = "*" Then Call eTweetXL_CLICK.ClrSetupBtn_Clk: Exit Function      '//clear post/tweet setup & Linker
        '//switches
        If InStr(1, Token, "-draft", vbTextCompare) Then Call eTweetXL_CLICK.DraftHdr_Clk: Exit Function      '//clear draft link
        If InStr(1, Token, "-linker", vbTextCompare) Then Call eTweetXL_CLICK.LinkerHdr_Clk: Exit Function      '//clear Linker
        If InStr(1, Token, "-post", vbTextCompare) Then Call eTweetXL_CLICK.PostHdr_Clk: Exit Function      '//clear draft link
        If InStr(1, Token, "-runtime", vbTextCompare) Then Call eTweetXL_CLICK.RuntimeHdr_Clk: Exit Function     '//clear time link
        If InStr(1, Token, "-user", vbTextCompare) Then Call eTweetXL_CLICK.UserHdr_Clk: Exit Function      '//clear user link
        
        '//default no parameter
        Call eTweetXL_CLICK.ClrSetupBtn_Clk
        Exit Function
        
        
'//Reverse Linker data
        ElseIf InStr(1, Token, "app.rev(", vbTextCompare) Or InStr(1, Token, "rev(", vbTextCompare) Then
        
        Dim xRevArr(5000) As String

        '//Reverse drafts
        If InStr(1, Token, "-draft", vbTextCompare) Then
                
        Dim xLinkArr(5000) As String: Dim oLink As Object
        Set oLink = ETWEETXLPOST.LinkerBox
        
        For Each Item In oLink.List
        If Item <> "" Then
        xLinkArr(x1) = Item
        x1 = x1 + 1
        End If
        Next Item
        x1 = x1 - 1
        
        Do Until xLinkArr(x2) = vbNullString
        xRevArr(x2) = xLinkArr(x1)
        x1 = x1 - 1: x2 = x2 + 1
        Loop
        x1 = 0: oLink.Clear
        
        Do Until xRevArr(x1) = vbNullString
        oLink.AddItem (xRevArr(x1))
        x1 = x1 + 1
        Loop
        
        Set oLink = Nothing
        Exit Function
        
        '//Reverse runtimes
        ElseIf InStr(1, Token, "-time", vbTextCompare) Then
        
        Dim xTimeArr(5000) As String: Dim oTime As Object
        Set oTime = ETWEETXLPOST.RuntimeBox
        
        For Each Item In oTime.List
        If Item <> "" Then
        xTimeArr(x1) = Item
        x1 = x1 + 1
        End If
        Next Item
        x1 = x1 - 1
         
        Do Until xTimeArr(x2) = vbNullString
        xRevArr(x2) = xTimeArr(x1)
        x1 = x1 - 1: x2 = x2 + 1
        Loop
        x1 = 0: oTime.Clear
                    
        Do Until xRevArr(x1) = vbNullString
        oTime.AddItem (xRevArr(x1))
        x1 = x1 + 1
        Loop
        
        Set oTime = Nothing
        Exit Function
    
        '//Reverse users
        ElseIf InStr(1, Token, "-user", vbTextCompare) Then
        
        Dim xUserArr(5000) As String: Dim oUser As Object
        Set oUser = ETWEETXLPOST.UserBox
        
        For Each Item In oUser.List
        If Item <> "" Then
        xUserArr(x1) = Item
        x1 = x1 + 1
        End If
        Next Item
        x1 = x1 - 1
         
        Do Until xUserArr(x2) = vbNullString
        xRevArr(x2) = xUserArr(x1)
        x1 = x1 - 1: x2 = x2 + 1
        Loop
        x1 = 0: oUser.Clear
                    
        Do Until xRevArr(x1) = vbNullString
        oUser.AddItem (xRevArr(x1))
        x1 = x1 + 1
        Loop
        
        Set oUser = Nothing
        Exit Function
        
    End If
    Exit Function
    
'//Save current post
        ElseIf InStr(1, Token, "app.save.post(", vbTextCompare) Or InStr(1, Token, "save.post(", vbTextCompare) Then
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.SavePostBtn_Clk")
        Exit Function
        
'//Save linker state
        ElseIf InStr(1, Token, "app.save.linker(", vbTextCompare) Or InStr(1, Token, "save.linker(", vbTextCompare) Then
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.SaveLinkerBtn_Clk")
        Exit Function

'//Split post
        ElseIf InStr(1, Token, "app.split.post(", vbTextCompare) Or InStr(1, Token, "split.post(", vbTextCompare) Then
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.SplitPostBtn_Clk")
        Exit Function
        
'//Trim post
        ElseIf InStr(1, Token, "app.trim.post(", vbTextCompare) Or InStr(1, Token, "trim.post(", vbTextCompare) Then
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.TrimPostBtn_Clk")
        Exit Function
    
'//View instanced media
        ElseIf InStr(1, Token, "app.view.media(", vbTextCompare) Or InStr(1, Token, "view.media(", vbTextCompare) Then
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.ViewMedBtn_Clk")
        Exit Function
        
        End If '//ADA end
        
'//nothing found
errLvl = 1

ErrEnd:
'//Article not found...
If errLvl <> 0 Then Token = Token & "(*/ERR)"
Workbooks(appEnv).Worksheets(appBlk).Range("xlasErrRef").Value = """" & Token & """"

End Function
Private Function libFlag$(Token, errLvl As Byte)

'/\_____________________________________
 '//
'//         FLAGS
'/\_____________________________________

On Error GoTo ErrEnd

Call getEnvironment(appEnv, appBlk)

        '//Run script silently (ignore any application messages)
        If InStr(1, Token, "--enablesilent", vbTextCompare) Then
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasSilent") = 1
        errLvl = 0
        If Len(Token) <= 5 Then Token = 1: Exit Function
        End If
        
        If InStr(1, Token, "--disablesilent", vbTextCompare) Then
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasSilent") = 0
        errLvl = 0
        If Len(Token) <= 5 Then Token = 1: Exit Function
        End If
        
        '//Use lesser features during import/load action
        If InStr(1, Token, "--enableloadless", vbTextCompare) Then
        Workbooks(appEnv).Worksheets(appBlk).Range("LoadLess").Value = 1
        errLvl = 0
        If Len(Token) <= 5 Then Token = 1: Exit Function
        End If
        
        If InStr(1, Token, "--disableloadless", vbTextCompare) Then
        Workbooks(appEnv).Worksheets(appBlk).Range("LoadLess").Value = 0
        errLvl = 0
        If Len(Token) <= 5 Then Token = 1: Exit Function
        End If
        
Exit Function

ErrEnd:
'//flag not found...
Token = "(*/ERR)"

End Function
Private Function libSwitch$(Token, errLvl As Byte)

'/\_____________________________________
 '//
'//         LIBRARY SWITCHES
'/\_____________________________________


End Function
