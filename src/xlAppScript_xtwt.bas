Attribute VB_Name = "xlAppScript_xtwt"
Public Function runLib$(xArt)
'/\____________________________________________________________________________________________________________________________
'//
'//       xtwt Library
'//         Version: 1.0.5
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
'//     Basic Lib Requirements: Windows 10, MS Excel Version 2107, PowerShell 5.1.19041.1023, eTweetXL v1.5.0
'//
'//                             (previous versions not tested &/or unsupported)
'/\____________________________________________________________________________________________________________________________
'//
'//     Latest Revision: 6/6/2022
'/\____________________________________________________________________________________________________________________________
'//
'//     Developer(s): anz7re
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
        Call modArtQ(xArt)
            
        '//Find application environment & block
        Call fndEnvironment(appEnv, appBlk)
        
        '//Find flags
        If InStr(1, xArt, "--") Or InStr(1, xArt, "++") Then _
        Call libFlag(xArt, errLvl): If xArt = 1 Then Exit Function Else _
        Call libSwitch(xArt, errLvl) '//Find switches
       
        '//Set library error level
        If Range("xlasLibErrLvl").Value2 = 0 Then On Error GoTo ErrMsg
        If Range("xlasLibErrLvl").Value2 = 1 Then On Error Resume Next
        
        '//Check for ADA Article
        If InStr(1, xArt, "app.") Then GoTo ADALink
                     
'/\_____________________________________
'//
'//         TWEET SETUP ARTICLES
'/\_____________________________________
'//
        
        '//Set Profile
        If InStr(1, xArt, "profile(", vbTextCompare) Then
        xArt = Replace(xArt, "profile(", vbNullString, , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        
        If InStr(1, xArt, "del.") Then GoTo DelProfile '//check for delete profile...
        If InStr(1, xArt, "mk.") Then GoTo NewProfile '//check for make new profile...
        
        Call fndWindow(xWin)
        xWin.ProfileListBox.Value = xArt
        Workbooks(appEnv).Worksheets(appBlk).Range("Profile").Value2 = xArt
        Exit Function
        
        '//Set User
        ElseIf InStr(1, xArt, "user(", vbTextCompare) Then
        
        '//make sure we're not performing a different user operation...
        If InStr(1, xArt, ".user(", vbTextCompare) Then
            GoTo UserLink
                End If

        xArt = Replace(xArt, "user(", vbNullString, , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        xUser = xArt
        Workbooks(appEnv).Worksheets(appBlk).Range("User").Value2 = xUser
        Call fndWindow(xWin)
        xWin.UserListBox.Value = xUser
        Call eTweetXL_CLICK.SetActive_Clk(xUser)
        Exit Function


        '//Add post info
        ElseIf InStr(1, xArt, "post(", vbTextCompare) Then
        xArt = Replace(xArt, "post(", vbNullString, , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        
        If xArt = vbNullString Then wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_Focus.shw_ETWEETXLPOST"): Exit Function      '//nothing entered show post setup
        
        '//special characters/wildcards
        If InStr(1, xArt, ",") Then
        xArtArr = Split(xArt, ",")
        '//switches
        If InStr(1, xArtArr(1), "-true", vbTextCompare) Or InStr(1, xArtArr(1), "1") Then S = "T"
        If InStr(1, xArtArr(1), "-false", vbTextCompare) Or InStr(1, xArtArr(1), "0") Then S = "F"
        Call fndWindow(xWin)
        xWin.PostBox.Value = xWin.PostBox.Value & xArtArr(0)
        If S = "T" Then Call eTweetXL_CLICK.SavePostBtn_Clk
        Exit Function
        End If
        
        Call fndWindow(xWin)
        xWin.PostBox.Value = xWin.PostBox.Value & xArt
        Exit Function
        
        '//Add draft name
        ElseIf InStr(1, xArt, "draft(", vbTextCompare) And InStr(1, xArt, ".draft(", vbTextCompare) = False Then
        xArt = Replace(xArt, "draft(", "", , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        ETWEETXLPOST.DraftBox.Value = xArt
        Exit Function
        
        '//Add/rmv media to post
        ElseIf InStr(1, xArt, "med(", vbTextCompare) Then
        xArt = Replace(xArt, "med(", "", , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        
        If InStr(1, xArt, "add.", vbTextCompare) Then M = "A" '//Check for add...
        If InStr(1, xArt, "rmv.", vbTextCompare) Then M = "R" '//Check for remove...
        '//special characters/wildcards
        If InStr(1, xArt, "*") Then M = M & "0" '//Check for all...
        If InStr(1, xArt, ",") Then M = M & "1" '//Check for and...
        
        xArt = Replace(xArt, "add.", "", , , vbTextCompare)
        xArt = Replace(xArt, "rmv.", "", , , vbTextCompare)
        xArt = Replace(xArt, "*", vbNullString)
         Call modArtP(xArt)
         
        '//check for multiple items
        If InStr(1, xArt, ",") Then
        xArtArr = Split(xArt, ",")
        xLast = UBound(xArtArr) - LBound(xArtArr)
        End If
        
        xMed = xArt
        
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
        xMed = xArtArr(X)
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
        xMed = xArtArr(X)
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.RmvPostMedBtn_Clk")
        Next
        Exit Function
        
RmvAllMedia:
        Do Until ETWEETXLPOST.MedLinkBox.Value = vbNullString
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.RmvPostMedBtn_Clk")
        Loop
        Exit Function
        
        '//Add/rmv thread
        ElseIf InStr(1, xArt, "thread(", vbTextCompare) Then
        xArt = Replace(xArt, "thread(", "", , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        
        If InStr(1, xArt, "add.", vbTextCompare) Then M = "A" '//Check for add...
        If InStr(1, xArt, "rmv.", vbTextCompare) Then M = "R" '//Check for remove...
        '//special characters/wildcards
        If InStr(1, xArt, "*") Then M = M & "0" '//Check for all...
        If InStr(1, xArt, ",") Then M = M & "1" '//Check for and...
        
        xArt = Replace(xArt, "add.", "", , , vbTextCompare)
        xArt = Replace(xArt, "rmv.", "", , , vbTextCompare)
        xArt = Replace(xArt, "*", vbNullString)
        Call modArtP(xArt)
        
        '//check for arguments...
        If InStr(1, xArt, ",") Then
        xArtArr = Split(xArt, ",")
        xLast = UBound(xArtArr) - LBound(xArtArr)
        
        '//add string to single thread
        If M = "1" Then
        Range("ThreadScrollPos").Value = xArtArr(0)
        Range("PostThread").Offset(xArtArr(0), 0).Value = xArtArr(1) '//2 arguments
        If xLast > 1 Then Range("MedThread").Offset(xArtArr(0), 0).Value = xArtArr(2) '//3 arguments
        Exit Function
        End If

        '//add string to all threads
        If M = "01" Then
        lRow = Cells(Rows.Count, "Y").End(xlUp).Row
        If lRow < 1 Then lRow = Cells(Rows.Count, "Z").End(xlUp).Row
        
        For X = 1 To lRow
        Range("PostThread").Offset(X, 0).Value2 = xArtArr(1) '//2 arguments
        If xLast > 1 Then Range("MedThread").Offset(X, 0).Value2 = xArtArr(2) '//3 arguments
        Next
        Exit Function
            End If
                End If
        
        If M = vbNullString And xArt = vbNullString Then GoTo AddThread
        If M = vbNullString And xArt <> vbNullString Then GoTo FindThread
        If M = "A" Then GoTo AddThread
        If M = "A1" Then GoTo AddMultiThread
        If M = "A01" Then GoTo AddMultiThread
        If M = "R" Then GoTo RmvThread
        If M = "R0" Then GoTo RmvAllThread
        If M = "R1" Then GoTo RmvMultiThread
        If M = "R01" Then GoTo RmvAllThread
        
        GoTo ErrMsg
        Exit Function
        
AddThread:
        '//add single thread
        If xArt <> vbNullString Then xLast = CInt(xArt): GoTo AddMultiThread
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.AddThreadBtn_Clk")
        Exit Function
        
AddMultiThread:
        '//add multi thread
        Call fndWindow(xWin)
        For X = 0 To xLast
        If xWin.PostBox.Value = vbNullString Then xWin.PostBox.Value = "thread" & X
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.AddThreadBtn_Clk")
        Next
        Exit Function

RmvThread:
        '//remove single focused thread
        If xArt <> vbNullString Then xLast = CInt(xArt): GoTo RmvMultiThread
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
        Call fndWindow(xWin)
        X = CInt(xArt)
        Range("ThreadScrollPos").Value = X
        xWin.ThreadCt.Caption = X
        xWin.PostBox.Value = Range("PostThread").Offset(X, 0).Value2  '//2 arguments
        If xLast > 1 Then xWin.MedLinkBox.Value = Range("MedThread").Offset(X, 0).Value2
        Exit Function
        
NewProfile:
xArt = Replace(xArt, "mk.", vbNullString)
Call modArtP(xArt)

'//add multiple users (no arguments)
'//switches
If InStr(1, xArt, "-list", vbTextCompare) Then
xArt = Replace(xArt, "-list", vbNullString, , , vbTextCompare)
xArtArr = Split(xArt, ",")
For X = 0 To UBound(xArtArr)
xInfo = xArtArr(X)
Call eTweetXL_CLICK.NewProfile_Clk(xInfo)
Next
Exit Function
End If

xInfo = xArt
Call eTweetXL_CLICK.NewProfile_Clk(xInfo)
Exit Function

DelProfile:
xArt = Replace(xArt, "del.", vbNullString)
Call modArtP(xArt)
If InStr(1, xArt, "-f", vbTextCompare) Then xArt = Replace(xArt, "-f", vbNullString, , , vbTextCompare): Range("xlasSilent").Value2 = 1 '//force deletion no prompt
If InStr(1, xArt, "*") Then GoTo DelAllProfile '//all switch

'//add multiple users (no arguments)
'//switches
If InStr(1, xArt, "-list", vbTextCompare) Then
xArt = Replace(xArt, "-list", vbNullString, , , vbTextCompare)
xArtArr = Split(xArt, ",")
For X = 0 To UBound(xArtArr)
xInfo = xArtArr(X)
Call eTweetXL_CLICK.RmvProfileBtn_Clk(xInfo)
Next
Exit Function
End If

xInfo = xArt
Call eTweetXL_CLICK.RmvProfileBtn_Clk(xInfo)
Exit Function

DelAllProfile:
Call eTweetXL_CLICK.RmvAllProfilesBtn_Clk
Exit Function


'/\_____________________________________
'//
'//     DRAFT ARTICLES
'/\_____________________________________
'//


        '//Add drafts to linker
        ElseIf InStr(1, xArt, "draft(", vbTextCompare) Then
        
        If InStr(1, xArt, "del.", vbTextCompare) Then M = "D" '//check for delete draft(s)...
        If InStr(1, xArt, "add.", vbTextCompare) Then M = "A" '//check for add draft(s)...
        If InStr(1, xArt, "rmv.", vbTextCompare) Then M = "R" '//check for remove draft(s)...
        '//special characters/wildcards
        If InStr(1, xArt, "*") Then M = M & "0" '//Check for all...
        If InStr(1, xArt, ",") Then M = M & "1" '//Check for and...
        If InStr(1, xArt, ":") Then M = M & "2" '//Check for through...
        
        xArt = Replace(xArt, "add.draft", "", , , vbTextCompare)
        xArt = Replace(xArt, "rmv.draft", "", , , vbTextCompare)
        xArt = Replace(xArt, "del.draft", "", , , vbTextCompare)
        xArt = Replace(xArt, "*", vbNullString)
        Call modArtP(xArt)
        
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

AddDraft:
xArt = Replace(xArt, """", vbNullString)
xPos = CDbl(xArt) '//position

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
        '//Remove specific draft from linker
        xArt = Replace(xArt, """", vbNullString)
        xPos = CDbl(xArt) '//position
        ETWEETXLPOST.LinkerBox.RemoveItem (xPos)
        Exit Function
        
RmvAllDraft:
        '//Remove all drafts from linker
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.DraftHdr_Clk")
        Exit Function
       
DeleteDraft:
'//Delete draft from archive
If xArt = vbNullString Then
wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.DelDraftBtn_Clk")
    Else
        ETWEETXLPOST.DraftBox.Value = xArt: wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.DelDraftBtn_Clk")
                End If
Exit Function
       
DeleteAllDraft:
'//Delete all drafts from archive
wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.DelAllDraftBtn_Clk")
Exit Function
       
'/\_____________________________________
'//
'//     USER ARTICLES
'/\_____________________________________
'//
        '//Add user to linker
        ElseIf InStr(1, xArt, "user(", vbTextCompare) Then
        
UserLink:
        xArt = Replace(xArt, "user(", vbNullString, , , vbTextCompare)
        If InStr(1, xArt, "mk.", vbTextCompare) Then GoTo NewUser '//check for make new user(s)...
        If InStr(1, xArt, "del.", vbTextCompare) Then GoTo DelUser '//check for delete user(s)...
        If InStr(1, xArt, "add.", vbTextCompare) Then M = "A" '//check for add user(s)...
        If InStr(1, xArt, "rmv.", vbTextCompare) Then M = "R" '//check for remove user(s)...
        '//special characters/wildcards
        If InStr(1, xArt, "*") Then M = M & "0" '//check for all...
        If InStr(1, xArt, ",") Then M = M & "1" '//check for and...
        If InStr(1, xArt, ":") Then M = M & "2" '//check for through...
        
        xArt = Replace(xArt, "add.", vbNullString, , , vbTextCompare)
        xArt = Replace(xArt, "rmv.", vbNullString, , , vbTextCompare)
        xArt = Replace(xArt, "*", vbNullString)
        Call modArtP(xArt)
        
        If M = "A" Then GoTo AddUser
        If M = "A0" Then GoTo AddAllUser
        If M = "A1" Then GoTo AddUser
        If M = "A2" Then GoTo AddUser
        If M = "A01" Then GoTo AddUser
        If M = "A02" Then GoTo AddUser
        If M = "R" Then GoTo RmvUser
        If M = "R0" Then GoTo RmvAllUser

AddUser:
Call modArtQ(xArt)
xPos = CDbl(xArt) '//position

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
        '//Remove specific user from Linker
        Call modArtQ(xArt)
        xPos = CDbl(xArt) '//position
        ETWEETXLPOST.UserBox.RemoveItem (xPos)
        Exit Function
        
RmvAllUser:
        '//Remove all users from linker
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.UserHdr_Clk")
        Exit Function
    
DelUser:
xArt = Replace(xArt, "del.", vbNullString)
Call modArtP(xArt)
'//special characters/wildcards
If InStr(1, xArt, "*") Then GoTo DelAllUser
'//switches
If InStr(1, xArt, "-f", vbTextCompare) Then xArt = Replace(xArt, "-f", vbNullString, , , vbTextCompare): Range("xlasSilent").Value2 = 1 '//force deletion no prompt
'//add multiple users (no arguments)
If InStr(1, xArt, "-list", vbTextCompare) Then
xArt = Replace(xArt, "-list", vbNullString, , , vbTextCompare)
xArtArr = Split(xArt, ",")
For X = 0 To UBound(xArtArr)
xInfo = xArtArr(X)
Call eTweetXL_CLICK.DelUserBtn_Clk(xInfo)
Next
Exit Function
End If

xInfo = xArt
Call eTweetXL_CLICK.DelUserBtn_Clk(xInfo)
Exit Function

DelAllUser:
Call eTweetXL_CLICK.DelAllUsersBtn_Clk
Exit Function

    
NewUser:
xArt = Replace(xArt, "mk.", "", , , vbTextCompare)
Call modArtP(xArt)

'//add multiple users (no arguments)
'//switches
If InStr(1, xArt, "-list", vbTextCompare) Then
xArt = Replace(xArt, "-list", vbNullString, , , vbTextCompare)
xArtArr = Split(xArt, ",")
For X = 0 To UBound(xArtArr)
xInfo = xArtArr(X)
Call eTweetXL_CLICK.NewUser_Clk(xInfo)
Next
Exit Function
End If

'//add single user w/ or w/o parameter(s)
    xInfo = xArt
    Call eTweetXL_CLICK.NewUser_Clk(xInfo)
    Exit Function

    
        
'/\_____________________________________
'//
'//     TIME ARTICLES
'/\_____________________________________
'//
        '//Add runtime to Linker
        ElseIf InStr(1, xArt, "time(", vbTextCompare) Then
            
            If InStr(1, xArt, "add.", vbTextCompare) Then M = "A" '//check for add time...
            If InStr(1, xArt, "rmv.", vbTextCompare) Then M = "R" '//check for remove time...
            '//special characters/wildcards
            If InStr(1, xArt, "*") Then M = M & "0" '//check for all...
            If InStr(1, xArt, ",") Then M = M & "1" '//check for and...
            If InStr(1, xArt, ":") Then M = M & "2" '//check for through...
            
            xArt = Replace(xArt, "add.time", "", , , vbTextCompare)
            xArt = Replace(xArt, "rmv.time", "", , , vbTextCompare)
            xArt = Replace(xArt, "*", vbNullString)
            Call modArtP(xArt)
        
            If InStr(1, M, "A") Or InStr(1, M, "R") = False Then GoTo AddTimeDirect
            
            If InStr(1, xArt, "h", vbTextCompare) Then
            xArtArr = Split(xArt, "h", , vbTextCompare)
            Offset = "H"
            xCount = xArtArr(0)
                End If
            If InStr(1, xArt, "m", vbTextCompare) Then
            xArtArr = Split(xArt, "m", , vbTextCompare)
            Offset = "M"
            xCount = xArtArr(0)
                End If
            If InStr(1, xArt, "s", vbTextCompare) Then
            xArtArr = Split(xArt, "s", , vbTextCompare)
            Offset = "S"
            xCount = xArtArr(0)
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
Call modArtQ(xArt)
xPos = 0
ETWEETXLPOST.TimeBox.Value = xArt
Call eTweetXL_CLICK.AddRuntimeBtn_Clk(xPos)
Exit Function

AddTimeOnce:
Call modArtQ(xArt)
xPos = CDbl(xArt) '//position

        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.TimerHdr_Clk") '//refresh current time
        
        Call eTweetXL_CLICK.AddRuntimeBtn_Clk(xPos) '//add runtime
        
        Exit Function


AddTime:
Call modArtQ(xArt)
If xArtArr(0) = vbNullString Then xPos = CDbl(xArt) Else xPos = xArtArr(0) '//position
 
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
        Call modArtQ(xArt)
        xPos = CDbl(xArt) '//position
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
ElseIf InStr(1, xArt, "me()", vbTextCompare) And Len(xArt) <= 4 Then MsgBox (Range("xlasWinForm").Value2): Exit Function

    '//Set window number...
        ElseIf InStr(1, xArt, "winform(", vbTextCompare) Then
        
    '//switches
        If InStr(1, xArt, "-last", vbTextCompare) Then _
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = _
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinFormLast").Value2: Exit Function       '//set to last window
        
        xArt = Replace(xArt, "winform(", "", , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        
        Call xlAppScript_lex.fndChar(xArt)
        If xArt = "(*Err)" Then Exit Function
        
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasWinForm").Value2 = xArt
        
        Exit Function
        
'//Show application windows...
ElseIf InStr(1, xArt, "sh.home", vbTextCompare) Or InStr(1, xArt, "home()", vbTextCompare) Then ETWEETXLHOME.Show: Exit Function
ElseIf InStr(1, xArt, "sh.setup", vbTextCompare) Or xArt = "setup()" Then ETWEETXLSETUP.Show: Exit Function
ElseIf InStr(1, xArt, "sh.post", vbTextCompare) Or xArt = "post()" Then ETWEETXLPOST.Show: Exit Function
ElseIf InStr(1, xArt, "sh.queue", vbTextCompare) Or xArt = "queue()" Then ETWEETXLQUEUE.Show: Exit Function
ElseIf InStr(1, xArt, "sh.apisetup", vbTextCompare) Or xArt = "apisetup()" Then ETWEETXLAPISETUP.Show: Exit Function
ElseIf InStr(1, xArt, "sh.me", vbTextCompare) Then Call fndWindow(xWin): xWin.Show: Exit Function
'//Hide application windows...
ElseIf InStr(1, xArt, "hd.home", vbTextCompare) Then ETWEETXLHOME.Hide: Exit Function
ElseIf InStr(1, xArt, "hd.setup", vbTextCompare) Then ETWEETXLSETUP.Hide: Exit Function
ElseIf InStr(1, xArt, "hd.post", vbTextCompare) Then ETWEETXLPOST.Hide: Exit Function
ElseIf InStr(1, xArt, "hd.queue", vbTextCompare) Then ETWEETXLQUEUE.Hide: Exit Function
ElseIf InStr(1, xArt, "hd.apisetup", vbTextCompare) Then ETWEETXLAPISETUP.Hide: Exit Function
ElseIf InStr(1, xArt, "hd.me", vbTextCompare) Then Call fndWindow(xWin): xWin.Hide: Exit Function
Exit Function

End If '//end

'/\_____________________________________________
'//
'//     ADA ARTICLES (APPLICATION DIRECT ACTION)
'/\_____________________________________________
'//

ADALink:

    '//Start application
        If InStr(1, xArt, "start(", vbTextCompare) Then
        If InStr(1, xArt, "app.", vbTextCompare) Or InStr(1, xArt, "start()", vbTextCompare) Then
        wbMacro = "eTweetXL_CLICK.StartBtn_Clk"
        Workbooks(appEnv).Application.Run "'" & appEnv & "'!" & wbMacro: End If
        Exit Function
        
    '//Break application
        ElseIf InStr(1, xArt, "break(", vbTextCompare) Then
        If InStr(1, xArt, "app.", vbTextCompare) Or InStr(1, xArt, "break()", vbTextCompare) Then
        wbMacro = "eTweetXL_CLICK.BreakBtn_Clk"
        Workbooks(appEnv).Application.Run "'" & appEnv & "'!" & wbMacro: End If
        Exit Function
        
    '//Set draft filter (single/threaded)
        ElseIf InStr(1, xArt, "app.dfilter(", vbTextCompare) Or InStr(1, xArt, "dfilter(", vbTextCompare) Then
        xArt = Replace(xArt, "app.", vbNullString, , , vbTextCompare)
        xArt = Replace(xArt, "dfilter(", vbNullString, , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        xFil = CByte(xArt)
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.DraftFilterBtn_Clk", (xFil))
        Exit Function
        
    '//Set dynamic offset
        ElseIf InStr(1, xArt, "app.dynoffset(", vbTextCompare) Or InStr(1, xArt, "dynoffset(", vbTextCompare) Then
        xArt = Replace(xArt, "app.", vbNullString, , , vbTextCompare)
        xArt = Replace(xArt, "dynoffset(", vbNullString, , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        xPos = xArt
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.DynOffset_Clk", (xPos))
        Exit Function
        
    '//Freeze/unfreeze application
        ElseIf InStr(1, xArt, "app.freeze(", vbTextCompare) Or InStr(1, xArt, "freeze(", vbTextCompare) Then
        xArt = Replace(xArt, "app.", vbNullString, , , vbTextCompare)
        xArt = Replace(xArt, "freeze(", vbNullString, , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        If xArt = 0 Then Range("AppState").Value2 = 1
        If xArt = 1 Then Range("AppState").Value2 = 2
        wbMacro = "eTweetXL_CLICK.FreezeBtn_Clk"
        Workbooks(appEnv).Application.Run "'" & appEnv & "'!" & wbMacro
        Exit Function
        
    '//Set application help wizard on/off
        ElseIf InStr(1, xArt, "app.help(", vbTextCompare) Or InStr(1, xArt, "help(", vbTextCompare) Then
        xArt = Replace(xArt, "app.", vbNullString, , , vbTextCompare)
        xArt = Replace(xArt, "help(", vbNullString, , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        If xArt = 0 Then Range("HelpStatus").Value2 = 0
        If xArt = 1 Then Range("HelpStatus").Value2 = 1
        xPos = xArt
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.HelpStatusBtn_Clk", (xPos))
        Exit Function
    
    '//Hide application
        ElseIf InStr(1, xArt, "app.hide(", vbTextCompare) Or InStr(1, xArt, "hide(", vbTextCompare) Then
        Call eTweetXL_CLICK.HideBtn_Clk
        Exit Function
        
    '//Load a designated link
        ElseIf InStr(1, xArt, "app.load.linker(", vbTextCompare) Or InStr(1, xArt, "load.linker(", vbTextCompare) Then
        
        '//switches
        If InStr(1, xArt, "-last", vbTextCompare) Then
        '//Reload last connected link
        xLink = AppLoc & "\mtsett\lastlink.tmp"
        Call eTweetXL_GET.getLink(xLink)
        Exit Function
        End If
        
        xArt = Replace(xArt, "app.load.linker(", vbNullString, , , vbTextCompare)
        xArt = Replace(xArt, "app.", vbNullString, , , vbTextCompare)
        xArt = Replace(xArt, "load.linker(", vbNullString, , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        If InStr(1, xArt, ",") Then GoTo LoadMultiLink
        xLink = xArt
        
        If xLink <> vbNullString Then
        Call eTweetXL_GET.getLink(xLink)
        Else
        '//Reload last imported link
        xLink = Range("RemLink").Value
        Call eTweetXL_GET.getLink(xLink)
        End If
        Exit Function

LoadMultiLink:
        xArtArr = Split(xArt, ",")
        
        For X = 0 To UBound(xArtArr)
        xLink = xArtArr(X)
        Call eTweetXL_GET.getLink(xLink)
        Next
        Exit Function

    '//Add offset
        ElseIf InStr(1, xArt, "app.offset(", vbTextCompare) Or InStr(1, xArt, "offset(", vbTextCompare) Then
        xArt = Replace(xArt, "app.(", vbNullString, , , vbTextCompare)
        xArt = Replace(xArt, "offset(", vbNullString, , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        ETWEETXLPOST.OffsetBox.Value = xArt
        Exit Function
      
    '//Set post for API send
        ElseIf InStr(1, xArt, "app.sendapi(", vbTextCompare) Or InStr(1, xArt, "sendapi(", vbTextCompare) Then
        xArt = Replace(xArt, "app.", vbNullString, , , vbTextCompare)
        xArt = Replace(xArt, "sendapi(", vbNullString, , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        xPos = xArt
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.SendAPI_Clk", (xPos))
        Exit Function
        
    '//Connect Linker
        ElseIf InStr(1, xArt, "app.connect(", vbTextCompare) Or InStr(1, xArt, "connect()", vbTextCompare) Then
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.ConnectBtn_Clk")
        Exit Function
        
    '//Clear application window
        ElseIf InStr(1, xArt, "app.clr(", vbTextCompare) Or InStr(1, xArt, "clr(", vbTextCompare) Then
        xArt = Replace(xArt, "app.", vbNullString, , , vbTextCompare)
        xArt = Replace(xArt, "clr.setup(", vbNullString, , , vbTextCompare)
        Call modArtP(xArt): Call modArtQ(xArt)
        
        '//special characters/wildcards
        If xArt = "*" Then Call eTweetXL_CLICK.ClrSetupBtn_Clk: Exit Function      '//clear post/tweet setup & Linker
        '//switches
        If InStr(1, xArt, "-draft", vbTextCompare) Then Call eTweetXL_CLICK.DraftHdr_Clk: Exit Function      '//clear draft link
        If InStr(1, xArt, "-linker", vbTextCompare) Then Call eTweetXL_CLICK.LinkerHdr_Clk: Exit Function      '//clear Linker
        If InStr(1, xArt, "-post", vbTextCompare) Then Call eTweetXL_CLICK.PostHdr_Clk: Exit Function      '//clear draft link
        If InStr(1, xArt, "-runtime", vbTextCompare) Then Call eTweetXL_CLICK.RuntimeHdr_Clk: Exit Function      '//clear time link
        If InStr(1, xArt, "-user", vbTextCompare) Then Call eTweetXL_CLICK.UserHdr_Clk: Exit Function      '//clear user link
        
        '//default no parameter
        Call eTweetXL_CLICK.ClrSetupBtn_Clk
        Exit Function
        
        
    '//Reverse Linker data
        ElseIf InStr(1, xArt, "app.rev(", vbTextCompare) Or InStr(1, xArt, "rev(", vbTextCompare) Then
        
        Dim xRevArr(5000) As String

        '//Reverse drafts
        If InStr(1, xArt, "-draft", vbTextCompare) Then
                
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
        ElseIf InStr(1, xArt, "-time", vbTextCompare) Then
        
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
        ElseIf InStr(1, xArt, "-user", vbTextCompare) Then
        
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
        ElseIf InStr(1, xArt, "app.save.post", vbTextCompare) Or InStr(1, xArt, "save.post", vbTextCompare) Then
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.SavePostBtn_Clk")
        Exit Function
        
    '//Save linker state
        ElseIf InStr(1, xArt, "app.save.linker", vbTextCompare) Or InStr(1, xArt, "save.linker", vbTextCompare) Then
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.SaveLinkerBtn_Clk")
        Exit Function
        
    '//View instanced media
        ElseIf InStr(1, xArt, "app.view.media", vbTextCompare) Or InStr(1, xArt, "view.media", vbTextCompare) Then
        wbMacro = Workbooks(appEnv).Application.Run("eTweetXL_CLICK.ViewMedBtn_Clk")
        Exit Function
        
        End If '//ADA end
        
'//nothing found
errLvl = 1

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

On Error GoTo ErrMsg

Call fndEnvironment(appEnv, appBlk)

        '//Run script silently (ignore any application messages)
        If InStr(1, xArt, "--s", vbTextCompare) Then
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasSilent") = 1
        errLvl = 0
        If Len(xArt) <= 5 Then xArt = 1: Exit Function
        End If
        
        If InStr(1, xArt, "++s", vbTextCompare) Then
        Workbooks(appEnv).Worksheets(appBlk).Range("xlasSilent") = 0
        errLvl = 0
        If Len(xArt) <= 5 Then xArt = 1: Exit Function
        End If
        
        '//Use lesser features during import/load action
        If InStr(1, xArt, "--l", vbTextCompare) Then
        Workbooks(appEnv).Worksheets(appBlk).Range("LoadLess").Value = 1
        errLvl = 0
        If Len(xArt) <= 5 Then xArt = 1: Exit Function
        End If
        
        If InStr(1, xArt, "++l", vbTextCompare) Then
        Workbooks(appEnv).Worksheets(appBlk).Range("LoadLess").Value = 0
        errLvl = 0
        If Len(xArt) <= 5 Then xArt = 1: Exit Function
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






End Function
'//=========================================================================================================================
'//
'//         CHANGE LOG
'/\_________________________________________________________________________________________________________________________
'
'
' Version: 1.0.5
'
'
' [ Date 6/6/2022 ]
'
' (1): Added "app.hide()" article for hiding the entire application
'
' (2): Fixed an issue where "ADA" articles couldn't be found b/c of there placement in the library
'
'
' [ Date 5/12/2022 ]
'
' (1): Associating names/syntax for xlAppScript parameters are now ".param" & modifiers "mod."
'
' Changes to variables/titling were made to reflect. Various application functions were also renamed.
'
'
' [ Date 5/9/2022 ]
'
' (1): Fixed issue w/ "errLvl" variable causing error b/c of mismatch in data types
'
'
' [ Date 5/5/2022 ]
'
' (1): Various library optimizations (mainly dealing w/ variables)
'
'
' Version: 1.0.4
'
'
'
' [ Date 3/27/2022 ]
'
' (1): Fixed an issue w/ flags not parsing b/c of a missing call to get the Runtime Environment, & Block
'
'
' [ Date 3/4/2022 ]
'
' (1): Changed "-rtime" switch to "-runtime" for clarity
'
'
'
'
' Version: 1.0.3
'
'
' [ Date 2/28/2022 ]
'
' (1): Updated library to utilize article cleaning function within lexer
'
' (2): Included additional updates made to lexer as well as addition of the "runtime block"
'
' Version: 1.0.2
'
' (1): Removed "app.dptrig" article. Initially was needed for performing actions after a profile's data was loaded.
'
' No longer needed due to recent bug fixes/changes.
'
' (2): Added "app.dfilter" article to set draft filter to single/threaded posts.
'
' Example: dfilter(0) = single | dfilter(1) = threaded
'
'
'
' [ Date 2/10/2022 ]
'
' (1): Changed "app.ptrig()" article to "app.dptrig()" for clarity. It still performs the same action of manually setting
' a pre/return code when pulling in specific application data.
'
' dp = Data Pull
'
'
' [ Date 2/8/2022 ]
'
' (1): Changed labeling for "SHOW WINDOW" & "HIDE WINDOW" articles to broader "WINFORM ARTICLES"
'
' (2): Removed "-re" switch from "load.linker()" article. Instead will default to reload if left empty.
'
' Example: load.linker() <---
'
' (3): Added "winform()" & "me()" articles back to xtwt library.
'
' (4): Added "-true" & "-false" switches for boolean operations and parameter values.
'
' (5): Added (,) switch to post() article. Second parameter decides if the post is saved or not using boolean a check.
'
' Example: post(insert text for your post here, -true)   <--- this would save your post to the current focused draft w/
' the text in the first parameter.
'
'
'
' [ Date 2/7/2022 ]
'
' (1): Added "app.dynoffset()" article for activating/deactivating the "Dynamic Offset" option.
'
' Examples: dynoffset(0) = Inactive | dynoffset(1) = Active
'
'
' (2): Added "app.media.show()" article for viewing currently instanced media from either the post or queue window.
'
' ***can be shortened to ---> show.media
'
' (3): Added "app.freeze()" article for pausing/unpausing remaining application automations.
'
'***If used after starting a run, the next automation(s) afterwards will be halted until unfrozen.
' The user will need to trigger a start again after the applications been unfrozen to resume the current run.
'
' Examples: freeze(0) = Unpaused | freeze(1) = Paused
'
'
'
' [ Date 2/6/2022 ]
'
' (1): "profile()" & "user()" articles can now be used across "Profile Setup" & "Tweet Setup" windows
'
' (2): Added "del.profile()" & "del.user()" as well as corresponding "mk.profile()" & "mk.user()" articles to deal w/
' creating & removing profiles/users from an archive.
'
' ***del. parameter supports (*) wildcard for removing all items
' ***Both del./mk. parameters support (,) character for listing items
'
' (3): Changed "SHORT COMMANDS" to "DIRECT ACTION" short for "APPLICATION DIRECT ACTION" or "ADA" to organize articles by "app." use/prefix.
' This may slightly speed up parsing but mainly changed these articles for readabilities sake.
'
' Added "-last" & "-re" switches to "load.linker()" article.
'
' *** -last = reload last connection
'
'
'
'
'
' [ Date 2/5/2022 ]
'
' (1): Created "add.thread()" & "rmv.thread() articles to deal w/ adding & removing threads (supports "*" wildcard for removing all threads)
'
' (2): Added "clr.post" & "clr.linker" for clearing post box & linker
'
' (3): Added "clr.setup"
'
'
' (4): Added "set.ptrig()" article for added control when switching through profiles & needing different users.
'
' ***The profile trigger helps stop instances of code from importing data multiple times during changes to certain window (namely profile/user changes)
'
' set.ptrig(0) = Inactive | set.ptrig(1) = Active
'
'
'
' (5): Updated corresponding setup commands for clarity:
'
' savepost = save.post | savelinker = save.linker | reload = re.linker | load() = load.linker()
'
'
'
' (6): Added "set.sendapi()" article for assigning posts for default or api send
'
' Example(s): set.sendapi(0) = default | set.sendapi(1) = send w/ api
'
'
'
' [ Date 2/4/2022 ]
'
' (1): Removed "form()" article from xtwt library to "xbas" library as it's become a much broader command
'
'
'
' Version: 1.0.1
'
' [ Date: 1/2/2022 ]
'
'(1): Added change log, library requirements, & license information. Edited library description.
'
'(2): Added "LoadLess" functionality which if set will ignore certain loading features the application would
'normally perform when pulling in data to a UserForm window (not capatible w/ eTweetXL versions prior to v1.4.1)
'
'Set "LoadLess" w/ "--l" switch.
'
'Set back to normal w/ "++l" switch.
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'

