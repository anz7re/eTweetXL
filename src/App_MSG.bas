Attribute VB_Name = "App_MSG"
'/###################################\
'//Application Error & Help Messages \\
'///#################################\\\

Function AppMsg(xMsg) As Integer

On Error Resume Next

Call CloseStrandedFiles

If Range("xlasSilent") <> 1 Then

Dim oFlowStrip As Object
nl = vbNewLine

'/1/xlFlowStrip syntax error
If xMsg = 1 Then
'//Check xlFlowStrip Window...
Call App_TOOLS.FindForm(xForm)
xForm.xlFlowStrip.ForeColor = vbRed
'//msg
MsgBox ("Syntax error"), vbExclamation, AppTag
Exit Function

'/2/File not found error
ElseIf xMsg = 2 Then
'//Check xlFlowStrip Window...
Call App_TOOLS.FindForm(xForm)

xForm.xlFlowStrip.Font.Color = vbRed
'//msg
MsgBox ("File not found"), vbExclamation, AppTag
Exit Function

'/3/Blank box error (information missing)
ElseIf xMsg = 3 Then
'//msg
MsgBox ("Information missing"), vbExclamation, AppTag
Exit Function

'/4/Invalid character
ElseIf xMsg = 4 Then
'//msg
MsgBox ("Invalid character entered"), vbExclamation, AppTag
Exit Function

'/5/Information not found
ElseIf xMsg = 5 Then
'//msg
MsgBox ("Couldn't find information for this user..."), vbInformation, AppTag
Exit Function

'/6/Missing connection before saving link
ElseIf xMsg = 6 Then
'//msg
MsgBox ("Connect your posts before saving!"), vbInformation, AppTag
Exit Function

'/7/API information not found for user
ElseIf xMsg = 7 Then
'//msg
MsgBox ("API information not found for this user"), vbCritical, AppTag
Exit Function

'/8/Invalid runtime
ElseIf xMsg = 8 Then
'//msg
MsgBox ("Invalid runtime entered"), vbExclamation, AppTag
Exit Function

'/9/No user set during run
ElseIf xMsg = 9 Then
'//msg
MsgBox ("A User hasn't been set!"), vbInformation, AppTag
Exit Function

'/10/Break
ElseIf xMsg = 10 Then
'//msg
MsgBox ("Break complete"), vbInformation, AppTag
Exit Function

'/11/Linker empty
ElseIf xMsg = 11 Then
'//msg
MsgBox ("Linker Emptied"), vbInformation, AppTag
Exit Function

'/12/Video is too large
ElseIf xMsg = 12 Then
'//msg
MsgBox ("This video is too large for Twitter."), vbInformation, AppTag
Exit Function

'/13/Gif is too large
ElseIf xMsg = 13 Then
'//msg
MsgBox ("This gif is too large for Twitter."), vbInformation, AppTag
Exit Function

'/14/Hit Twitter gif/video limit
ElseIf xMsg = 14 Then
'//msg
MsgBox ("Twitter only allows one gif/video per post."), vbInformation, AppTag
Exit Function

'/15/Hit Twitter Media limit
ElseIf xMsg = 15 Then
'//msg
MsgBox ("You've hit the media limit."), vbInformation, AppTag
Exit Function

'/16/Changes saved to post successfully
ElseIf xMsg = 16 Then
'//msg
MsgBox ("Changes saved"), vbInformation, AppTag
Exit Function

'/17/Changes saved to post successfully
ElseIf xMsg = 17 Then
'//msg
MsgBox ("Something's missing from the linker!"), vbInformation, AppTag
Exit Function

'/18/Username field empty
ElseIf xMsg = 18 Then
'//msg
MsgBox ("Username field empty"), vbInformation, AppTag
Exit Function

'/19/Password field empty
ElseIf xMsg = 19 Then
'//msg
MsgBox ("Password field empty"), vbInformation, AppTag
Exit Function

'/20/Password field empty
ElseIf xMsg = 20 Then
'//msg
MsgBox ("Profile field empty"), vbInformation, AppTag
Exit Function

'/21/Information not found
ElseIf xMsg = 21 Then
'//msg
MsgBox ("Information not found"), vbExclamation, AppTag
Exit Function

'/22/Exiting edit mode
ElseIf xMsg = 22 Then
'//msg
MsgBox ("EXITING EDIT MODE"), vbInformation, AppTag
Exit Function

'/23/Entering edit mode
ElseIf xMsg = 23 Then
'//msg
MsgBox ("EDIT MODE ACTIVE"), vbInformation, AppTag
Exit Function

'/24/Too many characters
ElseIf xMsg = 24 Then
'//msg
MsgBox ("This post has too many characters!"), vbInformation, AppTag
Exit Function

'/25/App currently running
ElseIf xMsg = 25 Then
'//msg
MsgBox ("The application is currently running..."), vbExclamation, AppTag
Exit Function

'/26/App currently frozen
ElseIf xMsg = 26 Then
'//msg
MsgBox ("The application is currently frozen..."), vbExclamation, AppTag
Exit Function

'/27/Error during start
ElseIf xMsg = 27 Then
'//msg
MsgBox ("There was an error starting the application." & nl & nl & "Please clear the Linker then retry." & nl & nl _
& "If the problem persists you may need to break &/or restart the application."), vbExclamation, AppTag
Exit Function

'/27/Help settings couldn't be changed
ElseIf xMsg = 28 Then
'//msg
MsgBox ("There was an error changing the help settings."), vbExclamation, AppTag
Exit Function

End If
    End If

End Function
Function HoverHelp(xMsg)

On Error Resume Next

If Range("xlasSilent").Value <> 1 And Range("HelpActive").Value <> 0 Then

'//check current hover position (exit if not @ same position to activate the help wizard)
If Range("HoverPos").Value <> xMsg Then Range("HoverActive").Value = 0: Range("HoverPos").Value = xMsg: Exit Function

nl = vbNewLine

Range("HoverActive").Value = Range("HoverActive").Value + 1

If Range("HoverActive").Value >= 20 Then

If xMsg = 1 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove ALL drafts from the Linker."
If xMsg = 2 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will add ALL drafts from this profile to the Linker."
If xMsg = 3 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will refresh the offset to " & """00:00:00""" & "."
If xMsg = 4 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will clear ALL text from the post box below."
If xMsg = 5 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove ALL times from the Linker."
If xMsg = 6 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will refresh the time to now."
If xMsg = 7 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove ALL users from the Linker."
If xMsg = 8 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will add a selected user to match the amount of drafts in the Linker."
If xMsg = 9 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove ALL currently added items from the Linker."
If xMsg = 10 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will freeze or unfreeze the application once its started."
If xMsg = 11 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will transport you to the home screen or if you're already there to the Queue."
If xMsg = 12 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will open " & """Control Box+""" & nl & nl & "Control Box+ is a Text Editor/IDE built for xlAppScript."
If xMsg = 13 Then ETWEETXLHELP.HelpMsgBox.Value = "Turning this on/active will assign users added to the Linker to be sent using the Twitter API."
If xMsg = 14 Then ETWEETXLHELP.HelpMsgBox.Value = "Turning this on/active will assign a random offset to each time added to the Linker."
If xMsg = 15 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will filter between your single [•] or threaded posts [...]"
If xMsg = 16 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove ALL drafts from the currently focused profile."
If xMsg = 17 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove the current draft from a profile."
If xMsg = 18 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will create a draft using the current name."
If xMsg = 19 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will extend the xlFlowStrip downwards."
If xMsg = 20 Then ETWEETXLHELP.HelpMsgBox.Value = "The current user set to send a post."
If xMsg = 21 Then ETWEETXLHELP.HelpMsgBox.Value = "Displays whether the application is running or not."
If xMsg = 22 Then ETWEETXLHELP.HelpMsgBox.Value = "Displays current run progression."
If xMsg = 23 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove the current profile from your archive."
If xMsg = 24 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove the current user from the focused profile."
If xMsg = 25 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will add the current profile to your archive."
If xMsg = 26 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will add the current user to the focused profile."
If xMsg = 27 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove ALL current profiles from your archive."
If xMsg = 28 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove ALL current users from the focused profile."
If xMsg = 29 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will allow you to add media to a post."
If xMsg = 30 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove currently focused media from a post."
If xMsg = 31 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will show the focused media."
If xMsg = 32 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will save the current post."
If xMsg = 33 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will add the current thread to a post."
If xMsg = 34 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove the current thread from a post."
If xMsg = 35 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove ALL threads from a post."
If xMsg = 36 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will arrange data sent to the Linker to be run."
If xMsg = 37 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will add the current user to the Linker."
If xMsg = 38 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove the last user from the Linker."
If xMsg = 39 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will add the current draft to the Linker."
If xMsg = 40 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove the last draft from the Linker."
If xMsg = 41 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will add a set time to the Linker."
If xMsg = 42 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will remove the last time from the Linker."
If xMsg = 43 Then ETWEETXLHELP.HelpMsgBox.Value = "Double-click or press Enter on a selected item to remove it from the Linker."
If xMsg = 44 Then ETWEETXLHELP.HelpMsgBox.Value = "Double-click or press Enter on a selected item to remove it from the Linker."
If xMsg = 45 Then ETWEETXLHELP.HelpMsgBox.Value = "Double-click on a selected time to change its value."
If xMsg = 46 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will allow you to convert & save your current Linker state into a link."
If xMsg = 47 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will allow you to import your links into the Linker."
If xMsg = 48 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will allow you to reload your last imported link."
If xMsg = 49 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will clean the ENTIRE Tweet Setup & Linker for a completely fresh environment."
If xMsg = 50 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will reload the last connected state from the Linker."
If xMsg = 51 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will force stop the application from ALL currently running automations & clean the environment."
If xMsg = 52 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will run the application once the Linker's connected."
If xMsg = 53 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will take you to the 'Queue' to manage running posts."
If xMsg = 54 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will take you to the 'Profile Setup' to edit profiles & user accounts."
If xMsg = 55 Then ETWEETXLHELP.HelpMsgBox.Value = "Clicking this will take you to the 'Tweet Setup' to manage drafts & links."

ETWEETXLHELP.Show

Range("HoverActive").Value = 0

End If
    End If
    
'//return current hover position
Range("HoverPos").Value = xMsg

End Function
