Attribute VB_Name = "zCleanup"
'/####################\
'//Application Cleanup\\
'///##################\\\

Sub clnDraft()

lastR = Cells(Rows.Count, "L").End(xlUp).Row

'//Clean DraftLink area
Range("L1:L" & lastR).ClearContents

lastR = Cells(Rows.Count, "P").End(xlUp).Row

'//Clean ProfileLink area
Range("P1:P" & lastR).ClearContents


End Sub
Sub clnProf()

lastR = Cells(Rows.Count, "A").End(xlUp).Row

'//Clean profile load area
Range("A2:B" & lastR + 1000).ClearContents

End Sub
Sub clnMain()

lastR = Cells(Rows.Count, "A").End(xlUp).Row

'//Clean main area (everything except latches & code)
Range("A2:AY" & lastR + 1000).ClearContents

End Sub
Sub clnMediaScroll()

lastR = Cells(Rows.Count, "I").End(xlUp).Row

'//Clean media scroll area
Range("I1:I" & lastR + 1).ClearContents

End Sub
Sub clnLatch()

lastR = Cells(Rows.Count, "AZ").End(xlUp).Row

'//Clean application Latch area
Range("AZ1:AZ15, AZ17:AZ" & lastR).ClearContents

End Sub
Sub clnLinker()

'//Clean Linker area
lastR = Cells(Rows.Count, "M").End(xlUp).Row

Range("L2:R" & lastR + 1000).ClearContents

End Sub
Sub clnLinker2()

    Call delAppData
    Call clnLatch
    Call clnSpec

    ETWEETXLHOME.ProgRatio = ""
    ETWEETXLHOME.AppStatus.Caption = "OFF"
    ETWEETXLHOME.AppStatus.ForeColor = vbRed
    ETWEETXLHOME.AppStatus.BackColor = -2147483633
    
    ETWEETXLHOME.ProgBar.Width = 0
    ETWEETXLPOST.UserBox.Clear
    ETWEETXLPOST.LinkerBox.Clear
    ETWEETXLPOST.RuntimeBox.Clear
    ETWEETXLPOST.SendAPI.Value = False
    
    '//Cleanup Queue
    ETWEETXLPOST.UserHdr.Caption = "User"
    ETWEETXLPOST.DraftHdr.Caption = "Draft"
    ETWEETXLPOST.RuntimeHdr.Caption = "Runtime"
    ETWEETXLQUEUE.UserHdr.Caption = "User"
    ETWEETXLQUEUE.QueueHdr.Caption = "Queued"
    ETWEETXLQUEUE.RuntimeHdr.Caption = "Runtime"
    ETWEETXLQUEUE.CurrRuntime.Value = vbNullString
    ETWEETXLQUEUE.CurrQueue.Value = vbNullString
    ETWEETXLQUEUE.QueueBox.Clear
    ETWEETXLQUEUE.RuntimeBox.Clear
    ETWEETXLQUEUE.UserBox.Clear
    
    '//Update application state
    Call eTweetXL_TOOLS.updAppState
    Call eTweetXL_GET.getQueueData

    If Range("xlasSilent").Value2 <> "1" Then
    '//Set application inactive...
    xMsg = 11: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)
    End If
    
    '//unlock flowstrip
    Call enlFlowStrip
    '//unfreeze application
    Call dfsFreeze
    
    ETWEETXLHOME.xlFlowStrip.Value = "Automation complete..."
    ETWEETXLSETUP.xlFlowStrip.Value = "Automation complete..."
    ETWEETXLPOST.xlFlowStrip.Value = "Automation complete..."
    ETWEETXLQUEUE.xlFlowStrip.Value = "Automation complete..."
    
End Sub
Sub clnSpec()

lastR = Cells(Rows.Count, "AM").End(xlUp).Row

'//Clean special Linker area
Range("AL2:AM" & lastR + 1000).ClearContents

lastR = Cells(Rows.Count, "M").End(xlUp).Row

'//Clean MainLink
Range("M2:N" & lastR + 1000).ClearContents

End Sub
Sub clnTwt()

lastR = Cells(Rows.Count, "C").End(xlUp).Row

'//Clean imported tweets area
Range("C2:E" & lastR + 1000).ClearContents

End Sub
Sub clnThr()

lastR = Cells(Rows.Count, "Y").End(xlUp).Row
If lastR < 1 Then lastR = Cells(Rows.Count, "Z").End(xlUp).Row

'//Clean imported tweets area
Range("Y1:Z" & lastR + 1000).ClearContents

End Sub
Sub clnRuntime()

lastR = Cells(Rows.Count, "R").End(xlUp).Row

'//Clean runtime area
Range("R1:R" & lastR).ClearContents

End Sub
Sub clnUser()

lastR = Cells(Rows.Count, "AL").End(xlUp).Row

'//Clean apiLink area
Range("AL1:AL" & lastR).ClearContents

lastR = Cells(Rows.Count, "AM").End(xlUp).Row

'//Clean TargetLink area
Range("AM1:AM" & lastR).ClearContents

lastR = Cells(Rows.Count, "M").End(xlUp).Row

'//Clean MainLink area
Range("M1:M" & lastR).ClearContents

lastR = Cells(Rows.Count, "Q").End(xlUp).Row

'//Clean UserLink area
Range("Q1:Q" & lastR).ClearContents

End Sub
Sub clnOnClose()

'//Cleanup & reset application
Call clnMain
Call clnLatch
Call clnLinker
Call clnRuntime
Call clnSpec
Call eTweetXL_TOOLS.delAppData
Range("ConnectTrig").Value2 = 0
Range("LinkTrig").Value = 0
Range("User").Value = vbNullString
ETWEETXLHOME.xlFlowStrip.Enabled = True
ETWEETXLPOST.xlFlowStrip.Enabled = True
ETWEETXLQUEUE.xlFlowStrip.Enabled = True
ETWEETXLSETUP.xlFlowStrip.Enabled = True
ETWEETXLHOME.ProgRatio = ""
ETWEETXLHOME.ActiveUser.Caption = ""
ETWEETXLHOME.ProgBar.Width = 0
ETWEETXLPOST.SendAPI.Value = False
ETWEETXLPOST.ActiveUser.Caption = ""
ETWEETXLPOST.UserBox.Clear
ETWEETXLPOST.LinkerBox.Clear
ETWEETXLPOST.RuntimeBox.Clear
ETWEETXLPOST.ProfileListBox.Value = ""
ETWEETXLPOST.UserListBox.Value = ""
ETWEETXLPOST.DraftBox.Value = ""
ETWEETXLPOST.UserHdr.Caption = "User"
ETWEETXLPOST.DraftHdr.Caption = "Draft"
ETWEETXLPOST.RuntimeHdr.Caption = "Runtime"
ETWEETXLQUEUE.ActiveUser.Caption = ""
ETWEETXLQUEUE.QueueBox.Clear
ETWEETXLQUEUE.RuntimeBox.Clear
ETWEETXLQUEUE.UserBox.Clear
ETWEETXLSETUP.ActiveUser.Caption = ""

ETWEETXLHOME.AppStatus.Caption = "OFF"
ETWEETXLHOME.AppStatus.ForeColor = vbRed
ETWEETXLHOME.AppStatus.BackColor = -2147483633

End Sub
Sub clnOpenFiles()

'//Close list of potentially opened files

Close #1
Close #2
Close #3
Close #4
Close #5
Close #6
Close #7

End Sub

