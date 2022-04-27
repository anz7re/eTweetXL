Attribute VB_Name = "CLEANUP"
'/####################\
'//Application Cleanup\\
'///##################\\\

Sub ClnDraftSpace()

lastRw = Cells(Rows.Count, "L").End(xlUp).Row

'//Clean DraftLink area
Range("L1:L" & lastRw).ClearContents

lastRw = Cells(Rows.Count, "P").End(xlUp).Row

'//Clean ProfileLink area
Range("P1:P" & lastRw).ClearContents


End Sub
Sub ClnLinkerSpace()

'//Clean Linker area
lastRw = Cells(Rows.Count, "M").End(xlUp).Row

Range("L2:R" & lastRw + 1000).ClearContents

End Sub
Sub ClnMainSpace()

lastRw = Cells(Rows.Count, "A").End(xlUp).Row

'//Clean main area (everything except latch & code space)
Range("A2:AY" & lastRw + 1000).ClearContents

End Sub
Sub ClnMediaScroll()

lastRw = Cells(Rows.Count, "I").End(xlUp).Row

'//Clean media scroll area
Range("I1:I" & lastRw + 1).ClearContents

End Sub
Sub ClnSpecSpace()

lastRw = Cells(Rows.Count, "AM").End(xlUp).Row

'//Clean special Linker area
Range("AL2:AM" & lastRw + 1000).ClearContents

lastRw = Cells(Rows.Count, "M").End(xlUp).Row

'//Clean mainlink
Range("M2:M" & lastRw + 1000).ClearContents

End Sub
Sub ClnProfSpace()

lastRw = Cells(Rows.Count, "A").End(xlUp).Row

'//Clean profile load area
Range("A2:C" & lastRw + 1000).ClearContents

End Sub
Sub ClnTwtSpace()

lastRw = Cells(Rows.Count, "D").End(xlUp).Row

'//Clean imported tweets area
Range("D2:K" & lastRw + 1000).ClearContents

End Sub
Sub ClnThrSpace()

lastRw = Cells(Rows.Count, "Y").End(xlUp).Row
If lastRw < 1 Then lastRw = Cells(Rows.Count, "Z").End(xlUp).Row

'//Clean imported tweets area
Range("Y1:Z" & lastRw + 1000).ClearContents

End Sub
Sub ClnLatchSpace()

lastRw = Cells(Rows.Count, "AZ").End(xlUp).Row

'//Clean application Latch area
Range("AZ1:AZ" & lastRw).ClearContents

End Sub
Sub ClnRuntimeSpace()

lastRw = Cells(Rows.Count, "R").End(xlUp).Row

'//Clean runtime area
Range("R1:R" & lastRw).ClearContents

End Sub
Sub ClnUserSpace()

lastRw = Cells(Rows.Count, "AL").End(xlUp).Row

'//Clean apiLink area
Range("AL1:AL" & lastRw).ClearContents

lastRw = Cells(Rows.Count, "AM").End(xlUp).Row

'//Clean PassLink area
Range("AM1:AM" & lastRw).ClearContents

lastRw = Cells(Rows.Count, "M").End(xlUp).Row

'//Clean MainLink area
Range("M1:M" & lastRw).ClearContents

lastRw = Cells(Rows.Count, "Q").End(xlUp).Row

'//Clean UserLink area
Range("Q1:Q" & lastRw).ClearContents

End Sub
Sub ClnOnClose()

'//Cleanup & reset application
Call Cleanup.ClnMainSpace
Call Cleanup.ClnLatchSpace
Call Cleanup.ClnLinkerSpace
Call Cleanup.ClnRuntimeSpace
Call Cleanup.ClnSpecSpace
Call App_TOOLS.DataKillSwitch
Range("ConnectTrig").Value = 0
Range("LinkTrig").Value = 0
Range("User").Value = ""
ETWEETXLHOME.xlFlowStrip.Enabled = True
ETWEETXLPOST.xlFlowStrip.Enabled = True
ETWEETXLQUEUE.xlFlowStrip.Enabled = True
ETWEETXLSETUP.xlFlowStrip.Enabled = True
ETWEETXLHOME.ProgRatio = ""
ETWEETXLHOME.ActivePresetBox.Caption = ""
ETWEETXLHOME.ProgBar.Width = 0
ETWEETXLPOST.SendAPI.Value = False
ETWEETXLPOST.ActivePresetBox.Caption = ""
ETWEETXLPOST.UserBox.Clear
ETWEETXLPOST.LinkerBox.Clear
ETWEETXLPOST.RuntimeBox.Clear
ETWEETXLPOST.ProfileListBox.Value = ""
ETWEETXLPOST.UserListBox.Value = ""
ETWEETXLPOST.DraftBox.Value = ""
ETWEETXLPOST.UserHdr.Caption = "User"
ETWEETXLPOST.DraftHdr.Caption = "Draft"
ETWEETXLPOST.RuntimeHdr.Caption = "Runtime"
ETWEETXLQUEUE.ActivePresetBox.Caption = ""
ETWEETXLQUEUE.QueueBox.Clear
ETWEETXLQUEUE.RuntimeBox.Clear
ETWEETXLQUEUE.UserBox.Clear
ETWEETXLSETUP.ActivePresetBox.Caption = ""

ETWEETXLHOME.LinkerActive.Caption = "OFF"
ETWEETXLHOME.LinkerActive.ForeColor = vbRed
ETWEETXLHOME.LinkerActive.BackColor = -2147483633

End Sub
Sub CloseStrandedFiles()

'//Close list of potentially opened files

Close #1
Close #2
Close #3
Close #4
Close #5
Close #6
Close #7

End Sub

