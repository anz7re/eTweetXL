Attribute VB_Name = "App_CHANGE"
'/############################\
'//Application Change Features\\
'///##########################\\\

Sub PostBox_Chg()

If Range("xlasWinForm").Value > 10 Then Exit Sub

Call App_TOOLS.FindForm(xForm)

On Error GoTo CleanBox

Dim PostCharCt As Integer

'//Get post char count
PostCharCt = Len(xForm.PostBox.Value)
xForm.CharCt.Caption = PostCharCt
If Len(xForm.PostBox.Value) < 280 Then
If xForm.PostBox.ForeColor <> vbBlack Then xForm.PostBox.ForeColor = vbBlack
xForm.CharCt.ForeColor = vbBlack
xForm.CharCt.BackColor = vbWhite
    Else
    If xForm.PostBox.ForeColor <> vbRed Then xForm.PostBox.ForeColor = vbRed
    xForm.CharCt.ForeColor = vbRed
    xForm.CharCt.BackColor = -2147483633
        End If

    xForm.PostBox.Value = Replace(xForm.PostBox.Value, "{ENTER};", Chr(10))
    xForm.PostBox.Value = Replace(xForm.PostBox.Value, "{SPACE};", " ")
 
Exit Sub

CleanBox:
ETWEETXLPOST.PostBox.Value = ""
ETWEETXLQUEUE.PostBox.Value = ""
                
End Sub
Sub DraftBox_Chg()

lastRw = Cells(Rows.Count, "I").End(xlUp).Row

'//Clear Media scroll space...
Range("I1:I" & lastRw).ClearContents

'//Refresh Media scroll...
Range("MedScrollPos").Value = 0
'//Refresh Gif/vid counter...
Range("GifCntr").Value = 0
Range("VidCntr").Value = 0

ETWEETXLPOST.PostBox.Value = ""
ETWEETXLPOST.MedLinkBox.Value = ""
xTwt = ETWEETXLPOST.DraftBox.Value
Call App_IMPORT.SelectedTweet(xTwt)
Call App_IMPORT.SelectedMedia

If InStr(1, ETWEETXLPOST.xlFlowStrip.Value, "-negate", vbTextCompare) = False Then
ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.DraftBox.Value & " selected..."
End If
               
End Sub
Sub TimeBox_Chg()

xChar = ETWEETXLPOST.TimeBox.Value

If ETWEETXLPOST.TimeBox.Value <> vbNullString Then Call App_TOOLS.CheckForChar(xChar)

If xChar = "(*Err)" Then Exit Sub

If ETWEETXLPOST.TimeBox.Value = "" Then
TimeHldr = Time
'//Check time for military conversion...
If InStr(1, TimeHldr, "PM") Then
TimeArr = Split(TimeHldr, ":")
    If TimeArr(0) <> 12 Then
        TimeArr(0) = TimeArr(0) + 12
            End If
                TimeHldr = TimeArr(0) & ":" & TimeArr(1) & ":" & TimeArr(2)
                        End If
    '//Check for 12AM
    TimeArr = Split(TimeHldr, ":")
    If TimeArr(0) = 12 Then
    TimeArr(0) = 0
    TimeHldr = TimeArr(0) & ":" & TimeArr(1) & ":" & TimeArr(2)
    End If
            
                
TimeHldr = Replace(TimeHldr, "AM", ""): TimeHldr = Replace(TimeHldr, "PM", "")

TimeHldr = Format$(TimeHldr, "hh:mm:ss")
ETWEETXLPOST.TimeBox.Value = TimeHldr
End If

If InStr(1, ETWEETXLPOST.TimeBox.Value, " ") Then ETWEETXLPOST.TimeBox.Value = Replace(ETWEETXLPOST.TimeBox.Value, " ", "")

If Len(ETWEETXLPOST.TimeBox.Value) > 8 Then
ETWEETXLPOST.TimeBox.Value = ""
End If

End Sub
Sub OffsetBox_Chg()

xChar = ETWEETXLPOST.OffsetBox.Value: If xChar = "00:00:00" Then Exit Sub

Call App_TOOLS.CheckForChar2(xChar)
If xChar = "(*Err)" Then Exit Sub


If ETWEETXLPOST.OffsetBox.Value = "" Then
    ETWEETXLPOST.OffsetBox.Value = "00:00:00"
        End If
        
If InStr(1, ETWEETXLPOST.OffsetBox.Value, " ") Then ETWEETXLPOST.OffsetBox.Value = Replace(ETWEETXLPOST.OffsetBox.Value, " ", "0")

If Len(ETWEETXLPOST.OffsetBox.Value) > 8 Then
ETWEETXLPOST.OffsetBox.Value = "00:00:00"
End If

End Sub
