Attribute VB_Name = "eTweetXL_CHANGE"
'/############################\
'//Application Change Features\\
'///##########################\\\

Sub PostBox_Chg()

If Range("xlasWinForm").Value2 > 10 Then Exit Sub

Call fndWindow(xWin)

On Error GoTo CleanBox

Dim PostCharCt As Integer

'//Get post char count
PostCharCt = Len(xWin.PostBox.Value)
xWin.CharCt.Caption = PostCharCt
If Len(xWin.PostBox.Value) < 280 Then
If Range("xlasBlkAddr97").Value2 = vbNullString Then xWin.PostBox.ForeColor = vbBlack
xWin.CharCt.ForeColor = vbBlack
xWin.CharCt.BackColor = vbWhite
    Else
    If Range("xlasBlkAddr97").Value2 = vbNullString Then xWin.PostBox.ForeColor = vbRed
    xWin.CharCt.ForeColor = vbRed
    xWin.CharCt.BackColor = -2147483633
        End If

    xWin.PostBox.Value = Replace(xWin.PostBox.Value, "{ENTER};", Chr(10))
    xWin.PostBox.Value = Replace(xWin.PostBox.Value, "{SPACE};", " ")
 
Exit Sub

CleanBox:
ETWEETXLPOST.PostBox.Value = ""
ETWEETXLQUEUE.PostBox.Value = ""
                
End Sub
Sub DraftBox_Chg()

Dim X As Byte

'//Clear Media scroll space...
For X = 0 To 4
Range("MediaScroll").Offset(X, 0).ClearContents
Next
'//Refresh Media scroll...
Range("MedScrollPos").Value = 0
'//Refresh Gif/vid counter...
Range("GifCntr").Value = 0
Range("VidCntr").Value = 0

ETWEETXLPOST.PostBox.Value = vbNullString
ETWEETXLPOST.MedLinkBox.Value = vbNullString
xTwt = ETWEETXLPOST.DraftBox.Value
Call eTweetXL_GET.getSelPost(xTwt)
Call eTweetXL_GET.getSelMedia

If Range("xlasSilent").Value2 <> 1 Then
ETWEETXLPOST.xlFlowStrip.Value = ETWEETXLPOST.DraftBox.Value & " selected..."
End If
               
End Sub
Sub TimeBox_Chg()

xChar = ETWEETXLPOST.TimeBox.Value

If ETWEETXLPOST.TimeBox.Value <> vbNullString Then Call eTweetXL_TOOLS.fndChar(xChar)

If xChar = "(*Err)" Then Exit Sub

If ETWEETXLPOST.TimeBox.Value = vbNullString Then
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
            
                
TimeHldr = Replace(TimeHldr, "AM", vbNullString): TimeHldr = Replace(TimeHldr, "PM", vbNullString)

TimeHldr = Format$(TimeHldr, "hh:mm:ss")
ETWEETXLPOST.TimeBox.Value = TimeHldr
End If

If InStr(1, ETWEETXLPOST.TimeBox.Value, " ") Then _
ETWEETXLPOST.TimeBox.Value = Replace(ETWEETXLPOST.TimeBox.Value, " ", vbNullString)

If Len(ETWEETXLPOST.TimeBox.Value) > 8 Then
ETWEETXLPOST.TimeBox.Value = vbNullString
End If

End Sub
Sub OffsetBox_Chg()

xChar = ETWEETXLPOST.OffsetBox.Value: If xChar = "00:00:00" Then Exit Sub

Call eTweetXL_TOOLS.fndChar2(xChar)
If xChar = "(*Err)" Then Exit Sub


If ETWEETXLPOST.OffsetBox.Value = "" Then
    ETWEETXLPOST.OffsetBox.Value = "00:00:00"
        End If
        
If InStr(1, ETWEETXLPOST.OffsetBox.Value, " ") Then ETWEETXLPOST.OffsetBox.Value = Replace(ETWEETXLPOST.OffsetBox.Value, " ", "0")

If Len(ETWEETXLPOST.OffsetBox.Value) > 8 Then
ETWEETXLPOST.OffsetBox.Value = "00:00:00"
End If

End Sub
