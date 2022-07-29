Attribute VB_Name = "eTweetXL_CHANGE"
'/############################\
'//Application Change Features\\
'///##########################\\\

Sub PostBox_Chg()

Dim oAppWin As Object
Dim PostCharCt As Integer

On Error GoTo CleanBox

Call getWindow(xWin)

Select Case Range("xlasWinForm").Value2
Case Is = 31: Set oAppWin = ETWEETXLPOST
Case Is = 41: Set oAppWin = ETWEETXLQUEUE
End Select

'//Get post char count
PostCharCt = Len(xWin.Value)

oAppWin.CharCt.Caption = PostCharCt

If Len(xWin.Value) <= 280 Then
If Range("xlasBlkAddr97").Value2 = vbNullString Then xWin.ForeColor = vbBlack
    oAppWin.CharCt.ForeColor = vbBlack
    oAppWin.CharCt.BackColor = vbWhite
    Else
    If Range("xlasBlkAddr97").Value2 = vbNullString Then xWin.ForeColor = vbRed
        oAppWin.CharCt.ForeColor = vbRed
        oAppWin.CharCt.BackColor = -2147483633
            End If

            xWin.Value = Replace(xWin.Value, "{ENTER};", Chr(10))
            xWin.Value = Replace(xWin.Value, "{SPACE};", " ")
 
Exit Sub

CleanBox:
ETWEETXLPOST.PostBox.Value = ""
ETWEETXLQUEUE.PostBox.Value = ""
                
End Sub
Sub DraftBox_Chg()

Dim x As Byte

'//Clear Media scroll space...
For x = 0 To 4
Range("MediaScroll").Offset(x, 0).ClearContents
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
