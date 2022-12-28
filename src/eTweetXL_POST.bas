Attribute VB_Name = "eTweetXL_POST"
'/############################\
'//Application Post Features  \\
'///##########################\\\

Public Sub pstLastLink()

'//For creating a backup .link file (connection state)
Dim X As Integer
X = 1

If Range("xlasSilent").Value2 <> 1 Then
'//Update connection state (xlFlowStrip)...
ETWEETXLPOST.xlFlowStrip.Value = "Saving backup link..."
xArt = "<lib>xbas;delayevent(5);$": Call lexKey(xArt)
    End If
        
lastR = Cells(Rows.Count, "L").End(xlUp).Row

'//Format sheet
Call eTweetXL_TOOLS.shtFormat

'//Select file save location
xPath = AppLoc & "\mtsett\lastlink.link"

Open xPath For Output As #1

Do Until X = lastR

Print #1, Range("MainLink").Offset(X, 0).Value2 & "," _
& Range("UserLink").Offset(X, 0).Value2 _
& Range("apiLink").Offset(X, 0).Value2 & "," _
& Range("ProfileLink").Offset(X, 0).Value2 & "," _
& Range("DraftLink").Offset(X, 0).Value & "," _
& Format(Range("Runtime").Offset(X, 0).Value2, "hh:mm:ss")
X = X + 1
Loop

Close #1

End Sub
Sub pstLastQueue()

'//For creating a backup .link file (queue state)
If Range("AppState").Value2 = 1 Then
xName = "lastqueue.link": xPath = AppLoc & "\mtsett\" & xName
Call eTweetXL_CLICK.SaveQueueBtn_Clk(xName, xPath)
End If

End Sub
Sub pstDraftData(ByVal xTwt As String)

On Error GoTo EndMacro

Dim ProfDbArr(100) As String
Dim draftFile As String: Dim theadFile As String: Dim xFile As String: Dim xStr As String
Dim oFSO As Object: Dim oFile As Object: Dim oFldr As Object

Call getWindow(xWin)
Call eTweetXL_LOC.xThrFile(thrFile)
Call eTweetXL_LOC.xTwtFile(twtFile)

If InStr(1, xTwt, " [...]") = False Then
If xWin.DraftFilterBtn.Caption <> "..." Or InStr(1, xTwt, " [•]") Then
xExt = ".twt": xT = " [•]"
xFile = twtFile
    Else: GoTo asThread
            End If
        Else
asThread:
            xExt = ".thr": xT = " [...]"
            xFile = thrFile
                End If
                    
 
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFldr = oFSO.GetFolder(xFile)

gNum = 1
'//Get profile names
For Each oFile In oFldr.Files
gNum = gNum + 1
Next oFile

'//Save post from tweet setup screen
If Range("xlasWinForm").Value2 = 13 Or Range("xlasWinForm").Value2 = 16 Then

'//check for thread
If Range("PostThread").Offset(1, 0).Value2 <> vbNullString Then GoTo postAsThread
            
If ETWEETXLPOST.DraftBox.Value = vbNullString Then
draftFile = twtFile & "draft_" & gNum & "_" & getDate & ".twt"
    Else
        draftFile = twtFile & ETWEETXLPOST.DraftBox.Value & ".twt"
            End If
            
            xStr = Replace(ETWEETXLPOST.PostBox.Value, Chr(10), "{ENTER};")
            xStr = Replace(xStr, " ", "{SPACE};")
            
            Open draftFile For Output As #1
            Print #1, xStr
            Print #1, "*-;"
            Print #1, "*-" & ETWEETXLPOST.MedLinkBox.Value
            Close #1
                    Exit Sub
                    
                
'//Save thread from tweet setup screen
postAsThread:

lastR = Cells(Rows.Count, "Y").End(xlUp).Row
If lastR < 1 Then lastR = Cells(Rows.Count, "Z").End(xlUp).Row

If ETWEETXLPOST.DraftBox.Value = vbNullString Then
threadFile = thrFile & "draft_" & gNum & "_" & getDate & ".thr"
    Else
        threadFile = thrFile & ETWEETXLPOST.DraftBox.Value & ".thr"
            End If

            Open threadFile For Output As #2
            
            For X = 1 To lastR - 1
            
            xStr = Replace(Range("PostThread").Offset(X, 0).Value2, Chr(10), "{ENTER};")
            xStr = Replace(xStr, " ", "{SPACE};")
            xMedArr = Split(Range("MedThread").Offset(X, 0).Value2, """ """)
            
            xMed = vbNullString
            For M = 0 To UBound(xMedArr)
            xMedHldr = xMedArr(M)
            xMedHldr = Replace(xMedHldr, """", vbNullString)
            If M >= 1 Then xMedHldr = " " & """" & xMedHldr & """" Else xMedHldr = """" & xMedHldr & """"
            xMed = xMed & xMedHldr
            Next
            
            Print #2, xStr
            Print #2, "*-;"
            Print #2, "*-" & xMed
            Print #2, "*-(" & X & ");"
            Next
            
            Close #2
            
                Exit Sub
                    
                    End If

'//Save single post from queue screen
If ETWEETXLQUEUE.QueueBox.Value <> "" Then

'//check for thread
If Range("PostThread").Offset(1, 0).Value2 <> vbNullString Then GoTo queueAsThread
xTwt = Replace(xTwt, xT, vbNullString)

draftFile = xFile & xTwt & xExt

xStr = Replace(ETWEETXLQUEUE.PostBox.Value, Chr(10), "{ENTER};")
xStr = Replace(xStr, " ", "{SPACE};")
            
Open draftFile For Output As #1
Print #1, xStr
Print #1, "*-;"
Print #1, "*-" & ETWEETXLQUEUE.MedLinkBox.Value
Close #1

Exit Sub
                

'//Save thread from queue screen
queueAsThread:

lastR = Cells(Rows.Count, "Y").End(xlUp).Row
If lastR < 1 Then lastR = Cells(Rows.Count, "Z").End(xlUp).Row
        
        xTwt = Replace(xTwt, xT, vbNullString)

        threadFile = thrFile & xTwt & xExt

            Open threadFile For Output As #2
            
            For X = 1 To lastR - 1
            
            xStr = Replace(Range("PostThread").Offset(X, 0).Value2, Chr(10), "{ENTER};")
            xStr = Replace(xStr, " ", "{SPACE};")
            xMedArr = Split(Range("MedThread").Offset(X, 0).Value2, """ """)
            
            xMed = vbNullString
            For M = 0 To UBound(xMedArr)
            xMedHldr = xMedArr(M)
            xMedHldr = Replace(xMedHldr, """", vbNullString)
            If M >= 1 And M < UBound(xMedArr) Then xMedHldr = """" & xMedHldr & """" & " " Else xMedHldr = """" & xMedHldr & """"
            xMed = xMed & xMedHldr
            Next
            
            Print #2, xStr
            Print #2, "*-;"
            Print #2, "*-" & xMed
            Print #2, "*-(" & X & ");"
            Next
            
            Close #2
            
            Exit Sub
              
                     End If
                    

Exit Sub

EndMacro:
xMsg = 21: Call eTweetXL_MSG.AppMsg(xMsg, errLvl)

End Sub
Sub pstPersData()

Call eTweetXL_LOC.xPersFile(persFile)

lastR = Cells(Rows.Count, "A").End(xlUp).Row

Open persFile & Range("Profile").Value2 & ".pers" For Output As #1

For xNum = 1 To lastR
Print #1, _
Range("Profile").Offset(xNum, 0).Value2 & ";" & _
Range("F1").Offset(xNum, 0).Value2 & ";" & _
Range("Browser").Value2 & ";" & _
Range("Scure").Offset(xNum, 0).Value2 & ";" & _
Range("Target").Offset(xNum, 0).Value2 & ";"
Next

Close #1

End Sub
