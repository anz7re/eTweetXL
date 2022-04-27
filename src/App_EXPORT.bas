Attribute VB_Name = "App_EXPORT"
'/############################\
'//Application Export Features\\
'///##########################\\\

Sub MyDraftData(xTwt)

On Error GoTo EndMacro

Dim ProfDbArr(100) As String
Dim draftFile, theadFile, xFile, xPost As String
Dim oFSO, oFile, oFldr As Object

Call FindForm(xForm)
Call App_Loc.xThrFile(thrFile)
Call App_Loc.xTwtFile(twtFile)

If InStr(1, xTwt, " [...]") = False Then
If xForm.DraftFilterBtn.Caption <> "..." Or InStr(1, xTwt, " [•]") Then
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
If Range("xlasWinForm").Value < 4 Then

'//check for thread
If Range("PostThread").Offset(1, 0).Value <> vbNullString Then GoTo postAsThread
            
If ETWEETXLPOST.DraftBox.Value = vbNullString Then
draftFile = twtFile & "draft_" & gNum & "_" & xDate & ".twt"
    Else
        draftFile = twtFile & ETWEETXLPOST.DraftBox.Value & ".twt"
            End If
            
            xPost = Replace(ETWEETXLPOST.PostBox.Value, Chr(10), "{ENTER};")
            xPost = Replace(xPost, " ", "{SPACE};")
            
            Open draftFile For Output As #1
            Print #1, xPost
            Print #1, "*-;"
            Print #1, "*-" & ETWEETXLPOST.MedLinkBox.Value
            Close #1
                    Exit Sub
                    
                
'//Save thread from tweet setup screen
postAsThread:

lastRw = Cells(Rows.Count, "Y").End(xlUp).Row
If lastRw < 1 Then lastRw = Cells(Rows.Count, "Z").End(xlUp).Row

If ETWEETXLPOST.DraftBox.Value = vbNullString Then
threadFile = thrFile & "draft_" & gNum & "_" & xDate & ".thr"
    Else
        threadFile = thrFile & ETWEETXLPOST.DraftBox.Value & ".thr"
            End If

            Open threadFile For Output As #2
            
            For X = 1 To lastRw - 1
            
            xPost = Replace(Range("PostThread").Offset(X, 0).Value, Chr(10), "{ENTER};")
            xPost = Replace(xPost, " ", "{SPACE};")
            xMedArr = Split(Range("MedThread").Offset(X, 0).Value, """ """)
            
            xMed = vbNullString
            For M = 0 To UBound(xMedArr)
            xMedHldr = xMedArr(M)
            xMedHldr = Replace(xMedHldr, """", vbNullString)
            If M >= 1 Then xMedHldr = " " & """" & xMedHldr & """" Else xMedHldr = """" & xMedHldr & """"
            xMed = xMed & xMedHldr
            Next
            
            Print #2, xPost
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
If Range("PostThread").Offset(1, 0).Value <> vbNullString Then GoTo queueAsThread
xTwt = Replace(xTwt, xT, vbNullString)

draftFile = xFile & xTwt & xExt

xPost = Replace(ETWEETXLQUEUE.PostBox.Value, Chr(10), "{ENTER};")
xPost = Replace(xPost, " ", "{SPACE};")
            
Open draftFile For Output As #1
Print #1, xPost
Print #1, "*-;"
Print #1, "*-" & ETWEETXLQUEUE.MedLinkBox.Value
Close #1

Exit Sub
                

'//Save thread from queue screen
queueAsThread:

lastRw = Cells(Rows.Count, "Y").End(xlUp).Row
If lastRw < 1 Then lastRw = Cells(Rows.Count, "Z").End(xlUp).Row
        
        xTwt = Replace(xTwt, xT, vbNullString)

        threadFile = thrFile & xTwt & xExt

            Open threadFile For Output As #2
            
            For X = 1 To lastRw - 1
            
            xPost = Replace(Range("PostThread").Offset(X, 0).Value, Chr(10), "{ENTER};")
            xPost = Replace(xPost, " ", "{SPACE};")
            xMedArr = Split(Range("MedThread").Offset(X, 0).Value, """ """)
            
            xMed = vbNullString
            For M = 0 To UBound(xMedArr)
            xMedHldr = xMedArr(M)
            xMedHldr = Replace(xMedHldr, """", vbNullString)
            If M >= 1 And M < UBound(xMedArr) Then xMedHldr = """" & xMedHldr & """" & " " Else xMedHldr = """" & xMedHldr & """"
            xMed = xMed & xMedHldr
            Next
            
            Print #2, xPost
            Print #2, "*-;"
            Print #2, "*-" & xMed
            Print #2, "*-(" & X & ");"
            Next
            
            Close #2
            
            Exit Sub
              
                     End If
                    

Exit Sub

EndMacro:
xMsg = 21: Call App_MSG.AppMsg(xMsg)

End Sub
Sub MyUserData()

Call App_Loc.xPersFile(persFile)

lastRw = Cells(Rows.Count, "A").End(xlUp).Row

Open persFile & Range("Profile").Value & ".pers" For Output As #1

For xNum = 1 To lastRw
Print #1, _
Range("Profile").Offset(xNum, 0).Value & ";" & _
Range("User").Offset(xNum, 0).Value & ";" & _
Range("Browser").Value & ";" & _
Range("Scure").Offset(xNum, 0).Value & ";" & _
Range("Target").Offset(xNum, 0).Value & ";"
Next

Close #1

End Sub
