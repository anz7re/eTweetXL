Option Explicit

On Error Resume Next

Dim  oFSO, oXL, oWs, oFile
Dim  xStr, xMsg, xOffset

Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oFile = oFSO.OpenTextFile("C:\Users\EDITHERE\.z7\autokit\etweetxl\mtsett\offset.mt", 1)

   xStr = oFile.ReadAll()
   xOffset = xStr
   oFile.Close
   Set oFile = Nothing

Set oXL = GetObject(, "Excel.Application")

If Not TypeName(oXL) = "Empty" Then
	If oXL.ActiveWorkbook.Name = "eTweetXL.xlsm" Then  
		Set oWs = oXL.ActiveWorkbook.Worksheets("Main")
		    oWs.Range("ActiveOffset").Value = xOffset
wscript.Sleep 10
				wscript.Quit
Else 
    xMsg = "eTweetXL NOT Running!"

	End If
		End If

wscript.Quit