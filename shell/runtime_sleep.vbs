Option Explicit

On Error Resume Next

Dim  oFSO, objFile, objXL, objWs, objRead
Dim  ReadStr, StrMsg, xOffset

Set oFSO = CreateObject("Scripting.FileSystemObject")

'//Get offset for sleep
Set objRead = oFSO.OpenTextFile("C:\Users\EDITHERE\.z7\autokit\etweetxl\mtsett\offset.mt", 1)

   ReadStr = objRead.ReadAll()
   xOffset = ReadStr
   objRead.Close
   Set objRead = Nothing

Set objXL = GetObject(, "Excel.Application")

If Not TypeName(objXL) = "Empty" Then
	If objXL.ActiveWorkbook.Name = "eTweetXL.xlsm" Then  
		Set objWs = objXL.ActiveWorkbook.Worksheets("Sheet1")
		    objWs.Cells(6, 52).value = xOffset
WScript.Sleep 50
				WScript.Quit
Else 
    strMsg = "eTweetXL NOT Running!"

	End If
		End If

WScript.Quit