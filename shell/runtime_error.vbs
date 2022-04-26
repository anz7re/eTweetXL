Option Explicit

On Error Resume Next

Dim  oFSO, objFile, objXL, objWs, objRead
Dim  ReadStr, StrMsg, xRtAction

Set oFSO = CreateObject("Scripting.FileSystemObject")

'//Get runtime message
Set objRead = oFSO.OpenTextFile("C:\Users\EDITHERE\.z7\autokit\etweetxl\debug\rt.err", 1)

   ReadStr = objRead.ReadAll()
   xRtAction = ReadStr
   objRead.Close
   Set objRead = Nothing

Set objXL = GetObject(, "Excel.Application")

If Not TypeName(objXL) = "Empty" Then
	If objXL.ActiveWorkbook.Name = "eTweetXL.xlsm" Then  
		Set objWs = objXL.ActiveWorkbook.Worksheets("Sheet1")
		    objWs.Cells(14, 52).value = xRtAction

WScript.Sleep 50
    			objXL.Application.Run "App_TOOLS.ShowRtAction"
			objXL.Application.Run "xlAppScript.DisableFlowStrip"
				WScript.Quit
Else 
    strMsg = "eTweetXL NOT Running!"

	End If
		End If

WScript.Quit