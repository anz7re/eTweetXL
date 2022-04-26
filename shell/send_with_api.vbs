Option Explicit

Dim objXL, objWb, strMsg, WsShell

Set objXL = GetObject(, "Excel.Application")

If Not TypeName(objXL) = "Empty" Then
	If objXL.ActiveWorkbook.Name = "eTweetXL.xlsm" Then    
    		objXL.Application.Run "App_TOOLS.RunPy"
			WScript.Quit
Else 
    strMsg = "eTweetXL NOT Running!"

	End If
		End If

Wscript.Quit