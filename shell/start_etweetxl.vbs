Option Explicit

On Error Resume Next

Dim objXL, objWb, strMsg

Set objXL = GetObject(, "Excel.Application")

If Not TypeName(objXL) = "Empty" Then
	If objXL.ActiveWorkbook.Name = "eTweetXL.xlsm" Then    
    		objXL.Application.Run "App_CLICK.Start_Clk"
			WScript.Quit
Else 
    strMsg = "eTweetXL NOT Running!"

	End If
		End If

WScript.Quit