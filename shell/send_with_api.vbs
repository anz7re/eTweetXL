Option Explicit

Dim oXL, oWb, xMsg

Set oXL = GetObject(, "Excel.Application")

If Not TypeName(oXL) = "Empty" Then
	If oXL.ActiveWorkbook.Name = "eTweetXL.xlsm" Then    
    		oXL.Application.Run "eTweetXL_TOOLS.runPy"
			wscript.Quit
Else 
    xMsg = "eTweetXL NOT Running!"

	End If
		End If

wscript.Quit
