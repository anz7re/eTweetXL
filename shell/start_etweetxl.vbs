Option Explicit

On Error Resume Next

Dim oXL, oWb, xMsg

Set oXL = GetObject(, "Excel.Application")

If Not TypeName(oXL) = "Empty" Then
	If oXL.ActiveWorkbook.Name = "eTweetXL.xlsm" Then    
    		oXL.Application.Run "eTweetXL_CLICK.StartBtn_Clk"
			wscript.Quit
Else 
    MsgBox("Application component not found!")

	End If
		End If

wscript.Quit
