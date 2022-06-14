Option Explicit

On Error Resume Next

Dim  oFSO, oFile, oXL, oWs
Dim  xStr, xMsg, xRtState

Set oFSO = CreateObject("Scripting.FileSystemObject")

'//GET OFFSET
Set oFile = oFSO.OpenTextFile("C:\Users\EDITHERE\.z7\autokit\etweetxl\debug\rt.err", 1)

   xStr = oFile.ReadAll()
   xRtState = xStr
   oFile.Close
   Set oFile = Nothing

Set oXL = GetObject(, "Excel.Application")

If Not TypeName(oXL) = "Empty" Then
	If oXL.ActiveWorkbook.Name = "eTweetXL.xlsm" Then  
		Set oWs = oXL.ActiveWorkbook.Worksheets("Main")
		    oWs.Range("RtState").Value = xRtState

wscript.Sleep 10
    			oXL.Application.Run "eTweetXL_GET.getRtState"
			oXL.Application.Run "eTweetXL_TOOLS.disableFlowStrip"
				wscript.Quit
Else 
    xMsg = "eTweetXL NOT Running!"

	End If
		End If

wscript.Quit
