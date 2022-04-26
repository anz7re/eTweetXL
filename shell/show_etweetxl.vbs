Option Explicit

Dim objXL, objWb

Set objXL = CreateObject("Excel.Application")

Set objWb = objXL.Workbooks.Open("C:\Users\EDITHERE\.z7\autokit\etweetxl\app\eTweetXL.xlsm")

WScript.Sleep 50

objWb.Save
objWb.Close
Set objWb = Nothing
Set objXL = Nothing

Wscript.Quit