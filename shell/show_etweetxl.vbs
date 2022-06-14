Option Explicit

Dim oXL, oWb

Set oXL = CreateObject("Excel.Application")

Set oWb = oXL.Workbooks.Open("C:\Users\EDITHERE\.z7\autokit\etweetxl\app\eTweetXL.xlsm")

wscript.Sleep 10

oWb.Save
oWb.Close
Set oWb = Nothing
Set oXL = Nothing

wscript.Quit
