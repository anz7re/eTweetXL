Option Explicit
Dim nL
nL = Chr(10)

MsgBox("ERROR!" & nL & nL & "If you're seeing this message there were multiple errors logging into your profile." & nL & nL & _
"Try restarting your selected browser and rerunning the application." & nL & nL & "If the problem persists, try adjusting the time in-between your posts (browsers are prone to breaking when you're using automation)."), vbExclamation, "eTweetXL v1.5.0"

WScript.Quit