Attribute VB_Name = "App_INFO"
'/##################################\
'//Important Application Information\\
'///################################\\\

Public Function AppLoc()

AppLoc = Env & AppPath

End Function
Public Function CtrlLoc()

CtrlLoc = Env & CtrlPath

End Function
Public Function AppPath()

AppPath = "\.z7\autokit\etweetxl"

End Function
Public Function CtrlPath()

CtrlPath = "\.z7\console\ctrl_box"

End Function
Public Function Env()

Env = Environ("USERPROFILE")

End Function
Public Function LinkTrigHldr()

LinkTrigHldr = Range("LinkTrig").Value

End Function
Public Function xProfile()

xProfile = ThisWorkbook.Worksheets("Main").Range("Profile").Value

If xProfile = "" Then
    xProfile = ETWEETXLSETUP.ProfileListBox.Value
        ElseIf xProfile = "" Then
            xProfile = ETWEETXLSETUP.ProfileNameBox.Value
                End If
                
End Function
Public Function AppWelcome()

AppWelcome = "Welcome to eTweetXL v1.5.0..."

End Function
Public Function AppTag()

AppTag = "eTweetXL v1.5.0"

End Function
Public Function WbAppName() As String

Dim wbName As String

wbName = ThisWorkbook.name

If InStr(1, wbName, ".xlsm") Then
wbName = Replace(ThisWorkbook.name, ".xlsm", "")
End If

If InStr(1, wbName, ".xlsm") Then
wbName = Replace(ThisWorkbook.name, ".xlsb", "")
End If

WbAppName = wbName

End Function

