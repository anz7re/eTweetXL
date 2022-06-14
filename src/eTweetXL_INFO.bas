Attribute VB_Name = "eTweetXL_INFO"
'/##################################\
'//Important Application Information\\
'///################################\\\

Public Function ActiveProfile() As String

ActiveProfile = ThisWorkbook.Worksheets("Main").Range("Profile").Value2

If ActiveProfile = vbNullString Then
    ActiveProfile = ETWEETXLSETUP.ProfileListBox.Value
        ElseIf ActiveProfile = vbNullString Then
            ActiveProfile = ETWEETXLSETUP.ProfileNameBox.Value
                End If
                
End Function
Public Function AppLoc() As String

AppLoc = Env & "\.z7\autokit\etweetxl"

End Function
Public Function AppTag() As String

AppTag = "eTweetXL v1.6.0"

End Function
Public Function AppWbName() As String

Dim wbName As String

wbName = ThisWorkbook.name

If InStr(1, wbName, ".xlsm") Then
wbName = Replace(ThisWorkbook.name, ".xlsm", "")
End If

If InStr(1, wbName, ".xlsm") Then
wbName = Replace(ThisWorkbook.name, ".xlsb", "")
End If

AppWbName = wbName

End Function
Public Function AppWelcome() As String

AppWelcome = "Welcome to eTweetXL v1.6.0..."

End Function
Public Function Env() As String

Env = Environ("USERPROFILE")

End Function
Public Function LinkTrigH() As Byte

LinkTrigH = ThisWorkbook.Worksheets("Main").Range("LinkTrig").Value2

End Function


