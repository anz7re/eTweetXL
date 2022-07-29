VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ETWEETXLAPISETUP 
   Caption         =   "eTweetXL"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7290
   OleObjectBlob   =   "ETWEETXLAPISETUP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ETWEETXLAPISETUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()

'//WinForm #
xWin = 15: Call setWindow(xWin)

Range("DataPullTrig").Value = 0

'//IMPORT PRESETS
Call eTweetXL_GET.getProfileNames

End Sub
Private Sub UserForm_Initialize()

Me.Caption = AppTag

End Sub
Private Sub accSecretBox_Change()

If Me.accSecretBox.Value <> "" Then Me.accSecretBox.BorderColor = &H8000000D

End Sub

Private Sub accTokenBox_Change()

If Me.accTokenBox.Value <> "" Then Me.accTokenBox.BorderColor = &H8000000D

End Sub

Private Sub apiKeyBox_Change()

If Me.apiKeyBox.Value <> "" Then Me.apiKeyBox.BorderColor = &H8000000D

End Sub

Private Sub apiSecretBox_Change()

If Me.apiSecretBox.Value <> "" Then Me.apiSecretBox.BorderColor = &H8000000D

End Sub
Private Sub ProfileListBox_Click()

Range("Profile").Value2 = ETWEETXLAPISETUP.ProfileListBox.Value
Range("DataPullTrig").Value = 0

Call eTweetXL_GET.getProfileData

End Sub
Private Sub SaveBtn_Click()

eTweetXL_CLICK.SaveBtn_Clk

End Sub

Private Sub UserListBox_Click()

Call eTweetXL_GET.getAPIData

End Sub

Private Sub UserListBox_Change()

'//CLEAR
apiKeyBox.Value = ""
apiSecretBox.Value = ""
accTokenBox.Value = ""
accSecretBox.Value = ""

Range("User").Value = Replace(ETWEETXLPOST.UserListBox.Value, Range("Scure").Value, "")
xUser = Range("User").Value

If xUser <> "" Then
Call eTweetXL_CLICK.SetActive_Clk(xUser)
End If

End Sub
Private Sub UserForm_Terminate()

Unload ETWEETXLAPISETUP
Range("xlasWinForm").Value2 = 12

End Sub

