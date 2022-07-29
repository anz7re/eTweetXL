VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CTRLBOXABOUT 
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "CTRLBOXABOUT.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CTRLBOXABOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub AppInformationBox_Change()

AppInformationBox.Value = CtrlLicense

End Sub
Private Sub UserForm_Activate()

AppInformationBox.Value = CtrlLicense

SocialLink.Font.Underline = True
WebsiteLink.Font.Underline = True

End Sub
Private Sub SocialLink_Click()

xLink = Replace(SocialLink.Caption, "@", vbNullString)

'//start w/ msedge
xArt = "<lib>xbas;sh(msedge.exe https://" & "twitter.com/" & xLink & ");$" '//xlas
Call lexKey(xArt)

End Sub
Private Sub WebsiteLink_Click()

'//start w/ msedge
xArt = "<lib>xbas;sh(msedge.exe https://" & WebsiteLink.Caption & ");$" '//xlas
Call lexKey(xArt)

End Sub

