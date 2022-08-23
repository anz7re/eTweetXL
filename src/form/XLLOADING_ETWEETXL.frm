VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} XLLOADING_ETWEETXL 
   ClientHeight    =   1530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6810
   OleObjectBlob   =   "XLLOADING_ETWEETXL.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "XLLOADING_ETWEETXL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()

'//wait...
xArt = "<lib>xbas;delayevent(25);$": Call lexKey(xArt)

'//close loading window...
If Range("xlasAppLoad").Value = 1 Then XLLOADING_ETWEETXL.Hide

'//set initial loading bar color
LoadBar.BackColor = RGB(254, 251, 1)

'//start loading...
For X = 3 To 100
LoadBarStatus = (X * 156) / 100 '//for incrementing the loading bar
LoadBar.Width = LoadBarStatus '//increment loading bar
xDot = xDot & ".": If Len(xDot) > 3 Then xDot = "." '//set dot to variable & check if there's more than 3 (reset to 1 if so)
LoadRatio.Caption = X & "%" & xDot '//percentage done
LoadStatus.Caption = "Please wait while the application loads" & xDot '//message while loading
R = 254 - (X * 2.5): G = 251: B = 1 '//color change formula
LoadBar.BackColor = RGB(R, G, B) '//change bar color
LogoBg.BackColor = RGB(B, G, R) '//change bar color
LoadBg1.BackColor = RGB(R, G, B) '//change bar color
LoadBg2.BackColor = RGB(B, G, R) '//change bar color
LoadBg2.Width = LoadBg2.Width - 2 '//change bar size
LoadBg3.BackColor = RGB(B, G, R) '//change bar color
LoadBg3.Width = LoadBg3.Width + 5 '//change bar size
XLLOADING_ETWEETXL.BackColor = RGB(R, G, B) '//change background color
X = X + 2 '//amount to increment each load iteration
xArt = "<lib>xbas;delayevent(2);$": Call lexKey(xArt) '//wait...
Next

'//loading complete...
If Range("xlasAppLoad").Value <> 1 Then Call xlLoadComplete

End Sub
Private Sub xlLoadComplete()

'//set application load trigger
Range("xlasAppLoad").Value = 1

'//set ending loading bar flicker
LoadBar.BackColor = &H8000000F

'//wait...
xArt = "<lib>xbas;delayevent(2);$": Call lexKey(xArt)

'//set ending loading bar color
LoadBar.BackColor = vbGreen

'//completion message
LoadStatus.Caption = "Loading complete..."

'//hide load window to refresh
XLLOADING_ETWEETXL.Hide

'//wait...
xArt = "<lib>xbas;delayevent(2);$": Call lexKey(xArt)

'//show refreshed load window
XLLOADING_ETWEETXL.Show

'//application/window to show after loading
ETWEETXLHOME.Show

End Sub



