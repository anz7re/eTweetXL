Attribute VB_Name = "eTweetXL_FOCUS"
'/##########################\
'//Application Focus Windows\\
'///########################\\\

Sub hde_ActiveForm()

On Error GoTo CloseAPISetup

If Range("xlasWinForm").Value2 = 11 Then ETWEETXLHOME.Hide
If Range("xlasWinForm").Value2 = 12 Then ETWEETXLSETUP.Hide
If Range("xlasWinForm").Value2 = 13 Then ETWEETXLPOST.Hide
If Range("xlasWinForm").Value2 = 14 Then ETWEETXLQUEUE.Hide
If Range("xlasWinForm").Value2 = 100 Then CTRLBOX.Hide
Exit Sub

CloseAPISetup:
Unload ETWEETXLAPISETUP

End Sub
Sub shw_ETWEETXLHOME()
    
On Error Resume Next

ETWEETXLHOME.Show

End Sub
Sub shw_CTRLBOX()
    
On Error Resume Next

CTRLBOX.Show

End Sub
Sub shw_ETWEETXLSETUP()
    
On Error Resume Next

ETWEETXLSETUP.Show

End Sub
Sub shw_ETWEETXLPOST()
    
On Error Resume Next

ETWEETXLPOST.Show

End Sub
Sub shw_ETWEETXLQUEUE()
    
On Error Resume Next

ETWEETXLQUEUE.Show

End Sub

