Attribute VB_Name = "App_Focus"
'/##########################\
'//Application Focus Windows\\
'///########################\\\

Sub HdForms()

On Error GoTo CloseAPISetup

If Range("xlasWinForm").Value = 1 Then ETWEETXLHOME.Hide
If Range("xlasWinForm").Value = 2 Then ETWEETXLSETUP.Hide
If Range("xlasWinForm").Value = 3 Then ETWEETXLPOST.Hide
If Range("xlasWinForm").Value = 4 Then ETWEETXLQUEUE.Hide
If Range("xlasWinForm").Value = 10 Then CTRLBOX.Hide
Exit Sub

CloseAPISetup:
Unload ETWEETXLAPISETUP

End Sub
Sub SH_ETWEETXLHOME()
    
On Error Resume Next

ETWEETXLHOME.Show

End Sub
Sub SH_CTRLBOX()
    
On Error Resume Next

CTRLBOX.Show

End Sub
Sub SH_ETWEETXLSETUP()
    
On Error Resume Next

ETWEETXLSETUP.Show

End Sub
Sub SH_ETWEETXLPOST()
    
On Error Resume Next

ETWEETXLPOST.Show

End Sub
Sub SH_ETWEETXLQUEUE()
    
On Error Resume Next

ETWEETXLQUEUE.Show

End Sub

