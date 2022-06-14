Attribute VB_Name = "xlAppScript_demos"
'//--------------------------------------------------------------------------------------------------------------------------------
'//     Method 1
'//
'//Run xlAppScript from VBA
'
Sub method_1()

'//Setup workbook if this is your first time running xlAppScript
Call connectWb

'//Enter some xlAppScript code & set to Lexer variable
xArt = "<lib>xbas;rng(A1).value(Testing123).bgcolor(gainsboro).fcolor(cornflowerblue);$"

'//Send code to the xlAppScript Lexer
Call lexKey(xArt)

'//All done!
'//_______________________________________________________________________________________________________________________________
End Sub
'//     Method 2
'//
'//Run xlAppScript from a file
'
Sub method_2()

'//Setup workbook if this is your first time running xlAppScript
Call connectWb

'//Set file containing xlAppScript to a variable
xFile = Environ("USERPROFILE") & "\documents\demo.txt"

Open xFile For Input As #1 '//open file
Do Until EOF(1) '//search until the end
Line Input #1, xStr '//set code in file to variable
xStrH = xStrH & xStr
Loop
Close #1

xArt = xStrH '//set article variable to code

'//Send code to the xlAppScript Lexer
Call lexKey(xArt)

'//All done!
'//_______________________________________________________________________________________________________________________________
Exit Sub
End Sub
'//     Method 3
'//
'//Send xlAppScript to another connected environment & run
'/(code can be adjusted for parsing through a file as seen in method 2)
'
Public Sub method_3()

'//set workbook active, & other lines of xlas code (this workbook was located in \documents)
xArt = "<lib>xbas;wb(xlasbook.xlsm).active;rng(A1).value(Testing123).bgcolor(gainsboro).fcolor(cornflowerblue);$"

'//send code to workbook for parsing using VBA (this will open the workbook)
X = Application.Run("'xlasbook.xlsm'!lexKey", (xArt))

End Sub
