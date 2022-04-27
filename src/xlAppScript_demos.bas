Attribute VB_Name = "xlAppScript_demos"
'//--------------------------------------------------------------------------------------------------------------------------------
'//     Method 1
'//
'//Run xlAppScript from VBA
'
Sub method_1()

'//Enter some xlAppScript code enclosed in quotes & set to lex variable 'xArt'
xArt = "<lib>xbas;rng(A1).value(100).bgcolor(gainsboro).fcolor(cornflowerblue);$"

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

'//Set file containing xlAppScript to a variable
xFile = Env & "\documents\demo.txt"

Open xFile For Input As #1 '//open file
Do Until EOF(1) '//search until the end
Line Input #1, xStr '//set code in file to variable
xStrHldr = xStrHldr & xStr
Loop
Close #1

xArt = xStrHldr '//set article variable to code

'//Send code to the xlAppScript Lexer
Call lexKey(xArt)

'//All done!
'//_______________________________________________________________________________________________________________________________
Exit Sub
End Sub
'//     Method 3
'//
'//Run xlAppScript from another connected workbook environment
'/(code can be adjusted for parsing through a file as seen in method 2)
'
Public Sub method_3()

'//open connected workbook environment, activate, & other lines of xlas code (this workbook was located in \documents)
xArt = "<lib>xbas;wb(xlasbook.xlsm).open;wb(xlasbook.xlsm).active;rng(A1).value(100).bgcolor(gainsboro).fcolor(cornflowerblue);$"

'//send code to workbook for parsing using VBA
X = Application.Run("'xlasbook.xlsm'!lexKey", (xArt))

End Sub
