Print "one" & CHR$(10) & "two" & CHR$(13) & "three" ' Displays three dialogs


Print "one", 'Do not print yet, ends with a comma
Print "two" 'Print "one two"
Print "three",'Do not print yet, ends with a comma
Print "four"; 'Do not print yet, ends with a semicolon
Print 'Print "three four"

' Display a simple message box with a new line.
Sub ExampleMsgBoxWithReturn
	MsgBox "one" & CHR$(10) & "two"
End Sub



Sub MsgBoxExamples()
	Dim i%
	Dim values
	values = Array(0, 1, 2, 3, 4, 5)
	For i = LBound(values) To UBound(values)
	MsgBox ("Dialog Type: " + values(i), values(i))
	Next
	values = Array(16, 32, 48, 64, 128, 256, 512)
	For i = LBound(values) To UBound(values)
	MsgBox ("Yes, No, Cancel, with Type: " + values(i), values(i) + 3)
	Next
End Sub