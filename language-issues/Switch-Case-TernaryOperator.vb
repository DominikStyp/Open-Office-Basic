' >>>> TERNARY OPERATOR (Immediate If) <<<
' Works like ternary operator in other languages
max_age = IIf(johns_age > bills_age, johns_age, bills_age)

' >>>> SWITCH CASE In-line  <<<
' The Choose statement returns a null if the expression is less than 1 or greater than the number of selection arguments. Choose returns “select_1” if the expression evaluates to 1, and “select_2” if the expression evaluates to 2
i% = 3
Print Choose (i%, 1/(i+1), 1/(i-1), 1/(i-2), 1/(i-3))


' Choose function 

Sub ExampleChoose
	Dim i%, v
	i% = CStr(InputBox("Enter an integer 1–4 (negative number is an error)"))
	v = Choose(i, "one", "two", "three", "four")
	If IsNull(v) Then
		Print "V is null"
	Else
		Print CStr(v)
	End If
End Sub



' >>>> SWITCH CASE <<<<
Select Case 2
	Case 1
		Print "One"
	Case 3
		Print "Three"
	Case Else
		Print "NONE"
End Select

' >>>> SWITCH CASE VARIATIONS <<<<
Select Case i
	Case 1, 3, 5
		Print "i is one, three, or five"
	Case 6 To 10
		Print "i is a value from 6 through 10"
	Case < -10
		Print "i is less than -10"
	Case IS > 10 'IS is optional
		Print "i is greater than 10"
	Case Else
		Print "No idea what i is"
End Select

' Multiple IS
Select Case i%
	Case 6, Is = 7, Is = 8, Is > 15, Is < 0
		Print "" & i & " matched"
	Case Else
		Print "" & i & " is out of range"
End Select