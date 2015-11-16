' >>>> IF / ELSEIF <<<<<<<< 
If i <> 3 Then
If k = 4 Then Print "k is four"
	If j = 7 Then
		Print "j is seven"
	ElseIf i = 1 Then
		Print "i is 1"
	End If
End If

'>>>>>>> IIF (Immidiate If)
If Condition Then
	object = TrueExpression
Else
	object = FalseExpression
End If
' .... can be replaced by following, that is kind of ternary operator
object = IIf (Condition, TrueExpression, FalseExpression)

'>>>>>>>>> CHOOSE 
' returns SELECT_1 if expression evaluates to 1
obj = Choose (expression, Select_1[, Select_2, ... [,Select_n]])


' GoSub
' The GoSub statement causes execution to jump to a defined label in the current routine. It isn’t possible to jump outside of the current routine.
Sub ExampleGoSub
	Dim i As Integer
	GoSub Line2 REM Jump to line 2 then return, i is 1
	GoSub [Line 1] REM Jump to line 1 then return, i is 2
	MsgBox "i = " + i, 0, "GoSub Example" REM i is now 2
	Exit Sub REM Leave the current subroutine.
	[Line 1]: REM this label has a space in it
	i = i + 1 REM Add one to i
	Return REM return to the calling location
	Line2: REM this label is more typical, no spaces
	i = 1 REM Set i to 1
	Return REM return to the calling location
End Sub

' GoTo
' The GoTo statement causes execution to jump to a defined label in the current routine. 
' It isn’t possible to jump outside of the current routine
Sub ExampleGoTo
	Dim i As Integer
	GoTo Line2 REM Okay, this looks easy enough
	Line1: REM but I am becoming confused
	i = i + 1 REM I wish that GoTo was not used
	GoTo TheEnd REM This is crazy, makes me think of spaghetti,
	Line2: REM Tangled strands going in and out; spaghetti code.
	i = 1 REM If you have to do it, you probably
	GoTo Line1 REM did something poorly.
	TheEnd: REM Do not use GoTo.
	MsgBox "i = " + i, 0, "GoTo Example"
End Sub

' On GoTo, On GoSub
Sub ExampleOnGoTo
	Dim i As Integer
	Dim s As String
	i = 1
	On i+1 GoSub Sub1, Sub2
	75
	s = s & Chr(13)
	On i GoTo Line1, Line2
	REM The exit causes us to exit if we do not continue execution
	Exit Sub
	Sub1:
	s = s & "In Sub 1" : Return
	Sub2:
	s = s & "In Sub 2" : Return
	Line1:
	s = s & "At Label 1" : GoTo TheEnd
	Line2:
	s = s & "At Label 2"
	TheEnd:
	MsgBox s, 0, "On GoTo Example"
End Sub


