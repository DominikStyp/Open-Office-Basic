
Private zero%
Sub ExampleErrorResumeNext
	On Error Resume Next
	Print 1/Zero%
	' Err = Error number, Error$ = error string, Erl = error line
	If Err <> 0 Then Print Error$ & " at line " & Erl 'Err was cleared
End Sub

' >>> DEFINED ERROR HANDLER
Private zero%
Private error_s$
Sub ExampleJumpErrorHandler
	On Error GoTo ExErrorHandler
	Print 1/Zero%
	MsgBox error_s, 0, "Jump Error Handler"
	Exit Sub
	ExErrorHandler: 'Handles division error
		error_s = error_s & "Error in MainJumpErrorHandler at line " & Erl() &_
		" : " & Error() & CHR$(10)
	Resume Next
End Sub

'>>>>>> ON ERROR

On Error Resume Next           'Ignore errors and continue running at the next line in the macro.
On Error GoTo    0             'Cancel the current error handler.
On Error GoTo    LabelName     'Transfer control to the specified label.

' >>> IGNORE ERRORS
On Error GoTo PropertiesDone 'Ignore any errors in this section.
a() = getProperties() 'If unable to get properties then
DisplayStuff(a(), "Properties") 'an error will prevent getting here.
PropertiesDone:
	On Error GoTo MethodsDone 'Ignore any errors in this section.
a() = getMethods()
DisplayStuff(a(), "Methods")
MethodsDone:
	On Error Goto 0 'Turn off current error handlers.

' On Error Resume subroutine in other spot
Sub ExampleResumeHandler
	Dim s$, z%
	On Error GoTo Handler1 'Add a message, then resume to Spot1
	s = "(0) 1/z = " & 1/z & CHR$(10) 'Divide by zero, so jump to Handler1
	Spot1: 'Got here from Handler1
	On Error GoTo Handler2 'Handler2 uses resume
	s = s & "(1) 1/z = "&1/z & CHR$(10) 'Fail the first time, work the second
	On Error GoTo Handler3 'Handler3 resumes the next line
	z = 0 'Allow for division by zero again
	s = s & "(2) 1/z = "&1/z & CHR$(10) 'Fail and call Handler3
	MsgBox s, 0, "Resume Handler"
	Exit Sub
	Handler1:
		s = s & "Handler1 called from line " & Erl() & CHR$(10)
	Resume Spot1
	Handler2:
		s = s & "Handler2 called from line " & Erl() & CHR$(10)
		z = 1 'Fix the error then do the line again
	Resume
	Handler3:
		s = s & "Handler3 called from line " & Erl() & CHR$(10)
		Resume Next
End Sub