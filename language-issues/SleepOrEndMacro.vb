Sub ExampleWait
	On Error Goto BadInput
	Dim nMillis As Long
	Dim nStart As Long
	Dim nEnd As Long
	Dim nElapsed As Long
	nMillis = CLng(InputBox("How many milliseconds to wait?"))
	nStart = GetSystemTicks()
	' Sleep for nMilis
	Wait(nMillis)
	nEnd = GetSystemTicks()
	nElapsed = nEnd - nStart
	MsgBox "You requested to wait for " & nMillis & " milliseconds" &_
	CHR$(10) & "Waited for " & nElapsed & " milliseconds", 0, "Example Wait"
	BadInput:
	If Err <> 0 Then
	Print Error() & " at line " & Erl
	End If
	On Error Goto 0
End Sub