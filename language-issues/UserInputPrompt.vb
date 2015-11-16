' >>>>> Dialog types of MsgBox
'0 Display OK button only.
'1 Display OK and Cancel buttons.
'2 Display Abort, Retry and Ignore buttons.
'3 Display Yes, No, and Cancel buttons.
'4 Display Yes and No buttons.
'5 Display Retry and Cancel buttons.
'16 Add the Stop icon to the dialog.
'32 Add the Question icon to the dialog.
'48 Add the Exclamation Point icon to the dialog.
'64 Add the Information icon to the dialog.
'128 First button in the dialog is the default button. This is the default behavior.
'256 Second button in the dialog is the default button.
'512 Third button in the dialog is the default button.

' >>>> Return values of MsgBox
'1 OK
'2 Cancel
'3 Abort
'4 Retry
'5 Ignore
'6 Yes
'7 No

MsgBox(Message)
MsgBox(Message, DialogType)
MsgBox(Message, DialogType, DialogTitle)

' >>>>>>>> Prompt for input from user 

InputBox(Message)
InputBox(Message, Title)
InputBox(Message, Title, Default)
InputBox(Message, Title, Default, x_pos, y_pos)

Sub ExampleInputBox
	Dim sReturn As String 'Return value
	Dim sMsg As String 'Holds the prompt
	Dim sTitle As String 'Window title
	Dim sDefault As String 'Default value
	Dim nXPos As Integer 'Twips from left edge
	Dim nYPos As Integer 'Twips from top edge
	nXPos = 1440 * 2 'Two inches from left edge of the window
	nYPos = 1440 * 4 'Four inches from top of the window
	sMsg = "Please enter some meaningful text"
	sTitle = "Meaningful Title"
	sDefault = "Hello"
	sReturn = InputBox(sMsg, sTitle, sDefault, nXPos, nYPos)
	If sReturn <> "" Then
		REM Print the entered string surrounded by double quotes
		Print "You entered """;sReturn;""""
	Else
		Print "You either entered an empty string or chose Cancel"
	End If
End Sub
