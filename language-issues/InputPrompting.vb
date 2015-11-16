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