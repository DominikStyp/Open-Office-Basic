' Choosing file by dialog
Function ChooseAFileName() As String
	Dim vFileDialog 'FilePicker service instance
	Dim vFileAccess 'SimpleFileAccess service instance
	Dim iAccept as Integer 'Response to the FilePicker
	Dim sInitPath as String 'Hold the initial path
	'Note: The following services MUST be called in the following order
	'or Basic will not remove the FileDialog Service
	vFileDialog = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
	vFileAccess = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	'Set the initial path here!
	sInitPath = ConvertToUrl(CurDir)
	If vFileAccess.Exists(sInitPath) Then
		vFileDialog.SetDisplayDirectory(sInitPath)
	End If
	iAccept = vFileDialog.Execute() 'Run the file chooser dialog
	If iAccept = 1 Then 'What was the return value?
		ChooseAFileName = vFileDialog.Files(0) 'Set file name if it was not canceled
	End If
	vFileDialog.Dispose() 'Dispose of the dialog
End Function


' Choose Directory by dialog
REM sInPath specifies the initial directory. If the initial directory
REM is not specified, then the user's default work directory is used.
REM The selected directory is returned as a URL.
Function ChooseADirectory(Optional sInPath$) As String
	Dim oDialog As Object
	Dim oSFA As Object
	Dim s As String
	Dim oPathSettings
	oDialog = CreateUnoService("com.sun.star.ui.dialogs.FolderPicker")
	'oDialog = CreateUnoService("com.sun.star.ui.dialogs.OfficeFolderPicker")
	oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")
	If IsMissing(sInPath) Then
		oPathSettings = CreateUnoService("com.sun.star.util.PathSettings")
		oDialog.setDisplayDirectory(oPathSettings.Work)
	ElseIf oSFA.Exists(sInPath) Then
		oDialog.setDisplayDirectory(sInPath)
	Else
		s = "Directory '" & sInPath & "' Does not exist"
	If MsgBox(s, 33, "Error") = 2 Then Exit Function
	End If
	If oDialog.Execute() = 1 Then
		ChooseADirectory() = oDialog.getDirectory()
	End If
End Function
