' >>>>>>>>>>>>>>>>> SIMPLE FILE ACCESS SERVICE <<<<<<<'
' >>>>>>>>>>>>>>>>> SIMPLE FILE ACCESS SERVICE <<<<<<<'
' >>>>>>>>>>>>>>>>> SIMPLE FILE ACCESS SERVICE <<<<<<<'
' >>>>>>>>>>>>>>>>> SIMPLE FILE ACCESS SERVICE <<<<<<<'

' >>> copy(fromURL, toURL) Copy a file.
' >>> move(fromURL, toURL) Move a file.
' >>> kill(url) Delete a file or directory, even if the folder is not empty.
' >>> isFolder(url) Return true if the URL represents a folder.
' >>> isReadOnly(url) Return true if the file is read-only.
' >>> setReadOnly(url, bool) Set file as read-only if the boolean argument is true, otherwise, clear the read-only flag.
' >>> createFolder(url) Creates a new Folder.
' >>> getSize(url) Returns the size of a file as a long integer.
' >>> getContentType(url) Return the content type of a file as a string. On my computer, an odt file has type application/vnd.sun.staroffice.fsys-file.
' >>> getDateTimeModified(url) Return the last modified date for the file as a com.sun.star.util.DateTime structure, which supports the properties: HundredthSeconds, Seconds, Minutes, Hours, Day, Month, and Year.
' >>> getFolderContents(url, bool) Returns the contents of a folder as an array of strings. Each string is the full path as a URL. If the bool is True, then files and directories are listed. If the bool is False, then only files are returned.
' >>> exists(url) Return true if a file or directory exists.
' >>> openFileRead(url) Open file to read, return an input stream.
' >>> openFileWrite(url) Open file to write, return an output stream.
' >>> openFileReadWrite(url) Open file to read and write, return a stream.
' >>> setInteractionHandler(handler) Set an interaction handler to be used for further operations. This is a more advanced topic and I will not discuss this here.
' >>> writeFile(toUrl, inputStream) Overwrite the file content with the given data.

Sub ExampleSimpleFileAccess
	Dim oSFA ' SimpleFileAccess service.
	Dim sFileName$ ' Name of file to open.
	Dim oStream ' Stream returned from SimpleFileAccess.
	Dim oTextStream ' TextStream service.
	Dim sStrings ' Strings to test write / read.
	Dim sInput$ ' The string that is read.
	Dim s$ ' Accumulate result to print.
	Dim i% ' Index variable.
	sStrings = Array("One", "UTF:Aa", "1@3")
	' File to use.
	sFileName = CurDir() & "/delme.out"
	' Create the SimpleFileAccess service.
	oSFA = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
	
	' >>>>>>>>> WRITING TO FILE <<<<<<<<<<< 
	'Create the Specialized stream.
	oTextStream = CreateUnoService("com.sun.star.io.TextOutputStream")
	'If the file already exists, delete it.
	If oSFA.exists(sFileName) Then
		oSFA.kill(sFileName)
	End If
	' Open the file for writing.
	oStream = oSFA.openFileWrite(sFileName)
	' Attach the simple stream to the text stream.
	' The text stream will use the simple stream.
	oTextStream.setOutputStream(oStream)
	' Write the strings.
	For i = LBound(sStrings) To UBound(sStrings)
		oTextStream.writeString(sStrings(i) & CHR$(10))
	Next
	' Close the stream.
	oTextStream.closeOutput()
	
	' >>>>>>>>> READING FILE <<<<<<<<<<< 
	oTextStream = CreateUnoService("com.sun.star.io.TextInputStream")
	oStream = oSFA.openFileRead(sFileName)
	oTextStream.setInputStream(oStream)
	For i = LBound(sStrings) To UBound(sStrings)
		sInput = oTextStream.readLine()
		s = s & CStr(i)
		' If the EOF is reached then the new line delimiters are
		' not removed. I consider this a bug.
		If oTextStream.isEOF() Then
		If Right(sInput, 1) = CHR$(10) Then
		sInput = Left(sInput, Len(sInput) - 1)
		End If
		End If
		' Verify that the read string is the same as the written string.
		If sInput <> sStrings(i) Then
			s = s & " : BAD "
		Else
			s = s & " : OK "
		End If
		s = s & "(" & sStrings(i) & ")"
		s = s & "(" & sInput & ")" & CHR$(10)
	Next
	oTextStream.closeInput()
	MsgBox s
End Sub
