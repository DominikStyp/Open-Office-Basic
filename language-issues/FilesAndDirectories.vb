
' >>> ChDir(path) Change the currently logged directory or drive. Deprecated; do not use.
' >>> ChDrive(path) Change the currently logged drive. Deprecated; do not use.
' >>> Close #n Close a previously opened file or files. Separate file numbers with a comma.
' >>> ConvertFromURL(str) Convert a path expressed as a URL to a system-specific path.
' >>> ConvertToURL(str) Convert a system-specific path to a URL.
' >>> CurDir
' >>> CurDir(drive) Return the current directory as a system-specific path. If the optional drive is specified, the current directory for the specified drive is returned.
' >>> Dir(path)
' >>> Dir(path, attr) Return a listing of files based on an included path. The path may contain a file specification - for example, “/home/andy/*.txt”. Optional attributes determine if a listing of files or directories is returned.
' >>> EOF(number) Return True if the file denoted by “number” is at the end of the file.
' >>> FileAttr(number, 1) Return the mode used to open the file given by “number”. The second argument specifies if the file-access or the operating-system mode is desired, but only the file mode is currently supported.
' >>> FileCopy(src, dest) Copy a file from source to destination.
' >>> FileDateTime(path) Return the date and time of the specified file as a string.
' >>> FileExists(path) Return True if the specified file or directory exists.
' >>> FileLen(path) Return the length of the specified file as a long.
' >>> FreeFile() Return the next available file number for use.
' >>> Get #number, variable
' >>> Get #number, pos, variable
' Read a record from a relative file, or a sequence of bytes from a binary file, into a variable. If the position argument is omitted, data is read from the current position in the file. For files opened in binary mode, the position is the byte position in the file.
' >>> GetAttr(path) Return a bit pattern identifying the file type.
' >>> GetPathSeparator() Return the system-specific path separator.
' >>> Input #number, var Sequentially read numeric or string records from an open file and assign the data to one or more variables. The carriage return (ASC=13), line feed (ASC=10), and comma act as delimiters. Input cannot read commas or quotation marks (") because they delimit the text. Use the Line Input statement if you must do this.
' >>> Kill(path) Delete a file from disk.
' >>> Line Input #number, var Sequentially read strings to a variable line-by-line up to the first carriage return (ASC=13) or line feed (ASC=10). Line end marks are not returned.
' >>> Loc(number) Return the current position in an open file.
' >>> LOF(number) Return the size of an open file, in bytes.
' >>> MkDir(path) Create the directory.
' >>> Name src As dest Rename a file or directory.
' >>> Open path For Mode As #n Open a data channel (file) for Mode (Input = read, Output = write)
' >>> Put #n, var
' >>> Put #n, pos, var
' >>> Write a record to a relative file or a sequence of bytes to a binary file.
' >>> Reset Close all open files and flush all files to disk.
' >>> RmDir(path) Remove a directory.
' >>> Seek #n, pos Set the position for the next writing or reading in a file.
' >>> SetAttr(path, attr) Set file attributes.
' >>> Write #n, string Write data to a file.

Sub ToFromURL
	Print ConvertToURL("/home/andy/logo.miff")
	Print ConvertFromURL("file:///home/andy/logo.miff") 'This requires UNIX
	Print ConvertToURL("c:\My Documents") 'This requires Windows
	Print ConvertFromURL("file:///c:/My%20Documents") 'This requires windows
End Sub

' Invoke shell command
Shell("C:\Windows\System32\calc.exe")


'Create and remove directories in OOME Work Directory
Sub ExampleCreateRmDirs
	If NOT CreateOOMEWorkDir() Then
	Exit Sub
	End If
	Dim sWorkDir$
	Dim sPath$
	sWorkDir = OOMEWorkDir()
	sPath = sWorkDir & "a" & GetPathSeparator() & "b"
	MkDir sPath
	Print "Created " & sPath
	RmOOMEWorkDir()
	Print "Removed " & sWorkDir
End Sub

' Rename file or directory
Name "C:\Joe.txt" As "C:\bill.txt" 'Rename a file
Name "C:\logs" As "C:\oldlogs" 'Rename a directory


' List files in directory (simple version)
Sub SimpleDirectoryListing
	sFileName = Dir("C:\", 0)
	Do While( sFileName <> "" )
		Print sFileName
		sFileName = Dir() 'Works like iterator
	Loop
End Sub


' List files in directory (advanced version)
Sub AdvancedDirectoryListing
	Dim s As String 'Temporary string
	Dim sFileName As String 'Last name returned from DIR
	Dim i As Integer 'Count number of dirs and files
	Dim sPath 'Current path with path separator at end
	sPath = CurDir & GetPathSeparator() 'With no separator, DIR returns the
	sFileName = Dir(sPath, 16) 'directory rather than what it contains
	i = 0 'Initialize the variable
	Do While (sFileName <> "") 'While something returned
		i = i + 1 'Count the directories
		s = s & "Dir " & CStr(i) &_
		" = " & sFileName & CHR$(10) 'Store in string for later printing
		sFileName = Dir() 'Get the next directory name
	Loop
	i = 0 'Start counting over for files
	sFileName = Dir(sPath, 0) 'Get files this time!
	Do While (sFileName <> "")
		i = i + 1
		s = s & "File " & CStr(i) & " = " & sFileName & " " &_
		PrettyFileLen(sPath & sFileName) & CHR$(10)
		sFileName = Dir()
	Loop
	MsgBox s, 0, ConvertToURL(sPath)
End Sub



' Write to file 30 lines (simple)
Sub WriteDataToFile
	FileName = "C:/!!TEST.txt"
	n = FreeFile() 'Next free file number
	Open FileName For Output Access Read Write As #n 'Open for read/write
	For i=1 To 30
		' File handle ALWAYS must begin on #
		Write #n "Test line " & i
	Next
	Close #n
Sub

' Read data from file (simple)
Sub ReadDataFromFile
	FileName = "C:/!!TEST.txt"
	n = FreeFile() 'Next free file number
	Dim s$, tmp$
	Open FileName For Input Access Read As #n 'Open for read
	Do While (NOT EOF(n)) ' Until loop reaches end of file
		' File handle ALWAYS must begin on #
		Input #n, tmp$ 'Read string to temporary variable
		s$ = s$ & CHR$(10) & tmp$
	Loop
	MsgBox s, 0, "Contents of the file"
	Close #n
End Sub
	
	
' Reading and writing data to file (advanced)
Sub WriteExampleGetOpenFileInfo
	Dim FileName As String 'Holds the file name
	Dim n As Integer 'Holds the file number
	Dim i As Integer 'Index variable
	Dim s As String 'Temporary string for input
	FileName = ConvertToURL(CurDir) & "/delme.txt"
	n = FreeFile() 'Next free file number
	Open FileName For Output Access Read Write As #n 'Open for read/write
	For i = 1 To 15032 'Write a lot of data
		Write #n, "This is line ",CStr(i),"or",i 'Write some text
	Next
	Seek #n, 1022 'Move the file pointer to location 1022
	For i = 1 To 100 'Read 100 pieces of data; this will set Loc
		Input #n, s 'Read one piece of data into the variable s
	Next
	MsgBox GetOpenFileInfo(n), 0, FileName
	Close #n
	Kill(FileName) 'Delete this file, I do not want it
End Sub


Function GetOpenFileInfo(n As Integer) As String
	Dim s As String
	Dim iAttr As Integer
	On Error GoTo BadFileNumber
	iAttr = FileAttr(n, 1)
	If iAttr = 0 Then
		s = "File handle " & CStr(n) & " is not currently open" & CHR$(10)
	Else
		s = "File handle " & CStr(n) & " was opened in mode:"
		If (iAttr AND 1) = 1 Then s = s & " Input"
		If (iAttr AND 2) = 2 Then s = s & " Output"
		If (iAttr AND 4) = 4 Then s = s & " Random"
		If (iAttr AND 8) = 8 Then s = s & " Append"
		If (iAttr AND 16) = 16 Then s = s & " Binary"
		iAttr = iAttr AND NOT (1 OR 2 OR 4 OR 8 OR 16)
		If iAttr AND NOT (1 OR 2 OR 4 OR 8 OR 16) <> 0 Then
			s = s & " unsupported attribute " & CStr(iAttr)
		End If
		s = s & CHR$(10)
		s = s & "File length = " & nPrettyFileLen(LOF(n)) & CHR$(10)
		s = s & "File location = " & CStr(LOC(n)) & CHR$(10)
		s = s & "Seek = " & CStr(Seek(n)) & CHR$(10)
		s = s & "End Of File = " & CStr(EOF(n)) & CHR$(10)
	End If
	AllDone:
		On Error GoTo 0
		GetOpenFileInfo = s
		Exit Function
	BadFileNumber:
		s = s & "Error with file handle " & CStr(n) & CHR$(10) &_
			"The file is probably not open" & CHR$(10) & Error()
		Resume AllDone
End Function

' Substitute path variables into string and re-substitute
Sub UsePathReSubstitution()
	Dim oPathSub ' PathSubstitution service.
	Dim s$ ' Accumulate the value to print.
	Dim sTemp$
	oPathSub = CreateUnoService( "com.sun.star.util.PathSubstitution" )
	' There are two variables to substitute.
	' False means do not generate an error
	' if an unknown variable is used.
	s = "$(temp)/OOME/ or $(work)"
	sTemp = oPathSub.substituteVariables(s, False)
	s = s & " = " & sTemp & CHR$(10)
	' This direction encodes the entire thing as though it were a single
	' path. This means that spaces are encoded in URL notation.
	s = s & sTemp & " = " & oPathSub.reSubstituteVariables(sTemp) & CHR$(10)
	MsgBox s
End Sub
