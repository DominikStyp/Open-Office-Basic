
'DLLs can by only used in Windows

Declare Sub MyMessageBeep Lib "user32.dll" Alias "MessageBeep" ( Long )
Declare Function CharUpper Lib "user32.dll" Alias "CharUpperA"_
(ByVal lpsz As String) As String
Sub ExampleCallDLL
	REM Convert a string to uppercase
	Dim strIn As String
	Dim strOut As String
	strIn = "i Have Upper and Lower"
	strOut = CharUpper(strIn)
	MsgBox "Converted:" & CHR$(10) & strIn & CHR$(10) &_
	"To:" & CHR$(10) & strOut, 0, "Call a DLL Function"
	REM On my computer, this plays a system sound
	Dim nBeepLen As Long
	nBeepLen = 5000
	MyMessageBeep(nBeepLen)
	FreeLibrary("user32.dll" )
End Sub