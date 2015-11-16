Sub ExampleShell
	Dim rc As Long
	rc = Shell("C:\andy\TSEProWin\g32.exe", 2, "c:\Macro.txt")
	Print "I just returned and the returned code is " & rc ' rc = 0
	Rem These two have spaces in their names
	Shell("file:///C|/Andy/My%20Documents/oo/tmp/h.bat",2)
	Shell("C:\Andy\My%20Documents\oo\tmp\h.bat",2)
End Sub