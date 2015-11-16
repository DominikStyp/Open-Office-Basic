'Dynamic Data Exchange (DDE) is a mechanism that allows information to be shared between programs. Data may be updated in real time or it may work as a request response.

' Use DDE as a Calc function to extract cell A1 from a document.
=DDE("soffice";"/home/andy/tstdoc.xls";"a1") 'DDE in Calc to obtain a cell
='file:///home/andy/TST.sxc'#$sheet1.A1 'Direct reference to a cell

Sub ExampleDDE
	Dim nDDEChannel As Integer
	Dim s As String
	REM OOo must have the file open or the channel will not be opened
	nDDEChannel = DDEInitiate("soffice", "c:\TST.sxc")
	If nDDEChannel = 0 Then
		Print "Sorry, failed to open a DDE channel"
	Else
		Print "Using channel " & nDDEChannel & " to request cell A1"
		s = DDERequest(nDDEChannel, "A1")
		Print "Returned " & s
		DDETerminate(nDDEChannel)
	End If
End Sub