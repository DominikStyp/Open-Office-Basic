'A pipe is an output stream and an input stream. Data written to the outputstream is buffered until it is read
'from the input stream. The Pipe service allows an outputstream to be converted into an input stream at the
'cost of an additional buffer. It is simple to create and close a pipe.

Function CreatePipe() As Object
	Dim oPipe ' Pipe Service.
	Dim oDataInp ' DataInputStream Service.
	Dim oDataOut ' DataOutputStream Service.
	oPipe = createUNOService ("com.sun.star.io.Pipe")
	oDataInp = createUNOService ("com.sun.star.io.DataInputStream")
	oDataOut = createUNOService ("com.sun.star.io.DataOutputStream")
	oDataInp.setInputStream(oPipe)
	oDataOut.setOutputStream(oPipe)
	CreatePipe = oPipe
End Function

Sub ClosePipe(oPipe)
	oPipe.Successor.closeInput
	oPipe.Predecessor.closeOutput
	oPipe.closeInput
	oPipe
End Sub