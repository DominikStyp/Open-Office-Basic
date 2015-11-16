' If the first argument is an array, the dimensions are determined.
' Special care is given to an empty array that was created using DimArray
' or Array.
' a : Variable to check
' sName : Name of the variable for a better looking string

Function arrayInfo(a, sName$) As String
	' First, verify that:
	' the variable is not NULL, an empty Object
	' the variable is not EMPTY, an uninitialized Variant
	' the variable is an array.
	If IsNull(a) Then
	arrayInfo = "Variable " & sName & " is Null"
	Exit Function
	End If
	If IsEmpty(a) Then
	arrayInfo = "Variable " & sName & " is Empty"
	Exit Function
	End If
	If Not IsArray(a) Then
	arrayInfo = "Variable " & sName & " is not an array"
	Exit Function
	End If
	' The variable is an array, so get ready to work
	Dim s As String 'Build the return value in s
	Dim iCurDim As Integer 'Current dimension
	Dim i%, j% 'Hold the LBound and UBound values
	On Error GoTo BadDimension 'Set up the error handler
	iCurDim = 1 'Ready to check the first dimension
	' Initial pretty return string
	s = "Array dimensioned as " & sName$ & "("
	Do While True 'Loop forever
	i = LBound(a(), iCurDim) 'Error if dimension is too large or
	j = UBound(a(), iCurDim) 'if invalid empty array
	If i > j Then Exit Do 'If empty array then get out
	If iCurDim > 1 Then s = s & ", " 'Separate dimensions with a comma
	s = s & i & " To " & j 'Add in the current dimensions
	iCurDim = iCurDim + 1 'Check the next dimension
	Loop
	' Only arrive here if the array is a valid empty array.
	' Otherwise, an error occurs when the dimension is too
	' large and a jump is made to the error handler
	' Include the type as returned from the TypeName function.
	' The type name includes a trailing "()" so remove this
	s = s & ") As " & Left(TypeName(a), LEN(TypeName(a))-2)
	arrayInfo = s
	Exit Function
	BadDimension:
	' Turn off the error handler
	On Error GoTo 0
	' Include the type as returned from the TypeName function.
	' The type name includes a trailing "()" so remove this
	s = s & ") As " & Left(TypeName(a), LEN(TypeName(a))-2)
	' If errored out on the first dimension then this must
	' be an invalid empty array.
	If iCurDim = 1 Then s = s & " *** INVALID Empty Array"
	arrayInfo = s
End Function


Sub UseArrayInfo
	Dim i As Integer, v
	Dim ia(1 To 3) As Integer
	Dim sa() As Single
	Dim m(3, 4, -4 To -1)
	Dim s As String
	s = s & arrayInfo(i, "i") & CHR$(10) 'Not an array
	s = s & arrayInfo(v, "v") & CHR$(10) 'Empty variant
	s = s & arrayInfo(sa(), "sa") & CHR$(10) 'Empty array
	s = s & arrayInfo(Array(), "Array") & CHR$(10) 'BAD empty array
	s = s & arrayInfo(ia(), "ia") & CHR$(10)
	s = s & arrayInfo(m(), "m") & CHR$(10)
	MsgBox s, 0, "Array Info"
End Sub