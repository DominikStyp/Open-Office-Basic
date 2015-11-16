' >> Array(args) Return a Variant array that contains the arguments.
' >> DimArray(args) Return an empty Variant array. The arguments specify the dimension.
' >> IsArray(var) Return True if this variable is an array, False otherwise.
' >> Join(array)
' >> Join(array, delimiter)
' >> Concatenate the array elements separated by the optional string delimiter and return as a String. The default delimiter is a single space.
' >> LBound(array)
' >> LBound(array, dimension)
' Return the lower bound of the array argument. The optional dimension specifies which dimension to check. The first dimension is 1.
' >> ReDim var(args) As Type Change the dimension of an array using the same syntax as the DIM statement. The keyword Preserve keeps existing data intact — ReDim Preserve x(1 To 4) As Integer.
' >> Split(str)
' >> Split(str, delimiter)
' >> Split(str, delimiter, n)
' Split the string argument into an array of strings. The default delimiter is a space. The optional argument “n” limits the number of strings returned.
' >> UBound(array)
' >> UBound(array, dimension) Return the upper bound of the array argument. The optional dimension specifies which dimension to check. The first dimension is 1.

Sub DeclaringArray
	Dim a(3) As Integer 'This array has four integer values, they are all zero
	Dim b(3) 'This array has four Variants, they are all Empty
	Dim c() 'This array has one dimension and no space Ubound < Lbound
	v = Array() 'This array has zero dimensions.
End Sub
Sub ExampleSimpleArray1
	Dim a(2) As Integer, b(-2 To 1) As Long
	Dim m(1 To 2, 3 To 4)
	REM Did you know that multiple statements can be placed
	REM on a single line if separated by a colon?
	a(0) = 0 : a(1) = 1 : a(2) = 2
	b(-2) = -2 : b(-1) = -1 : b(0) = 0 : b(1) = 1
	m(1, 3) = 3 : m(1, 4) = 4
	m(2, 3) = 6 : m(2, 4) = 8
	Print "m(2,3) = " & m(2,3)
	Print "b(-2) = " & b(-2)
End Sub

Sub ExampleArrayFunction
	Dim a, i%, s$
	a = Array("Zero", 1, Pi, Now)
	Rem String, Integer, Double, Date
	For i = LBound(a) To UBound(a)
	s$ = s$ & i & " : " & TypeName(a(i)) & " : " & a(i) & CHR$(10)
	Next
	MsgBox s$, 0, "Example of the Array Function"
End Sub

' Redimension array 
Sub ExampleDimArray
	Dim a(), i%
	Dim s$
	a = Array(10, 11, 12)
	s = "" & LBound(a()) & " " & UBound(a()) Rem 0 2
	a() = DimArray(3) REM the same as Dim a(3)
	a() = DimArray(2, 1) REM the same as Dim a(2,1)
	i = 4
	a = DimArray(3, i) Rem the same as Dim a(3,4)
	s = s & CHR$(10) & LBound(a(),1) & " " & UBound(a(),1) Rem 0, 3
	s = s & CHR$(10) & LBound(a(),2) & " " & UBound(a(),2) Rem 0, 4
	a() = DimArray() REM an empty array
	MsgBox s, 0, "Example Dim Array"
End Sub

' How to preserve data while truncating array
Sub RedimensionWithDataPreserve
	Dim a() As Integer
	ReDim a(3, 3, 3) As Integer
	a(1,1,1) = 1 : a(1, 1, 2) = 2 : a(2, 1, 1) = 3
	ReDim preserve a(-1 To 4, 4, 4) As Integer
	Print "(" & a(1,1,1) & ", " & a(1, 1, 2) & ", " & a(2, 1, 1) & ")"
End Sub

' Convert array to string
Function ArrayToString(a() As Variant) As String
	Dim i%, s$
	For i% = LBound(a()) To UBound(a())
		s$ = s$ & i% & " : " & a(i%) & CHR$(10)
	Next
	ArrayToString = s$
End Function

' Array are passed by reference, integers by value
Sub ExampleArrayCopyIsRef
	Dim a(5) As Integer, c(4) As Integer, s$
	c(0) = 4 : c(1) = 3 : c(2) = 2 : c(3) = 1 : c(4) = 0
	a() = c()
	a(1) = 7
	c(2) = 10
	s$ = "**** a() *****" & CHR$(10) & ArrayToString(a()) & CHR$(10) &_
	CHR$(10) & "**** c() *****" & CHR$(10) & ArrayToString(c())
	MsgBox s$, 0 , "Change One, Change Both"
End Sub

' Dimension the first array to have the same dimensions as the second.
' Perform an element-by-element copy of the array.
Sub SetIntArray(iArray() As Integer, v() As Variant)
	Dim i As Long
	ReDim iArray(LBound(v()) To UBound(v())) As Integer
	For i = LBound(v) To UBound(v)
	iArray(i) = v(i)
	Next
End Sub

' Multidimensional Arrays 
Sub CreateMultidimensionalArray
	i% = 7
	v = DimArray(3*i%) 'Same as Dim v(0 To 21)
	v = DimArray(i%, 4) 'Same as Dim v(0 To 7, 0 To 4)
	v = DimArray(1, 2)
	v(0, 0) = 1 : v(0, 1) = 2 : v(0, 2) = 3
	v(1, 0) = "one" : v(1, 1) = "two" : v(1, 2) = "three"
End Sub
