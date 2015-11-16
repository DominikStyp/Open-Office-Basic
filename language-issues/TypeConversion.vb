Sub ExampleVal
	' >>> To numbers
	Print Val(" 12 34") '1234
	Print Val("12 + 34") '12
	Print Val("-1.23e4") '-12300
	Print Fix(12.2) ' 12
	Print Int("12.5") ' 12
	Print CInt("&H" & Hex(747)) '747
	Print Hex(447) '1BF
	Print Oct(877) '1555
	' >>> To String conversion
	Dim n As Long, d As Double, b As Boolean
	n = 999999999 : d = EXP(1.0) : b = False
	Print "X" & CStr(b) 'XFalse
	Print "X" & CStr(n) 'X999999999
	Print "X" & CStr(d) 'X2.71828182845904
	Print "X" & CStr(Now)'X06/09/2010 20:24:24 (almost exactly 7 years after 1st edition)
	
End Sub

Sub StringConv
	Dim i As Integer
	i = "abc" 'Assigning a string with no numbers yields zero not an error
	Print i '0
	i = "3abc" 'Assigns 3, automatically converts as it can.
	Print i '3
	Print 4 + "abc" '4
End Sub