'ABS(number) The absolute value of a specified number.
'ATN(number) The angle, in radians, whose tangent is the specified number in the range of -Pi/2 through
'            Pi/2.
'CByte(expression) Round the String or numeric expression to a Byte.
'CCur(expression) Convert the expression to a Currency type.
'CDbl(expression) Convert a String or numeric expression to a Double.
'CDec(expression) Generate a Decimal type; implemented only on Windows.
'CInt(expression) Round the String or numeric expression to the nearest Integer.
'CLng(expression) Round the String or numeric expression to the nearest Long.
'COS(number) The cosine of the specified angle.
'CSng(expression) Convert a String or numeric expression to a Single.
'Exp(number) The base of natural logarithms raised to a power.
'Fix(number) Chop off the decimal portion.
'Format(obj, format) Fancy formatting, discussed in Chapter 6, “String Routines.”
'Hex(n)      Return the hexadecimal representation of a number as a String.
'Int(number) Round the number toward negative infinity.
'Log(number) The logarithm of a number. In Visual Basic .NET this method can be overloaded to return
'             either the natural (base e) logarithm or the logarithm of a specified base.
'Oct(number) Return the octal representation of a number as a String.
'Randomize(num) Initialize the random number generator. If num is ommitted, uses the system timer.
'Rnd         Return a random number as a Double from 0 through 1.
'Sgn(number) Integer value indicating the sign of a number.
'SIN(number) The sine of an angle.
'Sqr(number) The square root of a number.
'Str(number) Convert a number to a String with no localization.
'TAN(number) The tangent of an angle.
'Val(str)    Convert a String to a Double. This is very tolerant to non-numeric text.

' >>>>>>> Trigonometric examples
degrees = (radians * 180) / Pi
radians = (degrees * Pi) / 180
radians = (45° * Pi) / 180 = Pi / 4 = 3.141592654 / 4 = 0.785398163398

Sub ExampleTrigonometric
	Dim OppositeLeg As Double
	Dim AdjacentLeg As Double
	Dim Hypotenuse As Double
	Dim AngleInRadians As Double
	Dim AngleInDegrees As Double
	Dim s As String
	OppositeLeg = 3
	AdjacentLeg = 4
	AngleInRadians = ATN(3/4)
	AngleInDegrees = AngleInRadians * 180 / Pi
	s = "Opposite Leg = " & OppositeLeg & CHR$(10) &_
	"Adjacent Leg = " & AdjacentLeg & CHR$(10) &_
	"Angle in degrees from ATN = " & AngleInDegrees & CHR$(10) &_
	"Hypotenuse from COS = " & AdjacentLeg/COS(AngleInRadians) & CHR$(10) &_
	"Hypotenuse from SIN = " & OppositeLeg/SIN(AngleInRadians) & CHR$(10) &_
	"Opposite Leg from TAN = " & AdjacentLeg * TAN(AngleInRadians)
	MsgBox s, 0, "Trigonometric Functions"
End Sub

' >>>>>> Random number in range
Function RndRange(lowerBound As Double, upperBound As Double) As Double
	RndRange = lowerBound + Rnd() * (upperBound - lowerBound)
End Function





