' >>>>>>>>>>>>>>>>>>>>>>>>>> STRINGS <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' >>>>>>>>>>>>>>>>>>>>>>>>>> STRINGS <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' >>>>>>>>>>>>>>>>>>>>>>>>>> STRINGS <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' >>>>>>>>>>>>>>>>>>>>>>>>>> STRINGS <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' >>>>>>>>>>>>>>>>>>>>>>>>>> STRINGS <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' >>>>>>>>>>>>>>>>>>>>>>>>>> STRINGS <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

' >>> ASC(str) Return the integer ASCII value of the first character in the string. This supports 16-bit Unicode values as well.
' >>> CHR(n) Convert an ASCII number to a character.
' >>> CStr(obj) Convert standard types to a string.
' >>> Format(obj, format) Fancy formatting; works only for strings.
' >>> Hex(n) Return the hexadecimal representation of a number as a string.
' >>> InStr(str, str)
' >>> InStr(start, str, str)
' >>> InStr(start, str, str, mode)
' Attempt to find string 2 in string 1. Returns 0 if not found and starting location if it is found. The optional start argument indicates where to start looking. The default value for mode is 1 (case-insensitive comparison). Setting mode to 0 performs a case-sensitive comparison.
' >>> InStrRev(str, find, start, mode)
' Return the position of the first occurrence of one string within another, starting from the right side of the string. Only available with “Option VBASupport 1”. Start and mode are optional.
' >>> Join(s())
' >>> Join(s(), str)
' Concatenate the array elements, separated by the optional string delimiter, and return the value as a string. The default delimiter is a single space. Inverse of the Split function.
' >>> LCase(str) Return a lowercase copy of the string.
' >>> Left(str, n) Return the leftmost n characters from the string.
' >>> Len(str) Return the length of the string.
' >>> LSet str1 = str2 Left-justify a string into the space taken by another string.
' >>> LTrim(str) Return a copy of the string with all leading spaces removed.
' >>> Mid(str, start)
' >>> Mid(str, start, len)
' >>> Mid(str, start, len, str)
' Return the substring, starting at start. If the length is omitted, the entire end of the string is returned. If the final string argument is included, this replaces the specified portion of the first string with the last string.
' >>> Oct(n) Return the octal representation of a number as a string.
' >>> Replace(str, find, rpl, start, count, mode)
' Search str for find and replace it with rpl. Optionally, specify the start, count, and mode.
' >>> Right(str, n) Return the rightmost n characters.
' >>> RSet str1 = str2 Right-justify a string into the space taken by another string.
' >>> RTrim(str) Return a copy of the string with all trailing spaces removed.
' >>> Space(n) Return a string with the number of specified spaces.
' >>> Split(str)
' >>> Split(str, str)
' Split a string into an array based on an optional delimiter. Inverse of the Join function.
' >>> Str(n) Convert a number to a string with no localization.
' >>> StrComp(s1, s2)
' >>> StrComp(s1, s2, mode)
' Compare two strings returning -1, 0, or 1 if the first string is less than, equal to, or greater than the second in alphabetical order. Set the optional third argument to zero for a case-insensitive comparison. The default is 1 for a case-sensitive comparison.
' >>> StrConv(str, mode[, local])
' Converts a string based on the mode: 1=upper, 2=lower, 4=wide, 8=narrow, 16=Katakana, 32=Hiragana, 64=to unicode, 128=from unicode.
' >>> String(n, char)
' >>> String(n, ascii)
' Return a string with a single character repeated multiple times. The first argument is the number of times to repeat; the second argument is the character or ASCII value.
' >>> StrReverse Reverse a string. Must use “Option VBASupport 1”, or precede it with CompatibiltyMode(True).
' >>> Trim(str) Return a copy of the string with all leading and trailing spaces removed.
' >>> UCase(str) Return an uppercase copy of the string.
' >>> Val(str) Convert a string to a double. This is very tolerant to non-numeric text.

' String comparison
Print StrComp( "A", "AA") '-1 because "A" < "AA"
Print StrComp("AA", "AA") ' 0 because "AA" = "AA"
Print StrComp("AA", "A") ' 1 because "AA" > "A"
Print StrComp( "a", "A") ' 1 because "a" > "A"
Print StrComp( "a", "A", 1)' 1 because "a" > "A"
Print StrComp( "a", "A", 0)' 0 because "a" = "A" if case is ignored

' Change letter case
Print LCase("Las Vegas") REM Returns "las vegas"
Print UCase("Las Vegas") REM Returns "LAS VEGAS"

' Searching substring in string
Print InStr("CBAABC", "abc") '4 default to case insensitive
Print InStr(1, "CBAABC", "b") '2 first argument is 1 by default
Print InStr(2, "CBAABC", "b") '2 start with second character
Print InStr(3, "CBAABC", "b") '5 start with third character
Print InStr(1, "CBAABC", "b", 0) '0 case-sensitive comparison
Print InStr(1, "CBAABC", "b", 1) '2 case-insensitive comparison
Print InStr(1, "CBAABC", "B", 0) '2 case-sensitive comparison

' Extract substring from string
Print Left("12345", 8) '12345
Print Right("12345", 2) '45
Print Mid("123456", 3) '3456
Print Mid("123456", 3, 2) '34
Print Mid("123456789", 3, 5, "XX") 'Replace five characters with two: "12XX89"
Print Mid("123456789", 7, 12, "ABCDEFG") 'Add more than remove: "123456ABCDEFG"

' Align string
Dim s As String 'String variable to contain the result
s = String(10, "*") 'The result is 10 characters wide
RSet s = CStr(1.23) 'The number is not automatically converted to a string
LSet ss = s
Print "$" & s '$      1.23
Print ss & "$" '1.22     $

' Formatting output
Print Format(1223, "00.00") '1223.00
Print Format(1234.56789, "###00.00") '1234.57
' Formatting numbers
Sub ExampleFormat
	MsgBox Format(6328.2, "##,##0.00") REM 6,328.20
	MsgBox Format(123456789.5555, "##,##0.00") REM 123,456,789.56
	MsgBox Format(0.555, ".##") REM .56
	MsgBox Format(123.555, "#.##") REM 123.56
	MsgBox Format(123.555, ".##") REM 123.56
	MsgBox Format(0.555, "0.##") REM 0.56
	MsgBox Format(0.1255555, "%#.##") REM %12.56
	MsgBox Format(123.45678, "##E-####") REM 12E1
	MsgBox Format(.0012345678, "0.0E-####") REM 1.2E-003
	MsgBox Format(123.45678, "#.e-###") REM 1.e002
	MsgBox Format(.0012345678, "#.e-###") REM 1.e-003
	MsgBox Format(123.456789, "#.## is ###") REM 123.46
	MsgBox Format(8123.456789, "General Number") REM 8123.456789
	MsgBox Format(8123.456789, "Fixed") REM 8123.46
	MsgBox Format(8123.456789, "Currency") REM 8,123.46 $ (broken)
	MsgBox Format(8123.456789, "Standard") REM 8,123.46
	MsgBox Format(8123.456789, "Scientific") REM 8.12E+03
	MsgBox Format(0.00123456789, "Scientific") REM 1.23E-03
	MsgBox Format(0.00123456789, "Percent") REM 0.12%
End Sub



'Formatting dates

'>>> q The quarter of the year (1 through 4).
'>>> qq The quarter of the year as 1st quarter through 4th quarter
'>>> y The day in the year (1 through 365).
'>>> yy Two-digit year.
'>>> yyyy Complete four-digit year.
'>>> m Month number with no leading zero.
'>>> mm Two-digit month number; leading zeros are added as required.
'>>> mmm Month name abbreviated to three letters.
'>>> mmmm Full month name as text.
'>>> mmmmm First letter of month name.
'>>> d Day of the month with no leading zero.
'>>> dd Day of the month; leading zeros are added as required.
'>>> ddd Day as text abbreviated to three letters (Sun, Mon, Tue, Wed, Thu, Fri, Sat).
'>>> dddd Day as text (Sunday through Saturday).
'>>> ddddd Full date in a short date format.
'>>> dddddd Full date in a long format.
'>>> w Day of the week as returned by WeekDay (1 through 7).
'>>> ww Week in the year (1 though 52).
'>>> h Hour with no leading zero.
'>>> hh Two-digit hour; leading zeros are added as required.
'>>> n Minute with no leading zero.
'>>> nn Two-digit minute; leading zeros are added as required.
'>>> s Second with no leading zero.
'>>> ss Two-digit second; leading zeros are added as required.
'>>> ttttt Display complete time in a long time format.
'>>> c Display a complete date and time.
'>>> / Date separator. A locale-specific value is used.
'>>> : Time separator. A locale-specific value is used.

Sub FormatDateTimeStrings
	Dim i%
	Dim d As Date
	d = now()
	Dim s$
	Dim formats
	formats = Array("q", "qq", "y", "yy", "yyyy", _
	"m", "mm", "mmm", "mmmm", "mmmmm", _
	"d", "dd", "ddd", "dddd", "ddddd", "dddddd", _
	"w", "ww", "h", "hh", "n", "nn", "nnn", "s", "ss", _
	"ttttt", "c", "d/mmmm/yyyy h:nn:ss")
	For i = LBound(formats) To UBound(formats)
	s = s & formats(i) & " => " & Format(d, formats(i)) & CHR$(10)
	Next
	MsgBox s
End Sub

' String format specifiers
' < String in lowercase.
' > String in uppercase.
' @ Character placeholder. If the input character is empty, place a space in the outputstring. For example, “(@@@)” formats to “( )” with an empty string.
' & Character placeholder. If the input character is empty, place nothing in the output string. For example, “(&&&)” formats to “()” with an empty string.
' ! Normally, character placeholders are filled right to left; the ! forces the placeholders to be filled left to right.

Sub FormatStrings
	Dim i%
	Dim s$
	Dim formats
	formats = Array("<", ">", _
	"@@", "(@@@)", "[@@@@]", _
	"&&", "(&&&)", "[&&&&]", _
	)
	For i = LBound(formats) To UBound(formats)
	s = s & formats(i) & " => (" & Format("On", formats(i)) & ")" & CHR$(10)
	Next
	MsgBox s
End Sub