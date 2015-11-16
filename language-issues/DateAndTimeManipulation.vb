


' >>> CDate(expression) Date Convert a number or string to a date.
' >>> CDateFromIso(string) Date Convert to a date from an ISO 8601 date representation.
' >>> CDateToIso(date) String Convert a date to an ISO 8601 date representation.
' >>> Date() String Return the current date as a string.
' >>> DateAdd Date Add an interval to a date.
' >>> DateDiff Integer Returns the number of intervals between two dates.
' >>> DatePart Variant Obtain a specific part of a date value.
' >>> DateSerial(yr, mnth, day) Date Create a date from component pieces: Year, Month, Day.
' >>> DateValue(date) Date Extract the date from a date/time value by truncating the decimal portion.
' >>> Day(date) Integer Return the day of the month as an Integer from a Date value.
' >>> FormatDateTime String Format the date and time as a string. Requires OptionCompatible.
' >>> GetSystemTicks() Long Return the number of system ticks as a Long.
' >>> Hour(date) Integer Return the hour as an Integer from a Date value.
' >>> IsDate(value) Boolean Is this (value, converted to a string) a date?
' >>> Minute(date) Integer Return the minute as an Integer from a Date value.
' >>> Month(date) Integer Return the month as an Integer from a Date value.
' >>> MonthName String Return the name of the month based on an integer argument (1-12).
' >>> Now() Date Return the current date and time as a Date object.
' >>> Second(date) Integer Return the seconds as an Integer from a Date value.
' >>> Time() Date Return the time as a String in the format HH:MM:SS.
' >>> Timer() Date Return the number of seconds since midnight as a Date. Cast this to a Long.
' >>> TimeSerial(hour, min, sec) Date Create a date from component pieces: Hours, Minutes, Seconds.
' >>> TimeValue(“HH:MM:SS”) Date Extract the time value from a date; a pure time value between 0 and 1.
' >>> WeekDay(date) Integer Return the integer 1 through 7 corresponding to Sunday through Saturday.
' >>> WeekdayName String Return the day of the week based on an integer argument (1-7).
' >>> Year(date) Integer Return the year as an Integer from a Date value.

Sub ExampleDateType
	Dim tNow As Date, tToday As Date
	Dim tBirthDay As Date
	tNow = Now()
	tToday = Date()
	tBirthDay = DateSerial(1776, 7, 4)
	Print "Today = " & tToday
	Print "Now = " & tNow
	Print "A total of " & (tToday - tBirthDay) &_
	" days have passed since " & tBirthDay
End Sub


Sub PrintDate
	Print Date
	Print Time
	Print Now
End Sub

Sub IsDate
	Print IsDate("December 1, 1582 2:13:42") 'True
	Print IsDate("2:13:42") 'True
	Print IsDate("12/1/1582") 'True
	Print IsDate(Now) 'True
	Print IsDate("26:61:112") 'True: 112 seconds and 61 minutes!!!
	Print IsDate(True) 'False: Converts to string first
	Print IsDate(32686.22332) 'False: Converts to string first
	Print IsDate("02/29/2003") 'False: Only 28 days in February 03
End Sub

' Getting day of the week
Function WeekDayText(d) As String
	Select Case WeekDay(d)
	case 1
	WeekDayText="Sunday"
	case 2
	WeekDayText="Monday"
	case 3
	WeekDayText="Tuesday"
	case 4
	WeekDayText="Wednesday"
	case 5
	WeekDayText="Thursday"
	case 6
	WeekDayText="Friday"
	case 7
	WeekDayText="Saturday"
	End Select
End Function


' Getting month name
Sub ExampleMonthName
	Dim i%
	Dim s$
	For i = 1 To 12
	s = s & i & " = " & MonthName(i, True) & " = " & MonthName(i) & CHR$(10)
	Next
	MsgBox s, 0, "MonthName"
End Sub

' Getting day name
Sub ExampleWeekDayName
	Dim i%
	Dim s$
	CompatibilityMode(True)
	For i = 1 To 7
	s = s & i & " = " & WeekDayName(i, True) & " = " & WeekDayName(i) & CHR$(10)
	Next
	MsgBox s, 0, "WeekDayName"
End Sub


' Getting part of the date
Sub ExampleDatePart
	Dim TheDate As Date
	Dim f
	Dim i As Integer
	132
	Dim s$
	TheDate = Now
	' yyyy = year, q = quarter, m = month, y = day of the year
	' w = week day, ww = week of the year, d = day of the month
	' h = hour, n = minute, s = second
	f = Array("yyyy", "q", "m", "y", "w", "ww", "d", "h", "n", "s")
	s = "Now = " & TheDate & CHR$(10)
	For i = LBound(f) To UBound(f)
	s = s & "DatePart(" & f(i) & ", " & TheDate & ") = " & _
	DatePart(f(i), TheDate) & CHR$(10)
	Next
	MsgBox s
End Sub

' Formatting date and time
Sub ExampleFormatDateTime
	Dim s$, i%
	Dim d As Date
	d = Now
	CompatibilityMode(True)
	s = "FormatDateTime(d) = " & FormatDateTime(d) & CHR$(10)
	'0 = Default format with a short date and and a long time.
	'1 = Long date format with no time.
	'2 = Short date format.
	'3 = Time in the computer's regional settings.
	'4 = 24 hour hours and minutes as hh:mm.
	For i=0 To 4
		s = s & "FormatDateTime(d, " & i & ") = " & FormatDateTime(d, i) & CHR$(10)
	Next
	MsgBox s, 0, "FormatDateTime"
End Sub



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


' First day of the month
Function FirstDayOfMonth(d As Date) As Date
		FirstDayOfMonth() = DateSerial(Year(d), Month(d), 1)
End Function
' First day of this month
Sub FirstDayOfThisMonth()
		Dim d As Date
		d = FirstDayOfMonth(Now())
		MsgBox "First day of this month (" & d & ") is a " & WeekDayText(d)
End Sub

' Last day of the month
Function LastDayOfMonth(d As Date) As Date
	Dim nYear As Integer
	Dim nMonth As Integer
	nYear = Year(d) 'Current year
	nMonth = Month(d) + 1 'Next month, unless it was December.
	If nMonth > 12 Then 'If it is December then nMonth is now 13
	nMonth = 1 'Roll the month back to 1
	nYear = nYear + 1 'but increment the year
	End If
	LastDayOfMonth = CDate(DateSerial(nYear, nMonth, 1)-1)
End Function

' Last day of this month
Sub LastDayOfThisMonth()
	Dim d As Date
	d = LastDayOfMonth(Now())
	MsgBox "Last day of this month (" & d & ") is a " & WeekDayText(d)
End Sub



' Computing difference between dates (count difference between dates)
Sub ComputeDatesDiff
	Print DateAdd("d", 1, Now) 'Add one day.
	Print DateAdd("h", 1, Now) 'Add one hour.
	Print DateAdd("yyy", 1, Now) 'Add one year.
	Print DateDiff("yyyy", "03/13/1965", Date(Now)) 'Years from March 13, 1965 to now
	Print DateDiff("d", "03/13/1965", Date(Now)) 'Days from March 13, 1965 to now
End Sub

' Assembling date from parts
Sub AssembleDatesFromParts
	Print DateSerial(2003, 10, 1) '10/01/2003
	Print TimeSerial(13, 4, 45) '13:04:45
End Sub

' Elapsed time measurement
Sub ExampleElapsedTime
	Dim StartTicks As Long
	Dim EndTicks As Long
	Dim StartTime As Date
	Dim EndTime As Date
	StartTicks = GetSystemTicks()
	StartTime = Timer
	Wait(200) 'Pause execution for 0.2 seconds
	EndTicks = GetSystemTicks()
	EndTime = Timer
	MsgBox "After waiting 200 ms (0.2 seconds), " & CHR$(10) &_
	"System ticks = " & CStr(EndTicks - StartTicks) & CHR$(10) &_
	"Time elapsed = " & CStr((EndTime - StartTime)) &_
	" seconds" & CHR$(10), 0, "Elapsed Time"
End Sub
