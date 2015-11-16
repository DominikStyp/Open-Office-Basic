Private Function defineVars()
	Dim str$				REM String var
	Dim str1 As String		REM String var
	REM Using postfix to define variables types
	Dim Vcurrency@, Vdouble#, Vinteger%, Vlong&, Vstring$, Vsingle!
	Dim d As Date, b As Boolean, v As Variant, o As Object
	REM 1) omitting type means type is Variant
	REM 2) Single type is Decimal numbers in the range of +/-3.402823 x 10E38
	REM 3) String is up to 65536 characters
	REM 4) Currency is number like 0.0000 (4 decimal places)
	' In OpenOffice following are acceptable
	' Boolean, Date, Double, Integer, Long, Object, Variant
	DefBool o1, DefDate o2, DefDbl o3, DefInt o4, DefLng o5, DefOjb o6, DefVar v
	' Defining constants
	Const Gravity = 9.81
End Function

REM is a comment
Sub ExampleBooleanType
	Dim b as Boolean
	Dim s as String
	b = True
	b = False
	b = (5 = 3) REM Set to False
	s = "(5 = 3) => " & b
	b = (5 < 7) REM Set to True
	s = s & CHR$(10) & "(5 < 7) => " & b
	b = 7 REM Set to True because 7 is not 0
	s = s & CHR$(10) & "(7) => " & b
	MsgBox s
End Sub


Sub ExampleIntegerType
	Dim i1 As Integer, i2% REM i1 and i2 are both Integer variables
	Dim f2 As Double
	Dim s$
	f2= 3.5
	38
	i1= f2 REM i1 is rounded to 4
	s = "3.50 => " & i1
	f2= 3.49
	i2= f2 REM i2 is rounded to 3
	s = s & CHR$(10) & "3.49 => " & i2
	MsgBox
End Sub

Sub ExampleLongType
	Dim NumberOfDogs&, NumberOfCats As Long ' Both variables are Long
	Dim f2 As Double
	Dim s$
	f2= 3.5
	NumberOfDogs = f2 REM round to 4
	s = "3.50 => " & NumberOfDogs
	f2= 3.49
	NumberOfCats = f2 REM round to 3
	s = s & CHR$(10) & "3.49 => " & NumberOfCats
	MsgBox s
End Sub

Sub ExampleCurrencyType
	Dim Income@, CostPerDog As Currency
	Income@ = 22134.37
	CostPerDog = 100.0 / 3.0
	REM Prints as 22134.3700
	Print "Income = " & Income@
	REM Prints as 33.3333
	Print "Cost Per dog = " & CostPerDog
End Sub

Sub ExampleSingleType
	Dim GallonsUsed As Single, Miles As Single, mpg!
	GallonsUsed = 17.3
	Miles = 542.9
	mpg! = Miles / GallonsUsed
	Print "Fuel efficiency = " & mpg!
End Sub

Sub ExampleDoubleType
	Dim GallonsUsed As Double, Miles As Double, mpg#
	GallonsUsed = 17.3
	Miles = 542.9
	mpg# = Miles / GallonsUsed
	Print "Fuel efficiency = " & mpg#
End Sub

Sub ExampleStringType
	Dim FirstName As String, LastName$
	FirstName = "Andrew"
	LastName$ = "Pitonyak"
	Print "Hello " & FirstName & " " & LastName$
End Sub


' >>> PUBLIC VARIABLES:
' Use Public to declare a variable that is visible to all modules in the declaring library container. Outside the declaring library container, the public variables aren’t visible. Public variables are initialized every time a macro runs.

' >>> PRIVATE VARIABLES:
' Private or Dim can declare a variable in a module that should not be visible
' in another module.
' - Declaring a variable using Dim is equivalent to declaring a variable as Private.
' - Private variables are only private, however, only with CompatibilityMode(True).
' - Option Compatible has no affect on private variables.


'Keyword     Initialized       Dies               Scope
'-----------------------------------------------------------------------------
'Global      Compile time      Compile time       All modules and libraries.
'Public      Macro start       Macro finish       Declaring library container.
'Dim         Macro start       Macro finish       Declaring library container.
'Private     Macro start       Macro finish       Declaring module.



' >>>>>>>>> VARIABLES TYPES 

Sub ExampleTypes
	Dim b As Boolean
	Dim c As Currency
	Dim t As Date
	Dim d As Double
	Dim i As Integer
	Dim l As Long
	Dim o As Object
	Dim f As Single
	Dim s As String
	Dim v As Variant
	Dim n As Variant
	Dim ta()
	Dim ss$
	n = null
End Sub

' >>>>>>>>>>>>>>>>>>>>>> VARIABLES INSPECTION

' >>> IsArray Is the variable an array?
' >>> IsDate Does the string contain a valid date?
' >>> IsEmpty Is the variable an empty Variant variable?
' >>> IsMissing Is the variable a missing argument?
' >>> IsNull Is the variable an unassigned object?
' >>> IsNumeric Does the string contain a valid number?
' >>> IsObject Is the variable an object?
' >>> IsUnoStruct Is the variable a UNO structure?
' >>> TypeLen Space used by the variable type.
' >>> TypeName Return the type name of the object as a String.
' >>> VarType Return the variable type as an Integer.
