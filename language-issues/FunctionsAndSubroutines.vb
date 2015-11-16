' >>>>>>>>>>>>> SCOPE OF VARIABLES <<<<<<<<< 
' >>>>>>>>>>>>> SCOPE OF VARIABLES <<<<<<<<< 
' >>>>>>>>>>>>> SCOPE OF VARIABLES <<<<<<<<< 
' >>>>>>>>>>>>> SCOPE OF VARIABLES <<<<<<<<< 

' >>> LOCAL VARIABLES: Variables declared inside a subroutine or function are called local variables.


' DIFFERENCE BETWEEN SUBROUTINE AND FUNCTION
'A function is a subroutine that returns a value.

' >>> STATIC VARIABLES:
Sub ExampleStaticWorker
	Static iStatic1 As Integer
	Dim iNonStatic As Integer
	iNonStatic = iNonStatic + 1
	iStatic1 = iStatic1 + 1
	Msgbox "iNonStatic = " & iNonStatic & CHR$(10) &_
	"iStatic1 = " & iStatic1
End Sub

' >>>>>> GLOBAL VARIABLES:
' Global iNumberOfTimesRun
' Use Global to declare a variable that is available to every module in every library. The library containing the
' Global variable must be loaded for the variable to be visible.

' >>>>>> PRIVATE/PUBLIC  Subroutines/Functions

Private Sub PrivSub
	Print "In Private Sub"
	bbxx = 4
End Sub

'Using Option Compatibile is not sufficient to enable Private scope, CompatibilityMode(True) must be used.
Sub TestPrivateSub
  CompatibilityMode(False) 'Required only if CompatibilityMode(True) already used.
  Call PrivSub() 'This call works.
  CompatibilityMode(True) 'This is required, even if Option Compatible is used
  Call PrivSub() 'Runtime error (if PrivSub is in a different module).
End Sub




' >>>>>>>>> PASS ARGUMENTS BY REFERENCE VS BY VALUE
Sub ExampleArgumentValAndRef()
	Dim i1%, i2%
	i1 = 1 : i2 = 1
	ArgumentValAndRef(i1, i2)
	MsgBox "Argument passed by reference was 1 and is now " & i1 & CHR$(10) &_
	"Argument passed by value was 1 and is still " & i2 & CHR$(10)
End Sub
Sub ArgumentValAndRef(iRef%, ByVal iVal)
	iRef = iRef + 1 ' This will affect the caller
	iVal = iVal - 1 ' This will not affect the caller
End Sub


' >>>>>>>>>  Optional parameters and parameters existence check
Sub Test(A As Integer, Optional B As Integer)
  Dim B_Local As Integer
  ' Check whether B parameter is actually present         
  If Not IsMissing (B) Then   
    B_Local = B      ' B parameter present
  Else
    B_Local = 0      ' B parameter missing -> default value 0
  End If
  ' ... Start the actual function
End Sub

' >>>>>>>>> Predefined optional parameters
Sub DefaultExample(Optional n as Integer=100)
	REM If IsMissing(n) Then n = 100 'I will not have to do this anymore!
	Print n
End Sub



' FUNCTIONS >>>>>>>>>>>>>>>>>>>>>>

' Recursive call: This finally works in version 1.1
Function RecursiveFactorial(ByVal n As Long) As Long
	RecursiveFactorial = 1
	If n > 1 Then RecursiveFactorial = n * RecursiveFactorial(n-1)
End Function



