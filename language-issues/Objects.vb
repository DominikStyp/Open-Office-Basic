
' Your own type <<<<<<<
' Your own type <<<<<<<
Type PersonType
	FirstName As String
	LastName As String
End Type
Sub ExampleCreateNewType
	Dim Person As PersonType
	Person.FirstName = "Andrew"
	Person.LastName = "Pitonyak"
	PrintPerson(Person)
End Sub
Sub PrintPerson(x)
	Print "Person = " & x.FirstName & " " & x.LastName
End Sub
Sub DefineObject()
	Dim x As New PersonType ' The original way to do it.
	Dim y As PersonType ' New is no longer required.
	Dim z : z = CreateObject("PersonType") ' Create the object when desired.
End Sub