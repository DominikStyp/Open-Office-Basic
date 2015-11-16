Sub ExampleCreateNewType
Dim Person As PersonType
With Person
	.FirstName = "Andrew"
	.LastName = "Pitonyak"
End With
...
End Sub


Dim oProp As New com.sun.star.beans.PropertyValue
With oProp
	.Name = "Person" 'Set Name Property
	.Value = "Boy Bill" 'Set Value Property
End With