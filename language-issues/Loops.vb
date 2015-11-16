' >>>>>>>>>>>>>>>>> LOOPS <<<<<<<<<<<<<

' >>>  DO-WHILE
Do While condition
	Block
		[Exit Do]
	Block
Loop ' Loop word ends the loop

' >>>  DO-UNTIL
Do Until condition
	Block
		[Exit Do]
	Block
Loop

' >>>  DO
Do
	Block
		[Exit Do]
	Block	
Loop 

' >>>  WHILE-UNTIL
While condition
Do
	Block
		[Exit Do]
	Block
Loop Until condition

' >>>  FOR
For counter=start To end [Step stepValue]
	statement block1
		[Exit For]
	statement block2
Next [counter]

' >>>>>>>>>>>>>>>>>>> LOOP EXAMPLES <<<<<<<<

a() = Array(2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30)
x = Int(32 * Rnd) REM random integer between 0 and 32
i% = LBound(a())
Do While i% <= UBound(a()) 'From lower to higher value of array
	Print a(i%)
	i% = i% + 1
	If i% > 1000 Then Exit Do
Loop
Do Until i% > UBound(a()) 'From lower to higher value of array
	Print a(i%)
	i% = i% + 1
Loop
For i = 1 To 4 Step 2
	Print i ' Prints 1 then 3
Next i ' The i in this statement is optional.







