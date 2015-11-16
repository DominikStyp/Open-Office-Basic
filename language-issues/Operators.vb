
Dim i As Integer
i = "abc" 'Assigning a string with no numbers yields zero not an error
Print i '0
i = "3abc" 'Assigns 3, automatically converts as it can.
Print i '3
Print 4 + "abc" '4

' >>>> Modulo operator <<<<<
10 MOD 3


'>>>>>>>> MATH OPERATORS AND CONCATENATION

Print 123 + "3" REM 126 (Numeric)
Print "123" + 3 REM 1233 (String)
Print 123 & "3" REM 1233 (String)
Print "123" & 3 REM 1233 (String)
Print 123 & 3 REM Use at least one string or it will not work!
Print 123 + "3" & 4 '1264 Do addition then convert to String
Print 123 & "3" + 4 '12334 Do addition first but first operand is String
Print 123 & 3 + "4" '1237 Do addition first but first operand is Integer



' AND, OR, XOR, NOT
' Perform a logical  operation on Boolean values, and a bit-wise on numerical values

' EQV
' The EQV operator is a question of equivalence: Are the two operands the same? A logical EQV operation is performed for Boolean values, and a bit-wise EQV on numbers. If both operands have the same value, the result is True. If the operands don’t have the same value, the result is False.
' True EQV True = True
' True EQV False = False
' False EQV True = False
' False EQV False = True
' 1100 EQV 1010 = 1001


' IMP
' The IMP operator performs a logical implication. A logical IMP operation is performed on Boolean values, and a bit-wise IMP on numbers. As the name implies, “x IMP y” asks if the statement that “x implies y” is a true statement.
' True IMP True = True
' True IMP False = False
' False IMP True = True
' False IMP False = True
' 1100 IMP 1010 = 1011

' >>>>>>>> COMPARISON OPERATORS
Dim a$, b$, c$
a$ = "A" : b$ = "B" : c$ = "B"
Print a$ < b$ 'True
Print b$ = c$ 'True
Print c$ <= a$ 'False

Dim a$, i%, t$
a$ = "A" : t$ = "3" : i% = 3
Print a$ < "B" 'True, String compare
Print "B" < a$ 'False, String compare
Print i% = "3" 'True, String compare
Print i% = "+3" 'False, String compare
Print 3 = t$ 'True, String compare
Print i% < "2" 'False, String compare
Print i% > "22" 'True, String compare

' >>> Problems with constants
Print "A" < "B" '0=False, this is not correct
Print "B" < "A" '-1=True, this is not correct
Print 3 = "3" 'False, but this changes if a variable is used