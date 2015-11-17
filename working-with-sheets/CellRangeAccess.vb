
'Pick up Sheet1
'Get cells in range F13:F19
'Get cell's nr 1 value
' 
MsgBox ThisComponent.Sheets.getByName("Sheet1").getCellRangeByName("F13:F19").getCellByPosition( 0, 0 ).Value


' ALL UNO COMMANDS FOR fnDispatch() are here:
' https://wiki.openoffice.org/wiki/Framework/Article/OpenOffice.org_2.x_Commands

' Select current column (where cursor currently is
'    fnDispatch("SelectColumn")

' Insert String (where cursor currently is
'    fnDispatch("InsertText", array("Text","e"))
'    fnDispatch("EnterString",array("StringName","e"))


Sub proTest()
        Sheets("Sheet1").Select
        Range("C1").Select

        Do Until Selection.Offset(0, -2).Value = ""
                Selection.Value = Selection.Offset(0, -2).Value & " " & Selection.Offset(0, -1)
                Selection.Offset(1, 0).Select
        Loop
        Range("A1").Select

End Sub

' Return array with indexes 0 = row, 1  = column
' To get column number use:  MsgBox(getCurrentCellColumnAndRowNum()(1))
sub getCurrentCellColumnAndRowNum as Array
      Dim oCell,oDoc,column,row,addr
	  oDoc = ThisComponent
	  oCell = oDoc.getCurrentSelection
	  ' number of column indexed from 0
	  column = oCell.CellAddress.Column
	  ' number of row indexed from 0
	  row = oCell.CellAddress.Row
	  getCurrentCellColumnAndRowNum = Array(row,column)
end sub

' Gets current cell address like $F$10
sub getCurrentCellAbsoluteAddress as String
      Dim oCell,oDoc,arr,addr
	  oDoc = ThisComponent
	  oCell = oDoc.getCurrentSelection
	  'oCell.AbsoluteName gives full address like: $sheet1.$F$10
	  arr = Split(oCell.AbsoluteName,".")
	  addr = arr(1) ' here we have only $F$10 which we can put into goToCell function
	  getCurrentCellAbsoluteAddress = addr
end sub

' Goes to cell with specified address
sub goToCell(addr as String)
	fnDispatch("GoToCell", array("ToPoint",addr)) ' addr like "$J$1"
end sub


' Sort by specified column "J"
sub Sort
	SortBySpecifiedColumn("J")
end sub

' Automatically sorts by specified column
sub SortBySpecifiedColumn(column as String)
	Dim currentAddr
	' get current cell
	currentAddr = getCurrentCellAbsoluteAddress()
	' go to beginning of column like $J$1
	goToCell("$" + column + "$1")
    ' Sort column in descending order
	fnDispatch("SortDescending")
     ' return to previous cell
	goToCell(currentAddr)
end sub



' easily invoking UNO dispatcher
function fnDispatch(sCommand as string, optional mArgs)
       oFrame = ThisComponent.getCurrentController.getFrame
       oDispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
       'on error resume next
       if isMissing(mArgs) then
           fnDispatch = oDispatcher.executeDispatch(oFrame, ".uno:" & sCommand, "", 0, array())
       else
           nArgs = uBound(mArgs) \ 2
           dim Args(nArgs) as new com.sun.star.beans.PropertyValue
           for i = 0 to nArgs
               Args(i).name = mArgs(i * 2)
               Args(i).value = mArgs(i * 2 + 1)
           next
           fnDispatch = oDispatcher.executeDispatch(oFrame, ".uno:" & sCommand, "", 0, Args())
       end if
end function


Sub proFirst()
        Range("A1").Value = 34
        Range("A2").Value = 66
        Range("A3").Formula = "=A1+A2"
        Range("A1").Select
End Sub



' Reverse string
Function reverse(r As String) As String
    Dim t As String
    Dim c As String
    Dim i, l As Integer
   
    t = r
    l = Len(t)
    i = 1
   
    Do While i < l
        c = Mid(t, i, 1)
        Mid(t, i, 1) = Mid(t, l, 1)
        Mid(t, l, 1) = c
        i = i + 1
        l = l - 1
    Loop
    reverse = t
End Function


' Delete last line in every cell. USES: reverse()
sub deletelastline
   Dim oDoc As Object
   Dim oSelection As Object
   Dim s As String
   dim i As Long, j As Long
   Dim  c10 As Long
   
   oDoc = ThisComponent
   oSelection = oDoc.CurrentSelection
   if oSelection.supportsService("com.sun.star.sheet.SheetCellRange") then
      for i = 0 to oSelection.rows.count - 1
         for j = 0 to oSelection.columns.count - 1
            s = oSelection.getCellByPosition(j,i).String
            if Len(s) > 0 then
               s = reverse(s)
               c10 = InStr(s,Chr(10))
               if c10 > 0 then
                  s = Right(s,Len(s) - c10)
                  oSelection.getCellByPosition(j,i).setString(reverse(s))
               endif
            endif
         next j
      next i
   endif
               
end sub
