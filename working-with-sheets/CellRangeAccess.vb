
'Pick up Sheet1
'Get cells in range F13:F19
'Get cell's nr 1 value
' 
MsgBox ThisComponent.Sheets.getByName("Sheet1").getCellRangeByName("F13:F19").getCellByPosition( 0, 0 ).Value



Sub proTest()
        Sheets("Sheet1").Select
        Range("C1").Select

        Do Until Selection.Offset(0, -2).Value = ""
                Selection.Value = Selection.Offset(0, -2).Value & " " & Selection.Offset(0, -1)
                Selection.Offset(1, 0).Select
        Loop
        Range("A1").Select

End Sub


sub getCurrentCellColumnAndRowNum as String
      Dim oCell,oDoc,column,row,addr
	  oDoc = ThisComponent
	  oCell = oDoc.getCurrentSelection
	  ' number of column indexed from 0
	  column = oCell.CellAddress.Column
	  ' number of row indexed from 0
	  row = oCell.CellAddress.Row
end sub

sub getCurrentCellAbsoluteAddress as String
      Dim oCell,oDoc,arr,addr
	  oDoc = ThisComponent
	  oCell = oDoc.getCurrentSelection
	  'oCell.AbsoluteName gives full address like: $sheet1.$F$10
	  arr = Split(oCell.AbsoluteName,".")
	  addr = arr(1) ' here we have only $F$10 which we can put into goToCell function
	  getCurrentCellAddress = addr
end sub

sub goToCell(addr as String)
	fnDispatch("GoToCell", array("ToPoint",addr)) ' addr like "$J$1"
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