Option Explicit
Dim ExcelFileLocation, x
Wscript.Echo(TypeName(ExcelFileLocation))
ExcelFileLocation = "U:\WORK\misc\mip\MIPEval\data\Mip_all.xlsx"
Wscript.Echo(TypeName(ExcelFileLocation))
Wscript.Echo(Err.Number)

Function ExampleError
  On Error Resume Next
  x = 4/0
  Wscript.Echo(Err.Number)
  Wscript.Echo(x)
  
  Err.Clear
  x = 4/0
  Wscript.Echo(Err.Number)
End Function
ExampleError
  
Function DoNothing
  Wscript.Echo "AAA"
End Function
Function DoSomething
  Wscript.Echo "BBB"
End Function

Call DoSomething

Function ReadExcelFile(ByVal strFile)
  Dim objExcel, objWbook,objSheet,objCells
  Set objExcel = Wscript.CreateObject("Excel.Application")
  Set objWbook = objExcel.Workbooks.Open(strFile, False, True)
  Set objSheet = objWbook.Worksheets(1)
  objWbook.close()
  objExcel.quit()

End Function

ReadExcelFile(ExcelFileLocation)