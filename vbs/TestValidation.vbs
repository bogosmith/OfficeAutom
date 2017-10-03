Option Explicit
Dim  objExcel, objWbook1, objSheet1, objWbook2, objSheet2, rng, path1, path2, dataSheetName, pwd
path1 = "U:\WORK\develop\trunk\vbs\RIPtest.xlsm"
path2 = "U:\WORK\develop\trunk\vbs\EDFtest.xlsm"
dataSheetName = "Transactions"

Sub copyValidation(objSheet1, objSheet2, rng, pwd)
Dim val, fml
Set val = objSheet1.Range(rng).Validation
fml = val.Formula1
Wscript.echo fml

objSheet2.unprotect pwd
Set val = objSheet2.Range(rng).Validation
val.Delete
' expression.Add(Type, AlertStyle, Operator, Formula1, Formula2)
' where type 3 means drop-down list, see https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xldvtype-enumeration-excel
'val.Add 3,,,"=M3:M11"
val.Add 3,,,fml
End Sub

Wscript.StdOut.Write "Enter password: "
pwd = Wscript.StdIn.ReadLine

Set objExcel = Wscript.CreateObject("Excel.Application")
Set objWbook1 = objExcel.Workbooks.Open(path1, False, True)
Set objSheet1 = objWbook1.Worksheets(dataSheetName)
Set objWbook2 = objExcel.Workbooks.Open(path2, False, False)
Set objSheet2 = objWbook2.Worksheets(dataSheetName)
copyValidation objSheet1, objSheet2, "AS10:AS1999", pwd
'copyValidation objSheet1, objSheet2, "AW10:AW1999", pwd
copyValidation objSheet1, objSheet2, "P10:P1999", pwd
copyValidation objSheet1, objSheet2, "BK10:BK999", pwd
'copyValidation objSheet1, objSheet2, "P10", pwd
objWbook2.Save
objWbook1.close False
objWbook2.close False