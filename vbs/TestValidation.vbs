Option Explicit
Dim  objExcel, objWbook1, objSheet1, objWbook2, objSheet2, rng

Sub copyValidation(objSheet1, objSheet2, rng)
Dim val, fml
Set val = objSheet1.Range(rng).Validation
fml = val.Formula1
Wscript.echo fml
Set val = objSheet2.Range(rng).Validation
val.Delete
' expression.Add(Type, AlertStyle, Operator, Formula1, Formula2)
' where type 3 means drop-down list, see https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xldvtype-enumeration-excel
val.Add 3,,,fml
End Sub


Set objExcel = Wscript.CreateObject("Excel.Application")
Set objWbook1 = objExcel.Workbooks.Open("U:\WORK\develop\trunk\vbs\Excel1.xlsx", False, True)
Set objSheet1 = objWbook1.Worksheets("Data")
Set objWbook2 = objExcel.Workbooks.Open("U:\WORK\develop\trunk\vbs\Excel2.xlsx", False, False)
Set objSheet2 = objWbook2.Worksheets("Data")
copyValidation objSheet1, objSheet2, "D4:D6"
copyValidation objSheet1, objSheet2, "E4:E6"
objWbook2.Save
objWbook1.close False
objWbook2.close False

