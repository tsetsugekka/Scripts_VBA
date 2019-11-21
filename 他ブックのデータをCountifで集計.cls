Private Sub getValue_Click()
 
Dim MyWorkbook As Workbook
Set MyWorkbook = Application.Workbooks.Open("C:\Users\Shuai\Desktop\m1\m2.xlsx")
 
v = Application.WorksheetFunction.CountIf(MyWorkbook.Sheets("Sheet1").Range("A1:A18"), "a c")
ThisWorkbook.Sheets("Sheet1").Range("A1").Value = v

End Sub
