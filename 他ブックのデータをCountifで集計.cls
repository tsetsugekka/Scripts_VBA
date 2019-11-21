Private Sub getValue_Click()
 
Dim MyWorkbook As Workbook
Set MyWorkbook = Application.Workbooks.Open("C:\Users\Shuai\Desktop\m1\m2.xlsx")
 
v = Application.WorksheetFunction.CountIf(MyWorkbook.Sheets("Sheet1").Range("A1:A18"), "a c")
ThisWorkbook.Sheets("Sheet1").Range("A1").Value = v
MyWorkbook.Close SaveChanges:=False
  
End Sub

 
 

'モジュール化
 
Function CrossCount(FileLink, SheetNo, RangeArea, CountText)

Dim MyWorkbook As Workbook
Set MyWorkbook = Application.Workbooks.Open(FileLink)
 
CrossCount = Application.WorksheetFunction.CountIf(MyWorkbook.Sheets(SheetNo).Range(RangeArea), CountText)
MyWorkbook.Close SaveChanges:=False

End Function

 
'モジュール使用例
Private Sub getValue()

vm = CrossCount("C:\Users\Shuai\Desktop\m1\m2.xlsx", "Sheet1", "A1:A18", "a c")
Sheets("Sheet1").Range("A23").Value = vm


End Sub


