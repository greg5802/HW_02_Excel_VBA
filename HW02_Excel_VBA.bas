Attribute VB_Name = "Module1"
Sub Total_Volume()
Dim x As Long

'the number after "to" can be manipulated to account for extra rows or exclude rows
For x = 2 To 3169
Cells(x, 11).Value = Application.WorksheetFunction.SumIf(Range("A1:A797711"), Cells(x, 10).Value, Range("G1:G797711"))
Next
End Sub

