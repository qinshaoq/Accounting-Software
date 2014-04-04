Attribute VB_Name = "ReplaceMacros"
' 将一个Workbook中的所有单元格中的strDict的key替换为strDict的value
Sub ReplaceArrWorkbook(wb As Workbook, strDict As Scripting.Dictionary)
Dim i%
For i = 1 To wb.Sheets.Count
ReplaceArrSheet wb.Sheets(i), strDict
Next

End Sub

' 将一个Worksheet中的所有单元格中的strDict的key替换为strDict的value
Private Sub ReplaceArrSheet(ws As Worksheet, strDict As Scripting.Dictionary)
k = strDict.keys
v = strDict.Items
Dim i%
Dim key$, value$
For i = 0 To strDict.Count - 1
key = k(i)
value = v(i)
ReplaceSheet ws, key, value
Next i

End Sub


' 将一个Worksheet中的所有单元格中的sOld替换为sNew
Private Sub ReplaceSheet(ws As Worksheet, sOld As String, sNew As String)
ws.Cells.Replace sOld, sNew
End Sub

