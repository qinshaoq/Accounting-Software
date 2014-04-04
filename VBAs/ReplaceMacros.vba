Attribute VB_Name = "ReplaceMacros"
' ��һ��Workbook�е����е�Ԫ���е�strDict��key�滻ΪstrDict��value
Sub ReplaceArrWorkbook(wb As Workbook, strDict As Scripting.Dictionary)
Dim i%
For i = 1 To wb.Sheets.Count
ReplaceArrSheet wb.Sheets(i), strDict
Next

End Sub

' ��һ��Worksheet�е����е�Ԫ���е�strDict��key�滻ΪstrDict��value
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


' ��һ��Worksheet�е����е�Ԫ���е�sOld�滻ΪsNew
Private Sub ReplaceSheet(ws As Worksheet, sOld As String, sNew As String)
ws.Cells.Replace sOld, sNew
End Sub

