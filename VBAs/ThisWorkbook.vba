VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Test()
Dim strDict As New Scripting.Dictionary
strDict.Add "XXX公司", "谷歌"
strDict.Add "20YY年", "2014年"
ReplaceArrWorkbook ThisWorkbook, strDict
End Sub
Private Sub Workbook_Open()

ImportCodeModules

End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

SaveCodeModules

End Sub
