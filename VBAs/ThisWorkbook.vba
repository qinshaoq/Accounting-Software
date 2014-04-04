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
strDict.Add "XXX��˾", "�ȸ�"
strDict.Add "20YY��", "2014��"
ReplaceArrWorkbook ThisWorkbook, strDict
End Sub

Private Sub Workbook_Open()

ImportCodeModules

End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Not Me.Saved Then
        Msg = "�Ƿ񱣴�ԡ�"
        Msg = Msg & Me.Name & "���ĸ���?"
        Ans = MsgBox(Msg, vbQuestion + vbYesNoCancel)
        Select Case Ans
            Case vbYes
                Me.Save
                SaveCodeModules
            Case vbNo
                Me.Saved = True
            Case vbCancel
                Cancel = True
                Exit Sub
          End Select
    End If
End Sub
