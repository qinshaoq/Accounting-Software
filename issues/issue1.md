issue1: Excel关闭前执行代码
===================

##功能
已经修改的Excel关闭之前，如果用户需要保存，则将该文件中的VBA代码导出到固定路径
##问题
在Workbook_BeforeClose函数导出VBA代码，但是此时不知道用户是否想保存，如果用户不想保存，此时导出VBA代码就有问题。
##解决方法
在Workbook_BeforeClose函数中自己询问用户是否要保存，替代系统的询问对话框，然后根据用户的选择执行相应的操作。
##FIXME
自己弹出的询问是否对话框如何与系统语言保持一致？或者怎样能够调用到系统自己的询问保存对话框？
##参考
[Handling The Workbook Beforeclose Event][1]

	```
	Private Sub Workbook_BeforeClose(Cancel As Boolean)
		If Not Me.Saved Then
			Msg = "是否保存对“"
			Msg = Msg & Me.Name & "”的更改?"
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
	```

[1]: http://spreadsheetpage.com/index.php/site/tip/handling_the_workbook_beforeclose_event/