issue1: Excel�ر�ǰִ�д���
===================

##����
�Ѿ��޸ĵ�Excel�ر�֮ǰ������û���Ҫ���棬�򽫸��ļ��е�VBA���뵼�����̶�·��
##����
��Workbook_BeforeClose��������VBA���룬���Ǵ�ʱ��֪���û��Ƿ��뱣�棬����û����뱣�棬��ʱ����VBA����������⡣
##�������
��Workbook_BeforeClose�������Լ�ѯ���û��Ƿ�Ҫ���棬���ϵͳ��ѯ�ʶԻ���Ȼ������û���ѡ��ִ����Ӧ�Ĳ�����
##FIXME
�Լ�������ѯ���Ƿ�Ի��������ϵͳ���Ա���һ�£����������ܹ����õ�ϵͳ�Լ���ѯ�ʱ���Ի���
##�ο�
[Handling The Workbook Beforeclose Event][1]

	```
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
	```

[1]: http://spreadsheetpage.com/index.php/site/tip/handling_the_workbook_beforeclose_event/