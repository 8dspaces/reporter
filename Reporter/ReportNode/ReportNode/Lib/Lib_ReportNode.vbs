
'@Description ָ������־д��ڵ���
Public Function EnterNode(ByVal NodeName, ByVal NodeContent)
	' ��һ��Dictionary�������洢�ڵ����Ϣ
	Set dicMetaDescription = CreateObject("Scripting.Dictionary") 
	' ���ýڵ��״̬
	dicMetaDescription("Status") = MicDone
	' ���ýڵ������
	dicMetaDescription("PlainTextNodeName") = NodeName
	' ���ýڵ����ϸ������Ϣ������ʹ��HTML��ʽ��
	dicMetaDescription("StepHtmlInfo") = NodeContent
	' ���ýڵ��ͼ��
	dicMetaDescription("DllIconIndex") = 210
	dicMetaDescription("DllIconSelIndex") = 210
	' �ڵ�ͼ���ContextManager.dll���DLL�ļ��ж�ȡ
	dicMetaDescription("DllPAth") = Environment.Value("ProductDir") & "\bin\ContextManager.dll"
	'  ʹ��Reporter�����LogEventд���½ڵ�
	intContext = Reporter.LogEvent("User", dicMetaDescription, Reporter.GetContext) 
	' ����Reporter�����SetContext����д��Ľڵ���Ϊ���ڵ�
	Reporter.SetContext intContext	
End Function

'@Description �˳���ǰ��־�ڵ㣨��EnterNode���ʹ�ã�
Public Function ExitNode()
  '����Reporter�����UnSetContext��������һ��
  Reporter.UnSetContext
End Function
