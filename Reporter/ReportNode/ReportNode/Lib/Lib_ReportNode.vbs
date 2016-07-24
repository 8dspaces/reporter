
'@Description 指定把日志写入节点下
Public Function EnterNode(ByVal NodeName, ByVal NodeContent)
	' 用一个Dictionary对象来存储节点的信息
	Set dicMetaDescription = CreateObject("Scripting.Dictionary") 
	' 设置节点的状态
	dicMetaDescription("Status") = MicDone
	' 设置节点的名称
	dicMetaDescription("PlainTextNodeName") = NodeName
	' 设置节点的详细描述信息（可以使用HTML格式）
	dicMetaDescription("StepHtmlInfo") = NodeContent
	' 设置节点的图标
	dicMetaDescription("DllIconIndex") = 210
	dicMetaDescription("DllIconSelIndex") = 210
	' 节点图标从ContextManager.dll这个DLL文件中读取
	dicMetaDescription("DllPAth") = Environment.Value("ProductDir") & "\bin\ContextManager.dll"
	'  使用Reporter对象的LogEvent写入新节点
	intContext = Reporter.LogEvent("User", dicMetaDescription, Reporter.GetContext) 
	' 调用Reporter对象的SetContext把新写入的节点作为父节点
	Reporter.SetContext intContext	
End Function

'@Description 退出当前日志节点（与EnterNode配对使用）
Public Function ExitNode()
  '调用Reporter对象的UnSetContext，返回上一层
  Reporter.UnSetContext
End Function
