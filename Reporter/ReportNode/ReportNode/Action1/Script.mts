' 进入节点
EnterNode "父节点","<DIV align=left><H1>这是一个拥有孩子的节点</H1><b>Hello！</b> How are you！.</DIV>"
' 在节点内写Log
Reporter.ReportEvent MicPass, "Step1", "Step1 Pass！"
Reporter.ReportEvent MicWarning, "Step2", "Step2 Pass With Warning"
Reporter.ReportEvent MicFail, "Step2", "Step2 Fail！"
' 退出节点
ExitNode
' 在节点之外写Log
Reporter.ReportEvent MicPass, "Case2", "Case2 Pass!"
