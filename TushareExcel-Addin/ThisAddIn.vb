Public Class ThisAddIn
    ' 声明函数向导窗格
    Public FunctionGuidePane As Microsoft.Office.Tools.CustomTaskPane

    ' 声明AI聊天窗格
    Public AIChatPane As Microsoft.Office.Tools.CustomTaskPane

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        ' 创建函数向导控件实例
        Dim functionGuidePaneControl As New FunctionGuidePane()

        ' 创建自定义任务窗格
        FunctionGuidePane = Me.CustomTaskPanes.Add(functionGuidePaneControl, "Tushare 函数向导")

        ' 设置初始属性
        FunctionGuidePane.Width = 800
        FunctionGuidePane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        FunctionGuidePane.Visible = False



        ' 创建AI聊天控件实例
        Dim aiChatControl As New AIChatUserControl()
        ' 创建AI聊天任务窗格
        AIChatPane = Me.CustomTaskPanes.Add(aiChatControl, "AI 助手")
        ' 设置初始属性
        AIChatPane.Width = 800
        AIChatPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        AIChatPane.Visible = False

    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
        ' 清理资源
    End Sub
End Class