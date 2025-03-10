Imports System.Drawing
Imports System.Windows.Forms

Public Class FunctionGuidePane
    ' Tushare API管理器
    Private apiManager As New TushareApiManager()

    Private Sub FunctionGuidePane_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 加载API指南数据并初始化树视图
        InitializeTreeView()
    End Sub

    Private Sub InitializeTreeView()
        Try
            ' 使用API管理器填充树视图
            apiManager.PopulateTreeView(treeViewFunctions)
        Catch ex As Exception
            MessageBox.Show("初始化函数树状图时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub treeViewFunctions_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles treeViewFunctions.AfterSelect
        Try
            ' 获取选择的节点
            Dim selectedNode As TreeNode = e.Node

            ' 检查选择的是否是叶子节点（具体函数）
            If selectedNode.Nodes.Count = 0 AndAlso selectedNode.Parent IsNot Nothing Then
                ' 获取API名称
                Dim apiName As String = selectedNode.Tag?.ToString()

                If Not String.IsNullOrEmpty(apiName) Then
                    ' 显示API详情
                    txtFunctionDetails.Text = apiManager.GetApiDetails(apiName)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("选择函数时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 添加搜索功能
    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        ' 使用API管理器实现搜索逻辑
        apiManager.SearchApi(treeViewFunctions, txtSearch.Text.Trim())
    End Sub

    ' 双击节点可复制函数名称到剪贴板
    Private Sub treeViewFunctions_DoubleClick(sender As Object, e As EventArgs) Handles treeViewFunctions.DoubleClick
        Dim selectedNode As TreeNode = treeViewFunctions.SelectedNode

        If selectedNode IsNot Nothing AndAlso selectedNode.Nodes.Count = 0 AndAlso selectedNode.Parent IsNot Nothing Then
            Try
                ' 获取API名称
                Dim apiName As String = selectedNode.Tag?.ToString()

                If Not String.IsNullOrEmpty(apiName) Then
                    ' 复制函数名到剪贴板
                    Clipboard.SetText("TUSHARE_" & apiName.ToUpper())

                    ' 显示提示消息
                    MessageBox.Show("函数名 'TUSHARE_" & apiName.ToUpper() & "' 已复制到剪贴板",
                                   "复制成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Catch ex As Exception
                MessageBox.Show("复制函数名到剪贴板时出错: " & ex.Message,
                               "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
End Class