Imports System.IO
Imports System.Net
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class TushareApiManager

    ' JSON数据文件路径
    ' Private Const ApiGuideFilePath As String = "C:\Users\syesw\Desktop\tushare excel\TushareExcel\Data\tushare_api_guide.json"
    ' Private Const ApiGuideFilePath As String = "tushare_api_guide.json"
    Private ReadOnly ApiGuideFilePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tushare_api_guide.json")


    ' API指南数据
    Private apiGuideData As Dictionary(Of String, ApiInfo)

    ' API信息类
    Public Class ApiInfo
        Public Property Title As String
        Public Property Path As String
        Public Property Description As String
    End Class

    ' 构造函数
    Public Sub New()
        ' 初始化数据
        apiGuideData = New Dictionary(Of String, ApiInfo)
    End Sub

    ' 加载API指南数据
    Public Function LoadApiGuide() As Boolean
        Try
            ' 检查文件是否存在
            If Not File.Exists(ApiGuideFilePath) Then
                MessageBox.Show("API指南文件不存在: " & ApiGuideFilePath, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            ' 读取JSON文件
            Dim jsonText As String = File.ReadAllText(ApiGuideFilePath)

            ' 解析JSON数据
            apiGuideData = JsonConvert.DeserializeObject(Of Dictionary(Of String, ApiInfo))(jsonText)

            Return True
        Catch ex As Exception
            MessageBox.Show("加载API指南数据时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' 将API指南数据填充到TreeView控件
    Public Sub PopulateTreeView(treeView As TreeView)
        Try
            ' 清空树视图
            treeView.Nodes.Clear()

            ' 检查数据是否已加载
            If apiGuideData Is Nothing OrElse apiGuideData.Count = 0 Then
                If Not LoadApiGuide() Then
                    Return
                End If
            End If

            ' 创建根节点结构
            Dim rootNodes As New Dictionary(Of String, TreeNode)
            Dim subNodes As New Dictionary(Of String, TreeNode)

            ' 添加根节点 "数据接口"
            Dim rootNode As New TreeNode("Tushare API 函数")
            treeView.Nodes.Add(rootNode)

            ' 构建路径字典
            Dim pathDictionary As New Dictionary(Of String, List(Of KeyValuePair(Of String, ApiInfo)))

            ' 将API按路径分组
            For Each kvp In apiGuideData
                Dim apiName As String = kvp.Key
                Dim apiInfo As ApiInfo = kvp.Value
                Dim path As String = apiInfo.Path

                ' 分割路径
                Dim pathParts As String() = path.Split("/"c)
                If pathParts.Length >= 2 Then
                    ' 使用第二级路径作为主分类
                    Dim mainCategory As String = pathParts(1)

                    If Not pathDictionary.ContainsKey(mainCategory) Then
                        pathDictionary.Add(mainCategory, New List(Of KeyValuePair(Of String, ApiInfo)))
                    End If

                    pathDictionary(mainCategory).Add(New KeyValuePair(Of String, ApiInfo)(apiName, apiInfo))
                End If
            Next

            ' 为每个主分类创建节点
            For Each category In pathDictionary.Keys
                Dim categoryNode As New TreeNode(category)
                rootNode.Nodes.Add(categoryNode)

                ' 创建子分类字典
                Dim subCategoryDict As New Dictionary(Of String, TreeNode)

                ' 添加API到相应的分类中
                For Each apiPair In pathDictionary(category)
                    Dim apiName As String = apiPair.Key
                    Dim apiInfo As ApiInfo = apiPair.Value
                    Dim path As String = apiInfo.Path

                    ' 分割路径以获取子分类
                    Dim pathParts As String() = path.Split("/"c)
                    Dim subCategoryName As String = "其他"

                    If pathParts.Length >= 3 Then
                        subCategoryName = pathParts(2)
                    End If

                    ' 检查子分类是否已存在
                    If Not subCategoryDict.ContainsKey(subCategoryName) Then
                        Dim subCategoryNode As New TreeNode(subCategoryName)
                        categoryNode.Nodes.Add(subCategoryNode)
                        subCategoryDict.Add(subCategoryName, subCategoryNode)
                    End If

                    ' 创建API节点
                    Dim apiNode As New TreeNode(apiName & " - " & apiInfo.Title)
                    apiNode.Tag = apiName  ' 存储API名称以便后续使用
                    subCategoryDict(subCategoryName).Nodes.Add(apiNode)
                Next
            Next

            ' 展开根节点
            rootNode.Expand()
        Catch ex As Exception
            MessageBox.Show("填充树视图时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 获取API详细信息
    Public Function GetApiDetails(apiName As String) As String
        Try
            If apiGuideData.ContainsKey(apiName) Then
                Dim apiInfo As ApiInfo = apiGuideData(apiName)

                Dim details As New System.Text.StringBuilder()
                details.AppendLine("函数名称: " & apiName)
                details.AppendLine("标题: " & apiInfo.Title)
                details.AppendLine("路径: " & apiInfo.Path)

                If Not String.IsNullOrEmpty(apiInfo.Description) Then
                    details.AppendLine("描述: " & apiInfo.Description)
                End If

                details.AppendLine()
                details.AppendLine("用法示例:")
                details.AppendLine("=TUSHARE_" & apiName.ToUpper() & "(参数1, 参数2, ...)")

                Return details.ToString()
            Else
                Return "找不到API: " & apiName
            End If
        Catch ex As Exception
            Return "获取API详情时出错: " & ex.Message
        End Try
    End Function

    ' 搜索API
    Public Sub SearchApi(treeView As TreeView, searchText As String)
        Try
            If String.IsNullOrEmpty(searchText) Then
                ' 如果搜索文本为空，重置所有节点可见性
                SetAllNodesVisible(treeView.Nodes)
                Return
            End If

            ' 转为小写以进行不区分大小写的搜索
            searchText = searchText.ToLower()

            ' 搜索所有节点
            SearchNodes(treeView.Nodes, searchText)
        Catch ex As Exception
            MessageBox.Show("搜索API时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 递归搜索节点
    Private Sub SearchNodes(nodes As TreeNodeCollection, searchText As String)
        For Each node As TreeNode In nodes
            ' 递归搜索子节点
            SearchNodes(node.Nodes, searchText)

            ' 设置当前节点可见性
            Dim visible As Boolean = node.Text.ToLower().Contains(searchText)

            ' 如果有任何子节点可见，则当前节点也应该可见
            For Each childNode As TreeNode In node.Nodes
                If childNode.ForeColor <> System.Drawing.SystemColors.Window Then
                    visible = True
                    Exit For
                End If
            Next

            node.ForeColor = If(visible, System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window)
        Next
    End Sub

    ' 设置所有节点可见
    Private Sub SetAllNodesVisible(nodes As TreeNodeCollection)
        For Each node As TreeNode In nodes
            node.ForeColor = System.Drawing.SystemColors.WindowText
            SetAllNodesVisible(node.Nodes)
        Next
    End Sub

    ' 更新API指南数据
    Public Function UpdateApiGuide(token As String) As Boolean
        Try
            ' 使用提供的token从Tushare API获取最新的doctrees数据
            ' 处理数据并保存到本地文件

            ' 这里只是一个占位符，实际实现需要根据需求进行开发
            MessageBox.Show("此功能尚未实现。请使用Python脚本更新API指南数据。", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return False
        Catch ex As Exception
            MessageBox.Show("更新API指南数据时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function
End Class