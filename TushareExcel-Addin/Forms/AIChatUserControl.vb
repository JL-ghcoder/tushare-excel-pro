Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Excel = Microsoft.Office.Interop.Excel

Public Class AIChatUserControl
    ' API 配置 - 硬编码 (仅用于开发测试)
    ' 标记是否正在初始化
    Private isInitializing As Boolean = True

    ' Private Const GPT_API_URL As String = "https://api.openai.com/v1/chat/completions"
    ' Private Const GPT_API_KEY As String = ""

    ' Private Const DS_API_URL As String = "https://api.deepseek.com/chat/completions"
    ' Private Const DS_API_KEY As String = "" ' 替换为您的DeepSeek API密钥


    Dim GPT_API_URL = GetGPTApiUrl()
    Dim GPT_API_KEY = GetGPTToken()
    Dim DS_API_URL = GetDSApiUrl()
    Dim DS_API_KEY = GetDSToken()


    Private Const REG_PATH As String = "Software\TushareExcel"

    ''' <summary>
    ''' 获取GPT API URL，优先从设置中读取
    ''' </summary>
    Private Function GetGPTApiUrl() As String
        ' 尝试从注册表读取
        Try
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(REG_PATH)
                If key IsNot Nothing Then
                    Dim url = key.GetValue("GPTApiUrl")
                    If url IsNot Nothing Then
                        Return url.ToString()
                    End If
                End If
            End Using
        Catch ex As Exception
            ' 忽略错误，使用默认值
        End Try

        ' 使用默认值
        Return "https://api.openai.com/v1/chat/completions"
    End Function

    ''' <summary>
    ''' 获取GPT API Token，优先从设置中读取
    ''' </summary>
    Private Function GetGPTToken() As String
        ' 尝试从注册表读取
        Try
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(REG_PATH)
                If key IsNot Nothing Then
                    Dim token = key.GetValue("GPTToken")
                    If token IsNot Nothing Then
                        Return token.ToString()
                    End If
                End If
            End Using
        Catch ex As Exception
            ' 忽略错误，返回空字符串
        End Try

        ' 如果未设置Token，返回空字符串
        Return ""
    End Function

    Private Function GetDSApiUrl() As String
        ' 尝试从注册表读取
        Try
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(REG_PATH)
                If key IsNot Nothing Then
                    Dim url = key.GetValue("DSApiUrl")
                    If url IsNot Nothing Then
                        Return url.ToString()
                    End If
                End If
            End Using
        Catch ex As Exception
            ' 忽略错误，使用默认值
        End Try

        ' 使用默认值
        Return "https://api.deepseek.com/chat/completions"
    End Function

    ''' <summary>
    ''' 获取DS API Token，优先从设置中读取
    ''' </summary>
    Private Function GetDSToken() As String
        ' 尝试从注册表读取
        Try
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(REG_PATH)
                If key IsNot Nothing Then
                    Dim token = key.GetValue("DSToken")
                    If token IsNot Nothing Then
                        Return token.ToString()
                    End If
                End If
            End Using
        Catch ex As Exception
            ' 忽略错误，返回空字符串
        End Try

        ' 如果未设置Token，返回空字符串
        Return ""
    End Function


    ' 当前使用的API配置
    Private currentApiUrl As String = GPT_API_URL
    Private currentApiKey As String = GPT_API_KEY
    Private currentModelName As String = "gpt-4o"
    Private isUsingGPT As Boolean = True

    ' Tushare文档
    Private ReadOnly ApiGuideFilePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tushare_api_guide.json")
    Private includeTushareDocs As Boolean = False
    Private tushareApiGuide As JObject = Nothing

    ' 创建一个方法来切换是否引入 Tushare API 文档
    Public Sub ToggleIncludeTushareDocs(include As Boolean)
        includeTushareDocs = include

        ' 如果启用了 Tushare 文档功能，则加载 API 指南
        If include Then
            LoadTushareApiGuide()
        Else
            tushareApiGuide = Nothing
        End If
    End Sub

    ' 加载 Tushare API 指南文件
    Private Sub LoadTushareApiGuide()
        Try
            If File.Exists(ApiGuideFilePath) Then
                Dim jsonContent As String = File.ReadAllText(ApiGuideFilePath, Encoding.UTF8)
                tushareApiGuide = JObject.Parse(jsonContent)
                AddSystemMessage("已成功加载 Tushare API 指南文件。")
            Else
                AddSystemMessage("错误: Tushare API 指南文件不存在: " & ApiGuideFilePath)
                tushareApiGuide = Nothing
            End If
        Catch ex As Exception
            AddSystemMessage("加载 Tushare API 指南文件时出错: " & ex.Message)
            tushareApiGuide = Nothing
        End Try
    End Sub


    ' 选择是否引入Excel框选内容的checkbox状态
    Private includeSelectionInPrompt As Boolean = False

    ' 创建一个方法来切换是否引入Excel框选内容的状态
    Public Sub ToggleIncludeSelection(include As Boolean)
        includeSelectionInPrompt = include
    End Sub

    ' 保存聊天记录的富文本格式
    Private chatHistory As New List(Of ChatMessage)

    ' 保存代码块信息，用于运行VBA
    Private codeBlocks As New Dictionary(Of Integer, CodeBlockInfo)


    ' 获取当前Excel选区内容的方法
    Private Function GetExcelSelection() As String
        Try
            ' 获取Excel应用程序实例
            Dim excelApp As Excel.Application = DirectCast(Globals.ThisAddIn.Application, Excel.Application)

            ' 获取当前活动工作表
            Dim activeSheet As Excel.Worksheet = excelApp.ActiveSheet

            ' 获取当前选区
            Dim selection As Excel.Range = excelApp.Selection

            ' 检查选区是否为空
            If selection Is Nothing OrElse selection.Count = 0 Then
                Return "未选择任何单元格。"
            End If

            ' 准备用于存储选区数据的StringBuilder
            Dim sb As New StringBuilder()

            ' 添加标题行和选区信息
            sb.AppendLine("以下是用户选择的Excel数据:")
            sb.AppendLine($"选区地址: {selection.Address}")
            sb.AppendLine()

            ' 如果选区只有一个单元格
            If selection.Count = 1 Then
                sb.AppendLine("单元格 " & selection.Address & ": " & If(selection.Value Is Nothing, "", selection.Value.ToString()))
                Return sb.ToString()
            End If

            ' 获取选区的行数和列数
            Dim rowCount As Integer = selection.Rows.Count
            Dim colCount As Integer = selection.Columns.Count

            ' 创建表格表示 - 始终用表格表示多单元格数据
            sb.Append("| ")
            For c As Integer = 1 To colCount
                ' 列标题 - 使用Excel的列标题(A, B, C...)
                Dim colLetter As String = GetExcelColumnName(c)
                sb.Append(colLetter).Append(" | ")
            Next
            sb.AppendLine()

            ' 添加表格分隔行
            sb.Append("| ")
            For c As Integer = 1 To colCount
                sb.Append("--- | ")
            Next
            sb.AppendLine()

            ' 添加所有单元格的数据
            For r As Integer = 1 To rowCount
                sb.Append("| ")
                For c As Integer = 1 To colCount
                    ' 获取单元格值并处理可能的空值
                    Dim cellValue As String = ""
                    Try
                        Dim value = selection.Cells(r, c).Value
                        cellValue = If(value IsNot Nothing, value.ToString(), "")
                    Catch ex As Exception
                        cellValue = ""
                    End Try
                    sb.Append(cellValue).Append(" | ")
                Next
                sb.AppendLine()
            Next

            Return sb.ToString()
        Catch ex As Exception
            Return "获取Excel选区时出错: " & ex.Message
        End Try
    End Function

    ' 添加一个辅助方法来获取Excel列标题(A, B, C...)
    Private Function GetExcelColumnName(columnNumber As Integer) As String
        Dim dividend As Integer = columnNumber
        Dim columnName As String = String.Empty
        Dim modulo As Integer

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo) & columnName
            dividend = CInt((dividend - modulo) / 26)
        End While

        Return columnName
    End Function

    ' 代码块信息类
    Private Class CodeBlockInfo
        Public Property Language As String
        Public Property Code As String
        Public Property StartPosition As Integer
        Public Property EndPosition As Integer

        Public Sub New(language As String, code As String, startPos As Integer, endPos As Integer)
            Me.Language = language.ToLower()
            Me.Code = code
            Me.StartPosition = startPos
            Me.EndPosition = endPos
        End Sub

        Public Function IsVBA() As Boolean
            Return Language = "vba" OrElse Language = "vb" OrElse Language = "visual basic"
        End Function

        Public Function IsExcelFormula() As Boolean
            Return Language = "excel" AndAlso Code.Trim().StartsWith("=")
        End Function

    End Class

    ' 聊天消息类
    Private Class ChatMessage
        Public Property IsUser As Boolean
        Public Property Message As String
        Public Property Timestamp As DateTime

        Public Sub New(isUser As Boolean, message As String)
            Me.IsUser = isUser
            Me.Message = message
            Me.Timestamp = DateTime.Now
        End Sub
    End Class

    Private Sub AIChatUserControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 设置RichTextBox和其他控件的样式
        SetupChatStyles()

        ' 标记为初始化中
        isInitializing = True

        ' 初始化API设置 (不显示切换消息)
        chkUseGPT.Checked = True
        chkUseDS.Checked = False
        chkUseTushareDocs.Checked = False  ' 初始化 Tushare API 复选框状态
        UpdateApiSettingsInternal(False)

        ' 初始化完成
        isInitializing = False

        AddSystemMessage("欢迎使用Excel AI助手！请输入您的问题...")

    End Sub

    Private Sub SetupChatStyles()
        ' 设置文本框样式
        txtMessages.BackColor = Color.WhiteSmoke
        txtMessages.Font = New Font("Segoe UI", 9.5F)
        txtMessages.ReadOnly = True
        txtMessages.BorderStyle = BorderStyle.None

        ' 用户输入框样式
        txtUserInput.Font = New Font("Segoe UI", 9.5F)
        txtUserInput.BorderStyle = BorderStyle.FixedSingle
        txtUserInput.AcceptsReturn = True ' 允许多行输入

        ' 发送按钮样式
        btnSend.FlatStyle = FlatStyle.Flat
        btnSend.BackColor = Color.FromArgb(0, 120, 212)
        btnSend.ForeColor = Color.White
        btnSend.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
        btnSend.Cursor = Cursors.Hand
    End Sub

    ' 添加系统消息（灰色）
    Private Sub AddSystemMessage(message As String)
        txtMessages.SelectionStart = txtMessages.TextLength
        txtMessages.SelectionLength = 0
        txtMessages.SelectionColor = Color.Gray
        txtMessages.SelectionFont = New Font(txtMessages.Font, FontStyle.Italic)
        txtMessages.AppendText(message & vbCrLf & vbCrLf)

        ' 确保滚动到最新消息
        txtMessages.SelectionStart = txtMessages.TextLength
        txtMessages.ScrollToCaret()
    End Sub

    ' 添加用户消息（蓝色）
    Private Sub AddUserMessage(message As String)
        ' 保存消息到历史记录
        chatHistory.Add(New ChatMessage(True, message))

        txtMessages.SelectionStart = txtMessages.TextLength
        txtMessages.SelectionLength = 0

        ' 添加消息标签
        txtMessages.SelectionColor = Color.FromArgb(0, 120, 212)
        txtMessages.SelectionFont = New Font(txtMessages.Font, FontStyle.Bold)
        txtMessages.AppendText("您: ")

        ' 添加消息内容
        txtMessages.SelectionFont = New Font(txtMessages.Font, FontStyle.Regular)
        txtMessages.AppendText(message & vbCrLf & vbCrLf)
    End Sub

    ' 添加AI消息（绿色），支持代码格式化
    Private Sub AddAIMessage(message As String)
        ' 保存消息到历史记录
        chatHistory.Add(New ChatMessage(False, message))

        txtMessages.SelectionStart = txtMessages.TextLength
        txtMessages.SelectionLength = 0

        ' 添加消息标签
        txtMessages.SelectionColor = Color.FromArgb(0, 128, 0)
        txtMessages.SelectionFont = New Font(txtMessages.Font, FontStyle.Bold)
        txtMessages.AppendText("AI: ")

        ' 重置选择颜色和字体
        txtMessages.SelectionColor = Color.Black
        txtMessages.SelectionFont = New Font(txtMessages.Font, FontStyle.Regular)

        ' 检查消息是否包含代码块，并格式化它们
        If message.Contains("```") Then
            FormatMessageWithCode(message)
        Else
            ' 普通文本
            txtMessages.AppendText(message & vbCrLf & vbCrLf)
        End If
    End Sub

    ' 格式化包含代码块的消息
    ' 在FormatMessageWithCode方法中修改代码块处理逻辑
    Private Sub FormatMessageWithCode(message As String)
        ' 先清除旧的代码块信息
        codeBlocks.Clear()

        ' 分割消息，提取代码块
        Dim parts As New List(Of String)
        Dim isCodeBlock As Boolean = False
        Dim currentBlock As String = ""
        Dim codeLang As String = ""

        ' 分割字符串处理代码块
        Dim lines() As String = message.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)

        For Each line As String In lines
            If line.StartsWith("```") Then
                ' 代码块开始/结束标记
                If isCodeBlock Then
                    ' 结束当前代码块
                    parts.Add("CODE:" & codeLang & ":" & currentBlock)
                    currentBlock = ""
                    isCodeBlock = False
                    codeLang = ""
                Else
                    ' 开始新代码块
                    If currentBlock.Length > 0 Then
                        parts.Add("TEXT:" & currentBlock)
                        currentBlock = ""
                    End If
                    isCodeBlock = True

                    ' 获取代码语言标识（如果有）
                    codeLang = line.Replace("```", "").Trim()
                End If
            Else
                ' 常规文本或代码行
                If currentBlock.Length > 0 Then
                    currentBlock += vbCrLf
                End If
                currentBlock += line
            End If
        Next

        ' 添加最后一个块
        If currentBlock.Length > 0 Then
            If isCodeBlock Then
                parts.Add("CODE:" & codeLang & ":" & currentBlock)
            Else
                parts.Add("TEXT:" & currentBlock)
            End If
        End If

        ' 输出格式化后的块
        Dim codeBlockCount As Integer = 0

        For Each part As String In parts
            If part.StartsWith("CODE:") Then
                ' 代码块
                Dim codeInfo() As String = part.Split(New String() {":"}, 3, StringSplitOptions.None)
                Dim lang As String = codeInfo(1)
                Dim code As String = codeInfo(2)

                ' 记录此代码块的开始位置
                Dim blockStartPos As Integer = txtMessages.TextLength

                ' 添加代码语言标签和运行按钮（如果是VBA代码或Excel公式）
                Dim isVbaCode As Boolean = (lang.ToLower() = "vba" OrElse lang.ToLower() = "vb" OrElse lang.ToLower() = "visual basic")
                Dim isExcelFormula As Boolean = (lang.ToLower() = "excel" AndAlso code.Trim().StartsWith("="))

                txtMessages.SelectionStart = txtMessages.TextLength
                txtMessages.SelectionColor = Color.Gray
                txtMessages.SelectionFont = New Font("Segoe UI", 8.0F, FontStyle.Italic)

                If isVbaCode OrElse isExcelFormula Then
                    ' 创建一个唯一的代码块ID
                    codeBlockCount += 1

                    ' 添加语言标签和运行按钮
                    txtMessages.AppendText("[" & lang & "] ")

                    ' 记录"运行"按钮的起始位置
                    Dim runBtnStartPos As Integer = txtMessages.TextLength

                    ' 添加运行按钮文本（后面会通过点击事件处理）
                    txtMessages.SelectionColor = Color.Blue
                    txtMessages.SelectionFont = New Font("Segoe UI", 8.0F, FontStyle.Underline)

                    ' 对Excel公式使用不同的按钮文本
                    If isExcelFormula Then
                        txtMessages.AppendText("[ 应用此公式 ]")
                    Else
                        txtMessages.AppendText("[ 运行此代码 ]")
                    End If

                    ' 记录"运行"按钮的结束位置
                    Dim runBtnEndPos As Integer = txtMessages.TextLength

                    ' 将"运行"按钮区域添加到链接列表
                    Dim linkRange As New LinkLabel.Link(runBtnStartPos, runBtnEndPos - runBtnStartPos, codeBlockCount)
                    txtMessages.AppendText(vbCrLf)

                    ' 保存代码块信息
                    codeBlocks.Add(codeBlockCount, New CodeBlockInfo(lang, code, blockStartPos, txtMessages.TextLength))
                Else
                    txtMessages.AppendText("[" & lang & "]" & vbCrLf)
                End If

                ' 添加代码块背景
                txtMessages.SelectionStart = txtMessages.TextLength
                txtMessages.SelectionBackColor = Color.FromArgb(240, 240, 240)
                txtMessages.SelectionColor = Color.FromArgb(9, 9, 9)
                txtMessages.SelectionFont = New Font("Consolas", 9.0F)

                ' 记录代码开始位置，用于高亮
                Dim codeStartPos As Integer = txtMessages.TextLength

                ' 添加代码文本
                txtMessages.AppendText(code)

                ' 尝试应用代码高亮
                Try
                    ' 根据代码类型应用高亮
                    If isVbaCode OrElse isExcelFormula Then
                        ApplyCodeHighlighting(txtMessages, codeStartPos, code, lang)
                    End If
                Catch ex As Exception
                    ' 如果高亮过程出错，只记录错误但不中断程序
                    Debug.WriteLine("代码高亮错误: " & ex.Message)
                End Try

                ' 添加换行符
                txtMessages.AppendText(vbCrLf)

                ' 重置背景色
                txtMessages.SelectionStart = txtMessages.TextLength
                txtMessages.SelectionBackColor = txtMessages.BackColor
                txtMessages.AppendText(vbCrLf)
            Else
                ' 普通文本
                Dim text As String = part.Substring(5)
                txtMessages.SelectionColor = Color.Black
                txtMessages.SelectionFont = New Font(txtMessages.Font, FontStyle.Regular)
                txtMessages.AppendText(text & vbCrLf)
            End If
        Next

        txtMessages.AppendText(vbCrLf)
    End Sub

    Private Async Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        If String.IsNullOrWhiteSpace(txtUserInput.Text) Then
            Return
        End If

        ' 获取用户输入并清空输入框
        Dim userMessage As String = txtUserInput.Text
        txtUserInput.Text = ""

        ' 显示用户消息
        AddUserMessage(userMessage)

        ' 滚动到底部
        txtMessages.SelectionStart = txtMessages.Text.Length
        txtMessages.ScrollToCaret()

        ' 禁用发送按钮
        btnSend.Enabled = False
        btnSend.Text = "请稍候..."
        btnSend.BackColor = Color.Gray

        Try
            ' 调用API
            Dim response As String = Await SendToChatAIAsync(userMessage)

            ' 显示AI回复
            AddAIMessage(response)

            ' 将光标移动到文本末尾以显示最新消息
            txtMessages.SelectionStart = txtMessages.Text.Length
            txtMessages.ScrollToCaret()
        Catch ex As Exception
            AddSystemMessage("错误: " & ex.Message)
        Finally
            ' 重新启用发送按钮
            btnSend.Enabled = True
            btnSend.Text = "发送"
            btnSend.BackColor = Color.FromArgb(0, 120, 212)
        End Try
    End Sub

    Private Async Function SendToChatAIAsync(ByVal userMessage As String) As Task(Of String)
        Try
            ' 设置TLS版本
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12

            Using client As New HttpClient()
                ' 增加超时时间
                client.Timeout = TimeSpan.FromSeconds(60)

                ' 设置请求头
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & currentApiKey)

                ' 构建请求内容
                Dim requestData As New JObject()
                requestData.Add("model", currentModelName)

                Dim messages As New JArray()

                ' 系统消息
                Dim systemMessage As New JObject()

                ' 更新系统消息，添加 Tushare API 信息
                Dim systemPrompt As String = "你是一个Excel数据分析助手，善于分析数据和提供Excel公式和VBA/Python代码来解决问题。请使用Markdown格式输出代码，使用```vba, ```python等标记代码块。"

                ' 如果启用了 Tushare API 指南，添加相关信息
                If includeTushareDocs AndAlso tushareApiGuide IsNot Nothing Then
                    systemPrompt &= vbCrLf & vbCrLf & "你可以推荐使用以下 Tushare API 相关的 Excel 公式。每个公式的格式为 TUSHARE_ 加上 API 名称的大写形式，同时你还要注意作为公式传参必须把参数名和值都传入。例如，对于 stock_company API，Excel 公式为 ==TUSHARE_STOCK_COMPANY(""ts_code"", ""000001.SZ"")。以下是可用的 API 列表，如果要给我公式的话你要给我excel格式的公式代码:" & vbCrLf


                    ' 添加 Tushare API 信息
                    For Each api In tushareApiGuide.Properties()
                        Dim apiName As String = api.Name
                        Dim apiInfo As JObject = DirectCast(api.Value, JObject)

                        Dim title As String = If(apiInfo("title") IsNot Nothing, apiInfo("title").ToString(), "")
                        Dim description As String = If(apiInfo("description") IsNot Nothing, apiInfo("description").ToString(), "")
                        Dim excelFormula As String = "=TUSHARE_" & apiName.ToUpper()

                        systemPrompt &= vbCrLf & "- " & apiName & ": " & title & " - " & description & vbCrLf & "  Excel公式: " & excelFormula
                    Next
                End If

                systemMessage.Add("role", "system")
                systemMessage.Add("content", systemPrompt)
                messages.Add(systemMessage)

                ' 添加部分历史消息作为上下文（最多10条）
                ' 兼容旧版本.NET Framework，不使用TakeLast
                Dim historyToInclude As New List(Of ChatMessage)
                Dim startIndex As Integer = Math.Max(0, chatHistory.Count - 10)
                For i As Integer = startIndex To chatHistory.Count - 1
                    historyToInclude.Add(chatHistory(i))
                Next

                For Each msg In historyToInclude
                    Dim historyMsg As New JObject()
                    historyMsg.Add("role", If(msg.IsUser, "user", "assistant"))
                    historyMsg.Add("content", msg.Message)
                    messages.Add(historyMsg)
                Next

                ' 用户消息 - 如果已启用，则加入Excel选区内容
                Dim finalUserMessage As String = userMessage

                If includeSelectionInPrompt Then
                    Dim selectionContent As String = GetExcelSelection()
                    If Not String.IsNullOrEmpty(selectionContent) Then
                        finalUserMessage = selectionContent & vbCrLf & vbCrLf & "基于以上Excel数据，请回答以下问题:" & vbCrLf & userMessage
                    End If
                End If

                Dim userMsg As New JObject()
                userMsg.Add("role", "user")
                userMsg.Add("content", finalUserMessage)
                messages.Add(userMsg)

                requestData.Add("messages", messages)

                ' 根据不同模型设置其他参数
                If isUsingGPT Then
                    requestData.Add("temperature", 0.7)
                    requestData.Add("max_tokens", 2000)
                Else
                    requestData.Add("temperature", 0.7)
                    requestData.Add("stream", False)
                End If

                ' 将请求内容输出到控制台以便调试
                Debug.WriteLine("Request JSON: " & requestData.ToString())
                Debug.WriteLine("Using API: " & currentApiUrl)

                ' 发送请求
                Dim content As New StringContent(requestData.ToString(), Encoding.UTF8, "application/json")

                Try
                    Dim response As HttpResponseMessage = Await client.PostAsync(currentApiUrl, content)

                    ' 检查响应
                    If response.IsSuccessStatusCode Then
                        Dim responseString As String = Await response.Content.ReadAsStringAsync()
                        Dim responseObject As JObject = JObject.Parse(responseString)

                        ' 提取AI回复
                        If isUsingGPT Then
                            Return responseObject("choices")(0)("message")("content").ToString()
                        Else
                            Return responseObject("choices")(0)("message")("content").ToString()
                        End If
                    Else
                        Dim errorResponse As String = Await response.Content.ReadAsStringAsync()
                        Throw New Exception("API响应错误: " & response.StatusCode & " - " & errorResponse)
                    End If
                Catch ex As HttpRequestException
                    Throw New Exception("网络请求错误: " & ex.Message)
                Catch ex As JsonException
                    Throw New Exception("JSON解析错误: " & ex.Message)
                End Try
            End Using
        Catch ex As Exception
            Return "发送请求时出错: " & ex.Message
        End Try
    End Function


    ' 键盘快捷键处理 - 按Enter发送消息
    Private Sub txtUserInput_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUserInput.KeyDown
        ' 如果按下的是Enter键且没有同时按Shift键 (Shift+Enter用于换行)
        If e.KeyCode = Keys.Enter AndAlso Not e.Shift Then
            e.SuppressKeyPress = True  ' 阻止默认的Enter键行为
            If btnSend.Enabled Then
                btnSend_Click(sender, EventArgs.Empty)
            End If
        End If
    End Sub

    ' 右键菜单 - 复制和清除聊天
    Private Sub txtMessages_MouseUp(sender As Object, e As MouseEventArgs) Handles txtMessages.MouseUp
        If e.Button = MouseButtons.Right Then
            Dim contextMenu As New ContextMenuStrip()

            ' 复制选中文本
            Dim copyItem As New ToolStripMenuItem("复制")
            AddHandler copyItem.Click, Sub(s, args)
                                           If txtMessages.SelectionLength > 0 Then
                                               Clipboard.SetText(txtMessages.SelectedText)
                                           End If
                                       End Sub
            contextMenu.Items.Add(copyItem)

            ' 复制全部文本
            Dim copyAllItem As New ToolStripMenuItem("复制全部")
            AddHandler copyAllItem.Click, Sub(s, args)
                                              Clipboard.SetText(txtMessages.Text)
                                          End Sub
            contextMenu.Items.Add(copyAllItem)

            ' 分隔线
            contextMenu.Items.Add(New ToolStripSeparator())

            ' 清除聊天
            Dim clearItem As New ToolStripMenuItem("清除聊天")
            AddHandler clearItem.Click, Sub(s, args)
                                            If MessageBox.Show("确定要清除所有聊天记录吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                                                txtMessages.Clear()
                                                chatHistory.Clear()
                                                codeBlocks.Clear()
                                                AddSystemMessage("聊天已清除。")
                                            End If
                                        End Sub
            contextMenu.Items.Add(clearItem)

            ' 显示菜单
            contextMenu.Show(txtMessages, e.Location)
        End If
    End Sub

    ' 处理聊天框中的点击事件，用于运行VBA代码
    ' 修改点击事件处理，根据代码类型执行不同操作
    Private Sub txtMessages_MouseClick(sender As Object, e As MouseEventArgs) Handles txtMessages.MouseClick
        If e.Button = MouseButtons.Left Then
            ' 获取点击位置的字符索引
            Dim index As Integer = txtMessages.GetCharIndexFromPosition(New Point(e.X, e.Y))

            ' 检查是否点击了任何代码块的"运行"按钮
            For Each kvp In codeBlocks
                Dim id As Integer = kvp.Key
                Dim codeInfo As CodeBlockInfo = kvp.Value

                ' 检查点击位置是否在"运行"按钮文本区域内
                ' 此处简化处理，考虑到"[ 运行此代码 ]"或"[ 应用此公式 ]"文本的位置
                Dim runBtnStartPos As Integer = codeInfo.StartPosition

                ' 确定按钮文本
                Dim runBtnText As String
                If codeInfo.IsExcelFormula() Then
                    runBtnText = "[ 应用此公式 ]"
                Else
                    runBtnText = "[ 运行此代码 ]"
                End If

                Dim runBtnLength As Integer = runBtnText.Length

                ' 查找"运行/应用"按钮文本位置
                Dim btnPos As Integer = txtMessages.Text.IndexOf(runBtnText, runBtnStartPos)

                If btnPos >= 0 AndAlso index >= btnPos AndAlso index < btnPos + runBtnLength Then
                    ' 点击了按钮
                    If codeInfo.IsVBA() Then
                        RunVBACode(codeInfo.Code)
                    ElseIf codeInfo.IsExcelFormula() Then
                        ApplyExcelFormula(codeInfo.Code)
                    End If
                    Exit For
                End If
            Next
        End If
    End Sub

    ' 新增方法：应用Excel公式到当前选中单元格
    Private Sub ApplyExcelFormula(formula As String)
        Try
            ' 获取Excel应用程序实例
            Dim excelApp As Excel.Application = DirectCast(Globals.ThisAddIn.Application, Excel.Application)

            ' 获取当前活动单元格
            Dim activeCell As Excel.Range = excelApp.ActiveCell

            ' 检查是否有选中单元格
            If activeCell Is Nothing Then
                AddSystemMessage("错误: 没有选中的单元格。请先在Excel中选择一个单元格。")
                Return
            End If

            ' 清理公式 (移除前后空白)
            Dim cleanFormula As String = formula.Trim()

            ' 确保公式以等号开头
            If Not cleanFormula.StartsWith("=") Then
                cleanFormula = "=" & cleanFormula
            End If

            ' 显示正在应用公式的消息
            AddSystemMessage("正在应用公式到单元格 " & activeCell.Address & "...")

            ' 应用公式到单元格
            activeCell.Formula = cleanFormula

            ' 应用完成消息
            AddSystemMessage("公式已成功应用到单元格 " & activeCell.Address)

        Catch ex As Exception
            AddSystemMessage("应用Excel公式时出错: " & ex.Message)
        End Try
    End Sub

    ' 执行VBA代码
    Private Sub RunVBACode(code As String)
        Try
            ' 获取Excel应用程序实例
            Dim excelApp As Excel.Application = DirectCast(Globals.ThisAddIn.Application, Excel.Application)

            ' 在当前Excel实例中创建一个临时VBA模块并执行代码
            Dim vbProject As Object = excelApp.ActiveWorkbook.VBProject
            Dim vbComponents As Object = vbProject.VBComponents

            ' 创建一个唯一的模块名
            Dim moduleName As String = "AIModule_" & DateTime.Now.Ticks.ToString()

            ' 添加一个新的标准模块
            Dim vbComponent As Object = vbComponents.Add(1) ' 1 = vbext_ct_StdModule

            ' 设置模块名称
            vbComponent.Name = moduleName

            ' 获取代码模块
            Dim codeModule As Object = vbComponent.CodeModule

            ' 向模块添加代码
            ' 首先添加一个Sub包装我们的代码
            Dim procName As String = "ExecuteAICode_" & DateTime.Now.Ticks.ToString()

            ' 检查代码是否已经包含了Sub/Function声明
            If code.Trim().StartsWith("Sub ") OrElse code.Trim().StartsWith("Function ") Then
                ' 代码已经包含了过程声明，直接使用
                codeModule.AddFromString(code)

                ' 提取过程名
                Dim match As Match = Regex.Match(code, "(Sub|Function)\s+(\w+)")
                If match.Success Then
                    procName = match.Groups(2).Value
                Else
                    Throw New Exception("无法识别代码中的过程名")
                End If
            Else
                ' 添加自定义过程包装代码
                codeModule.AddFromString("Sub " & procName & "()" & vbCrLf &
                                        code & vbCrLf &
                                        "End Sub")
            End If

            ' 执行该过程
            AddSystemMessage("正在执行VBA代码...")
            excelApp.Run(procName)
            AddSystemMessage("VBA代码执行完成！")

            ' 稍后删除这个临时模块
            Dim tempModule As Object = vbComponent
            System.Threading.Tasks.Task.Run(Sub()
                                                Try
                                                    ' 给一些时间让代码完成执行
                                                    System.Threading.Thread.Sleep(1000)
                                                    vbComponents.Remove(tempModule)
                                                Catch ex As Exception
                                                    ' 忽略清理错误
                                                End Try
                                            End Sub)

        Catch ex As Exception
            AddSystemMessage("执行VBA代码时出错: " & ex.Message)
        End Try
    End Sub

    ' 处理包含选区内容复选框状态变化
    Private Sub chkIncludeSelection_CheckedChanged(sender As Object, e As EventArgs) Handles chkIncludeSelection.CheckedChanged
        ToggleIncludeSelection(chkIncludeSelection.Checked)

        If chkIncludeSelection.Checked Then
            AddSystemMessage("已启用选区内容导入：AI将分析您在Excel中选择的数据。")
        Else
            AddSystemMessage("已禁用选区内容导入：AI不会读取您的Excel选区。")
        End If
    End Sub

    ' GPT复选框状态变化
    Private Sub chkUseGPT_CheckedChanged(sender As Object, e As EventArgs) Handles chkUseGPT.CheckedChanged
        ' 如果正在初始化，则不处理事件
        If isInitializing Then Return

        If chkUseGPT.Checked Then
            ' 选中GPT时，取消选中DeepSeek
            chkUseDS.Checked = False
            UpdateApiSettingsInternal(True)
        ElseIf Not chkUseDS.Checked Then
            ' 如果两个都没选中，则默认选中GPT
            chkUseGPT.Checked = True
        End If
    End Sub

    ' DeepSeek复选框状态变化
    Private Sub chkUseDS_CheckedChanged(sender As Object, e As EventArgs) Handles chkUseDS.CheckedChanged
        ' 如果正在初始化，则不处理事件
        If isInitializing Then Return

        If chkUseDS.Checked Then
            ' 选中DeepSeek时，取消选中GPT
            chkUseGPT.Checked = False
            UpdateApiSettingsInternal(True)
        ElseIf Not chkUseGPT.Checked Then
            ' 如果两个都没选中，则默认选中DeepSeek
            chkUseDS.Checked = True
        End If
    End Sub

    ' 添加 Tushare 复选框状态变化处理方法
    Private Sub chkUseTushareDocs_CheckedChanged(sender As Object, e As EventArgs) Handles chkUseTushareDocs.CheckedChanged
        ' 如果正在初始化，则不处理事件
        If isInitializing Then Return

        ToggleIncludeTushareDocs(chkUseTushareDocs.Checked)

        If chkUseTushareDocs.Checked Then
            AddSystemMessage("已启用 Tushare API 指南：AI将获取可用的 Tushare 接口信息。")
        Else
            AddSystemMessage("已禁用 Tushare API 指南：AI不会获取 Tushare 接口信息。")
        End If
    End Sub


    ' 更新API设置 - 内部方法，带显示消息参数
    Private Sub UpdateApiSettingsInternal(showMessage As Boolean)
        If chkUseGPT.Checked Then
            currentApiUrl = GPT_API_URL
            currentApiKey = GPT_API_KEY
            currentModelName = "gpt-4o"
            isUsingGPT = True
            If showMessage Then
                AddSystemMessage("已切换至OpenAI GPT模型")
            End If
        Else
            currentApiUrl = DS_API_URL
            currentApiKey = DS_API_KEY
            currentModelName = "deepseek-chat"
            isUsingGPT = False
            If showMessage Then
                AddSystemMessage("已切换至DeepSeek Chat模型")
            End If
        End If
    End Sub




    ' 代码高亮的主入口函数 - 简化版
    Private Sub ApplyCodeHighlighting(ByRef rtb As RichTextBox, startPos As Integer, code As String, language As String)
        ' 保存当前的选择位置
        Dim originalSelectionStart As Integer = rtb.SelectionStart
        Dim originalSelectionLength As Integer = rtb.SelectionLength

        Try
            ' 基于语言应用不同的高亮规则
            If language.ToLower() = "vba" OrElse language.ToLower() = "vb" OrElse language.ToLower() = "visual basic" Then
                HighlightVBACode(rtb, startPos, code)
            ElseIf language.ToLower() = "excel" Then
                HighlightExcelFormula(rtb, startPos, code)
            End If
        Catch ex As Exception
            ' 忽略高亮过程中的错误
            Debug.WriteLine("高亮处理错误: " & ex.Message)
        End Try

        ' 恢复原来的选择
        rtb.SelectionStart = originalSelectionStart
        rtb.SelectionLength = originalSelectionLength
    End Sub

    ' 简化的VBA代码高亮 - 只高亮关键字
    Private Sub HighlightVBACode(ByRef rtb As RichTextBox, startPos As Integer, code As String)
        ' VBA关键字列表 (蓝色)
        Dim keywords As String() = {"Sub", "Function", "End", "If", "Then", "Else", "ElseIf", "For", "To", "Next",
                              "While", "Wend", "Do", "Loop", "Until", "Select", "Case", "Dim", "As", "Set",
                              "New", "Public", "Private", "With", "Call", "Exit", "On Error", "Resume", "GoTo"}

        ' 高亮关键字
        For Each keyword In keywords
            HighlightKeyword(rtb, startPos, code, keyword, Color.Blue)
        Next

        ' 高亮MsgBox (紫色)
        ' HighlightKeyword(rtb, startPos, code, "MsgBox", Color.FromArgb(128, 0, 128))

        ' 高亮Selection (青色)
        ' HighlightKeyword(rtb, startPos, code, "Selection", Color.FromArgb(0, 128, 128))

        ' 高亮Not, Is, Nothing (蓝色)
        ' HighlightKeyword(rtb, startPos, code, "Not", Color.Blue)
        ' HighlightKeyword(rtb, startPos, code, "Is", Color.Blue)
        ' HighlightKeyword(rtb, startPos, code, "Nothing", Color.Blue)

        ' 高亮True, False (蓝色)
        ' HighlightKeyword(rtb, startPos, code, "True", Color.Blue)
        ' HighlightKeyword(rtb, startPos, code, "False", Color.Blue)
    End Sub

    ' 高亮特定关键字
    Private Sub HighlightKeyword(ByRef rtb As RichTextBox, startPos As Integer, text As String, keyword As String, color As Color)
        Dim searchText As String = text
        Dim index As Integer = 0

        While True
            ' 查找关键字的位置
            index = searchText.IndexOf(keyword, index, StringComparison.OrdinalIgnoreCase)
            If index = -1 Then Exit While

            ' 检查是否为独立的关键字 (前后为非字母/数字/下划线)
            Dim isWordStart As Boolean = (index = 0 OrElse Not IsWordChar(text.Chars(index - 1)))
            Dim isWordEnd As Boolean = (index + keyword.Length >= text.Length OrElse Not IsWordChar(text.Chars(index + keyword.Length)))

            If isWordStart AndAlso isWordEnd Then
                ' 高亮关键字
                rtb.SelectionStart = startPos + index
                rtb.SelectionLength = keyword.Length
                rtb.SelectionColor = color
            End If

            ' 继续搜索
            index += keyword.Length
        End While
    End Sub

    ' 检查字符是否为单词字符 (字母、数字或下划线)
    Private Function IsWordChar(c As Char) As Boolean
        Return Char.IsLetterOrDigit(c) OrElse c = "_"c
    End Function


    ' 简化的Excel公式高亮
    Private Sub HighlightExcelFormula(ByRef rtb As RichTextBox, startPos As Integer, code As String)
        ' 高亮等号 (蓝色)
        For i As Integer = 0 To code.Length - 1
            If code(i) = "="c Then
                rtb.SelectionStart = startPos + i
                rtb.SelectionLength = 1
                rtb.SelectionColor = Color.FromArgb(0, 0, 128) ' 深蓝色
            End If
        Next

        ' 高亮引号内的字符串 (红色)
        Dim inString As Boolean = False
        Dim stringStart As Integer = -1

        For i As Integer = 0 To code.Length - 1
            If code(i) = """"c Then
                If inString Then
                    ' 结束字符串
                    rtb.SelectionStart = startPos + stringStart
                    rtb.SelectionLength = i - stringStart + 1
                    rtb.SelectionColor = Color.Red
                    inString = False
                Else
                    ' 开始字符串
                    stringStart = i
                    inString = True
                End If
            End If
        Next

        ' 高亮常用Excel函数 (蓝色)
        Dim excelFunctions As String() = {"SUM", "AVERAGE", "COUNT", "MAX", "MIN", "IF", "VLOOKUP", "HLOOKUP", "INDEX", "MATCH"}

        For Each func In excelFunctions
            HighlightKeyword(rtb, startPos, code, func, Color.Blue)
        Next

        ' 高亮TUSHARE函数 (紫色)
        If code.IndexOf("TUSHARE_", StringComparison.OrdinalIgnoreCase) >= 0 Then
            ' 查找所有TUSHARE_开头的函数
            Dim index As Integer = 0

            While True
                index = code.IndexOf("TUSHARE_", index, StringComparison.OrdinalIgnoreCase)
                If index = -1 Then Exit While

                ' 查找函数名的结束位置
                Dim funcEnd As Integer = index + 8 ' "TUSHARE_"的长度
                While funcEnd < code.Length AndAlso (Char.IsLetterOrDigit(code(funcEnd)) OrElse code(funcEnd) = "_"c)
                    funcEnd += 1
                End While

                ' 高亮整个函数名
                rtb.SelectionStart = startPos + index
                rtb.SelectionLength = funcEnd - index
                rtb.SelectionColor = Color.FromArgb(128, 0, 128) ' 紫色

                ' 继续搜索
                index = funcEnd
            End While
        End If
    End Sub

End Class