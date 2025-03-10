Imports System.Net
Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelDna.Integration
Imports System.Windows.Forms
Imports Microsoft.Win32

Partial Public Module TushareFunctions

    Private Const REG_PATH As String = "Software\TushareExcel"

    ''' <summary>
    ''' 获取API URL，优先从设置中读取
    ''' </summary>
    Private Function GetApiUrl() As String
        ' 尝试从注册表读取
        Try
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(REG_PATH)
                If key IsNot Nothing Then
                    Dim url = key.GetValue("ApiUrl")
                    If url IsNot Nothing Then
                        Return url.ToString()
                    End If
                End If
            End Using
        Catch ex As Exception
            ' 忽略错误，使用默认值
        End Try

        ' 使用默认值
        Return "http://api.tushare.pro"
    End Function

    ''' <summary>
    ''' 获取API Token，优先从设置中读取
    ''' </summary>
    Private Function GetToken() As String
        ' 尝试从注册表读取
        Try
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(REG_PATH)
                If key IsNot Nothing Then
                    Dim token = key.GetValue("Token")
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

    ''' <summary>
    ''' 处理Tushare API请求并返回结果数组
    ''' </summary>
    Private Function ProcessTushareRequest(requestJson As String, showHeader As Boolean, fields As String) As Object
        ' 获取API URL
        Dim apiUrl = GetApiUrl()

        Try
            ' 创建HTTP请求
            Dim request As HttpWebRequest = WebRequest.Create(apiUrl)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.Timeout = 30000 ' 30秒超时

            ' 写入请求数据
            Using streamWriter As New StreamWriter(request.GetRequestStream())
                streamWriter.Write(requestJson)
            End Using

            ' 发送请求并获取响应
            Using response As HttpWebResponse = request.GetResponse()
                Using streamReader As New StreamReader(response.GetResponseStream())
                    Dim responseJson As String = streamReader.ReadToEnd()

                    ' 解析响应
                    Dim jsonResponse As JObject = JObject.Parse(responseJson)

                    ' 检查响应码
                    If jsonResponse("code").Value(Of Integer)() = 0 Then
                        Return ExtractDataArray(jsonResponse, showHeader, fields)
                    Else
                        ' 处理错误
                        Return "错误: " & jsonResponse("msg").ToString()
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' 异常处理
            Return "异常: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 从JSON响应中提取数据数组
    ''' </summary>
    Private Function ExtractDataArray(jsonResponse As JObject, showHeader As Boolean, fieldsStr As String) As Object(,)
        ' 获取数据和字段
        Dim data As JObject = jsonResponse("data")
        Dim items As JArray = data("items")
        Dim fields As JArray = data("fields")

        ' 处理指定字段筛选
        Dim selectedFields As New List(Of String)
        Dim selectedIndexes As New List(Of Integer)

        If Not String.IsNullOrEmpty(fieldsStr) Then
            ' 解析用户指定的字段
            selectedFields = fieldsStr.Split(","c).Select(Function(f) f.Trim()).ToList()

            ' 找出指定字段在原始字段中的索引位置
            For i As Integer = 0 To fields.Count - 1
                Dim fieldName As String = fields(i).ToString()
                If selectedFields.Contains(fieldName) Then
                    selectedIndexes.Add(i)
                End If
            Next
        Else
            ' 如果未指定字段，则使用所有字段
            For i As Integer = 0 To fields.Count - 1
                selectedIndexes.Add(i)
            Next
        End If

        ' 获取行列数
        Dim rowCount As Integer = items.Count
        Dim colCount As Integer = selectedIndexes.Count

        ' 创建结果数组
        Dim resultRowCount As Integer = rowCount + If(showHeader, 1, 0)
        Dim result(resultRowCount - 1, colCount - 1) As Object

        ' 如果需要显示表头，则添加表头
        If showHeader Then
            For j As Integer = 0 To selectedIndexes.Count - 1
                result(0, j) = fields(selectedIndexes(j)).ToString()
            Next
        End If

        ' 添加数据行
        Dim resultRowStart As Integer = If(showHeader, 1, 0)
        For i As Integer = 0 To items.Count - 1
            Dim row As JArray = items(i)
            For j As Integer = 0 To selectedIndexes.Count - 1
                ' 提取实际值
                result(i + resultRowStart, j) = GetActualValue(row(selectedIndexes(j)))
            Next
        Next

        Return result
    End Function

    ''' <summary>
    ''' 从JSON值中提取实际值
    ''' </summary>
    Private Function GetActualValue(value As Object) As Object
        If value Is Nothing Then
            Return Nothing
        ElseIf TypeOf value Is JValue Then
            Return DirectCast(value, JValue).Value
        Else
            Return value.ToString()
        End If
    End Function

    ''' <summary>
    ''' 通用Tushare函数 - 支持最多10个参数
    ''' </summary>
    <ExcelFunction(Description:="通用Tushare API调用，可传入最多10个参数", Category:="Tushare")>
    Public Function TUSHARE_API(
            <ExcelArgument(Description:="API名称，如trade_cal、daily等")> apiName As String,
            <ExcelArgument(Description:="参数名1")> Optional param1Name As String = "",
            <ExcelArgument(Description:="参数值1")> Optional param1Value As String = "",
            <ExcelArgument(Description:="参数名2")> Optional param2Name As String = "",
            <ExcelArgument(Description:="参数值2")> Optional param2Value As String = "",
            <ExcelArgument(Description:="参数名3")> Optional param3Name As String = "",
            <ExcelArgument(Description:="参数值3")> Optional param3Value As String = "",
            <ExcelArgument(Description:="参数名4")> Optional param4Name As String = "",
            <ExcelArgument(Description:="参数值4")> Optional param4Value As String = "",
            <ExcelArgument(Description:="参数名5")> Optional param5Name As String = "",
            <ExcelArgument(Description:="参数值5")> Optional param5Value As String = "",
            <ExcelArgument(Description:="参数名6")> Optional param6Name As String = "",
            <ExcelArgument(Description:="参数值6")> Optional param6Value As String = "",
            <ExcelArgument(Description:="参数名7")> Optional param7Name As String = "",
            <ExcelArgument(Description:="参数值7")> Optional param7Value As String = "",
            <ExcelArgument(Description:="参数名8")> Optional param8Name As String = "",
            <ExcelArgument(Description:="参数值8")> Optional param8Value As String = "",
            <ExcelArgument(Description:="参数名9")> Optional param9Name As String = "",
            <ExcelArgument(Description:="参数值9")> Optional param9Value As String = "",
            <ExcelArgument(Description:="参数名10")> Optional param10Name As String = "",
            <ExcelArgument(Description:="参数值10")> Optional param10Value As String = "") As Object
        Try
            ' 获取Token
            Dim token As String = GetToken()
            If String.IsNullOrEmpty(token) Then
                Return "错误: 未设置Tushare API Token，请先使用配置窗口设置Token"
            End If

            ' 创建JSON对象 - 更简洁的方式
            Dim paramObj As New JObject()
            Dim requestObj As New JObject()
            Dim showHeader As Boolean = True
            Dim fields As String = ""

            ' 添加参数对到参数对象，同时检查是否有控制参数
            AddParamIfValidAndCheckControl(paramObj, param1Name, param1Value, showHeader, fields)
            AddParamIfValidAndCheckControl(paramObj, param2Name, param2Value, showHeader, fields)
            AddParamIfValidAndCheckControl(paramObj, param3Name, param3Value, showHeader, fields)
            AddParamIfValidAndCheckControl(paramObj, param4Name, param4Value, showHeader, fields)
            AddParamIfValidAndCheckControl(paramObj, param5Name, param5Value, showHeader, fields)
            AddParamIfValidAndCheckControl(paramObj, param6Name, param6Value, showHeader, fields)
            AddParamIfValidAndCheckControl(paramObj, param7Name, param7Value, showHeader, fields)
            AddParamIfValidAndCheckControl(paramObj, param8Name, param8Value, showHeader, fields)
            AddParamIfValidAndCheckControl(paramObj, param9Name, param9Value, showHeader, fields)
            AddParamIfValidAndCheckControl(paramObj, param10Name, param10Value, showHeader, fields)

            ' 构建请求对象
            requestObj("api_name") = apiName
            requestObj("token") = token
            requestObj("params") = paramObj

            ' 转换为JSON字符串
            Dim requestJson As String = requestObj.ToString(Formatting.None)

            ' 调用API并返回结果
            Return ProcessTushareRequest(requestJson, showHeader, fields)

        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 添加有效的参数到JSON对象，并检查是否为控制参数
    ''' </summary>
    Private Sub AddParamIfValidAndCheckControl(ByRef paramObj As JObject, paramName As String, paramValue As String, ByRef showHeader As Boolean, ByRef fields As String)
        If Not String.IsNullOrEmpty(paramName) AndAlso Not String.IsNullOrEmpty(paramValue) Then
            ' 检查是否为控制参数
            If paramName.ToLower() = "showheader" Then
                ' 尝试将值转换为布尔类型
                Dim headerValue As Boolean
                If Boolean.TryParse(paramValue, headerValue) Then
                    showHeader = headerValue
                End If
            ElseIf paramName.ToLower() = "fields" Then
                ' 保存字段列表
                fields = paramValue
            Else
                ' 常规API参数，添加到请求对象
                paramObj(paramName) = paramValue
            End If
        End If
    End Sub

End Module