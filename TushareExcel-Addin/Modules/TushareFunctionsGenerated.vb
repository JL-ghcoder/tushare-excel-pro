Imports System.Net
Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelDna.Integration

Partial Public Module TushareFunctions
    ' 这一部分代码全部用程序生成
    ' 如果更新了接口就直接替换这一部分代码

    ''' <summary>
    ''' 基金列表
    ''' 获取公募基金数据列表，包括场内和场外基金
    ''' </summary>
    <ExcelFunction(Description:="基金列表", Category:="Tushare")>
    Public Function TUSHARE_FUND_BASIC(
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
            ' 调用通用API函数，传入fund_basic API名称和所有参数
            Return TUSHARE_API("fund_basic", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 股票列表
    ''' 获取基础信息数据，包括股票代码、名称、上市日期、退市日期等
    ''' </summary>
    <ExcelFunction(Description:="股票列表", Category:="Tushare")>
    Public Function TUSHARE_STOCK_BASIC(
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
            ' 调用通用API函数，传入stock_basic API名称和所有参数
            Return TUSHARE_API("stock_basic", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 交易日历
    ''' 获取各大期货交易所交易日历数据
    ''' </summary>
    <ExcelFunction(Description:="交易日历", Category:="Tushare")>
    Public Function TUSHARE_TRADE_CAL(
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
            ' 调用通用API函数，传入trade_cal API名称和所有参数
            Return TUSHARE_API("trade_cal", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 日线行情
    ''' 获取股票行情数据，或通过[**通用行情接口**]( https://tushare.pro/document/2?doc_id=109)获取数据，包含了前后复权数据
    ''' </summary>
    <ExcelFunction(Description:="日线行情", Category:="Tushare")>
    Public Function TUSHARE_DAILY(
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
            ' 调用通用API函数，传入daily API名称和所有参数
            Return TUSHARE_API("daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 复权因子
    ''' 获取股票复权因子，可提取单只股票全部历史复权因子，也可以提取单日全部股票的复权因子。
    ''' </summary>
    <ExcelFunction(Description:="复权因子", Category:="Tushare")>
    Public Function TUSHARE_ADJ_FACTOR(
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
            ' 调用通用API函数，传入adj_factor API名称和所有参数
            Return TUSHARE_API("adj_factor", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 每日指标
    ''' 获取全部股票每日重要的基本面指标，可用于选股分析、报表展示等。
    ''' </summary>
    <ExcelFunction(Description:="每日指标", Category:="Tushare")>
    Public Function TUSHARE_DAILY_BASIC(
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
            ' 调用通用API函数，传入daily_basic API名称和所有参数
            Return TUSHARE_API("daily_basic", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 利润表
    ''' 获取上市公司财务利润表数据
    ''' </summary>
    <ExcelFunction(Description:="利润表", Category:="Tushare")>
    Public Function TUSHARE_INCOME(
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
            ' 调用通用API函数，传入income API名称和所有参数
            Return TUSHARE_API("income", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 资产负债表
    ''' 获取上市公司资产负债表
    ''' </summary>
    <ExcelFunction(Description:="资产负债表", Category:="Tushare")>
    Public Function TUSHARE_BALANCESHEET(
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
            ' 调用通用API函数，传入balancesheet API名称和所有参数
            Return TUSHARE_API("balancesheet", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 现金流量表
    ''' 获取上市公司现金流量表
    ''' </summary>
    <ExcelFunction(Description:="现金流量表", Category:="Tushare")>
    Public Function TUSHARE_CASHFLOW(
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
            ' 调用通用API函数，传入cashflow API名称和所有参数
            Return TUSHARE_API("cashflow", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 业绩预告
    ''' 获取业绩预告数据
    ''' </summary>
    <ExcelFunction(Description:="业绩预告", Category:="Tushare")>
    Public Function TUSHARE_FORECAST(
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
            ' 调用通用API函数，传入forecast API名称和所有参数
            Return TUSHARE_API("forecast", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 业绩快报
    ''' 获取上市公司业绩快报
    ''' </summary>
    <ExcelFunction(Description:="业绩快报", Category:="Tushare")>
    Public Function TUSHARE_EXPRESS(
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
            ' 调用通用API函数，传入express API名称和所有参数
            Return TUSHARE_API("express", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 沪深港通资金流向
    ''' 获取沪股通、深股通、港股通每日资金流向数据，每次最多返回300条记录，总量不限制。每天18~20点之间完成当日更新
    ''' </summary>
    <ExcelFunction(Description:="沪深港通资金流向", Category:="Tushare")>
    Public Function TUSHARE_MONEYFLOW_HSGT(
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
            ' 调用通用API函数，传入moneyflow_hsgt API名称和所有参数
            Return TUSHARE_API("moneyflow_hsgt", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 沪深股通十大成交股
    ''' 获取沪股通、深股通每日前十大成交详细数据，每天18~20点之间完成当日更新
    ''' </summary>
    <ExcelFunction(Description:="沪深股通十大成交股", Category:="Tushare")>
    Public Function TUSHARE_HSGT_TOP10(
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
            ' 调用通用API函数，传入hsgt_top10 API名称和所有参数
            Return TUSHARE_API("hsgt_top10", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 港股通十大成交股
    ''' 获取港股通每日成交数据，其中包括沪市、深市详细数据，每天18~20点之间完成当日更新
    ''' </summary>
    <ExcelFunction(Description:="港股通十大成交股", Category:="Tushare")>
    Public Function TUSHARE_GGT_TOP10(
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
            ' 调用通用API函数，传入ggt_top10 API名称和所有参数
            Return TUSHARE_API("ggt_top10", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 融资融券交易汇总
    ''' 获取融资融券每日交易汇总数据
    ''' </summary>
    <ExcelFunction(Description:="融资融券交易汇总", Category:="Tushare")>
    Public Function TUSHARE_MARGIN(
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
            ' 调用通用API函数，传入margin API名称和所有参数
            Return TUSHARE_API("margin", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 融资融券交易明细
    ''' 获取沪深两市每日融资融券明细
    ''' </summary>
    <ExcelFunction(Description:="融资融券交易明细", Category:="Tushare")>
    Public Function TUSHARE_MARGIN_DETAIL(
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
            ' 调用通用API函数，传入margin_detail API名称和所有参数
            Return TUSHARE_API("margin_detail", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 前十大股东
    ''' 获取上市公司前十大股东数据，包括持有数量和比例等信息
    ''' </summary>
    <ExcelFunction(Description:="前十大股东", Category:="Tushare")>
    Public Function TUSHARE_TOP10_HOLDERS(
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
            ' 调用通用API函数，传入top10_holders API名称和所有参数
            Return TUSHARE_API("top10_holders", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 前十大流通股东
    ''' 获取上市公司前十大流通股东数据
    ''' </summary>
    <ExcelFunction(Description:="前十大流通股东", Category:="Tushare")>
    Public Function TUSHARE_TOP10_FLOATHOLDERS(
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
            ' 调用通用API函数，传入top10_floatholders API名称和所有参数
            Return TUSHARE_API("top10_floatholders", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 财务指标数据
    ''' 获取上市公司财务指标数据，为避免服务器压力，现阶段每次请求最多返回60条记录，可通过设置日期多次请求获取更多数据。
    ''' </summary>
    <ExcelFunction(Description:="财务指标数据", Category:="Tushare")>
    Public Function TUSHARE_FINA_INDICATOR(
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
            ' 调用通用API函数，传入fina_indicator API名称和所有参数
            Return TUSHARE_API("fina_indicator", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 财务审计意见
    ''' 获取上市公司定期财务审计意见数据
    ''' </summary>
    <ExcelFunction(Description:="财务审计意见", Category:="Tushare")>
    Public Function TUSHARE_FINA_AUDIT(
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
            ' 调用通用API函数，传入fina_audit API名称和所有参数
            Return TUSHARE_API("fina_audit", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 主营业务构成
    ''' 获得上市公司主营业务构成，分地区和产品两种方式
    ''' </summary>
    <ExcelFunction(Description:="主营业务构成", Category:="Tushare")>
    Public Function TUSHARE_FINA_MAINBZ(
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
            ' 调用通用API函数，传入fina_mainbz API名称和所有参数
            Return TUSHARE_API("fina_mainbz", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 台湾电子产业月营收明细
    ''' 获取台湾TMT行业上市公司各类产品月度营收情况。
    ''' </summary>
    <ExcelFunction(Description:="台湾电子产业月营收明细", Category:="Tushare")>
    Public Function TUSHARE_TMT_TWINCOMEDETAIL(
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
            ' 调用通用API函数，传入tmt_twincomedetail API名称和所有参数
            Return TUSHARE_API("tmt_twincomedetail", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 台湾电子产业月营收
    ''' 获取台湾TMT电子产业领域各类产品月度营收数据。
    ''' </summary>
    <ExcelFunction(Description:="台湾电子产业月营收", Category:="Tushare")>
    Public Function TUSHARE_TMT_TWINCOME(
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
            ' 调用通用API函数，传入tmt_twincome API名称和所有参数
            Return TUSHARE_API("tmt_twincome", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 指数基本信息
    ''' 获取指数基础信息。
    ''' </summary>
    <ExcelFunction(Description:="指数基本信息", Category:="Tushare")>
    Public Function TUSHARE_INDEX_BASIC(
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
            ' 调用通用API函数，传入index_basic API名称和所有参数
            Return TUSHARE_API("index_basic", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 南华期货指数行情
    ''' 获取南华指数每日行情，指数行情也可以通过[**通用行情接口**]( https://tushare.pro/document/2?doc_id=109)获取数据．
    ''' </summary>
    <ExcelFunction(Description:="南华期货指数行情", Category:="Tushare")>
    Public Function TUSHARE_INDEX_DAILY(
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
            ' 调用通用API函数，传入index_daily API名称和所有参数
            Return TUSHARE_API("index_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 指数成分和权重
    ''' 获取各类指数成分和权重，**月度数据** ，建议输入参数里开始日期和结束日分别输入当月第一天和最后一天的日期。
    ''' </summary>
    <ExcelFunction(Description:="指数成分和权重", Category:="Tushare")>
    Public Function TUSHARE_INDEX_WEIGHT(
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
            ' 调用通用API函数，传入index_weight API名称和所有参数
            Return TUSHARE_API("index_weight", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 股票曾用名
    ''' 历史名称变更记录
    ''' </summary>
    <ExcelFunction(Description:="股票曾用名", Category:="Tushare")>
    Public Function TUSHARE_NAMECHANGE(
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
            ' 调用通用API函数，传入namechange API名称和所有参数
            Return TUSHARE_API("namechange", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 分红送股数据
    ''' 分红送股数据
    ''' </summary>
    <ExcelFunction(Description:="分红送股数据", Category:="Tushare")>
    Public Function TUSHARE_DIVIDEND(
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
            ' 调用通用API函数，传入dividend API名称和所有参数
            Return TUSHARE_API("dividend", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 沪深股通成分股
    ''' 获取沪股通、深股通成分数据
    ''' </summary>
    <ExcelFunction(Description:="沪深股通成分股", Category:="Tushare")>
    Public Function TUSHARE_HS_CONST(
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
            ' 调用通用API函数，传入hs_const API名称和所有参数
            Return TUSHARE_API("hs_const", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 龙虎榜每日统计单
    ''' 龙虎榜每日交易明细
    ''' </summary>
    <ExcelFunction(Description:="龙虎榜每日统计单", Category:="Tushare")>
    Public Function TUSHARE_TOP_LIST(
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
            ' 调用通用API函数，传入top_list API名称和所有参数
            Return TUSHARE_API("top_list", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 龙虎榜机构交易单
    ''' 龙虎榜机构成交明细
    ''' </summary>
    <ExcelFunction(Description:="龙虎榜机构交易单", Category:="Tushare")>
    Public Function TUSHARE_TOP_INST(
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
            ' 调用通用API函数，传入top_inst API名称和所有参数
            Return TUSHARE_API("top_inst", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 股权质押统计数据
    ''' 获取股票质押统计数据
    ''' </summary>
    <ExcelFunction(Description:="股权质押统计数据", Category:="Tushare")>
    Public Function TUSHARE_PLEDGE_STAT(
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
            ' 调用通用API函数，传入pledge_stat API名称和所有参数
            Return TUSHARE_API("pledge_stat", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 股权质押明细数据
    ''' 获取股票质押明细数据
    ''' </summary>
    <ExcelFunction(Description:="股权质押明细数据", Category:="Tushare")>
    Public Function TUSHARE_PLEDGE_DETAIL(
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
            ' 调用通用API函数，传入pledge_detail API名称和所有参数
            Return TUSHARE_API("pledge_detail", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 上市公司基本信息
    ''' 获取上市公司基础信息，单次提取4500条，可以根据交易所分批提取
    ''' </summary>
    <ExcelFunction(Description:="上市公司基本信息", Category:="Tushare")>
    Public Function TUSHARE_STOCK_COMPANY(
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
            ' 调用通用API函数，传入stock_company API名称和所有参数
            Return TUSHARE_API("stock_company", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 电影月度票房
    ''' 获取电影月度票房数据
    ''' </summary>
    <ExcelFunction(Description:="电影月度票房", Category:="Tushare")>
    Public Function TUSHARE_BO_MONTHLY(
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
            ' 调用通用API函数，传入bo_monthly API名称和所有参数
            Return TUSHARE_API("bo_monthly", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 电影周度票房
    ''' 获取周度票房数据
    ''' </summary>
    <ExcelFunction(Description:="电影周度票房", Category:="Tushare")>
    Public Function TUSHARE_BO_WEEKLY(
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
            ' 调用通用API函数，传入bo_weekly API名称和所有参数
            Return TUSHARE_API("bo_weekly", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 电影日度票房
    ''' 获取电影日度票房
    ''' </summary>
    <ExcelFunction(Description:="电影日度票房", Category:="Tushare")>
    Public Function TUSHARE_BO_DAILY(
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
            ' 调用通用API函数，传入bo_daily API名称和所有参数
            Return TUSHARE_API("bo_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 影院日度票房
    ''' 获取每日各影院的票房数据
    ''' </summary>
    <ExcelFunction(Description:="影院日度票房", Category:="Tushare")>
    Public Function TUSHARE_BO_CINEMA(
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
            ' 调用通用API函数，传入bo_cinema API名称和所有参数
            Return TUSHARE_API("bo_cinema", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 基金管理人
    ''' 获取公募基金管理人列表
    ''' </summary>
    <ExcelFunction(Description:="基金管理人", Category:="Tushare")>
    Public Function TUSHARE_FUND_COMPANY(
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
            ' 调用通用API函数，传入fund_company API名称和所有参数
            Return TUSHARE_API("fund_company", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 基金净值
    ''' 获取公募基金净值数据
    ''' </summary>
    <ExcelFunction(Description:="基金净值", Category:="Tushare")>
    Public Function TUSHARE_FUND_NAV(
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
            ' 调用通用API函数，传入fund_nav API名称和所有参数
            Return TUSHARE_API("fund_nav", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 基金分红
    ''' 获取公募基金分红数据
    ''' </summary>
    <ExcelFunction(Description:="基金分红", Category:="Tushare")>
    Public Function TUSHARE_FUND_DIV(
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
            ' 调用通用API函数，传入fund_div API名称和所有参数
            Return TUSHARE_API("fund_div", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 基金持仓
    ''' 获取公募基金持仓数据，季度更新
    ''' </summary>
    <ExcelFunction(Description:="基金持仓", Category:="Tushare")>
    Public Function TUSHARE_FUND_PORTFOLIO(
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
            ' 调用通用API函数，传入fund_portfolio API名称和所有参数
            Return TUSHARE_API("fund_portfolio", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' IPO新股上市
    ''' 获取新股上市列表数据
    ''' </summary>
    <ExcelFunction(Description:="IPO新股上市", Category:="Tushare")>
    Public Function TUSHARE_NEW_SHARE(
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
            ' 调用通用API函数，传入new_share API名称和所有参数
            Return TUSHARE_API("new_share", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 股票回购
    ''' 获取上市公司回购股票数据
    ''' </summary>
    <ExcelFunction(Description:="股票回购", Category:="Tushare")>
    Public Function TUSHARE_REPURCHASE(
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
            ' 调用通用API函数，传入repurchase API名称和所有参数
            Return TUSHARE_API("repurchase", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 概念股分类表
    ''' 获取概念股分类，目前只有ts一个来源，未来将逐步增加来源
    ''' </summary>
    <ExcelFunction(Description:="概念股分类表", Category:="Tushare")>
    Public Function TUSHARE_CONCEPT(
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
            ' 调用通用API函数，传入concept API名称和所有参数
            Return TUSHARE_API("concept", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 概念股明细列表
    ''' 获取概念股分类明细数据
    ''' </summary>
    <ExcelFunction(Description:="概念股明细列表", Category:="Tushare")>
    Public Function TUSHARE_CONCEPT_DETAIL(
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
            ' 调用通用API函数，传入concept_detail API名称和所有参数
            Return TUSHARE_API("concept_detail", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 基金行情（含ETF）
    ''' 获取场内基金日线行情，类似股票日行情，包括ETF行情
    ''' </summary>
    <ExcelFunction(Description:="基金行情（含ETF）", Category:="Tushare")>
    Public Function TUSHARE_FUND_DAILY(
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
            ' 调用通用API函数，传入fund_daily API名称和所有参数
            Return TUSHARE_API("fund_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 大盘指数每日指标
    ''' 目前只提供上证综指，深证成指，上证50，中证500，中小板指，创业板指的每日指标数据
    ''' </summary>
    <ExcelFunction(Description:="大盘指数每日指标", Category:="Tushare")>
    Public Function TUSHARE_INDEX_DAILYBASIC(
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
            ' 调用通用API函数，传入index_dailybasic API名称和所有参数
            Return TUSHARE_API("index_dailybasic", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 合约信息
    ''' 获取期货合约列表数据
    ''' </summary>
    <ExcelFunction(Description:="合约信息", Category:="Tushare")>
    Public Function TUSHARE_FUT_BASIC(
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
            ' 调用通用API函数，传入fut_basic API名称和所有参数
            Return TUSHARE_API("fut_basic", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 日线行情
    ''' 期货日线行情数据
    ''' </summary>
    <ExcelFunction(Description:="日线行情", Category:="Tushare")>
    Public Function TUSHARE_FUT_DAILY(
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
            ' 调用通用API函数，传入fut_daily API名称和所有参数
            Return TUSHARE_API("fut_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 每日持仓排名
    ''' 获取每日成交持仓排名数据
    ''' </summary>
    <ExcelFunction(Description:="每日持仓排名", Category:="Tushare")>
    Public Function TUSHARE_FUT_HOLDING(
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
            ' 调用通用API函数，传入fut_holding API名称和所有参数
            Return TUSHARE_API("fut_holding", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 仓单日报
    ''' 获取仓单日报数据，了解各仓库/厂库的仓单变化
    ''' </summary>
    <ExcelFunction(Description:="仓单日报", Category:="Tushare")>
    Public Function TUSHARE_FUT_WSR(
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
            ' 调用通用API函数，传入fut_wsr API名称和所有参数
            Return TUSHARE_API("fut_wsr", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 每日结算参数
    ''' 获取每日结算参数数据，包括交易和交割费率等
    ''' </summary>
    <ExcelFunction(Description:="每日结算参数", Category:="Tushare")>
    Public Function TUSHARE_FUT_SETTLE(
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
            ' 调用通用API函数，传入fut_settle API名称和所有参数
            Return TUSHARE_API("fut_settle", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 新闻快讯
    ''' 获取主流新闻网站的快讯新闻数据
    ''' </summary>
    <ExcelFunction(Description:="新闻快讯", Category:="Tushare")>
    Public Function TUSHARE_NEWS(
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
            ' 调用通用API函数，传入news API名称和所有参数
            Return TUSHARE_API("news", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 周线行情
    ''' 获取A股周线行情
    ''' </summary>
    <ExcelFunction(Description:="周线行情", Category:="Tushare")>
    Public Function TUSHARE_WEEKLY(
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
            ' 调用通用API函数，传入weekly API名称和所有参数
            Return TUSHARE_API("weekly", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 月线行情
    ''' 获取A股月线数据
    ''' </summary>
    <ExcelFunction(Description:="月线行情", Category:="Tushare")>
    Public Function TUSHARE_MONTHLY(
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
            ' 调用通用API函数，传入monthly API名称和所有参数
            Return TUSHARE_API("monthly", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' Shibor利率
    ''' shibor利率
    ''' </summary>
    <ExcelFunction(Description:="Shibor利率", Category:="Tushare")>
    Public Function TUSHARE_SHIBOR(
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
            ' 调用通用API函数，传入shibor API名称和所有参数
            Return TUSHARE_API("shibor", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' Shibor报价数据
    ''' Shibor报价数据
    ''' </summary>
    <ExcelFunction(Description:="Shibor报价数据", Category:="Tushare")>
    Public Function TUSHARE_SHIBOR_QUOTE(
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
            ' 调用通用API函数，传入shibor_quote API名称和所有参数
            Return TUSHARE_API("shibor_quote", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' LPR贷款基础利率
    ''' LPR贷款基础利率
    ''' </summary>
    <ExcelFunction(Description:="LPR贷款基础利率", Category:="Tushare")>
    Public Function TUSHARE_SHIBOR_LPR(
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
            ' 调用通用API函数，传入shibor_lpr API名称和所有参数
            Return TUSHARE_API("shibor_lpr", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' Libor利率
    ''' Libor拆借利率
    ''' </summary>
    <ExcelFunction(Description:="Libor利率", Category:="Tushare")>
    Public Function TUSHARE_LIBOR(
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
            ' 调用通用API函数，传入libor API名称和所有参数
            Return TUSHARE_API("libor", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' Hibor利率
    ''' Hibor利率
    ''' </summary>
    <ExcelFunction(Description:="Hibor利率", Category:="Tushare")>
    Public Function TUSHARE_HIBOR(
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
            ' 调用通用API函数，传入hibor API名称和所有参数
            Return TUSHARE_API("hibor", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 新闻联播文字稿
    ''' 获取新闻联播文字稿数据，数据开始于2006年6月，超过12年历史
    ''' </summary>
    <ExcelFunction(Description:="新闻联播文字稿", Category:="Tushare")>
    Public Function TUSHARE_CCTV_NEWS(
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
            ' 调用通用API函数，传入cctv_news API名称和所有参数
            Return TUSHARE_API("cctv_news", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 全国电影剧本备案数据
    ''' 获取全国电影剧本备案的公示数据
    ''' </summary>
    <ExcelFunction(Description:="全国电影剧本备案数据", Category:="Tushare")>
    Public Function TUSHARE_FILM_RECORD(
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
            ' 调用通用API函数，传入film_record API名称和所有参数
            Return TUSHARE_API("film_record", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 期权合约信息
    ''' 获取期权合约信息
    ''' </summary>
    <ExcelFunction(Description:="期权合约信息", Category:="Tushare")>
    Public Function TUSHARE_OPT_BASIC(
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
            ' 调用通用API函数，传入opt_basic API名称和所有参数
            Return TUSHARE_API("opt_basic", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 期权日线行情
    ''' 获取期权日线行情
    ''' </summary>
    <ExcelFunction(Description:="期权日线行情", Category:="Tushare")>
    Public Function TUSHARE_OPT_DAILY(
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
            ' 调用通用API函数，传入opt_daily API名称和所有参数
            Return TUSHARE_API("opt_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 限售股解禁
    ''' 获取限售股解禁
    ''' </summary>
    <ExcelFunction(Description:="限售股解禁", Category:="Tushare")>
    Public Function TUSHARE_SHARE_FLOAT(
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
            ' 调用通用API函数，传入share_float API名称和所有参数
            Return TUSHARE_API("share_float", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 大宗交易
    ''' 大宗交易
    ''' </summary>
    <ExcelFunction(Description:="大宗交易", Category:="Tushare")>
    Public Function TUSHARE_BLOCK_TRADE(
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
            ' 调用通用API函数，传入block_trade API名称和所有参数
            Return TUSHARE_API("block_trade", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 财报披露日期表
    ''' 获取财报披露计划日期
    ''' </summary>
    <ExcelFunction(Description:="财报披露日期表", Category:="Tushare")>
    Public Function TUSHARE_DISCLOSURE_DATE(
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
            ' 调用通用API函数，传入disclosure_date API名称和所有参数
            Return TUSHARE_API("disclosure_date", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 股票开户数据（停）
    ''' 获取股票账户开户数据，统计周期为一周
    ''' </summary>
    <ExcelFunction(Description:="股票开户数据（停）", Category:="Tushare")>
    Public Function TUSHARE_STK_ACCOUNT(
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
            ' 调用通用API函数，传入stk_account API名称和所有参数
            Return TUSHARE_API("stk_account", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 股票开户数据（旧）
    ''' 获取股票账户开户数据旧版格式数据，数据从2008年1月开始，到2015年5月29，新数据请通过[股票开户数据](https://tushare.pro/document/2?doc_id=164)获取。
    ''' </summary>
    <ExcelFunction(Description:="股票开户数据（旧）", Category:="Tushare")>
    Public Function TUSHARE_STK_ACCOUNT_OLD(
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
            ' 调用通用API函数，传入stk_account_old API名称和所有参数
            Return TUSHARE_API("stk_account_old", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 股东人数
    ''' 获取上市公司股东户数数据，数据不定期公布
    ''' </summary>
    <ExcelFunction(Description:="股东人数", Category:="Tushare")>
    Public Function TUSHARE_STK_HOLDERNUMBER(
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
            ' 调用通用API函数，传入stk_holdernumber API名称和所有参数
            Return TUSHARE_API("stk_holdernumber", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 个股资金流向
    ''' 获取沪深A股票资金流向数据，分析大单小单成交情况，用于判别资金动向，数据开始于2010年。
    ''' </summary>
    <ExcelFunction(Description:="个股资金流向", Category:="Tushare")>
    Public Function TUSHARE_MONEYFLOW(
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
            ' 调用通用API函数，传入moneyflow API名称和所有参数
            Return TUSHARE_API("moneyflow", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 指数周线行情
    ''' 获取指数周线行情
    ''' </summary>
    <ExcelFunction(Description:="指数周线行情", Category:="Tushare")>
    Public Function TUSHARE_INDEX_WEEKLY(
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
            ' 调用通用API函数，传入index_weekly API名称和所有参数
            Return TUSHARE_API("index_weekly", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 指数月线行情
    ''' 获取指数月线行情,每月更新一次
    ''' </summary>
    <ExcelFunction(Description:="指数月线行情", Category:="Tushare")>
    Public Function TUSHARE_INDEX_MONTHLY(
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
            ' 调用通用API函数，传入index_monthly API名称和所有参数
            Return TUSHARE_API("index_monthly", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 温州民间借贷利率
    ''' 温州民间借贷利率，即温州指数
    ''' </summary>
    <ExcelFunction(Description:="温州民间借贷利率", Category:="Tushare")>
    Public Function TUSHARE_WZ_INDEX(
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
            ' 调用通用API函数，传入wz_index API名称和所有参数
            Return TUSHARE_API("wz_index", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 广州民间借贷利率
    ''' 广州民间借贷利率
    ''' </summary>
    <ExcelFunction(Description:="广州民间借贷利率", Category:="Tushare")>
    Public Function TUSHARE_GZ_INDEX(
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
            ' 调用通用API函数，传入gz_index API名称和所有参数
            Return TUSHARE_API("gz_index", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 股东增减持
    ''' 获取上市公司增减持数据，了解重要股东近期及历史上的股份增减变化
    ''' </summary>
    <ExcelFunction(Description:="股东增减持", Category:="Tushare")>
    Public Function TUSHARE_STK_HOLDERTRADE(
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
            ' 调用通用API函数，传入stk_holdertrade API名称和所有参数
            Return TUSHARE_API("stk_holdertrade", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 上市公司公告
    ''' 获取全量公告数据，提供pdf下载URL
    ''' </summary>
    <ExcelFunction(Description:="上市公司公告", Category:="Tushare")>
    Public Function TUSHARE_ANNS_D(
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
            ' 调用通用API函数，传入anns_d API名称和所有参数
            Return TUSHARE_API("anns_d", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 外汇基础信息（海外）
    ''' 获取海外外汇基础信息，目前只有FXCM交易商的数据
    ''' </summary>
    <ExcelFunction(Description:="外汇基础信息（海外）", Category:="Tushare")>
    Public Function TUSHARE_FX_OBASIC(
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
            ' 调用通用API函数，传入fx_obasic API名称和所有参数
            Return TUSHARE_API("fx_obasic", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 外汇日线行情
    ''' 获取外汇日线行情
    ''' </summary>
    <ExcelFunction(Description:="外汇日线行情", Category:="Tushare")>
    Public Function TUSHARE_FX_DAILY(
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
            ' 调用通用API函数，传入fx_daily API名称和所有参数
            Return TUSHARE_API("fx_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 全国电视剧备案公示数据
    ''' 获取2009年以来全国拍摄制作电视剧备案公示数据
    ''' </summary>
    <ExcelFunction(Description:="全国电视剧备案公示数据", Category:="Tushare")>
    Public Function TUSHARE_TELEPLAY_RECORD(
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
            ' 调用通用API函数，传入teleplay_record API名称和所有参数
            Return TUSHARE_API("teleplay_record", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 申万行业分类
    ''' 获取申万行业分类，可以获取申万2014年版本（28个一级分类，104个二级分类，227个三级分类）和2021年本版（31个一级分类，134个二级分类，346个三级分类）列表信息
    ''' </summary>
    <ExcelFunction(Description:="申万行业分类", Category:="Tushare")>
    Public Function TUSHARE_INDEX_CLASSIFY(
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
            ' 调用通用API函数，传入index_classify API名称和所有参数
            Return TUSHARE_API("index_classify", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 每日涨跌停价格
    ''' 获取全市场（包含A/B股和基金）每日涨跌停价格，包括涨停价格，跌停价格等，每个交易日8点40左右更新当日股票涨跌停价格。
    ''' </summary>
    <ExcelFunction(Description:="每日涨跌停价格", Category:="Tushare")>
    Public Function TUSHARE_STK_LIMIT(
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
            ' 调用通用API函数，传入stk_limit API名称和所有参数
            Return TUSHARE_API("stk_limit", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 可转债基础信息
    ''' 获取可转债基本信息
    ''' </summary>
    <ExcelFunction(Description:="可转债基础信息", Category:="Tushare")>
    Public Function TUSHARE_CB_BASIC(
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
            ' 调用通用API函数，传入cb_basic API名称和所有参数
            Return TUSHARE_API("cb_basic", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 可转债发行
    ''' 获取可转债发行数据
    ''' </summary>
    <ExcelFunction(Description:="可转债发行", Category:="Tushare")>
    Public Function TUSHARE_CB_ISSUE(
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
            ' 调用通用API函数，传入cb_issue API名称和所有参数
            Return TUSHARE_API("cb_issue", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 可转债行情
    ''' 获取可转债行情
    ''' </summary>
    <ExcelFunction(Description:="可转债行情", Category:="Tushare")>
    Public Function TUSHARE_CB_DAILY(
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
            ' 调用通用API函数，传入cb_daily API名称和所有参数
            Return TUSHARE_API("cb_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 沪深股通持股明细
    ''' 获取沪深港股通持股明细，数据来源港交所。
    ''' </summary>
    <ExcelFunction(Description:="沪深股通持股明细", Category:="Tushare")>
    Public Function TUSHARE_HK_HOLD(
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
            ' 调用通用API函数，传入hk_hold API名称和所有参数
            Return TUSHARE_API("hk_hold", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 期货主力与连续合约
    ''' 获取期货主力（或连续）合约与月合约映射数据
    ''' </summary>
    <ExcelFunction(Description:="期货主力与连续合约", Category:="Tushare")>
    Public Function TUSHARE_FUT_MAPPING(
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
            ' 调用通用API函数，传入fut_mapping API名称和所有参数
            Return TUSHARE_API("fut_mapping", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 港股基础信息
    ''' 获取港股列表信息
    ''' </summary>
    <ExcelFunction(Description:="港股基础信息", Category:="Tushare")>
    Public Function TUSHARE_HK_BASIC(
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
            ' 调用通用API函数，传入hk_basic API名称和所有参数
            Return TUSHARE_API("hk_basic", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 港股日线行情
    ''' 获取港股每日增量和历史行情，每日18点左右更新当日数据
    ''' </summary>
    <ExcelFunction(Description:="港股日线行情", Category:="Tushare")>
    Public Function TUSHARE_HK_DAILY(
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
            ' 调用通用API函数，传入hk_daily API名称和所有参数
            Return TUSHARE_API("hk_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 上市公司管理层
    ''' 获取上市公司管理层
    ''' </summary>
    <ExcelFunction(Description:="上市公司管理层", Category:="Tushare")>
    Public Function TUSHARE_STK_MANAGERS(
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
            ' 调用通用API函数，传入stk_managers API名称和所有参数
            Return TUSHARE_API("stk_managers", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 管理层薪酬和持股
    ''' 获取上市公司管理层薪酬和持股
    ''' </summary>
    <ExcelFunction(Description:="管理层薪酬和持股", Category:="Tushare")>
    Public Function TUSHARE_STK_REWARDS(
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
            ' 调用通用API函数，传入stk_rewards API名称和所有参数
            Return TUSHARE_API("stk_rewards", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 新闻通讯（长篇）
    ''' 获取长篇通讯信息，覆盖主要新闻资讯网站
    ''' </summary>
    <ExcelFunction(Description:="新闻通讯（长篇）", Category:="Tushare")>
    Public Function TUSHARE_MAJOR_NEWS(
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
            ' 调用通用API函数，传入major_news API名称和所有参数
            Return TUSHARE_API("major_news", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 港股通每日成交统计
    ''' 获取港股通每日成交信息，数据从2014年开始
    ''' </summary>
    <ExcelFunction(Description:="港股通每日成交统计", Category:="Tushare")>
    Public Function TUSHARE_GGT_DAILY(
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
            ' 调用通用API函数，传入ggt_daily API名称和所有参数
            Return TUSHARE_API("ggt_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 港股通每月成交统计
    ''' 港股通每月成交信息，数据从2014年开始
    ''' </summary>
    <ExcelFunction(Description:="港股通每月成交统计", Category:="Tushare")>
    Public Function TUSHARE_GGT_MONTHLY(
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
            ' 调用通用API函数，传入ggt_monthly API名称和所有参数
            Return TUSHARE_API("ggt_monthly", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 复权因子
    ''' 获取基金复权因子，用于计算基金复权行情
    ''' </summary>
    <ExcelFunction(Description:="复权因子", Category:="Tushare")>
    Public Function TUSHARE_FUND_ADJ(
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
            ' 调用通用API函数，传入fund_adj API名称和所有参数
            Return TUSHARE_API("fund_adj", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 国债收益率曲线
    ''' 获取中债收益率曲线，目前可获取中债国债收益率曲线即期和到期收益率曲线数据
    ''' </summary>
    <ExcelFunction(Description:="国债收益率曲线", Category:="Tushare")>
    Public Function TUSHARE_YC_CB(
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
            ' 调用通用API函数，传入yc_cb API名称和所有参数
            Return TUSHARE_API("yc_cb", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 基金规模
    ''' 获取基金规模数据，包含上海和深圳ETF基金
    ''' </summary>
    <ExcelFunction(Description:="基金规模", Category:="Tushare")>
    Public Function TUSHARE_FUND_SHARE(
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
            ' 调用通用API函数，传入fund_share API名称和所有参数
            Return TUSHARE_API("fund_share", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 基金经理
    ''' 获取公募基金经理数据，包括基金经理简历等数据
    ''' </summary>
    <ExcelFunction(Description:="基金经理", Category:="Tushare")>
    Public Function TUSHARE_FUND_MANAGER(
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
            ' 调用通用API函数，传入fund_manager API名称和所有参数
            Return TUSHARE_API("fund_manager", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 国际主要指数
    ''' 获取国际主要指数日线行情
    ''' </summary>
    <ExcelFunction(Description:="国际主要指数", Category:="Tushare")>
    Public Function TUSHARE_INDEX_GLOBAL(
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
            ' 调用通用API函数，传入index_global API名称和所有参数
            Return TUSHARE_API("index_global", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 每日停复牌信息
    ''' 按日期方式获取股票每日停复牌信息
    ''' </summary>
    <ExcelFunction(Description:="每日停复牌信息", Category:="Tushare")>
    Public Function TUSHARE_SUSPEND_D(
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
            ' 调用通用API函数，传入suspend_d API名称和所有参数
            Return TUSHARE_API("suspend_d", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 沪深市场每日交易统计
    ''' 获取交易所股票交易统计，包括各板块明细
    ''' </summary>
    <ExcelFunction(Description:="沪深市场每日交易统计", Category:="Tushare")>
    Public Function TUSHARE_DAILY_INFO(
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
            ' 调用通用API函数，传入daily_info API名称和所有参数
            Return TUSHARE_API("daily_info", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 期货主要品种交易周报
    ''' 获取期货交易所主要品种每周交易统计信息，数据从2010年3月开始
    ''' </summary>
    <ExcelFunction(Description:="期货主要品种交易周报", Category:="Tushare")>
    Public Function TUSHARE_FUT_WEEKLY_DETAIL(
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
            ' 调用通用API函数，传入fut_weekly_detail API名称和所有参数
            Return TUSHARE_API("fut_weekly_detail", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 国债收益率曲线利率
    ''' 获取美国每日国债收益率曲线利率
    ''' </summary>
    <ExcelFunction(Description:="国债收益率曲线利率", Category:="Tushare")>
    Public Function TUSHARE_US_TYCR(
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
            ' 调用通用API函数，传入us_tycr API名称和所有参数
            Return TUSHARE_API("us_tycr", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 国债实际收益率曲线利率
    ''' 国债实际收益率曲线利率
    ''' </summary>
    <ExcelFunction(Description:="国债实际收益率曲线利率", Category:="Tushare")>
    Public Function TUSHARE_US_TRYCR(
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
            ' 调用通用API函数，传入us_trycr API名称和所有参数
            Return TUSHARE_API("us_trycr", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 短期国债利率
    ''' 获取美国短期国债利率数据
    ''' </summary>
    <ExcelFunction(Description:="短期国债利率", Category:="Tushare")>
    Public Function TUSHARE_US_TBR(
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
            ' 调用通用API函数，传入us_tbr API名称和所有参数
            Return TUSHARE_API("us_tbr", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 国债长期利率
    ''' 国债长期利率
    ''' </summary>
    <ExcelFunction(Description:="国债长期利率", Category:="Tushare")>
    Public Function TUSHARE_US_TLTR(
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
            ' 调用通用API函数，传入us_tltr API名称和所有参数
            Return TUSHARE_API("us_tltr", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 国债长期利率平均值
    ''' 国债实际长期利率平均值
    ''' </summary>
    <ExcelFunction(Description:="国债长期利率平均值", Category:="Tushare")>
    Public Function TUSHARE_US_TRLTR(
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
            ' 调用通用API函数，传入us_trltr API名称和所有参数
            Return TUSHARE_API("us_trltr", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 国内生产总值（GDP）
    ''' 获取国民经济之GDP数据
    ''' </summary>
    <ExcelFunction(Description:="国内生产总值（GDP）", Category:="Tushare")>
    Public Function TUSHARE_CN_GDP(
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
            ' 调用通用API函数，传入cn_gdp API名称和所有参数
            Return TUSHARE_API("cn_gdp", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 居民消费价格指数（CPI）
    ''' 获取CPI居民消费价格数据，包括全国、城市和农村的数据
    ''' </summary>
    <ExcelFunction(Description:="居民消费价格指数（CPI）", Category:="Tushare")>
    Public Function TUSHARE_CN_CPI(
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
            ' 调用通用API函数，传入cn_cpi API名称和所有参数
            Return TUSHARE_API("cn_cpi", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 全球财经事件
    ''' 获取全球财经日历、包括经济事件数据更新
    ''' </summary>
    <ExcelFunction(Description:="全球财经事件", Category:="Tushare")>
    Public Function TUSHARE_ECO_CAL(
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
            ' 调用通用API函数，传入eco_cal API名称和所有参数
            Return TUSHARE_API("eco_cal", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 货币供应量（月）
    ''' 获取货币供应量之月度数据
    ''' </summary>
    <ExcelFunction(Description:="货币供应量（月）", Category:="Tushare")>
    Public Function TUSHARE_CN_M(
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
            ' 调用通用API函数，传入cn_m API名称和所有参数
            Return TUSHARE_API("cn_m", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 工业生产者出厂价格指数（PPI）
    ''' 获取PPI工业生产者出厂价格指数数据
    ''' </summary>
    <ExcelFunction(Description:="工业生产者出厂价格指数（PPI）", Category:="Tushare")>
    Public Function TUSHARE_CN_PPI(
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
            ' 调用通用API函数，传入cn_ppi API名称和所有参数
            Return TUSHARE_API("cn_ppi", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 可转债转股价变动
    ''' 获取可转债转股价变动
    ''' </summary>
    <ExcelFunction(Description:="可转债转股价变动", Category:="Tushare")>
    Public Function TUSHARE_CB_PRICE_CHG(
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
            ' 调用通用API函数，传入cb_price_chg API名称和所有参数
            Return TUSHARE_API("cb_price_chg", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 可转债转股结果
    ''' 获取可转债转股结果
    ''' </summary>
    <ExcelFunction(Description:="可转债转股结果", Category:="Tushare")>
    Public Function TUSHARE_CB_SHARE(
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
            ' 调用通用API函数，传入cb_share API名称和所有参数
            Return TUSHARE_API("cb_share", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 港股交易日历
    ''' 获取交易日历
    ''' </summary>
    <ExcelFunction(Description:="港股交易日历", Category:="Tushare")>
    Public Function TUSHARE_HK_TRADECAL(
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
            ' 调用通用API函数，传入hk_tradecal API名称和所有参数
            Return TUSHARE_API("hk_tradecal", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 美股基础信息
    ''' 获取美股列表信息
    ''' </summary>
    <ExcelFunction(Description:="美股基础信息", Category:="Tushare")>
    Public Function TUSHARE_US_BASIC(
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
            ' 调用通用API函数，传入us_basic API名称和所有参数
            Return TUSHARE_API("us_basic", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 美股交易日历
    ''' 获取美股交易日历信息
    ''' </summary>
    <ExcelFunction(Description:="美股交易日历", Category:="Tushare")>
    Public Function TUSHARE_US_TRADECAL(
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
            ' 调用通用API函数，传入us_tradecal API名称和所有参数
            Return TUSHARE_API("us_tradecal", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 美股日线行情
    ''' 获取美股行情（未复权），包括全部股票全历史行情，以及重要的市场和估值指标
    ''' </summary>
    <ExcelFunction(Description:="美股日线行情", Category:="Tushare")>
    Public Function TUSHARE_US_DAILY(
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
            ' 调用通用API函数，传入us_daily API名称和所有参数
            Return TUSHARE_API("us_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 备用行情
    ''' 获取备用行情，包括特定的行情指标(数据从2017年中左右开始，早期有几天数据缺失，近期正常)
    ''' </summary>
    <ExcelFunction(Description:="备用行情", Category:="Tushare")>
    Public Function TUSHARE_BAK_DAILY(
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
            ' 调用通用API函数，传入bak_daily API名称和所有参数
            Return TUSHARE_API("bak_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 债券回购日行情
    ''' 债券回购日行情
    ''' </summary>
    <ExcelFunction(Description:="债券回购日行情", Category:="Tushare")>
    Public Function TUSHARE_REPO_DAILY(
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
            ' 调用通用API函数，传入repo_daily API名称和所有参数
            Return TUSHARE_API("repo_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 同花顺行业概念板块
    ''' 获取同花顺板块指数。注：数据版权归属同花顺，如做商业用途，请主动联系同花顺，如需帮助请联系微信：waditu_a
    ''' </summary>
    <ExcelFunction(Description:="同花顺行业概念板块", Category:="Tushare")>
    Public Function TUSHARE_THS_INDEX(
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
            ' 调用通用API函数，传入ths_index API名称和所有参数
            Return TUSHARE_API("ths_index", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 同花顺概念和行业指数行情
    ''' 获取同花顺板块指数行情。注：数据版权归属同花顺，如做商业用途，请主动联系同花顺，如需帮助请联系微信：waditu_a
    ''' </summary>
    <ExcelFunction(Description:="同花顺概念和行业指数行情", Category:="Tushare")>
    Public Function TUSHARE_THS_DAILY(
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
            ' 调用通用API函数，传入ths_daily API名称和所有参数
            Return TUSHARE_API("ths_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 同花顺行业概念成分
    ''' 获取同花顺概念板块成分列表注：数据版权归属同花顺，如做商业用途，请主动联系同花顺。
    ''' </summary>
    <ExcelFunction(Description:="同花顺行业概念成分", Category:="Tushare")>
    Public Function TUSHARE_THS_MEMBER(
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
            ' 调用通用API函数，传入ths_member API名称和所有参数
            Return TUSHARE_API("ths_member", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 股票历史列表
    ''' 获取备用基础列表，数据从2016年开始
    ''' </summary>
    <ExcelFunction(Description:="股票历史列表", Category:="Tushare")>
    Public Function TUSHARE_BAK_BASIC(
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
            ' 调用通用API函数，传入bak_basic API名称和所有参数
            Return TUSHARE_API("bak_basic", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 各渠道公募基金销售保有规模占比
    ''' 获取各渠道公募基金销售保有规模占比数据，年度更新
    ''' </summary>
    <ExcelFunction(Description:="各渠道公募基金销售保有规模占比", Category:="Tushare")>
    Public Function TUSHARE_FUND_SALES_RATIO(
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
            ' 调用通用API函数，传入fund_sales_ratio API名称和所有参数
            Return TUSHARE_API("fund_sales_ratio", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 销售机构公募基金销售保有规模
    ''' 获取销售机构公募基金销售保有规模数据，本数据从2021年Q1开始公布，季度更新
    ''' </summary>
    <ExcelFunction(Description:="销售机构公募基金销售保有规模", Category:="Tushare")>
    Public Function TUSHARE_FUND_SALES_VOL(
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
            ' 调用通用API函数，传入fund_sales_vol API名称和所有参数
            Return TUSHARE_API("fund_sales_vol", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 券商月度金股
    ''' 获取券商月度金股，一般1日~3日内更新当月数据
    ''' </summary>
    <ExcelFunction(Description:="券商月度金股", Category:="Tushare")>
    Public Function TUSHARE_BROKER_RECOMMEND(
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
            ' 调用通用API函数，传入broker_recommend API名称和所有参数
            Return TUSHARE_API("broker_recommend", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 深圳市场每日交易情况
    ''' 获取深圳市场每日交易概况
    ''' </summary>
    <ExcelFunction(Description:="深圳市场每日交易情况", Category:="Tushare")>
    Public Function TUSHARE_SZ_DAILY_INFO(
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
            ' 调用通用API函数，传入sz_daily_info API名称和所有参数
            Return TUSHARE_API("sz_daily_info", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 可转债赎回信息
    ''' 获取可转债到期赎回、强制赎回等信息。数据来源于公开披露渠道，供个人和机构研究使用，请不要用于数据商业目的。
    ''' </summary>
    <ExcelFunction(Description:="可转债赎回信息", Category:="Tushare")>
    Public Function TUSHARE_CB_CALL(
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
            ' 调用通用API函数，传入cb_call API名称和所有参数
            Return TUSHARE_API("cb_call", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 大宗交易
    ''' 获取沪深交易所债券大宗交易数据，可以通过[**数据工具**](https://tushare.pro/webclient/)调试和查看数据。
    ''' </summary>
    <ExcelFunction(Description:="大宗交易", Category:="Tushare")>
    Public Function TUSHARE_BOND_BLK(
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
            ' 调用通用API函数，传入bond_blk API名称和所有参数
            Return TUSHARE_API("bond_blk", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 大宗交易明细
    ''' 获取沪深交易所债券大宗交易数据，可以通过[**数据工具**](https://tushare.pro/webclient/)调试和查看数据。
    ''' </summary>
    <ExcelFunction(Description:="大宗交易明细", Category:="Tushare")>
    Public Function TUSHARE_BOND_BLK_DETAIL(
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
            ' 调用通用API函数，传入bond_blk_detail API名称和所有参数
            Return TUSHARE_API("bond_blk_detail", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 中央结算系统持股明细
    ''' 获取中央结算系统机构席位持股明细，数据覆盖**全历史**，根据交易所披露时间，当日数据在下一交易日早上9点前完成
    ''' </summary>
    <ExcelFunction(Description:="中央结算系统持股明细", Category:="Tushare")>
    Public Function TUSHARE_CCASS_HOLD_DETAIL(
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
            ' 调用通用API函数，传入ccass_hold_detail API名称和所有参数
            Return TUSHARE_API("ccass_hold_detail", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 机构调研数据
    ''' 获取上市公司机构调研记录数据
    ''' </summary>
    <ExcelFunction(Description:="机构调研数据", Category:="Tushare")>
    Public Function TUSHARE_STK_SURV(
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
            ' 调用通用API函数，传入stk_surv API名称和所有参数
            Return TUSHARE_API("stk_surv", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 上海黄金基础信息
    ''' 获取上海黄金交易所现货合约基础信息
    ''' </summary>
    <ExcelFunction(Description:="上海黄金基础信息", Category:="Tushare")>
    Public Function TUSHARE_SGE_BASIC(
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
            ' 调用通用API函数，传入sge_basic API名称和所有参数
            Return TUSHARE_API("sge_basic", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 上海黄金现货日行情
    ''' 获取上海黄金交易所现货合约日线行情
    ''' </summary>
    <ExcelFunction(Description:="上海黄金现货日行情", Category:="Tushare")>
    Public Function TUSHARE_SGE_DAILY(
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
            ' 调用通用API函数，传入sge_daily API名称和所有参数
            Return TUSHARE_API("sge_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 券商盈利预测数据
    ''' 获取券商（卖方）每天研报的盈利预测数据，数据从2010年开始，每晚19~22点更新当日数据
    ''' </summary>
    <ExcelFunction(Description:="券商盈利预测数据", Category:="Tushare")>
    Public Function TUSHARE_REPORT_RC(
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
            ' 调用通用API函数，传入report_rc API名称和所有参数
            Return TUSHARE_API("report_rc", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 每日筹码及胜率
    ''' 获取A股每日筹码平均成本和胜率情况，每天17~18点左右更新，数据从2018年开始
    ''' </summary>
    <ExcelFunction(Description:="每日筹码及胜率", Category:="Tushare")>
    Public Function TUSHARE_CYQ_PERF(
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
            ' 调用通用API函数，传入cyq_perf API名称和所有参数
            Return TUSHARE_API("cyq_perf", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 每日筹码分布
    ''' 获取A股每日的筹码分布情况，提供各价位占比，数据从2018年开始，每天17~18点之间更新当日数据
    ''' </summary>
    <ExcelFunction(Description:="每日筹码分布", Category:="Tushare")>
    Public Function TUSHARE_CYQ_CHIPS(
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
            ' 调用通用API函数，传入cyq_chips API名称和所有参数
            Return TUSHARE_API("cyq_chips", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 中央结算系统持股统计
    ''' 获取中央结算系统持股汇总数据，覆盖全部历史数据，根据交易所披露时间，当日数据在下一交易日早上9点前完成入库
    ''' </summary>
    <ExcelFunction(Description:="中央结算系统持股统计", Category:="Tushare")>
    Public Function TUSHARE_CCASS_HOLD(
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
            ' 调用通用API函数，传入ccass_hold API名称和所有参数
            Return TUSHARE_API("ccass_hold", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 股票技术面因子
    ''' 获取股票每日技术面因子数据，用于跟踪股票当前走势情况，数据由Tushare社区自产，覆盖全历史
    ''' </summary>
    <ExcelFunction(Description:="股票技术面因子", Category:="Tushare")>
    Public Function TUSHARE_STK_FACTOR(
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
            ' 调用通用API函数，传入stk_factor API名称和所有参数
            Return TUSHARE_API("stk_factor", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 涨跌停和炸板数据
    ''' 获取A股每日涨跌停、炸板数据情况，数据从2020年开始（不提供ST股票的统计）
    ''' </summary>
    <ExcelFunction(Description:="涨跌停和炸板数据", Category:="Tushare")>
    Public Function TUSHARE_LIMIT_LIST_D(
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
            ' 调用通用API函数，传入limit_list_d API名称和所有参数
            Return TUSHARE_API("limit_list_d", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 动能因子
    ''' 获取小佩数据动量因子数据，可以获取股票动能评级数据，包括最新及过去历史数据
    ''' </summary>
    <ExcelFunction(Description:="动能因子", Category:="Tushare")>
    Public Function TUSHARE_STOCK_MX(
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
            ' 调用通用API函数，传入stock_mx API名称和所有参数
            Return TUSHARE_API("stock_mx", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 估值因子
    ''' 小沛估值因子
    ''' </summary>
    <ExcelFunction(Description:="估值因子", Category:="Tushare")>
    Public Function TUSHARE_STOCK_VX(
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
            ' 调用通用API函数，传入stock_vx API名称和所有参数
            Return TUSHARE_API("stock_vx", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 港股分钟行情
    ''' 港股分钟数据，支持1min/5min/15min/30min/60min行情，提供Python SDK和 http Restful API两种方式
    ''' </summary>
    <ExcelFunction(Description:="港股分钟行情", Category:="Tushare")>
    Public Function TUSHARE_HK_MINS(
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
            ' 调用通用API函数，传入hk_mins API名称和所有参数
            Return TUSHARE_API("hk_mins", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 可转债票面利率
    ''' 获取可转债票面利率
    ''' </summary>
    <ExcelFunction(Description:="可转债票面利率", Category:="Tushare")>
    Public Function TUSHARE_CB_RATE(
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
            ' 调用通用API函数，传入cb_rate API名称和所有参数
            Return TUSHARE_API("cb_rate", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 中信行业指数日行情
    ''' 获取中信行业指数日线行情
    ''' </summary>
    <ExcelFunction(Description:="中信行业指数日行情", Category:="Tushare")>
    Public Function TUSHARE_CI_DAILY(
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
            ' 调用通用API函数，传入ci_daily API名称和所有参数
            Return TUSHARE_API("ci_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 社融增量（月度）
    ''' 获取月度社会融资数据
    ''' </summary>
    <ExcelFunction(Description:="社融增量（月度）", Category:="Tushare")>
    Public Function TUSHARE_SF_MONTH(
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
            ' 调用通用API函数，传入sf_month API名称和所有参数
            Return TUSHARE_API("sf_month", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 市场游资最全名录
    ''' 获取游资分类名录信息
    ''' </summary>
    <ExcelFunction(Description:="市场游资最全名录", Category:="Tushare")>
    Public Function TUSHARE_HM_LIST(
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
            ' 调用通用API函数，传入hm_list API名称和所有参数
            Return TUSHARE_API("hm_list", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 游资交易每日明细
    ''' 获取每日游资交易明细，数据开始于2022年8。游资分类名录，请点击<a href=""https://tushare.pro/document/2?doc_id=311"">游资名录</a>
    ''' </summary>
    <ExcelFunction(Description:="游资交易每日明细", Category:="Tushare")>
    Public Function TUSHARE_HM_DETAIL(
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
            ' 调用通用API函数，传入hm_detail API名称和所有参数
            Return TUSHARE_API("hm_detail", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 历史分钟行情
    ''' 获取全市场期货合约分钟数据，支持1min/5min/15min/30min/60min行情，提供Python SDK和 http Restful API两种方式，如果需要主力合约分钟，请先通过主力[mapping](https://tushare.pro/document/2?doc_id=189)接口获取对应的合约代码后提取分钟。
    ''' </summary>
    <ExcelFunction(Description:="历史分钟行情", Category:="Tushare")>
    Public Function TUSHARE_FT_MINS(
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
            ' 调用通用API函数，传入ft_mins API名称和所有参数
            Return TUSHARE_API("ft_mins", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 实时行情（爬虫）
    ''' 本接口是tushare org版实时接口的顺延，数据来自网络，且不进入tushare服务器，属于爬虫接口，请将tushare升级到1.3.3版本以上。
    ''' </summary>
    <ExcelFunction(Description:="实时行情（爬虫）", Category:="Tushare")>
    Public Function TUSHARE_REALTIME_QUOTE(
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
            ' 调用通用API函数，传入realtime_quote API名称和所有参数
            Return TUSHARE_API("realtime_quote", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 实时成交（爬虫）
    ''' 本接口是tushare org版实时接口的顺延，数据来自网络，且不进入tushare服务器，属于爬虫接口，数据包括该股票当日开盘以来的所有分笔成交数据。
    ''' </summary>
    <ExcelFunction(Description:="实时成交（爬虫）", Category:="Tushare")>
    Public Function TUSHARE_REALTIME_TICK(
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
            ' 调用通用API函数，传入realtime_tick API名称和所有参数
            Return TUSHARE_API("realtime_tick", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 实时排名（爬虫）
    ''' 本接口是tushare org版实时接口的顺延，数据来自网络，且不进入tushare服务器，属于爬虫接口，数据包括该股票当日开盘以来的所有分笔成交数据。
    ''' </summary>
    <ExcelFunction(Description:="实时排名（爬虫）", Category:="Tushare")>
    Public Function TUSHARE_REALTIME_LIST(
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
            ' 调用通用API函数，传入realtime_list API名称和所有参数
            Return TUSHARE_API("realtime_list", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 同花顺App热榜数
    ''' 获取同花顺App热榜数据，包括热股、概念板块、ETF、可转债、港美股等等，每日盘中提取4次，收盘后4次，最晚22点提取一次。
    ''' </summary>
    <ExcelFunction(Description:="同花顺App热榜数", Category:="Tushare")>
    Public Function TUSHARE_THS_HOT(
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
            ' 调用通用API函数，传入ths_hot API名称和所有参数
            Return TUSHARE_API("ths_hot", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 东方财富App热榜
    ''' 获取东方财富App热榜数据，包括A股市场、ETF基金、港股市场、美股市场等等，每日盘中提取4次，收盘后4次，最晚22点提取一次。
    ''' </summary>
    <ExcelFunction(Description:="东方财富App热榜", Category:="Tushare")>
    Public Function TUSHARE_DC_HOT(
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
            ' 调用通用API函数，传入dc_hot API名称和所有参数
            Return TUSHARE_API("dc_hot", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 柜台流通式债券报价
    ''' 柜台流通式债券报价
    ''' </summary>
    <ExcelFunction(Description:="柜台流通式债券报价", Category:="Tushare")>
    Public Function TUSHARE_BC_OTCQT(
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
            ' 调用通用API函数，传入bc_otcqt API名称和所有参数
            Return TUSHARE_API("bc_otcqt", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 柜台流通式债券最优报价
    ''' 柜台流通式债券最优报价
    ''' </summary>
    <ExcelFunction(Description:="柜台流通式债券最优报价", Category:="Tushare")>
    Public Function TUSHARE_BC_BESTOTCQT(
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
            ' 调用通用API函数，传入bc_bestotcqt API名称和所有参数
            Return TUSHARE_API("bc_bestotcqt", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 采购经理指数（PMI）
    ''' 采购经理人指数
    ''' </summary>
    <ExcelFunction(Description:="采购经理指数（PMI）", Category:="Tushare")>
    Public Function TUSHARE_CN_PMI(
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
            ' 调用通用API函数，传入cn_pmi API名称和所有参数
            Return TUSHARE_API("cn_pmi", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 融资融券标的（盘前）
    ''' 获取沪深京三大交易所融资融券标的（包括ETF），每天盘前更新
    ''' </summary>
    <ExcelFunction(Description:="融资融券标的（盘前）", Category:="Tushare")>
    Public Function TUSHARE_MARGIN_SECS(
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
            ' 调用通用API函数，传入margin_secs API名称和所有参数
            Return TUSHARE_API("margin_secs", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 申万行业指数日行情
    ''' 获取申万行业日线行情（默认是申万2021版行情）
    ''' </summary>
    <ExcelFunction(Description:="申万行业指数日行情", Category:="Tushare")>
    Public Function TUSHARE_SW_DAILY(
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
            ' 调用通用API函数，传入sw_daily API名称和所有参数
            Return TUSHARE_API("sw_daily", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 股票技术面因子(专业版）
    ''' 获取股票每日技术面因子数据，用于跟踪股票当前走势情况，数据由Tushare社区自产，覆盖全历史；输出参数_bfq表示不复权，_qfq表示前复权 _hfq表示后复权，描述中说明了因子的默认传参，如需要特殊参数或者更多因子可以联系管理员评估
    ''' </summary>
    <ExcelFunction(Description:="股票技术面因子(专业版）", Category:="Tushare")>
    Public Function TUSHARE_STK_FACTOR_PRO(
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
            ' 调用通用API函数，传入stk_factor_pro API名称和所有参数
            Return TUSHARE_API("stk_factor_pro", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 每日股本（盘前）
    ''' 每日开盘前获取当日股票的股本情况，包括总股本和流通股本，涨跌停价格等。
    ''' </summary>
    <ExcelFunction(Description:="每日股本（盘前）", Category:="Tushare")>
    Public Function TUSHARE_STK_PREMARKET(
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
            ' 调用通用API函数，传入stk_premarket API名称和所有参数
            Return TUSHARE_API("stk_premarket", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 转融资交易汇总
    ''' 转融通融资汇总
    ''' </summary>
    <ExcelFunction(Description:="转融资交易汇总", Category:="Tushare")>
    Public Function TUSHARE_SLB_LEN(
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
            ' 调用通用API函数，传入slb_len API名称和所有参数
            Return TUSHARE_API("slb_len", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 转融券交易汇总
    ''' 转融通转融券交易汇总
    ''' </summary>
    <ExcelFunction(Description:="转融券交易汇总", Category:="Tushare")>
    Public Function TUSHARE_SLB_SEC(
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
            ' 调用通用API函数，传入slb_sec API名称和所有参数
            Return TUSHARE_API("slb_sec", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 转融券交易明细
    ''' 转融券交易明细
    ''' </summary>
    <ExcelFunction(Description:="转融券交易明细", Category:="Tushare")>
    Public Function TUSHARE_SLB_SEC_DETAIL(
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
            ' 调用通用API函数，传入slb_sec_detail API名称和所有参数
            Return TUSHARE_API("slb_sec_detail", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 做市借券交易汇总
    ''' 做市借券交易汇总
    ''' </summary>
    <ExcelFunction(Description:="做市借券交易汇总", Category:="Tushare")>
    Public Function TUSHARE_SLB_LEN_MM(
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
            ' 调用通用API函数，传入slb_len_mm API名称和所有参数
            Return TUSHARE_API("slb_len_mm", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 申万行业成分（分级）
    ''' 按三级分类提取申万行业成分，可提供某个分类的所有成分，也可按股票代码提取所属分类，参数灵活
    ''' </summary>
    <ExcelFunction(Description:="申万行业成分（分级）", Category:="Tushare")>
    Public Function TUSHARE_INDEX_MEMBER_ALL(
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
            ' 调用通用API函数，传入index_member_all API名称和所有参数
            Return TUSHARE_API("index_member_all", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 周/月线行情(每日更新)
    ''' 股票周/月线行情(每日更新)
    ''' </summary>
    <ExcelFunction(Description:="周/月线行情(每日更新)", Category:="Tushare")>
    Public Function TUSHARE_STK_WEEKLY_MONTHLY(
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
            ' 调用通用API函数，传入stk_weekly_monthly API名称和所有参数
            Return TUSHARE_API("stk_weekly_monthly", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 期货周/月线行情(每日更新)
    ''' 期货周/月线行情(每日更新)
    ''' </summary>
    <ExcelFunction(Description:="期货周/月线行情(每日更新)", Category:="Tushare")>
    Public Function TUSHARE_FUT_WEEKLY_MONTHLY(
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
            ' 调用通用API函数，传入fut_weekly_monthly API名称和所有参数
            Return TUSHARE_API("fut_weekly_monthly", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 美股复权行情
    ''' 获取美股复权行情，支持美股全市场股票，提供股本、市值、复权因子和成交信息等多个数据指标
    ''' </summary>
    <ExcelFunction(Description:="美股复权行情", Category:="Tushare")>
    Public Function TUSHARE_US_DAILY_ADJ(
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
            ' 调用通用API函数，传入us_daily_adj API名称和所有参数
            Return TUSHARE_API("us_daily_adj", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 港股复权行情
    ''' 获取港股复权行情，提供股票股本、市值和成交及换手多个数据指标
    ''' </summary>
    <ExcelFunction(Description:="港股复权行情", Category:="Tushare")>
    Public Function TUSHARE_HK_DAILY_ADJ(
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
            ' 调用通用API函数，传入hk_daily_adj API名称和所有参数
            Return TUSHARE_API("hk_daily_adj", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 实时分钟行情
    ''' 获取全市场期货合约实时分钟数据，支持1min/5min/15min/30min/60min行情，提供Python SDK、 http Restful API和websocket三种方式，如果需要主力合约分钟，请先通过主力[mapping](https://tushare.pro/document/2?doc_id=189)接口获取对应的合约代码后提取分钟。
    ''' </summary>
    <ExcelFunction(Description:="实时分钟行情", Category:="Tushare")>
    Public Function TUSHARE_RT_FUT_MIN(
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
            ' 调用通用API函数，传入rt_fut_min API名称和所有参数
            Return TUSHARE_API("rt_fut_min", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 期权分钟行情
    ''' 获取全市场期权合约分钟数据，支持1min/5min/15min/30min/60min行情，提供Python SDK和 http Restful API两种方式。
    ''' </summary>
    <ExcelFunction(Description:="期权分钟行情", Category:="Tushare")>
    Public Function TUSHARE_OPT_MINS(
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
            ' 调用通用API函数，传入opt_mins API名称和所有参数
            Return TUSHARE_API("opt_mins", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 行业资金流向（THS）
    ''' 获取同花顺行业板块资金流向，每日盘后更新
    ''' </summary>
    <ExcelFunction(Description:="行业资金流向（THS）", Category:="Tushare")>
    Public Function TUSHARE_MONEYFLOW_IND_THS(
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
            ' 调用通用API函数，传入moneyflow_ind_ths API名称和所有参数
            Return TUSHARE_API("moneyflow_ind_ths", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 板块资金流向（DC）
    ''' 获取东方财富板块资金流向，每天盘后更新
    ''' </summary>
    <ExcelFunction(Description:="板块资金流向（DC）", Category:="Tushare")>
    Public Function TUSHARE_MONEYFLOW_IND_DC(
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
            ' 调用通用API函数，传入moneyflow_ind_dc API名称和所有参数
            Return TUSHARE_API("moneyflow_ind_dc", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 大盘资金流向（DC）
    ''' 获取东方财富大盘资金流向数据，每日盘后更新
    ''' </summary>
    <ExcelFunction(Description:="大盘资金流向（DC）", Category:="Tushare")>
    Public Function TUSHARE_MONEYFLOW_MKT_DC(
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
            ' 调用通用API函数，传入moneyflow_mkt_dc API名称和所有参数
            Return TUSHARE_API("moneyflow_mkt_dc", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 榜单数据（开盘啦）
    ''' 获取开盘啦涨停、跌停、炸板等榜单数据
    ''' </summary>
    <ExcelFunction(Description:="榜单数据（开盘啦）", Category:="Tushare")>
    Public Function TUSHARE_KPL_LIST(
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
            ' 调用通用API函数，传入kpl_list API名称和所有参数
            Return TUSHARE_API("kpl_list", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 个股资金流向（THS）
    ''' 获取同花顺个股资金流向数据，每日盘后更新
    ''' </summary>
    <ExcelFunction(Description:="个股资金流向（THS）", Category:="Tushare")>
    Public Function TUSHARE_MONEYFLOW_THS(
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
            ' 调用通用API函数，传入moneyflow_ths API名称和所有参数
            Return TUSHARE_API("moneyflow_ths", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 个股资金流向（DC）
    ''' 获取东方财富个股资金流向数据，每日盘后更新，数据开始于20230911
    ''' </summary>
    <ExcelFunction(Description:="个股资金流向（DC）", Category:="Tushare")>
    Public Function TUSHARE_MONEYFLOW_DC(
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
            ' 调用通用API函数，传入moneyflow_dc API名称和所有参数
            Return TUSHARE_API("moneyflow_dc", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 题材数据（开盘啦）
    ''' 获取开盘啦概念题材列表，每天盘后更新
    ''' </summary>
    <ExcelFunction(Description:="题材数据（开盘啦）", Category:="Tushare")>
    Public Function TUSHARE_KPL_CONCEPT(
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
            ' 调用通用API函数，传入kpl_concept API名称和所有参数
            Return TUSHARE_API("kpl_concept", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 题材成分（开盘啦）
    ''' 获取开盘啦概念题材的成分股
    ''' </summary>
    <ExcelFunction(Description:="题材成分（开盘啦）", Category:="Tushare")>
    Public Function TUSHARE_KPL_CONCEPT_CONS(
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
            ' 调用通用API函数，传入kpl_concept_cons API名称和所有参数
            Return TUSHARE_API("kpl_concept_cons", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 股票开盘集合竞价数据
    ''' 股票开盘9:30集合竞价数据，每天盘后更新
    ''' </summary>
    <ExcelFunction(Description:="股票开盘集合竞价数据", Category:="Tushare")>
    Public Function TUSHARE_STK_AUCTION_O(
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
            ' 调用通用API函数，传入stk_auction_o API名称和所有参数
            Return TUSHARE_API("stk_auction_o", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 股票收盘集合竞价数据
    ''' 股票收盘15:00集合竞价数据，每天盘后更新
    ''' </summary>
    <ExcelFunction(Description:="股票收盘集合竞价数据", Category:="Tushare")>
    Public Function TUSHARE_STK_AUCTION_C(
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
            ' 调用通用API函数，传入stk_auction_c API名称和所有参数
            Return TUSHARE_API("stk_auction_c", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 同花顺涨跌停榜单
    ''' 获取同花顺每日涨跌停榜单数据，历史数据从20231101开始提供，增量每天16点左右更新
    ''' </summary>
    <ExcelFunction(Description:="同花顺涨跌停榜单", Category:="Tushare")>
    Public Function TUSHARE_LIMIT_LIST_THS(
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
            ' 调用通用API函数，传入limit_list_ths API名称和所有参数
            Return TUSHARE_API("limit_list_ths", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 涨停股票连板天梯
    ''' 获取每天连板个数晋级的股票，可以分析出每天连续涨停进阶个数，判断强势热度
    ''' </summary>
    <ExcelFunction(Description:="涨停股票连板天梯", Category:="Tushare")>
    Public Function TUSHARE_LIMIT_STEP(
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
            ' 调用通用API函数，传入limit_step API名称和所有参数
            Return TUSHARE_API("limit_step", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 涨停最强板块统计
    ''' 获取每天涨停股票最多最强的概念板块，可以分析强势板块的轮动，判断资金动向
    ''' </summary>
    <ExcelFunction(Description:="涨停最强板块统计", Category:="Tushare")>
    Public Function TUSHARE_LIMIT_CPT_LIST(
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
            ' 调用通用API函数，传入limit_cpt_list API名称和所有参数
            Return TUSHARE_API("limit_cpt_list", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 指数技术面因子(专业版)
    ''' 获取指数每日技术面因子数据，用于跟踪指数当前走势情况，数据由Tushare社区自产，覆盖全历史；输出参数_bfq表示不复权描述中说明了因子的默认传参，如需要特殊参数或者更多因子可以联系管理员评估，指数包括大盘指数 申万行业指数 中信指数
    ''' </summary>
    <ExcelFunction(Description:="指数技术面因子(专业版)", Category:="Tushare")>
    Public Function TUSHARE_IDX_FACTOR_PRO(
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
            ' 调用通用API函数，传入idx_factor_pro API名称和所有参数
            Return TUSHARE_API("idx_factor_pro", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 基金技术面因子(专业版)
    ''' 获取场内基金每日技术面因子数据，用于跟踪场内基金当前走势情况，数据由Tushare社区自产，覆盖全历史；输出参数_bfq表示不复权，描述中说明了因子的默认传参，如需要特殊参数或者更多因子可以联系管理员评估
    ''' </summary>
    <ExcelFunction(Description:="基金技术面因子(专业版)", Category:="Tushare")>
    Public Function TUSHARE_FUND_FACTOR_PRO(
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
            ' 调用通用API函数，传入fund_factor_pro API名称和所有参数
            Return TUSHARE_API("fund_factor_pro", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 东方财富概念板块
    ''' 获取东方财富每个交易日的概念板块数据，支持按日期查询
    ''' </summary>
    <ExcelFunction(Description:="东方财富概念板块", Category:="Tushare")>
    Public Function TUSHARE_DC_INDEX(
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
            ' 调用通用API函数，传入dc_index API名称和所有参数
            Return TUSHARE_API("dc_index", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 东方财富概念成分
    ''' 获取东方财富板块每日成分数据，可以根据概念板块代码和交易日期，获取历史成分
    ''' </summary>
    <ExcelFunction(Description:="东方财富概念成分", Category:="Tushare")>
    Public Function TUSHARE_DC_MEMBER(
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
            ' 调用通用API函数，传入dc_member API名称和所有参数
            Return TUSHARE_API("dc_member", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 神奇九转指标
    ''' 神奇九转（又称“九转序列”）是一种基于技术分析的股票趋势反转指标，其思想来源于技术分析大师汤姆·迪马克（Tom DeMark）的TD序列。该指标的核心功能是通过识别股价在上涨或下跌过程中连续9天的特定走势，来判断股价的潜在反转点，从而帮助投资者提高抄底和逃顶的成功率，日线级别配合60min的九转效果更好，数据从20230101开始。
    ''' </summary>
    <ExcelFunction(Description:="神奇九转指标", Category:="Tushare")>
    Public Function TUSHARE_STK_NINETURN(
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
            ' 调用通用API函数，传入stk_nineturn API名称和所有参数
            Return TUSHARE_API("stk_nineturn", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 周/月线复权行情(每日更新)
    ''' 股票周/月线行情(复权--每日更新)
    ''' </summary>
    <ExcelFunction(Description:="周/月线复权行情(每日更新)", Category:="Tushare")>
    Public Function TUSHARE_STK_WEEK_MONTH_ADJ(
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
            ' 调用通用API函数，传入stk_week_month_adj API名称和所有参数
            Return TUSHARE_API("stk_week_month_adj", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 期货合约涨跌停价格
    ''' 获取所有期货合约每天的涨跌停价格及最低保证金率，数据开始于2005年。
    ''' </summary>
    <ExcelFunction(Description:="期货合约涨跌停价格", Category:="Tushare")>
    Public Function TUSHARE_FT_LIMIT(
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
            ' 调用通用API函数，传入ft_limit API名称和所有参数
            Return TUSHARE_API("ft_limit", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 开盘竞价成交（当日）
    ''' 获取当日个股和ETF的集合竞价成交情况，每天9点25后可以获取当日的集合竞价成交数据
    ''' </summary>
    <ExcelFunction(Description:="开盘竞价成交（当日）", Category:="Tushare")>
    Public Function TUSHARE_STK_AUCTION(
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
            ' 调用通用API函数，传入stk_auction API名称和所有参数
            Return TUSHARE_API("stk_auction", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 分钟行情
    ''' 获取A股分钟数据，支持1min/5min/15min/30min/60min行情，提供Python SDK和 http Restful API两种方式
    ''' </summary>
    <ExcelFunction(Description:="分钟行情", Category:="Tushare")>
    Public Function TUSHARE_STK_MINS(
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
            ' 调用通用API函数，传入stk_mins API名称和所有参数
            Return TUSHARE_API("stk_mins", param1Name, param1Value, param2Name, param2Value,
                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,
                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,
                              param9Name, param9Value, param10Name, param10Value)
        Catch ex As Exception
            ' 异常处理
            Return "错误: " & ex.Message
        End Try
    End Function
End Module