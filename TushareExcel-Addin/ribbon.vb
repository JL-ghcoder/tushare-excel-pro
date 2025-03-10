Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports Microsoft.Office.Tools.Ribbon
Public Class ribbon
    Private Sub ribbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
    End Sub

    Private Sub settings_Click(sender As Object, e As RibbonControlEventArgs) Handles settings.Click
        Try
            ' 创建并显示配置窗口
            Dim configForm As New ConfigForm()
            ' 显示对话框
            If configForm.ShowDialog() = DialogResult.OK Then
                ' 配置已保存，可以更新应用程序状态
                MessageBox.Show("Tushare API设置已更新", "设置", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("显示设置窗口时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub btnDeepseek_Click(sender As Object, e As RibbonControlEventArgs) Handles btnDeepseek.Click
        Try
            ' 创建并显示配置窗口
            Dim configForm As New ConfigFormDS()
            ' 显示对话框
            If configForm.ShowDialog() = DialogResult.OK Then
                ' 配置已保存，可以更新应用程序状态
                MessageBox.Show("DeepSeek API设置已更新", "设置", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("显示设置窗口时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnGPT_Click(sender As Object, e As RibbonControlEventArgs) Handles btnGPT.Click
        Try
            ' 创建并显示配置窗口
            Dim configForm As New ConfigFormGPT()
            ' 显示对话框
            If configForm.ShowDialog() = DialogResult.OK Then
                ' 配置已保存，可以更新应用程序状态
                MessageBox.Show("ChatGPT API设置已更新", "设置", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("显示设置窗口时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub docs_Click(sender As Object, e As RibbonControlEventArgs) Handles docs.Click
        Try
            ' 切换函数向导任务窗格的可见性
            Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = Globals.ThisAddIn.FunctionGuidePane
            If taskPane IsNot Nothing Then
                taskPane.Visible = Not taskPane.Visible
                ' 可选：如果窗格变为可见，可以设置焦点
                If taskPane.Visible Then
                    ' 将焦点设置到窗格上 (如果需要)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("显示函数向导时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 共用的插入文本方法
    Private Sub InsertTextUsingClipboard(text As String)
        Try
            ' 将文本复制到剪贴板
            Clipboard.SetText(text)
            ' 发送粘贴命令
            System.Windows.Forms.SendKeys.SendWait("^v")
        Catch ex As Exception
            MessageBox.Show("使用剪贴板插入文本时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDaily_Click(sender As Object, e As RibbonControlEventArgs) Handles btnDaily.Click
        ' MsgBox("按钮成功")
        InsertTextUsingClipboard("TUSHARE_DAILY(""ts_code"", """", ""start_date"", """", ""end_date"", """")")
    End Sub

    Private Sub btnMins_Click(sender As Object, e As RibbonControlEventArgs) Handles btnMins.Click
        InsertTextUsingClipboard("TUSHARE_STK_MINS(""ts_code"", """", ""freq"", """", ""start_date"", """", ""end_date"", """")")
    End Sub

    Private Sub btnWeekly_Click(sender As Object, e As RibbonControlEventArgs) Handles btnWeekly.Click
        InsertTextUsingClipboard("TUSHARE_WEEKLY(""ts_code"", """", ""start_date"", """", ""end_date"", """")")
    End Sub

    Private Sub btnMonthly_Click(sender As Object, e As RibbonControlEventArgs) Handles btnMonthly.Click
        InsertTextUsingClipboard("TUSHARE_MONTHLY(""ts_code"", """", ""start_date"", """", ""end_date"", """")")
    End Sub

    Private Sub btnDailyBasic_Click(sender As Object, e As RibbonControlEventArgs) Handles btnDailyBasic.Click
        InsertTextUsingClipboard("TUSHARE_DAILY_BASIC(""ts_code"", """", ""start_date"", """", ""end_date"", """")")
    End Sub

    Private Sub btnIncome_Click(sender As Object, e As RibbonControlEventArgs) Handles btnIncome.Click
        InsertTextUsingClipboard("TUSHARE_INCOME(""ts_code"", """", ""start_date"", """", ""end_date"", """")")
    End Sub

    Private Sub btnBalancesheet_Click(sender As Object, e As RibbonControlEventArgs) Handles btnBalancesheet.Click
        InsertTextUsingClipboard("TUSHARE_BALANCESHEET(""ts_code"", """", ""start_date"", """", ""end_date"", """")")
    End Sub

    Private Sub btnCashflow_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCashflow.Click
        InsertTextUsingClipboard("TUSHARE_CASHFLOW(""ts_code"", """", ""start_date"", """", ""end_date"", """")")
    End Sub

    Private Sub btnFinaIndicator_Click(sender As Object, e As RibbonControlEventArgs) Handles btnFinaIndicator.Click
        InsertTextUsingClipboard("TUSHARE_FINAL_INDICATOR(""ts_code"", """", ""start_date"", """", ""end_date"", """")")
    End Sub

    Private Sub btnIndexBasic_Click(sender As Object, e As RibbonControlEventArgs) Handles btnIndexBasic.Click
        InsertTextUsingClipboard("TUSHARE_INDEX_BASIC(""ts_code"", """")")
    End Sub

    Private Sub btnIndexDaily_Click(sender As Object, e As RibbonControlEventArgs) Handles btnIndexDaily.Click
        InsertTextUsingClipboard("TUSHARE_INDEX_DAILY(""ts_code"", """", ""start_date"", """", ""end_date"", """")")
    End Sub

    Private Sub btnIndexWeekly_Click(sender As Object, e As RibbonControlEventArgs) Handles btnIndexWeekly.Click
        InsertTextUsingClipboard("TUSHARE_INDEX_WEEKLY(""ts_code"", """", ""start_date"", """", ""end_date"", """")")
    End Sub

    Private Sub btnIndexMonthly_Click(sender As Object, e As RibbonControlEventArgs) Handles btnIndexMonthly.Click
        InsertTextUsingClipboard("TUSHARE_INDEX_MONTHLY(""ts_code"", """", ""start_date"", """", ""end_date"", """")")
    End Sub

    Private Sub btnIndexWeight_Click(sender As Object, e As RibbonControlEventArgs) Handles btnIndexWeight.Click
        InsertTextUsingClipboard("TUSHARE_INDEX_WEIGHT(""index_code"", """", ""start_date"", """", ""end_date"", """")")
    End Sub

    Private Sub btnIndexGlobal_Click(sender As Object, e As RibbonControlEventArgs) Handles btnIndexGlobal.Click
        InsertTextUsingClipboard("TUSHARE_INDEX_GLOBAL(""ts_code"", """", ""start_date"", """", ""end_date"", """")")
    End Sub

    Private Sub btnAI_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAI.Click
        Try
            ' 切换AI助手窗格的可见性
            If Globals.ThisAddIn.AIChatPane IsNot Nothing Then
                Globals.ThisAddIn.AIChatPane.Visible = Not Globals.ThisAddIn.AIChatPane.Visible
            End If
        Catch ex As Exception
            MessageBox.Show("显示AI助手时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub intro_Click(sender As Object, e As RibbonControlEventArgs) Handles intro.Click
        ' 创建并显示AboutPlugin窗体
        Dim aboutForm As New AboutPlugin()
        aboutForm.ShowDialog()
    End Sub

    Private Sub btnReload_Click(sender As Object, e As RibbonControlEventArgs) Handles btnReload.Click
        Try
            ' 获取当前活动工作表
            Dim activeSheet As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

            ' 在状态栏显示操作进行中的消息
            Globals.ThisAddIn.Application.StatusBar = "正在刷新公式..."

            ' 禁用屏幕更新以提高性能
            Globals.ThisAddIn.Application.ScreenUpdating = False

            ' 获取使用区域
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            ' 循环遍历使用区域中的每个单元格
            Dim cell As Excel.Range
            Dim formulaCount As Integer = 0

            For Each cell In usedRange
                ' 检查单元格是否包含公式
                If Not String.IsNullOrEmpty(cell.Formula) AndAlso cell.Formula.StartsWith("=") Then
                    ' 获取原始公式
                    Dim formula As String = cell.Formula

                    ' 临时清除公式然后重新应用
                    cell.Formula = formula

                    formulaCount += 1
                End If
            Next

            ' 恢复屏幕更新
            Globals.ThisAddIn.Application.ScreenUpdating = True

            ' 显示完成消息
            Globals.ThisAddIn.Application.StatusBar = "公式刷新完成！共刷新了 " & formulaCount & " 个公式。"

            ' 3秒后清除状态栏消息
            System.Threading.Thread.Sleep(3000)
            Globals.ThisAddIn.Application.StatusBar = False

        Catch ex As Exception
            ' 恢复屏幕更新
            Globals.ThisAddIn.Application.ScreenUpdating = True

            ' 恢复状态栏
            Globals.ThisAddIn.Application.StatusBar = False

            ' 显示错误消息
            MessageBox.Show("刷新公式时发生错误: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRemove_Click(sender As Object, e As RibbonControlEventArgs) Handles btnRemove.Click
        Try
            Dim excel As Excel.Application = Globals.ThisAddIn.Application

            ' 显示状态消息
            excel.StatusBar = "请选择要转换的公式区域..."

            ' 提示用户选择区域
            Dim selectedRange As Excel.Range = Nothing

            ' 使用Excel内置的InputBox来获取用户选择的区域
            ' 参数8表示需要一个单元格引用
            selectedRange = excel.InputBox("请选择包含公式的区域：", "转换公式为值", Type:=8)

            ' 检查用户是否取消了选择
            If selectedRange Is Nothing Then
                excel.StatusBar = "操作已取消。"
                System.Threading.Thread.Sleep(1500)
                excel.StatusBar = False
                Exit Sub
            End If

            ' 禁用屏幕更新以提高性能
            excel.ScreenUpdating = False

            ' 计算包含公式的单元格数量
            Dim formulaCount As Integer = 0
            Dim cell As Excel.Range

            For Each cell In selectedRange
                If Not String.IsNullOrEmpty(cell.Formula) AndAlso cell.Formula.StartsWith("=") Then
                    formulaCount += 1
                End If
            Next

            If formulaCount = 0 Then
                MessageBox.Show("所选区域中没有包含公式的单元格。", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                excel.ScreenUpdating = True
                excel.StatusBar = False
                Exit Sub
            End If

            ' 直接使用Value = Value将公式转换为值
            selectedRange.Value = selectedRange.Value

            ' 恢复屏幕更新
            excel.ScreenUpdating = True

            ' 显示完成消息
            excel.StatusBar = "已成功将 " & formulaCount & " 个公式转换为值。"

            ' 3秒后清除状态栏消息
            System.Threading.Thread.Sleep(3000)
            excel.StatusBar = False

        Catch ex As Exception
            ' 恢复Excel状态
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.StatusBar = False

            ' 如果用户取消了操作，不显示错误
            If TypeOf ex Is System.Runtime.InteropServices.COMException AndAlso ex.Message.Contains("4390") Then
                Exit Sub
            End If

            ' 显示错误消息
            MessageBox.Show("转换公式为值时发生错误: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnBeautify_Click(sender As Object, e As RibbonControlEventArgs) Handles btnBeautify.Click
        Try
            Dim excel As Excel.Application = Globals.ThisAddIn.Application
            Dim activeSheet As Excel.Worksheet = excel.ActiveSheet

            ' 在状态栏显示操作正在进行中
            excel.StatusBar = "正在美化表格..."

            ' 禁用屏幕更新以提高性能
            excel.ScreenUpdating = False

            ' 获取使用的区域
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            ' 如果使用的区域为空，则退出
            If usedRange Is Nothing OrElse usedRange.Cells.Count = 1 AndAlso String.IsNullOrEmpty(usedRange.Value) Then
                MessageBox.Show("当前工作表没有数据需要美化。", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                excel.ScreenUpdating = True
                excel.StatusBar = False
                Exit Sub
            End If

            ' 统一字体和字号
            usedRange.Font.Name = "微软雅黑"
            usedRange.Font.Size = 10

            ' 将所有单元格文本居中对齐
            usedRange.HorizontalAlignment = -4108 ' xlCenter = -4108
            usedRange.VerticalAlignment = -4108 ' xlCenter = -4108

            ' 自动调整列宽以适应内容
            usedRange.Columns.AutoFit()

            ' 为每列调整适当宽度，确保不超过255的最大限制
            For i As Integer = 1 To usedRange.Columns.Count
                Dim column As Excel.Range = usedRange.Columns(i)
                Dim newWidth As Double = column.ColumnWidth * 1.05 ' 增加5%的宽度

                ' 确保宽度不超过Excel限制(255)
                If newWidth > 250 Then ' 留一些余量
                    newWidth = 250
                End If

                column.ColumnWidth = newWidth
            Next

            ' 为数字列设置数字格式
            For i As Integer = 1 To usedRange.Columns.Count
                Dim column As Excel.Range = usedRange.Columns(i)
                Dim isNumericColumn As Boolean = True

                ' 检查列是否包含数值数据
                For Each cell As Excel.Range In column.Cells
                    If Not IsNothing(cell.Value) AndAlso Not IsNumeric(cell.Value) Then
                        isNumericColumn = False
                        Exit For
                    End If
                Next

                ' 如果是数值列，设置数字格式
                If isNumericColumn Then
                    column.NumberFormat = "#,##0.00" ' 使用千位分隔符和两位小数
                End If
            Next

            ' 调整行高，确保所有行有统一且合适的高度
            usedRange.RowHeight = 20 ' 设置统一的行高

            ' 恢复屏幕更新
            excel.ScreenUpdating = True

            ' 选择A1单元格
            activeSheet.Range("A1").Select()

            ' 显示完成消息
            excel.StatusBar = "表格美化完成！"
            System.Threading.Thread.Sleep(3000)
            excel.StatusBar = False

        Catch ex As Exception
            ' 恢复Excel状态
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.StatusBar = False

            ' 显示错误消息
            MessageBox.Show("美化表格时发生错误: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class