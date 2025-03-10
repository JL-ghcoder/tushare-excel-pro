Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Drawing

Public Class AboutPlugin
    Private Sub AboutPlugin_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' 设置RichTextBox的内容
        rtbPluginInfo.Clear()

        ' 为RichTextBox设置富文本内容
        rtbPluginInfo.SelectionFont = New Font("Microsoft YaHei UI", 12, FontStyle.Bold)
        rtbPluginInfo.AppendText("Tushare Excel Pro 专业版" & vbCrLf)

        rtbPluginInfo.SelectionFont = New Font("Microsoft YaHei UI", 9, FontStyle.Regular)
        rtbPluginInfo.AppendText("版本: v1.0.0" & vbCrLf)
        rtbPluginInfo.AppendText("发布日期: 2025-03-12" & vbCrLf)

        rtbPluginInfo.SelectionFont = New Font("Microsoft YaHei UI", 10, FontStyle.Bold)
        rtbPluginInfo.AppendText("作者信息:" & vbCrLf)

        rtbPluginInfo.SelectionFont = New Font("Microsoft YaHei UI", 9, FontStyle.Regular)
        rtbPluginInfo.AppendText("姓名: Peng" & vbCrLf)
        rtbPluginInfo.AppendText("邮箱: ispierce.zhou@gmail.com" & vbCrLf)

        rtbPluginInfo.SelectionFont = New Font("Microsoft YaHei UI", 9, FontStyle.Regular)
        rtbPluginInfo.AppendText("姓名: Jun" & vbCrLf)
        rtbPluginInfo.AppendText("邮箱: isjun.liu@gmail.com" & vbCrLf)

        rtbPluginInfo.SelectionFont = New Font("Microsoft YaHei UI", 10, FontStyle.Bold)
        rtbPluginInfo.AppendText("功能介绍:" & vbCrLf)

        rtbPluginInfo.SelectionFont = New Font("Microsoft YaHei UI", 9, FontStyle.Regular)
        rtbPluginInfo.AppendText("• 支持完整的tushare官方接口" & vbCrLf)
        rtbPluginInfo.AppendText("• 便捷的侧边栏公式查询" & vbCrLf)
        rtbPluginInfo.AppendText("• 内置ChatGPT和DeepSeek" & vbCrLf)
        rtbPluginInfo.AppendText("• TuShare接口作为AI知识库" & vbCrLf)
        rtbPluginInfo.AppendText("• 通过AI自动运行VBA代码和Excel公式" & vbCrLf & vbCrLf)

        rtbPluginInfo.SelectionFont = New Font("Microsoft YaHei UI", 10, FontStyle.Bold)
        rtbPluginInfo.AppendText("开源信息:" & vbCrLf)

        rtbPluginInfo.SelectionFont = New Font("Microsoft YaHei UI", 9, FontStyle.Regular)
        rtbPluginInfo.AppendText("本项目遵循MIT许可证开源" & vbCrLf)
        rtbPluginInfo.AppendText("开源地址: https://github.com/JL-ghcoder/tushare-excel-pro" & vbCrLf & vbCrLf)

        rtbPluginInfo.SelectionFont = New Font("Microsoft YaHei UI", 9, FontStyle.Italic)
        rtbPluginInfo.AppendText("附注：更多问题可以通过github或者邮件联系到作者。")

        ' 将光标移动到开头
        rtbPluginInfo.SelectionStart = 0
        rtbPluginInfo.SelectionLength = 0
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        ' 点击确定按钮关闭窗体
        Me.Close()
    End Sub

    ' 使RichTextBox中的链接可点击（可选功能）
    Private Sub rtbPluginInfo_LinkClicked(sender As Object, e As LinkClickedEventArgs) Handles rtbPluginInfo.LinkClicked
        ' 打开链接
        Try
            Process.Start(e.LinkText)
        Catch ex As Exception
            MessageBox.Show("无法打开链接: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class