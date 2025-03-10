Imports System.Windows.Forms
Imports Microsoft.Win32

Public Class ConfigFormDS
    ' 用于存储和访问设置的属性
    Public Property DSApiUrl As String
    Public Property DSToken As String

    ' 注册表路径常量
    Private Const REG_PATH As String = "Software\TushareExcel"

    ' 窗体加载时初始化
    Private Sub ConfigForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 设置默认值
        txtApiUrl.Text = "https://api.deepseek.com/chat/completions"

        ' 从注册表读取配置
        Try
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(REG_PATH)
                If key IsNot Nothing Then
                    Dim savedUrl = key.GetValue("DSApiUrl")
                    Dim savedToken = key.GetValue("DSToken")

                    If savedUrl IsNot Nothing Then
                        txtApiUrl.Text = savedUrl.ToString()
                    End If

                    If savedToken IsNot Nothing Then
                        txtToken.Text = savedToken.ToString()
                    End If
                End If
            End Using
        Catch ex As Exception
            ' 读取配置时出错，使用默认值
        End Try
    End Sub

    ' 保存按钮点击事件
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' 验证输入
        If txtApiUrl.Text.Trim() = "" Then
            MessageBox.Show("请输入API URL", "验证错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtApiUrl.Focus()
            Return
        End If

        If txtToken.Text.Trim() = "" Then
            MessageBox.Show("请输入API Token", "验证错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtToken.Focus()
            Return
        End If

        ' 保存值到属性
        DSApiUrl = txtApiUrl.Text.Trim()
        DSToken = txtToken.Text.Trim()

        ' 保存到注册表
        Try
            Using key As RegistryKey = Registry.CurrentUser.CreateSubKey(REG_PATH)
                key.SetValue("DSApiUrl", DSApiUrl)
                key.SetValue("DSToken", DSToken)
            End Using

            MessageBox.Show("设置已保存", "保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("保存设置时出错: " & ex.Message, "保存错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' 关闭窗体
        DialogResult = DialogResult.OK
        Close()
    End Sub

    ' 取消按钮点击事件
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        DialogResult = DialogResult.Cancel
        Close()
    End Sub
End Class