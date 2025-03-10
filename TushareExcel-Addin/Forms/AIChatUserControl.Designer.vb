<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AIChatUserControl
    Inherits System.Windows.Forms.UserControl

    'UserControl 重写释放以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtUserInput = New System.Windows.Forms.TextBox()
        Me.btnSend = New System.Windows.Forms.Button()
        Me.txtMessages = New System.Windows.Forms.RichTextBox()
        Me.chkIncludeSelection = New System.Windows.Forms.CheckBox()
        Me.chkUseGPT = New System.Windows.Forms.CheckBox()
        Me.chkUseDS = New System.Windows.Forms.CheckBox()
        Me.chkUseTushareDocs = New System.Windows.Forms.CheckBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtUserInput
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.txtUserInput, 3)
        Me.txtUserInput.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtUserInput.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.txtUserInput.Location = New System.Drawing.Point(3, 749)
        Me.txtUserInput.Multiline = True
        Me.txtUserInput.Name = "txtUserInput"
        Me.txtUserInput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtUserInput.Size = New System.Drawing.Size(477, 78)
        Me.txtUserInput.TabIndex = 0
        '
        'btnSend
        '
        Me.btnSend.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnSend.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.btnSend.Location = New System.Drawing.Point(486, 749)
        Me.btnSend.Name = "btnSend"
        Me.btnSend.Size = New System.Drawing.Size(158, 78)
        Me.btnSend.TabIndex = 1
        Me.btnSend.Text = "发送"
        Me.btnSend.UseVisualStyleBackColor = True
        '
        'txtMessages
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.txtMessages, 4)
        Me.txtMessages.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtMessages.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.txtMessages.Location = New System.Drawing.Point(3, 3)
        Me.txtMessages.Name = "txtMessages"
        Me.txtMessages.ReadOnly = True
        Me.txtMessages.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.txtMessages.Size = New System.Drawing.Size(641, 658)
        Me.txtMessages.TabIndex = 2
        Me.txtMessages.Text = ""
        '
        'chkIncludeSelection
        '
        Me.chkIncludeSelection.AutoSize = True
        Me.chkIncludeSelection.Dock = System.Windows.Forms.DockStyle.Fill
        Me.chkIncludeSelection.Location = New System.Drawing.Point(164, 708)
        Me.chkIncludeSelection.Name = "chkIncludeSelection"
        Me.chkIncludeSelection.Size = New System.Drawing.Size(155, 35)
        Me.chkIncludeSelection.TabIndex = 3
        Me.chkIncludeSelection.Text = "引入所选单元格内容"
        Me.chkIncludeSelection.UseVisualStyleBackColor = True
        '
        'chkUseGPT
        '
        Me.chkUseGPT.AutoSize = True
        Me.chkUseGPT.Checked = True
        Me.chkUseGPT.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkUseGPT.Dock = System.Windows.Forms.DockStyle.Fill
        Me.chkUseGPT.Location = New System.Drawing.Point(164, 667)
        Me.chkUseGPT.Name = "chkUseGPT"
        Me.chkUseGPT.Size = New System.Drawing.Size(155, 35)
        Me.chkUseGPT.TabIndex = 4
        Me.chkUseGPT.Text = "使用ChatGPT"
        Me.chkUseGPT.UseVisualStyleBackColor = True
        '
        'chkUseDS
        '
        Me.chkUseDS.AutoSize = True
        Me.chkUseDS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.chkUseDS.Location = New System.Drawing.Point(325, 667)
        Me.chkUseDS.Name = "chkUseDS"
        Me.chkUseDS.Size = New System.Drawing.Size(155, 35)
        Me.chkUseDS.TabIndex = 5
        Me.chkUseDS.Text = "使用Deepseek"
        Me.chkUseDS.UseVisualStyleBackColor = True
        '
        'chkUseTushareDocs
        '
        Me.chkUseTushareDocs.AutoSize = True
        Me.chkUseTushareDocs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.chkUseTushareDocs.Location = New System.Drawing.Point(325, 708)
        Me.chkUseTushareDocs.Name = "chkUseTushareDocs"
        Me.chkUseTushareDocs.Size = New System.Drawing.Size(155, 35)
        Me.chkUseTushareDocs.TabIndex = 6
        Me.chkUseTushareDocs.Text = "使用Tushare文档"
        Me.chkUseTushareDocs.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 4
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.chkUseTushareDocs, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtMessages, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.txtUserInput, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.btnSend, 3, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.chkUseGPT, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.chkIncludeSelection, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.chkUseDS, 2, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 4
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 80.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(647, 830)
        Me.TableLayoutPanel1.TabIndex = 7
        '
        'AIChatUserControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "AIChatUserControl"
        Me.Size = New System.Drawing.Size(647, 830)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtUserInput As Windows.Forms.TextBox
    Friend WithEvents btnSend As Windows.Forms.Button
    Friend WithEvents txtMessages As Windows.Forms.RichTextBox
    Friend WithEvents chkIncludeSelection As Windows.Forms.CheckBox
    Friend WithEvents chkUseGPT As Windows.Forms.CheckBox
    Friend WithEvents chkUseDS As Windows.Forms.CheckBox
    Friend WithEvents chkUseTushareDocs As Windows.Forms.CheckBox
    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
End Class
