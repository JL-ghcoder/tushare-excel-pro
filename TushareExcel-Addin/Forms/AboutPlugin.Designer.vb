<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AboutPlugin
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.rtbPluginInfo = New System.Windows.Forms.RichTextBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("黑体", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(250, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(280, 28)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Tushare Excel 插件"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("黑体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(306, 109)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(151, 24)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "版本 v1.0.0"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("黑体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(272, 179)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(228, 24)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "由 Jun, Peng 制作"
        '
        'rtbPluginInfo
        '
        Me.rtbPluginInfo.Font = New System.Drawing.Font("宋体", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.rtbPluginInfo.Location = New System.Drawing.Point(61, 241)
        Me.rtbPluginInfo.Name = "rtbPluginInfo"
        Me.rtbPluginInfo.ReadOnly = True
        Me.rtbPluginInfo.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.rtbPluginInfo.Size = New System.Drawing.Size(665, 458)
        Me.rtbPluginInfo.TabIndex = 3
        Me.rtbPluginInfo.Text = ""
        '
        'btnOK
        '
        Me.btnOK.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.btnOK.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.btnOK.Location = New System.Drawing.Point(328, 724)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(109, 54)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "确定"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'AboutPlugin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(792, 802)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.rtbPluginInfo)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "AboutPlugin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "关于插件"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents rtbPluginInfo As Windows.Forms.RichTextBox
    Friend WithEvents btnOK As Windows.Forms.Button
End Class
