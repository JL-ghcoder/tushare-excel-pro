<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FunctionGuidePane
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
        Me.treeViewFunctions = New System.Windows.Forms.TreeView()
        Me.txtFunctionDetails = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'treeViewFunctions
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.treeViewFunctions, 2)
        Me.treeViewFunctions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.treeViewFunctions.Font = New System.Drawing.Font("宋体", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.treeViewFunctions.Location = New System.Drawing.Point(3, 44)
        Me.treeViewFunctions.Name = "treeViewFunctions"
        Me.treeViewFunctions.Size = New System.Drawing.Size(202, 649)
        Me.treeViewFunctions.TabIndex = 0
        '
        'txtFunctionDetails
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.txtFunctionDetails, 2)
        Me.txtFunctionDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtFunctionDetails.Font = New System.Drawing.Font("宋体", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.txtFunctionDetails.Location = New System.Drawing.Point(211, 44)
        Me.txtFunctionDetails.Multiline = True
        Me.txtFunctionDetails.Name = "txtFunctionDetails"
        Me.txtFunctionDetails.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtFunctionDetails.Size = New System.Drawing.Size(236, 649)
        Me.txtFunctionDetails.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("宋体", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 22)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "搜索："
        '
        'txtSearch
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.txtSearch, 3)
        Me.txtSearch.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.txtSearch.Font = New System.Drawing.Font("宋体", 11.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.txtSearch.Location = New System.Drawing.Point(88, 5)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(359, 33)
        Me.txtSearch.TabIndex = 3
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 4
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.txtFunctionDetails, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtSearch, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.treeViewFunctions, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 6.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 94.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(450, 696)
        Me.TableLayoutPanel1.TabIndex = 4
        '
        'FunctionGuidePane
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "FunctionGuidePane"
        Me.Size = New System.Drawing.Size(450, 696)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents treeViewFunctions As Windows.Forms.TreeView
    Friend WithEvents txtFunctionDetails As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents txtSearch As Windows.Forms.TextBox
    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
End Class
