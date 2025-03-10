Partial Class ribbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
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

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ribbon))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.settings = Me.Factory.CreateRibbonButton
        Me.docs = Me.Factory.CreateRibbonButton
        Me.intro = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.HistoricalData = Me.Factory.CreateRibbonSplitButton
        Me.btnDaily = Me.Factory.CreateRibbonButton
        Me.btnMins = Me.Factory.CreateRibbonButton
        Me.btnWeekly = Me.Factory.CreateRibbonButton
        Me.btnMonthly = Me.Factory.CreateRibbonButton
        Me.btnDailyBasic = Me.Factory.CreateRibbonButton
        Me.SplitButton1 = Me.Factory.CreateRibbonSplitButton
        Me.btnIncome = Me.Factory.CreateRibbonButton
        Me.btnBalancesheet = Me.Factory.CreateRibbonButton
        Me.btnCashflow = Me.Factory.CreateRibbonButton
        Me.btnFinaIndicator = Me.Factory.CreateRibbonButton
        Me.SplitButton2 = Me.Factory.CreateRibbonSplitButton
        Me.btnIndexBasic = Me.Factory.CreateRibbonButton
        Me.btnIndexDaily = Me.Factory.CreateRibbonButton
        Me.btnIndexWeekly = Me.Factory.CreateRibbonButton
        Me.btnIndexMonthly = Me.Factory.CreateRibbonButton
        Me.btnIndexWeight = Me.Factory.CreateRibbonButton
        Me.btnIndexGlobal = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.btnReload = Me.Factory.CreateRibbonButton
        Me.btnRemove = Me.Factory.CreateRibbonButton
        Me.btnBeautify = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.btnDeepseek = Me.Factory.CreateRibbonButton
        Me.btnGPT = Me.Factory.CreateRibbonButton
        Me.btnAI = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Label = "Tushare Excel"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.settings)
        Me.Group1.Items.Add(Me.docs)
        Me.Group1.Items.Add(Me.intro)
        Me.Group1.Label = "设置"
        Me.Group1.Name = "Group1"
        '
        'settings
        '
        Me.settings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.settings.Image = CType(resources.GetObject("settings.Image"), System.Drawing.Image)
        Me.settings.Label = "配置"
        Me.settings.Name = "settings"
        Me.settings.ShowImage = True
        '
        'docs
        '
        Me.docs.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.docs.Image = CType(resources.GetObject("docs.Image"), System.Drawing.Image)
        Me.docs.Label = "文档"
        Me.docs.Name = "docs"
        Me.docs.ShowImage = True
        '
        'intro
        '
        Me.intro.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.intro.Image = Global.TushareExcel.My.Resources.Resources.teamintro
        Me.intro.Label = "介绍"
        Me.intro.Name = "intro"
        Me.intro.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.HistoricalData)
        Me.Group2.Items.Add(Me.SplitButton1)
        Me.Group2.Items.Add(Me.SplitButton2)
        Me.Group2.Label = "快捷公式"
        Me.Group2.Name = "Group2"
        '
        'HistoricalData
        '
        Me.HistoricalData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.HistoricalData.Image = Global.TushareExcel.My.Resources.Resources.formula
        Me.HistoricalData.Items.Add(Me.btnDaily)
        Me.HistoricalData.Items.Add(Me.btnMins)
        Me.HistoricalData.Items.Add(Me.btnWeekly)
        Me.HistoricalData.Items.Add(Me.btnMonthly)
        Me.HistoricalData.Items.Add(Me.btnDailyBasic)
        Me.HistoricalData.Label = "历史行情"
        Me.HistoricalData.Name = "HistoricalData"
        Me.HistoricalData.Tag = ""
        '
        'btnDaily
        '
        Me.btnDaily.Label = "日线行情"
        Me.btnDaily.Name = "btnDaily"
        Me.btnDaily.ScreenTip = "获取股票日线历史行情数据"
        Me.btnDaily.ShowImage = True
        '
        'btnMins
        '
        Me.btnMins.Label = "分钟行情"
        Me.btnMins.Name = "btnMins"
        Me.btnMins.ScreenTip = "获取股票分钟历史行情数据"
        Me.btnMins.ShowImage = True
        '
        'btnWeekly
        '
        Me.btnWeekly.Label = "周线行情"
        Me.btnWeekly.Name = "btnWeekly"
        Me.btnWeekly.ScreenTip = "获取股票每周历史行情数据"
        Me.btnWeekly.ShowImage = True
        '
        'btnMonthly
        '
        Me.btnMonthly.Label = "月线行情"
        Me.btnMonthly.Name = "btnMonthly"
        Me.btnMonthly.ScreenTip = "获取股票每月历史行情数据"
        Me.btnMonthly.ShowImage = True
        '
        'btnDailyBasic
        '
        Me.btnDailyBasic.Label = "每日指标"
        Me.btnDailyBasic.Name = "btnDailyBasic"
        Me.btnDailyBasic.ScreenTip = "获取股票每日重要的基本面指标"
        Me.btnDailyBasic.ShowImage = True
        '
        'SplitButton1
        '
        Me.SplitButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.SplitButton1.Image = Global.TushareExcel.My.Resources.Resources.formula
        Me.SplitButton1.Items.Add(Me.btnIncome)
        Me.SplitButton1.Items.Add(Me.btnBalancesheet)
        Me.SplitButton1.Items.Add(Me.btnCashflow)
        Me.SplitButton1.Items.Add(Me.btnFinaIndicator)
        Me.SplitButton1.Label = "财务数据"
        Me.SplitButton1.Name = "SplitButton1"
        Me.SplitButton1.Tag = ""
        '
        'btnIncome
        '
        Me.btnIncome.Label = "利润表"
        Me.btnIncome.Name = "btnIncome"
        Me.btnIncome.ScreenTip = "获取股票利润表数据"
        Me.btnIncome.ShowImage = True
        '
        'btnBalancesheet
        '
        Me.btnBalancesheet.Label = "资产负债表"
        Me.btnBalancesheet.Name = "btnBalancesheet"
        Me.btnBalancesheet.ScreenTip = "获取股票资产负债表数据"
        Me.btnBalancesheet.ShowImage = True
        '
        'btnCashflow
        '
        Me.btnCashflow.Label = "现金流量表"
        Me.btnCashflow.Name = "btnCashflow"
        Me.btnCashflow.ScreenTip = "获取股票现金流量表数据"
        Me.btnCashflow.ShowImage = True
        '
        'btnFinaIndicator
        '
        Me.btnFinaIndicator.Label = "财务指标数据"
        Me.btnFinaIndicator.Name = "btnFinaIndicator"
        Me.btnFinaIndicator.ScreenTip = "获取股票财务指标数据"
        Me.btnFinaIndicator.ShowImage = True
        '
        'SplitButton2
        '
        Me.SplitButton2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.SplitButton2.Image = Global.TushareExcel.My.Resources.Resources.formula
        Me.SplitButton2.Items.Add(Me.btnIndexBasic)
        Me.SplitButton2.Items.Add(Me.btnIndexDaily)
        Me.SplitButton2.Items.Add(Me.btnIndexWeekly)
        Me.SplitButton2.Items.Add(Me.btnIndexMonthly)
        Me.SplitButton2.Items.Add(Me.btnIndexWeight)
        Me.SplitButton2.Items.Add(Me.btnIndexGlobal)
        Me.SplitButton2.Label = "指数数据"
        Me.SplitButton2.Name = "SplitButton2"
        Me.SplitButton2.Tag = ""
        '
        'btnIndexBasic
        '
        Me.btnIndexBasic.Label = "指数信息"
        Me.btnIndexBasic.Name = "btnIndexBasic"
        Me.btnIndexBasic.ScreenTip = "获取指数基本信息"
        Me.btnIndexBasic.ShowImage = True
        '
        'btnIndexDaily
        '
        Me.btnIndexDaily.Label = "指数日线行情"
        Me.btnIndexDaily.Name = "btnIndexDaily"
        Me.btnIndexDaily.ScreenTip = "获取指数日线行情"
        Me.btnIndexDaily.ShowImage = True
        '
        'btnIndexWeekly
        '
        Me.btnIndexWeekly.Label = "指数周线行情"
        Me.btnIndexWeekly.Name = "btnIndexWeekly"
        Me.btnIndexWeekly.ScreenTip = "获取指数周线行情"
        Me.btnIndexWeekly.ShowImage = True
        '
        'btnIndexMonthly
        '
        Me.btnIndexMonthly.Label = "指数月线行情"
        Me.btnIndexMonthly.Name = "btnIndexMonthly"
        Me.btnIndexMonthly.ScreenTip = "获取指数月线行情"
        Me.btnIndexMonthly.ShowImage = True
        '
        'btnIndexWeight
        '
        Me.btnIndexWeight.Label = "指数成分和权重"
        Me.btnIndexWeight.Name = "btnIndexWeight"
        Me.btnIndexWeight.ScreenTip = "获取指数成分和权重"
        Me.btnIndexWeight.ShowImage = True
        '
        'btnIndexGlobal
        '
        Me.btnIndexGlobal.Label = "国际主要指数"
        Me.btnIndexGlobal.Name = "btnIndexGlobal"
        Me.btnIndexGlobal.ScreenTip = "获取国际主要指数"
        Me.btnIndexGlobal.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.btnReload)
        Me.Group3.Items.Add(Me.btnRemove)
        Me.Group3.Items.Add(Me.btnBeautify)
        Me.Group3.Label = "辅助功能"
        Me.Group3.Name = "Group3"
        '
        'btnReload
        '
        Me.btnReload.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnReload.Image = Global.TushareExcel.My.Resources.Resources.reload
        Me.btnReload.Label = "刷新公式"
        Me.btnReload.Name = "btnReload"
        Me.btnReload.ShowImage = True
        '
        'btnRemove
        '
        Me.btnRemove.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnRemove.Image = Global.TushareExcel.My.Resources.Resources.remove
        Me.btnRemove.Label = "删除公式"
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.ShowImage = True
        '
        'btnBeautify
        '
        Me.btnBeautify.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnBeautify.Image = Global.TushareExcel.My.Resources.Resources.beautify
        Me.btnBeautify.Label = "美化表格"
        Me.btnBeautify.Name = "btnBeautify"
        Me.btnBeautify.ShowImage = True
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.btnDeepseek)
        Me.Group4.Items.Add(Me.btnGPT)
        Me.Group4.Items.Add(Me.btnAI)
        Me.Group4.Label = "AI"
        Me.Group4.Name = "Group4"
        '
        'btnDeepseek
        '
        Me.btnDeepseek.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnDeepseek.Image = Global.TushareExcel.My.Resources.Resources.deepseek
        Me.btnDeepseek.Label = "DeepSeek配置"
        Me.btnDeepseek.Name = "btnDeepseek"
        Me.btnDeepseek.ShowImage = True
        '
        'btnGPT
        '
        Me.btnGPT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnGPT.Image = Global.TushareExcel.My.Resources.Resources.gpt
        Me.btnGPT.Label = "ChatGPT配置"
        Me.btnGPT.Name = "btnGPT"
        Me.btnGPT.ShowImage = True
        '
        'btnAI
        '
        Me.btnAI.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnAI.Image = Global.TushareExcel.My.Resources.Resources.ai
        Me.btnAI.Label = "开启AI"
        Me.btnAI.Name = "btnAI"
        Me.btnAI.ShowImage = True
        '
        'ribbon
        '
        Me.Name = "ribbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents settings As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents docs As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents intro As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents HistoricalData As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents btnDaily As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnWeekly As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnMins As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnMonthly As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnDailyBasic As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton1 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents btnIncome As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnBalancesheet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnCashflow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnFinaIndicator As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton2 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents btnIndexBasic As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnIndexDaily As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnIndexWeekly As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnIndexMonthly As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnIndexWeight As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnIndexGlobal As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnReload As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnRemove As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnDeepseek As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnGPT As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnAI As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnBeautify As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property ribbon() As ribbon
        Get
            Return Me.GetRibbon(Of ribbon)()
        End Get
    End Property
End Class
