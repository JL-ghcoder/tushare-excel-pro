## 概览
**Tushare Excel Pro** 是一款专为金融数据分析打造的 Excel 插件，让用户无需编程即可高效使用 **Tushare** 进行数据查询与分析。无论是金融研究员、投资者，还是量化分析师，这款插件都能帮助您在 Excel 中轻松调用**Tushare** 的全部接口，实现高效的数据获取、分析与可视化。     
<br>

## 核心功能
✅ **一键调用 Tushare 全部接口**

+ 在 Excel 单元格中直接输入 **"TUSHARE_ " + "官方接口名称"** 公式，即可调用 Tushare 的全部 API 接口，快速获取股票、指数、基金、期货、宏观经济等数据。
+ 支持参数配置，自定义筛选数据，提高数据处理的灵活性。

<br><br>
✅ **内置侧边栏显示官方接口文档**

+ 侧边栏实时展示 **Tushare 官方 API 文档**，无需频繁切换浏览器查阅文档，提高工作效率。
+ 直接点击文档示例，一键填充 Excel 公式，提高数据调用的便捷性。

<br><br>
✅ **内置 AI 数据分析助手**

+ 插件内置了ChatGPT与DeepSeek模型，可以使用AI自动分析获取的数据。
+ 结合 VBA 和 Excel 公式，智能识别数据模式，给出数据解读与优化建议。

<br><br>
✅ **AI 直接运行 VBA 代码 & Excel 公式**

+ 无需手动编写代码，AI 可根据您的需求自动生成 **VBA 代码** 并直接运行，极大提升 Excel 自动化能力。
+ 支持 AI 阅读 Tushare 接口文档，并优化 **Excel 公式**，自动推荐最优公式写法，让表格数据使用更高效。  

<br><br>
✅ **Excel 接口文档作为 AI 知识库**

+ AI 内置 Tushare API 知识库，支持自然语言查询接口用法，帮助用户快速掌握 API 的使用方式。
+ 可针对不同数据类型，智能推荐适合的 API 调用方式，降低使用门槛。

<br><br>
## 安装方法
该插件依赖于 Windows 环境，Mac 用户是无法进行安装的。

首先确保你的 **Excel 版本大于2019**，并且电脑具备至少 **.Net Framework 4.6.2** 及以上环境。

(如果你使用 Excel 2019甚至更低的版本将无法正确应用公式)


<br><br>
从 Release 或者 项目文件里直接下载 **TushareExcelPro-setup.exe** 安装文件



可以安装到任何你想要的位置，安装成功后会自动运行 vsto 配置程序，配置结束后重启你的 Excel 就能看到出现了一个新的 Tab标签 - **Tushare Excel**



![](https://cdn.nlark.com/yuque/0/2025/png/25506303/1741622742030-64866e09-be38-42ad-bf1f-bdd6b552a8e4.png)


<br><br>
在安装好 Ribbon 后还需要导入内置的 Tushare 公式，在 开发工具 -> Excel 加载项 -> 浏览 里选择安装目录下 **TushareExcel-AddIn64-packed.xll** 这个文件。

(如果你的 Excel 没有开发工具这个标签，你需要自己手动打开)



![](https://cdn.nlark.com/yuque/0/2025/png/25506303/1741622919535-5c078475-7649-4092-8d61-51606d7249b3.png)



这样当你在 Excel 工作表里打出 **=TUSHARE** 时就能发现你已经能够调用 Tushare 的所有相关接口。



![](https://cdn.nlark.com/yuque/0/2025/png/25506303/1741623029503-e4f5e367-5cac-47e5-b8bd-e466a2e3da83.png)


<br><br>
接下来在 Tushare Excel标签 里分别点击 **配置**，**DeepSeek配置**，**ChatGPT配置**就可以配置你的Tushare，ChatGPT以及DeepSeek API url/key。当你点击保存时配置文件会自动写入注册表，所以你下次打开时就不用再重新输入了。



![](https://cdn.nlark.com/yuque/0/2025/png/25506303/1741625216142-2d9a9d93-3859-4d9c-809d-46a064d7e200.png)


<br><br>
## 使用方法
### Excel公式使用方法
内置的公式全部大写以 **TUSHARE_+接口名称** 组成。

例如：`=TUSHARE_ADJ_FACTOR()`

<br><br>
详细的接口名称与功能可以在 [Tushare 官网](https://tushare.pro/document/2) 查询，也可以询问AI，或者你可以点击标签栏里的文档查看。



对于参数，请你以 `"参数名", "值"` 的方式来编写公式，例如获取茅台10月的数据可以编写成这样：

`=TUSHARE_DAILY("ts_code", "600519.SH", "start_date", "20241001", "end_date", "20241031") `


<br><br>
其中**参数名和官网的参数一一对应**，同时每一个公式都可以通过两个额外的参数进行排版：

showheader：布尔变量用于控制是否要输出表头

fields：用于控制输出哪几列



例如获取茅台10月的 trade_date, close 列数据并且不要表头可以编写成：

`=TUSHARE_DAILY("ts_code", "600519.SH", "start_date", "20241001", "end_date", "20241031", "showheader", FALSE, "fields", "trade_date,close")`


<br><br>
### AI功能使用方法
要使用AI前你需要先配置好你自己的API Key



点击 **开启AI** 使用AI功能时有4个选项：



![](https://cdn.nlark.com/yuque/0/2025/png/25506303/1741623357370-c8b326d4-7995-4126-86f3-a0eb6bbc09ca.png)



前两个是用于选择你要使用的AI


<br><br>
**引入所选单元格内容** 勾选后会在你点击发送时把你在工作表里框选的内容一并传给AI，例如我可以框住工作表里的股票数据，然后让AI告诉我在过去一个月里股价最高的是哪一天。



**使用Tushare文档** 勾选后AI会自动读取当前版本的 Tushare Excel 接口，例如我勾选后可以询问AI：<u>能不能用excel公式把茅台2024年10月的ohlcv数据拿给我</u>


<br><br>
![](https://cdn.nlark.com/yuque/0/2025/png/25506303/1741623584677-7d7b63da-77c3-4e47-86c1-6a9709b720f2.png)


<br><br>
你可以点击 **应用此公式**，它会自动在你工作表所在单元格里粘贴这个公式 (你可能需要手动去除开头的'@'符号，这是Excel的自动保护)


除此之外你也可以让AI**自动编写并且执行vba宏程序**，例如你可以询问AI：<u>给我写一个VBA程序，帮我把所选单元格内容根据分数列从高到底进行排序</u>。



![](https://cdn.nlark.com/yuque/0/2025/png/25506303/1741624079235-493657fa-f5d0-4bb2-8800-07a18dfaab90.png)



你可以点击运行此代码，AI会自动执行该宏程序。

注意：**执行宏程序是不可撤销的！**


<br><br>
## 联系我们
Tushare Excel Pro 会跟随 Tushare 官方同步更新。

你可以通过github issues或者邮件联系到我们:  

ispierce.zhou@gmail.com

isjun.liu@gmail.com