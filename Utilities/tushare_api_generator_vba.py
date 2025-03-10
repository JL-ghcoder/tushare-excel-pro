import json
import os

def generate_specific_tushare_functions(json_file, output_file):
    """
    从tushare_api_guide.json生成特定格式的VB.NET代码，用于Excel-DNA加载项
    """

    with open(json_file, 'r', encoding='utf-8') as f:
        api_guide = json.load(f)
    
    # 输出VB代码
    with open(output_file, 'w', encoding='utf-8') as f:
        # 写入必要的引用和模块定义
        f.write("Imports System.Net\n")
        f.write("Imports System.IO\n")
        f.write("Imports System.Text\n")
        f.write("Imports Newtonsoft.Json\n")
        f.write("Imports Newtonsoft.Json.Linq\n")
        f.write("Imports ExcelDna.Integration\n\n")
        
        f.write("Partial Public Module TushareFunctions\n")
        f.write("    ' 这一部分代码全部用程序生成\n")
        f.write("    ' 如果更新了接口就直接替换这一部分代码\n")
        
        # 为每个API生成Excel函数
        for api_name, api_info in api_guide.items():
            function_name = f"TUSHARE_{api_name.upper()}"
            description = api_info.get('description', '').replace('\n', ' ').replace("\"", "\"\"")
            title = api_info.get('title', '').replace('\n', ' ').replace("\"", "\"\"")
            
            # 写入函数注释和Excel函数特性
            f.write(f"\n    ''' <summary>\n")
            f.write(f"    ''' {title}\n")
            f.write(f"    ''' {description}\n")
            f.write(f"    ''' </summary>\n")
            f.write(f"    <ExcelFunction(Description:=\"{title}\", Category:=\"Tushare\")>\n")
            f.write(f"    Public Function {function_name}(\n")
            
            # 添加10个参数
            for i in range(1, 11):
                if i < 10:
                    f.write(f"            <ExcelArgument(Description:=\"参数名{i}\")> Optional param{i}Name As String = \"\",\n")
                    f.write(f"            <ExcelArgument(Description:=\"参数值{i}\")> Optional param{i}Value As String = \"\",\n")
                else:
                    f.write(f"            <ExcelArgument(Description:=\"参数名{i}\")> Optional param{i}Name As String = \"\",\n")
                    f.write(f"            <ExcelArgument(Description:=\"参数值{i}\")> Optional param{i}Value As String = \"\") As Object\n")
            
            # 函数体
            f.write(f"        Try\n")
            f.write(f"            ' 调用通用API函数，传入{api_name} API名称和所有参数\n")
            f.write(f"            Return TUSHARE_API(\"{api_name}\", param1Name, param1Value, param2Name, param2Value,\n")
            f.write(f"                              param3Name, param3Value, param4Name, param4Value, param5Name, param5Value,\n")
            f.write(f"                              param6Name, param6Value, param7Name, param7Value, param8Name, param8Value,\n")
            f.write(f"                              param9Name, param9Value, param10Name, param10Value)\n")
            f.write(f"        Catch ex As Exception\n")
            f.write(f"            ' 异常处理\n")
            f.write(f"            Return \"错误: \" & ex.Message\n")
            f.write(f"        End Try\n")
            f.write(f"    End Function\n")
        
        # 结束模块定义
        f.write("End Module")


if __name__ == "__main__":
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        json_file = os.path.join(current_dir, "tushare_api_guide.json")
        output_file = os.path.join(current_dir, "TushareFunctionsGenerated.vb")
        
        # 生成API
        generate_specific_tushare_functions(json_file, output_file)
        print(f"生成的VB代码已保存到: {output_file}")
    except Exception as e:
        print(f"错误: {str(e)}")