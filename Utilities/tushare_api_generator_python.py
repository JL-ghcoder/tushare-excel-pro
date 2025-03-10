import json
import os

def generate_tushare_apis(json_file, output_file):
    """
    从tushare_api_guide.json生成python接口函数代码
    """

    with open(json_file, 'r', encoding='utf-8') as f:
        api_guide = json.load(f)
    
    # 输出Python代码
    with open(output_file, 'w', encoding='utf-8') as f:
        # 写入导入语句
        f.write("import requests\n")
        f.write("import json\n")
        f.write("import pandas as pd\n\n")
        
        # 写入通用API调用函数
        f.write("def call_tushare_api(api_name, token, **params):\n")
        f.write("    \"\"\"通用Tushare API调用函数\n")
        f.write("    \n")
        f.write("    Args:\n")
        f.write("        api_name (str): API名称\n")
        f.write("        token (str): Tushare API令牌\n")
        f.write("        **params: API参数\n")
        f.write("    \n")
        f.write("    Returns:\n")
        f.write("        pandas.DataFrame: API返回数据\n")
        f.write("    \"\"\"\n")
        f.write("    url = \"http://api.tushare.pro\"\n")
        f.write("    \n")
        f.write("    # 构建请求数据\n")
        f.write("    payload = {\n")
        f.write("        \"api_name\": api_name,\n")
        f.write("        \"token\": token,\n")
        f.write("        \"params\": params,\n")
        f.write("        \"fields\": \"\"\n")
        f.write("    }\n")
        f.write("    \n")
        f.write("    # 发送请求\n")
        f.write("    headers = {\n")
        f.write("        \"Content-Type\": \"application/json\",\n")
        f.write("        \"Accept\": \"application/json\"\n")
        f.write("    }\n")
        f.write("    \n")
        f.write("    try:\n")
        f.write("        response = requests.post(url, data=json.dumps(payload), headers=headers)\n")
        f.write("        response.raise_for_status()  # 检查HTTP错误\n")
        f.write("        \n")
        f.write("        # 解析响应\n")
        f.write("        result = response.json()\n")
        f.write("        \n")
        f.write("        if result.get('code') == 0:\n")
        f.write("            # 成功获取数据\n")
        f.write("            data = result.get('data')\n")
        f.write("            if data and 'items' in data and 'fields' in data:\n")
        f.write("                # 转换为DataFrame\n")
        f.write("                df = pd.DataFrame(data['items'], columns=data['fields'])\n")
        f.write("                return df\n")
        f.write("            else:\n")
        f.write("                raise ValueError(\"API返回的数据格式不正确\")\n")
        f.write("        else:\n")
        f.write("            # API返回错误\n")
        f.write("            raise ValueError(f\"API错误: {result.get('msg')}\")\n")
        f.write("    \n")
        f.write("    except Exception as e:\n")
        f.write("        raise Exception(f\"调用Tushare API时出错: {str(e)}\")\n\n")
        
        # 为每个API生成函数
        for api_name, api_info in api_guide.items():
            function_name = f"tushare_{api_name}"
            description = api_info.get('description', '').replace('\n', ' ')
            title = api_info.get('title', '').replace('\n', ' ')
            
            f.write(f"\ndef {function_name}(token, **params):\n")
            f.write(f"    \"\"\"\n")
            f.write(f"    {title}\n")
            f.write(f"    \n")
            f.write(f"    {description}\n")
            f.write(f"    \n")
            f.write(f"    Args:\n")
            f.write(f"        token (str): Tushare API令牌\n")
            f.write(f"        **params: API参数\n")
            f.write(f"    \n")
            f.write(f"    Returns:\n")
            f.write(f"        pandas.DataFrame: API返回数据\n")
            f.write(f"    \"\"\"\n")
            f.write(f"    return call_tushare_api(\"{api_name}\", token, **params)\n")
        
        # 生成函数映射字典
        f.write("\n# 函数映射字典\nAPI_FUNCTIONS = {\n")
        for api_name in api_guide.keys():
            function_name = f"tushare_{api_name}"
            f.write(f"    '{api_name}': {function_name},\n")
        f.write("}\n")

if __name__ == "__main__":
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        json_file = os.path.join(current_dir, "tushare_api_guide.json")
        output_file = os.path.join(current_dir, "tushare_api.py")
        
        # 生成API
        generate_tushare_apis(json_file, output_file)
        print(f"Generated Tushare API code to: {output_file}")
    except Exception as e:
        print(f"error: {str(e)}")