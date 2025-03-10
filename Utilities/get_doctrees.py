import pandas as pd
import json
import requests
from anytree import Node, RenderTree
from anytree.exporter import DotExporter

def get_doctrees(token):
    url = "http://api.tushare.pro/" # Tushare官方doctrees接口
    payload = {
        "api_name": "doctrees",
        "token": token,
        "params": {},
        "fields": ""
    }
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json"
    }
    
    response = requests.post(url, data=json.dumps(payload), headers=headers)
    
    if response.status_code == 200:
        result = response.json()
        if result.get('code') == 0:
            data = result.get('data')
            if data:
                df = pd.DataFrame(data.get('items'), columns=data.get('fields'))
                return df
            else:
                print("没有返回数据")
                return None
        else:
            print(f"API请求错误: {result.get('msg')}")
            return None
    else:
        print(f"HTTP请求失败，状态码: {response.status_code}")
        return None

def create_tree_structure(df):

    nodes = {}
    
    # 首先创建所有节点
    for _, row in df.iterrows():
        node_id = row['id']
        parent_id = row['parent_id']
        title = row['title']
        level = row['level']
        path = row['path']
        
        # 清理title（去除Markdown标记）
        clean_title = title.split('\n')[0].replace('#', '').replace('*', '').strip()
        if '----' in clean_title:
            clean_title = clean_title.split('----')[0].strip()
        
        nodes[node_id] = {
            'id': node_id,
            'parent_id': parent_id,
            'title': clean_title,
            'level': level,
            'path': path,
            'node_obj': None,
            'children': []
        }
    
    # 构建树结构
    root = None
    for node_id, node_info in nodes.items():
        parent_id = node_info['parent_id']
        
        # 如果没有父节点，则设为根节点
        if parent_id is None or parent_id == 0 or parent_id not in nodes:
            node_info['node_obj'] = Node(node_info['title'], id=node_id)
            if root is None:  # 设置第一个遇到的顶级节点为根节点
                root = node_info['node_obj']
        else:
            # 如果父节点存在，则添加为其子节点
            parent_node = nodes[parent_id]['node_obj']
            if parent_node:
                node_info['node_obj'] = Node(node_info['title'], parent=parent_node, id=node_id)
                nodes[parent_id]['children'].append(node_id)
    
    return root, nodes

def print_tree(root):
    # 打印树结构
    print("\n树状结构展示：")
    for pre, _, node in RenderTree(root):
        print(f"{pre}{node.name}")

def generate_api_guide(df, nodes):
    # 生成API指南
    api_guide = {}
    
    for _, row in df.iterrows():
        node_id = row['id']
        title = row['title']
        path = row['path']
        content = row['src_content']
        
        # 如果内容中包含"接口："，则提取API信息
        if content and '接口：' in content:
            # 提取API名称
            api_name_start = content.find('接口：') + 3
            api_name_end = content.find('\n', api_name_start)
            if api_name_end == -1:  # 如果没有换行符，则取整个字符串
                api_name_end = len(content)
            
            # 提取基本的API名称（如果有英文逗号或中文逗号，只保留逗号前的内容）
            api_name_raw = content[api_name_start:api_name_end].strip()
            if ',' in api_name_raw:
                api_name = api_name_raw.split(',')[0].strip()
            elif '，' in api_name_raw:
                api_name = api_name_raw.split('，')[0].strip()
            else:
                api_name = api_name_raw
            
            # 提取描述（如果有）
            description = ""
            if '描述：' in content:
                desc_start = content.find('描述：') + 3
                desc_end = content.find('\n', desc_start)
                if desc_end == -1:
                    desc_end = len(content)
                description = content[desc_start:desc_end].strip()
            
            api_guide[api_name] = {
                'title': title.split('\n')[0].replace('#', '').replace('*', '').strip(),
                'path': path,
                'description': description
            }
    
    return api_guide

def main():
    # 替换为你的实际token
    token = "TOKEN"
    df = get_doctrees(token)
    
    if df is None:
        print("无法获取数据")
        return
    
    # 清理数据
    df = df.fillna('')  # 将NaN替换为空字符串
    
    # 创建树结构
    root, nodes = create_tree_structure(df)
    
    # 打印树结构
    print_tree(root)
    
    # 生成API指南
    api_guide = generate_api_guide(df, nodes)
    
    # 打印API指南
    print("\nAPI指南：")
    for api_name, info in api_guide.items():
        print(f"\n接口名称: {api_name}")
        print(f"标题: {info['title']}")
        print(f"路径: {info['path']}")
        if info['description']:
            print(f"描述: {info['description']}")
    
    # 保存为JSON
    with open('tushare_api_guide.json', 'w', encoding='utf-8') as f:
        json.dump(api_guide, f, ensure_ascii=False, indent=2)
    print("\nAPI指南已保存至 tushare_api_guide.json")
    
    # 创建分类索引
    categories = {}
    for node_id, node_info in nodes.items():
        if node_info['level'] == 1:  # 第一级分类
            categories[node_id] = {
                'name': node_info['title'],
                'subcategories': {}
            }
            
            # 添加子分类
            for child_id in node_info['children']:
                child_info = nodes[child_id]
                categories[node_id]['subcategories'][child_id] = {
                    'name': child_info['title'],
                    'apis': []
                }
                
                # 添加API
                for api_name, api_info in api_guide.items():
                    if api_info['path'].startswith(child_info['path']):
                        categories[node_id]['subcategories'][child_id]['apis'].append({
                            'name': api_name,
                            'title': api_info['title'],
                            'description': api_info['description']
                        })
    
    # 保存分类索引
    with open('tushare_categories.json', 'w', encoding='utf-8') as f:
        json.dump(categories, f, ensure_ascii=False, indent=2)
    print("分类索引已保存至 tushare_categories.json")

if __name__ == "__main__":
    main()