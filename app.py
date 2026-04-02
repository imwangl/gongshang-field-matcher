import os
import io
import re
import json
import time
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
import requests
import Levenshtein

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

VERSION = "1.0.2"
SHEET_ID = "1V6uygE_6POZjS8kHuvtGpWxn5LdpUdwxRg9g87RLWuE"

# 缓存数据
SHEET_CACHE = None
CACHE_TIME = 0
CACHE_TTL = 3600  # 缓存1小时

def get_sheet_data_cached():
    """从Google Sheets获取数据，使用缓存"""
    global SHEET_CACHE, CACHE_TIME
    
    current_time = time.time()
    
    # 检查缓存是否有效
    if SHEET_CACHE is not None and (current_time - CACHE_TIME) < CACHE_TTL:
        print(f"使用缓存数据，{int(CACHE_TTL - (current_time - CACHE_TIME))}秒后过期")
        return SHEET_CACHE
    
    # 重新获取数据
    print("从Google Sheets获取数据...")
    try:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet=%E7%9B%AE%E5%BD%95"
        response = requests.get(url, timeout=60)
        
        if response.status_code == 200:
            # 解析CSV
            lines = response.text.strip().split('\n')
            if len(lines) >= 2:
                headers = lines[0].split(',')
                data = []
                for line in lines[1:]:
                    if line.strip():
                        values = line.split(',')
                        row = {}
                        for i, h in enumerate(headers):
                            row[h.strip()] = values[i].strip() if i < len(values) else ''
                        data.append(row)
                
                SHEET_CACHE = data
                CACHE_TIME = current_time
                print(f"获取到 {len(data)} 行数据")
                return data
    except Exception as e:
        print(f"获取数据失败: {e}")
    
    return []

def search_in_sheet(keyword):
    """在目录Sheet中搜索关键词"""
    data = get_sheet_data_cached()
    if not data:
        return None
    
    keyword = keyword.strip()
    keyword_normalized = keyword.replace("信息", "").replace("数据", "").replace("内容", "")
    
    results = []
    for row in data:
        search_cols = ['对应数据名称', '数据表']
        
        for col in search_cols:
            cell_value = str(row.get(col, '')).strip()
            if not cell_value or cell_value == '':
                continue
            
            # 精确匹配
            if cell_value == keyword:
                results.append({'match': cell_value, 'score': 100, 'source': '目录'})
                continue
            
            # 包含匹配（双向）
            if keyword in cell_value or cell_value in keyword:
                results.append({'match': cell_value, 'score': 85, 'source': '目录'})
                continue
            
            # 部分前缀/后缀匹配
            cell_clean = cell_value.replace("工商-", "").replace("企业", "").replace("公司", "")
            key_clean = keyword.replace("工商-", "").replace("企业", "").replace("公司", "")
            if cell_clean in key_clean or key_clean in cell_clean:
                results.append({'match': cell_value, 'score': 70, 'source': '目录'})
                continue
            
            # ��义相似度
            try:
                sim1 = Levenshtein.ratio(keyword, cell_value)
                sim2 = Levenshtein.ratio(key_clean, cell_clean)
                sim = max(sim1, sim2)
                if sim >= 0.5:
                    results.append({'match': cell_value, 'score': int(sim * 100), 'source': '目录'})
            except:
                pass
    
    if results:
        results.sort(key=lambda x: x['score'], reverse=True)
        return results[0]
    return None

def find_match(user_field):
    """查找匹配"""
    return search_in_sheet(user_field)

def parse_user_fields(filepath):
    """解析用户上传的文件"""
    ext = os.path.splitext(filepath)[1].lower()
    fields = []
    
    if ext in ['.xlsx', '.xls']:
        df = pd.read_excel(filepath)
        fields = df.iloc[1:, 0].dropna().tolist()
    elif ext in ['.txt']:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                
                line = line.rstrip('；').rstrip(';')
                if '；' in line or ';' in line:
                    line = line.replace(';', '、').replace('；', '、')
                
                match = re.match(r'^\d+、[^：]+：(.+)$', line)
                if match:
                    content = match.group(1)
                    parts = content.split('、')
                    fields.extend([p.strip() for p in parts if p.strip()])
                else:
                    for sep in ['、', '，', ',']:
                        if sep in line:
                            fields.extend([p.strip() for p in line.split(sep) if p.strip()])
                            break
                    else:
                        if line.strip():
                            fields.append(line.strip())
    
    return fields

@app.route('/')
def index():
    return render_template('index.html', version=VERSION)

@app.route('/template/<type>')
def download_template(type):
    if type == 'excel':
        df = pd.DataFrame({'表/字段名': []})
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        return send_file(output, download_name='工商字段匹配模板.xlsx', as_attachment=True)
    elif type == 'txt':
        content = """1、公司概况：基本信息、联系方式、变更记录、主要人员；
2、股东和对外投资：股东信息、对外投资；
"""
        output = io.BytesIO(content.encode('utf-8'))
        return send_file(output, download_name='工商字段匹配模板.txt', as_attachment=True)
    return "模板不存在", 404

@app.route('/match', methods=['POST'])
def match_fields():
    try:
        if 'file' not in request.files:
            return jsonify({'error': '请上传文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '请选择文件'}), 400
        
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        
        # 解析字段
        user_fields = parse_user_fields(filepath)
        
        if not user_fields:
            return jsonify({'error': '未能解析出字段'}), 400
        
        # 匹配
        results = []
        for field in user_fields:
            result = find_match(field)
            
            if result:
                score = result['score']
                if score == 100:
                    match_type = '完全匹配'
                else:
                    match_type = '推荐'
                results.append({
                    'user_field': field,
                    'matched': result['match'],
                    'source': result['source'],
                    'match_type': match_type,
                    'score': score
                })
            else:
                results.append({
                    'user_field': field,
                    'matched': '-',
                    'source': '-',
                    'match_type': '匹配不到',
                    'score': 0
                })
        
        # 统计
        total = len(results)
        exact = len([r for r in results if r['match_type'] == '完全匹配'])
        recommend = len([r for r in results if r['match_type'] == '推荐'])
        no_match = len([r for r in results if r['match_type'] == '匹配不到'])
        
        # 保存结果
        result_df = pd.DataFrame(results)
        output = io.BytesIO()
        result_df.to_excel(output, index=False)
        output.seek(0)
        
        result_path = os.path.join(app.config['OUTPUT_FOLDER'], 'matching_result.xlsx')
        with open(result_path, 'wb') as f:
            f.write(output.getvalue())
        
        return jsonify({
            'success': True,
            'stats': {'total': total, 'exact': exact, 'recommend': recommend, 'no_match': no_match},
            'results': results[:100]
        })
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/download')
def download_result():
    result_path = os.path.join(app.config['OUTPUT_FOLDER'], 'matching_result.xlsx')
    if os.path.exists(result_path):
        return send_file(result_path, as_attachment=True)
    return "文件未找到", 404

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)