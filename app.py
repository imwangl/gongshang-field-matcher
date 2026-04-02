import os
import io
import re
import json
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

VERSION = "1.0.0"
SHEET_ID = "1V6uygE_6POZjS8kHuvtGpWxn5LdpUdwxRg9g87RLWuE"

def get_sheet_data(sheet_name):
    """从Google Sheets获取数据"""
    # Sheet名称URL编码
    import urllib.parse
    encoded_name = urllib.parse.quote(sheet_name)
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={encoded_name}"
    
    try:
        response = requests.get(url, timeout=30)
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
                return data
    except Exception as e:
        print(f"获取失败: {e}")
    return []

def search_in_sheet(sheet_name, keyword, search_columns):
    """在指定Sheet的列中搜索关键词"""
    data = get_sheet_data(sheet_name)
    if not data:
        return None
    
    keyword = keyword.strip()
    keyword_normalized = keyword.replace("信息", "").replace("数据", "").replace("内容", "")
    
    results = []
    for row in data:
        for col in search_columns:
            if col not in row:
                continue
            cell_value = str(row.get(col, '')).strip()
            if not cell_value:
                continue
            
            # 精确匹配
            if cell_value == keyword:
                results.append({'match': cell_value, 'score': 100, 'source': sheet_name})
                continue
            
            # 包含匹配
            if keyword in cell_value or cell_value in keyword:
                results.append({'match': cell_value, 'score': 85, 'source': sheet_name})
                continue
            
            # 语义相似度
            try:
                sim = Levenshtein.ratio(keyword_normalized, cell_value.replace("信息", ""))
                if sim >= 0.6:
                    results.append({'match': cell_value, 'score': int(sim * 100), 'source': sheet_name})
            except:
                pass
    
    # 返回最佳匹配
    if results:
        results.sort(key=lambda x: x['score'], reverse=True)
        return results[0]
    return None

def find_match(user_field):
    """在所有Sheet中查找匹配"""
    # 搜索顺序：目录 -> 其他Sheet -> Sheet1
    search_order = [
        ('目录', ['D', 'C']),  # 目录Sheet的D列是中文表名
        ('开庭公告', ['A']),
        ('法院公告', ['A']),
        ('终本案件', ['A']),
        ('限制消费令', ['A']),
        ('法律诉讼', ['A']),
        ('破产重整', ['A']),
        ('被执行信息', ['A']),
        ('失信信息', ['A']),
        ('股权出质', ['A']),
        ('动产抵押', ['A']),
        ('经营异常', ['A']),
        ('行政��罚', ['A']),
        ('Sheet1', ['A', 'B']),  # Sheet1的A列是英文表名，B列是中文表名
    ]
    
    for sheet_name, cols in search_order:
        result = search_in_sheet(sheet_name, user_field, cols)
        if result:
            return result
    
    return None

def parse_user_fields(filepath):
    """解析用户上传的文件"""
    ext = os.path.splitext(filepath)[1].lower()
    fields = []
    
    if ext in ['.xlsx', '.xls']:
        df = pd.read_excel(filepath)
        # 跳过第一行表头，从第二行开始
        fields = df.iloc[1:, 0].dropna().tolist()
    elif ext in ['.txt']:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                
                # 清理行尾分号
                line = line.rstrip('；').rstrip(';')
                if '；' in line or ';' in line:
                    line = line.replace(';', '、').replace('；', '、')
                
                # 结构化格式 "1、公司概况：基本信息、联系方式"
                match = re.match(r'^\d+、[^：]+：(.+)$', line)
                if match:
                    content = match.group(1)
                    parts = content.split('、')
                    fields.extend([p.strip() for p in parts if p.strip()])
                else:
                    # 普通分隔
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
    """下载模板文件"""
    if type == 'excel':
        # Excel模板
        df = pd.DataFrame({'表/字段名': []})
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        return send_file(output, download_name='工商字段匹配模板.xlsx', as_attachment=True)
    elif type == 'txt':
        # TXT模板
        content = """1、公司概况：基本信息、联系方式、变更记录、主要人员；
2、股东和对外投资：股东信息、对外投资；
3、法定代表人信息：担任法定代表人的企业；
4、风险信息：开庭公告、法律诉讼；
5、经营信息：行政许可、税务评级；
"""
        output = io.BytesIO(content.encode('utf-8'))
        return send_file(output, download_name='工商字段匹配模板.txt', as_attachment=True)
    
    return "模板不存在", 404

@app.route('/match', methods=['POST'])
def match_fields():
    """匹配字段"""
    if 'file' not in request.files:
        return jsonify({'error': '请上传文件'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '请选择文件'}), 400
    
    # 保存文件
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)
    
    try:
        # 解析用户字段
        user_fields = parse_user_fields(filepath)
        
        if not user_fields:
            return jsonify({'error': '未能解析出字段'}), 400
        
        # 匹配
        results = []
        for field in user_fields:
            match_result = find_match(field)
            
            if match_result:
                score = match_result['score']
                if score == 100:
                    match_type = '完全匹配'
                else:
                    match_type = '推荐'
                results.append({
                    'user_field': field,
                    'matched': match_result['match'],
                    'source': match_result['source'],
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
            'results': results
        })
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/download')
def download_result():
    """下载结果"""
    result_path = os.path.join(app.config['OUTPUT_FOLDER'], 'matching_result.xlsx')
    if os.path.exists(result_path):
        return send_file(result_path, as_attachment=True)
    return "文件未找到", 404

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)