import os
import io
import re
import time
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
import Levenshtein

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

VERSION = "1.1.0"

# 加载匹配数据
SHEET2_D = []  # Sheet2的D列
SHEET_G = {}   # Sheet3+的G列，key是sheet名

def load_match_data():
    """加载匹配数据"""
    global SHEET2_D, SHEET_G
    
    local_file = os.path.join(os.path.dirname(__file__), 'templates', '工商库.xlsx')
    if not os.path.exists(local_file):
        print("本地文件不存在")
        return
    
    try:
        xl = pd.ExcelFile(local_file)
        print(f"工作表: {xl.sheet_names[:10]}")
        
        # Sheet2的D列
        if 'Sheet2' in xl.sheet_names:
            df2 = pd.read_excel(local_file, sheet_name='Sheet2')
            if 'D' in df2.columns:
                SHEET2_D = df2['D'].dropna().astype(str).tolist()
                SHEET2_D = [x.strip() for x in SHEET2_D if x.strip()]
                print(f"Sheet2 D列: {len(SHEET2_D)} 条")
        
        # Sheet3及以后，每个sheet的G列
        for sheet in xl.sheet_names:
            if sheet.startswith('Sheet') and sheet != 'Sheet2':
                try:
                    df = pd.read_excel(local_file, sheet_name=sheet)
                    if 'G' in df.columns:
                        g_data = df['G'].dropna().astype(str).tolist()
                        g_data = [x.strip() for x in g_data if x.strip()]
                        if g_data:
                            SHEET_G[sheet] = g_data
                            print(f"{sheet} G列: {len(g_data)} 条")
                except:
                    pass
        
        print(f"总Sheet数: {len(SHEET_G)}")
    
    except Exception as e:
        print(f"加载失败: {e}")

# 启动时加载数据
load_match_data()

def parse_user_fields(filepath):
    """解析用户上传的文件"""
    ext = os.path.splitext(filepath)[1].lower()
    fields = []
    
    if ext in ['.xlsx', '.xls']:
        df = pd.read_excel(filepath)
        # 假设第一列是字段名
        fields = df.iloc[1:, 0].dropna().astype(str).tolist()
        fields = [x.strip() for x in fields if x.strip()]
    elif ext in ['.txt']:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                # 清理
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
                    for sep in ['、', '，', ',']:
                        if sep in line:
                            fields.extend([p.strip() for p in line.split(sep) if p.strip()])
                            break
                    else:
                        if line.strip():
                            fields.append(line.strip())
    
    return fields

def find_match(user_field):
    """匹配字段"""
    user_field = str(user_field).strip()
    if not user_field:
        return None
    
    user_clean = user_field.replace("信息", "").replace("数据", "").replace("记录", "").replace(" ", "")
    
    # 1. 先匹配Sheet2的D列
    for target in SHEET2_D:
        target = str(target).strip()
        if not target:
            continue
        
        # 精确匹配
        if user_field == target:
            return {'matched': target, 'source': 'Sheet2-D', 'type': '完全匹配', 'score': 100}
        
        # 语义相似
        try:
            sim1 = Levenshtein.ratio(user_field, target)
            sim2 = Levenshtein.ratio(user_clean, target.replace("信息", "").replace("数据", ""))
            sim = max(sim1, sim2)
            if sim >= 0.7:
                return {'matched': target, 'source': 'Sheet2-D', 'type': '推荐', 'score': int(sim*100)}
        except:
            pass
    
    # 2. 再匹配Sheet3+的G列
    for sheet_name, g_data in SHEET_G.items():
        for target in g_data:
            target = str(target).strip()
            if not target:
                continue
            
            if user_field == target:
                return {'matched': target, 'source': f'{sheet_name}-G', 'type': '完全匹配', 'score': 100}
            
            try:
                sim1 = Levenshtein.ratio(user_field, target)
                sim2 = Levenshtein.ratio(user_clean, target.replace("信息", "").replace("数据", ""))
                sim = max(sim1, sim2)
                if sim >= 0.7:
                    return {'matched': target, 'source': f'{sheet_name}-G', 'type': '推荐', 'score': int(sim*100)}
            except:
                pass
    
    return None

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
        return send_file(output, download_name='模板.xlsx', as_attachment=True)
    elif type == 'txt':
        content = "1、公司概况：基本信息、联系方式、变更记录、主要人员；\n2、股东信息：股东信息、对外投资；"
        output = io.BytesIO(content.encode('utf-8'))
        return send_file(output, download_name='模板.txt', as_attachment=True)
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
                results.append({
                    'user_field': field,
                    'matched': result['matched'],
                    'source': result['source'],
                    'match_type': result['type'],
                    'score': result['score']
                })
            else:
                results.append({
                    'user_field': field,
                    'matched': '-',
                    'source': '-',
                    'match_type': '匹配失���',
                    'score': 0
                })
        
        # 统计
        total = len(results)
        exact = len([r for r in results if r['match_type'] == '完全匹配'])
        recommend = len([r for r in results if r['match_type'] == '推荐'])
        failed = len([r for r in results if r['match_type'] == '匹配失败'])
        
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
            'stats': {'total': total, 'exact': exact, 'recommend': recommend, 'failed': failed},
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