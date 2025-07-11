import os
from flask import Flask, render_template, request, redirect, flash, redirect, url_for, jsonify
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
from werkzeug.utils import secure_filename
import webbrowser
from datetime import datetime
import sqlite3

app = Flask(__name__)
app.secret_key = 'your_secret_key' 
app.config['UPLOAD_FOLDER'] = 'Uploads'

# 添加以下两行定义
MODEL_DIR = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'model')
DB_PATH = os.path.join(MODEL_DIR, 'ath.db')

app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + DB_PATH.replace('\\', '/')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# 确保目录存在
try:
    os.makedirs(MODEL_DIR, exist_ok=True)
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    print(f"创建/验证目录: {MODEL_DIR}, {app.config['UPLOAD_FOLDER']}")
except OSError as e:
    print(f"无法创建目录: {e}")
    raise

# 测试 SQLite 连接
try:
    conn = sqlite3.connect(DB_PATH)
    print(f"成功连接到数据库: {DB_PATH}")
    conn.close()
except sqlite3.OperationalError as e:
    print(f"无法连接到数据库: {e}")
    raise

db = SQLAlchemy(app)

# 数据库模型
class AthStu(db.Model):
    __tablename__ = 'ath_stu'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    gender = db.Column(db.String(10), nullable=False)
    birth = db.Column(db.Date, nullable=False)
    group = db.Column(db.String(50), nullable=False)
    program = db.Column(db.String(50), nullable=False)
    school = db.Column(db.String(100))
    district = db.Column(db.String(50))
    emergency_phone_call = db.Column(db.String(20), nullable=False)
    win_num = db.Column(db.Integer, default=0)

class AthAdult(db.Model):
    __tablename__ = 'ath_adult'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    gender = db.Column(db.String(10), nullable=False)
    birth = db.Column(db.Date, nullable=False)
    group = db.Column(db.String(50), nullable=False)
    program = db.Column(db.String(50), nullable=False)
    dojo = db.Column(db.String(100))
    emergency_phone_call = db.Column(db.String(20), nullable=False)
    belt = db.Column(db.String(50))
    win_num = db.Column(db.Integer, default=0)

# 初始化数据库（不创建表，因为已手动创建）
def init_db():
    with app.app_context():
        print(f"数据库路径: {app.config['SQLALCHEMY_DATABASE_URI']}")

# 在应用启动前调用 init_db
init_db()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/school')
def school():
    groups = categorize_players()
    return render_template('school.html', groups=groups)

@app.route('/ibko')
def ibok():
    return render_template('ibko.html')

def map_excel_to_db(row, name_counter=None):
    #将Excel行数据映射到数据库字段
    gender = str(row['性别']).strip()
    if gender not in ['男', '女']:
        raise ValueError("性别必须是'男'或'女'")
    
    # 处理日期格式，确保返回ISO格式日期字符串
    try:
        birth_date = pd.to_datetime(row['出生日期'], errors='coerce')
        if pd.isna(birth_date):
            birth_date = datetime.strptime(str(row['出生日期']).split()[0], '%Y-%m-%d')
        # 转换为ISO格式字符串(YYYY-MM-DD)
        birth_date = birth_date.strftime('%Y-%m-%d')
    except:
        birth_date = datetime.now().strftime('%Y-%m-%d')  
    
    # 处理重名情况
    original_name = str(row['姓名']).strip()
    name = original_name
    if name_counter and name_counter.get(original_name, 0) > 0:
        name = f"{original_name}{name_counter[original_name]}"
    
    return {
        'name': name,
        'gender': gender,
        'birth': birth_date,
        'group': str(row['组别']).strip(),
        'program': str(row['项目']).strip(),
        'school': str(row['所属学校']).strip() if '所属学校' in row and pd.notna(row['所属学校']) else None,
        'district': str(row['所属区']).strip() if '所属区' in row and pd.notna(row['所属区']) else None,
        'emergency_phone_call': str(row['紧急联系人']).strip(),
        'win_num': 0
    }

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('未选择文件', 'error')
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '':
            flash('未选择文件', 'error')
            return redirect(request.url)
        
        if file and file.filename.endswith('.xlsx'):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            try:
                file.save(file_path)
                df = pd.read_excel(file_path)
                
                # 检查重名并计数
                name_counts = {}
                name_counter = {}
                for name in df['姓名']:
                    name = str(name).strip()
                    name_counts[name] = name_counts.get(name, 0) + 1
                
                # 映射并筛选需要的字段
                mapped_data = []
                for _, row in df.iterrows():
                    name = str(row['姓名']).strip()
                    if name_counts[name] > 1:
                        name_counter[name] = name_counter.get(name, 0) + 1
                        mapped_data.append(map_excel_to_db(row, name_counter))
                    else:
                        mapped_data.append(map_excel_to_db(row))
                
                # 检查数据库表是否存在
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                cursor.execute("PRAGMA table_info(ath_stu)")
                columns = [column[1] for column in cursor.fetchall()]
                
                filtered_df = pd.DataFrame(mapped_data)
                filtered_df = filtered_df[[col for col in filtered_df.columns if col in columns]]
                
                # 将数据写入ath_stu表
                filtered_df.to_sql('ath_stu', conn, if_exists='append', index=False)
                
                conn.close()
                flash('数据导入成功！', 'success')
                return redirect('/school')
                
            except Exception as e:
                flash(f'导入失败: {str(e)}', 'error')
                return redirect(request.url)
            finally:
                if os.path.exists(file_path):
                    os.remove(file_path)
        else:
            flash('请上传有效的Excel文件(.xlsx)', 'error')
            return redirect(request.url)
    return render_template('upload.html')

@app.route('/stu_generate_matching')
def stu_generate_matching():
    groups = categorize_players()
    print("分组数据:", groups)  # 添加调试信息
    matches = []
    
    # 仅为学生（ath_stu）生成对阵
    for gender in groups.get('students', {}):
        for group in groups['students'][gender]:
            for program in groups['students'][gender][group]:
                for subgroup in groups['students'][gender][group][program]:
                    players = groups['students'][gender][group][program][subgroup]
                    # 用于记录已经参与对阵的选手
                    matched_players = set()
                    
                    for i in range(len(players)):
                        if players[i].id not in matched_players:
                            for j in range(i + 1, len(players)):
                                if players[j].id not in matched_players and players[i].school != players[j].school:
                                    matches.append({
                                        'player1': players[i].name,
                                        'player2': players[j].name,
                                        'gender': gender,
                                        'group': group,
                                        'program': program,
                                        'subgroup': subgroup,
                                        'school1': players[i].school,
                                        'school2': players[j].school
                                    })
                                    matched_players.add(players[i].id)
                                    matched_players.add(players[j].id)
                                    break
    print("生成的对阵数据:", matches)  # 添加调试信息
    return render_template('stu_generate_matching.html', matches=matches)

def categorize_players():
    groups = {
        'students': {'男': {}, '女': {}, '未知': {}},
        'adults': {'男': {}, '女': {}, '未知': {}}
    }
    
    # 学生选手按胜场数排序
    students = AthStu.query.order_by(AthStu.win_num.desc()).all()
    for player in students:
        gender = player.gender or '未知'
        group = player.group or '其它'
        program = '型' if player.program and '型' in player.program else '组手'
        subgroup = '其它'
        if program == '组手' and player.program:
            if 'A' in player.program:
                subgroup = 'A'
            elif 'B' in player.program:
                subgroup = 'B'
        
        if group not in groups['students'][gender]:
            groups['students'][gender][group] = {'型': {}, '组手': {}}
        if program not in groups['students'][gender][group]:
            groups['students'][gender][group][program] = {}
        if subgroup not in groups['students'][gender][group][program]:
            groups['students'][gender][group][program][subgroup] = []
        
        groups['students'][gender][group][program][subgroup].append(player)
    
    # 成人选手按胜场数排序
    adults = AthAdult.query.order_by(AthAdult.win_num.desc()).all()
    for player in adults:
        gender = player.gender or '未知'
        group = player.group or '其它'
        program = '型' if player.program and '型' in player.program else '组手'
        subgroup = '其它'
        if program == '组手' and player.program:
            if 'A' in player.program:
                subgroup = 'A'
            elif 'B' in player.program:
                subgroup = 'B'
        
        if group not in groups['adults'][gender]:
            groups['adults'][gender][group] = {'型': {}, '组手': {}}
        if program not in groups['adults'][gender][group]:
            groups['adults'][gender][group][program] = {}
        if subgroup not in groups['adults'][gender][group][program]:
            groups['adults'][gender][group][program][subgroup] = []
        
        groups['adults'][gender][group][program][subgroup].append(player)
    
    for category in ['students', 'adults']:
        for gender in list(groups[category].keys()):
            for group in list(groups[category][gender].keys()):
                for program in list(groups[category][gender][group].keys()):
                    for subgroup in list(groups[category][gender][group][program].keys()):
                        if not groups[category][gender][group][program][subgroup]:
                            del groups[category][gender][group][program][subgroup]
                    if not groups[category][gender][group][program]:
                        del groups[category][gender][group][program]
                if not groups[category][gender][group]:
                    del groups[category][gender][group]
            if not groups[category][gender]:
                del groups[category][gender]
    
    return groups

def find_available_port(start_port=5000, max_attempts=10):
    """查找可用的端口"""
    import socket
    for port in range(start_port, start_port + max_attempts):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(('127.0.0.1', port))
                return port
        except OSError:
            continue
    raise OSError(f"无法在端口{start_port}-{start_port+max_attempts-1}范围内找到可用端口")

@app.route('/api/player/<int:player_id>')
def get_player_details(player_id):
    player = AthStu.query.get(player_id)
    if not player:
        player = AthAdult.query.get(player_id)
        if not player:
            return jsonify({'error': '选手不存在'}), 404
    
    return jsonify({
        'name': player.name,
        'gender': player.gender,
        'birth': player.birth,
        'group': player.group,
        'program': player.program,
        'school': player.school,
        'district': player.district,
        'emergency_phone_call': player.emergency_phone_call,
        'win_num': player.win_num
    })

@app.route('/update_winner', methods=['POST'])
def update_winner():
    data = request.get_json()
    winner_name = data['winner']
    
    # 更新学生表
    student = AthStu.query.filter_by(name=winner_name).first()
    if student:
        student.win_num += 1
    else:
        # 更新成人表
        adult = AthAdult.query.filter_by(name=winner_name).first()
        if adult:
            adult.win_num += 1
    
    db.session.commit()
    return jsonify({'status': 'success'})

if __name__ == '__main__':
    # main function here
    app.run(host='127.0.0.1', port=5000, debug=True)