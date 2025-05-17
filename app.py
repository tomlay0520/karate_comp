import os
from flask import Flask, render_template, request, redirect, flash
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
from werkzeug.utils import secure_filename
import webbrowser
from datetime import datetime
import sqlite3

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # 生产环境中请替换为安全的密钥
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
        # 表已手动创建，仅绑定模型
        # db.create_all()  # 已注释，避免表创建

# 在应用启动前调用 init_db
init_db()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/school')
def school():
    groups = categorize_players()
    return render_template('school.html', groups=groups)

@app.route('/ibok')
def ibok():
    return render_template('ibko.html')

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
                required_columns = ['姓名', '性别', '出生日期', '组别', '项目', '紧急联系电话']
                if not all(col in df.columns for col in required_columns):
                    flash('Excel文件缺少必要列: 姓名, 性别, 出生日期, 组别, 项目, 紧急联系电话', 'error')
                    return redirect(request.url)
                
                # 清空现有数据
                AthStu.query.delete()
                AthAdult.query.delete()
                db.session.commit()
                
                # 插入数据
                for _, row in df.iterrows():
                    name = str(row['姓名']).strip() if pd.notna(row['姓名']) else '未知'
                    gender = str(row['性别']).strip() if pd.notna(row['性别']) else '未知'
                    if gender not in ['男', '女']:
                        gender = '未知'
                    birth = pd.to_datetime(row['出生日期'], errors='coerce') if pd.notna(row['出生日期']) else None
                    if birth is None:
                        flash(f'无效的出生日期格式: {row["出生日期"]}', 'error')
                        continue
                    group = str(row['组别']).strip() if pd.notna(row['组别']) else '未知'
                    program = str(row['项目']).strip() if pd.notna(row['项目']) else '未知'
                    emergency_phone = str(row['紧急联系电话']).strip() if pd.notna(row['紧急联系电话']) else '未知'
                    win_num = int(row['获奖次数']) if '获奖次数' in df.columns and pd.notna(row['获奖次数']) and str(row['获奖次数']).isdigit() else 0
                    
                    # 根据组别或年龄判断是学生还是成人
                    is_student = '小学' in group or '初中' in group or '高中' in group or '少儿' in group
                    age = (datetime.now().date() - birth.date()).days // 365
                    if not is_student and age < 18:
                        is_student = True
                    
                    if is_student:
                        school = str(row['学校']).strip() if '学校' in df.columns and pd.notna(row['学校']) else None
                        district = str(row['区']).strip() if '区' in df.columns and pd.notna(row['区']) else None
                        player = AthStu(
                            name=name,
                            gender=gender,
                            birth=birth,
                            group=group,
                            program=program,
                            school=school,
                            district=district,
                            emergency_phone_call=emergency_phone,
                            win_num=win_num
                        )
                    else:
                        dojo = str(row['道馆']).strip() if '道馆' in df.columns and pd.notna(row['道馆']) else None
                        belt = str(row['带别']).strip() if '带别' in df.columns and pd.notna(row['带别']) else None
                        player = AthAdult(
                            name=name,
                            gender=gender,
                            birth=birth,
                            group=group,
                            program=program,
                            dojo=dojo,
                            emergency_phone_call=emergency_phone,
                            belt=belt,
                            win_num=win_num
                        )
                    db.session.add(player)
                
                db.session.commit()
                flash('选手信息上传成功', 'success')
                return redirect('/school')
            except Exception as e:
                db.session.rollback()
                flash(f'文件处理错误: {str(e)}', 'error')
                return redirect(request.url)
            finally:
                if os.path.exists(file_path):
                    os.remove(file_path)
        else:
            flash('请上传.xlsx文件', 'error')
            return redirect(request.url)
    return render_template('upload.html')

@app.route('/stu_generate_matching')
def stu_generate_matching():
    groups = categorize_players()
    matches = []
    # 仅为学生（ath_stu）生成对阵
    for gender in groups.get('students', {}):
        for group in groups['students'][gender]:
            for program in groups['students'][gender][group]:
                players = groups['students'][gender][group][program]
                for i in range(0, len(players), 2):
                    if i + 1 < len(players):
                        matches.append({
                            'player1': players[i].name,
                            'player2': players[i+1].name,
                            'gender': gender,
                            'group': group,
                            'program': program
                        })
    return render_template('stu_generate_matching.html', matches=matches)

def categorize_players():
    groups = {
        'students': {'男': {}, '女': {}, '未知': {}},
        'adults': {'男': {}, '女': {}, '未知': {}}
    }
    
    # 查询学生
    students = AthStu.query.all()
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
    
    # 查询成人
    adults = AthAdult.query.all()
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
    
    # 清理空分组
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

if __name__ == '__main__':
    try:
        url = 'http://127.0.0.1:5001'
        webbrowser.open(url)
        app.run(host='127.0.0.1', port=5001, debug=True)
    except OSError as e:
        print(f"端口 5001 被占用，尝试 5002 端口...")
        url = 'http://127.0.0.1:5002'
        webbrowser.open(url)
        app.run(host='127.0.0.1', port=5002, debug=True)