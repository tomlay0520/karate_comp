import os
from flask import Flask, render_template, request, redirect, flash, url_for, jsonify
from flask_sqlalchemy import SQLAlchemy
from flask_socketio import SocketIO, emit
import pandas as pd
from werkzeug.utils import secure_filename
from datetime import datetime
import sqlite3
import threading
import time
import socket

app = Flask(__name__)

# 处理 favicon 请求
@app.route('/favicon.ico')
def favicon():
    return '', 204
app.secret_key = 'your_secret_key'
app.config['UPLOAD_FOLDER'] = 'Uploads'


current_round = 1 #比赛轮次
lost_players = [] #记录比赛输掉的选手

# 定义路径
MODEL_DIR = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'model')
DB_PATH = os.path.join(MODEL_DIR, 'ath.db')
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + DB_PATH.replace('\\', '/')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# 初始化 SocketIO
socketio = SocketIO(app, cors_allowed_origins="*")

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

class Match(db.Model):
    __tablename__ = 'matches'
    id = db.Column(db.Integer, primary_key=True)
    player1 = db.Column(db.String(100), nullable=False)
    player2 = db.Column(db.String(100), nullable=False)
    gender = db.Column(db.String(10), nullable=False)
    group_name = db.Column(db.String(50), nullable=False)
    program = db.Column(db.String(50), nullable=False)
    subgroup = db.Column(db.String(50), nullable=False)
    school1 = db.Column(db.String(100))
    school2 = db.Column(db.String(100))
    winner = db.Column(db.String(100))

# 倒计时状态
current_timer = {'remaining_seconds': 180, 'is_paused': True, 'match_active': False}
timer_thread = None

# 初始化数据库
def init_db():
    with app.app_context():
        db.create_all()
        print(f"数据库路径: {app.config['SQLALCHEMY_DATABASE_URI']}")

# 启动倒计时线程
def start_timer_thread():
    global current_timer, timer_thread
    if timer_thread and timer_thread.is_alive():
        return
    timer_thread = threading.Thread(target=run_timer)
    timer_thread.daemon = True
    timer_thread.start()

def run_timer():
    global current_timer
    while current_timer['match_active'] and current_timer['remaining_seconds'] > 0:
        if not current_timer['is_paused']:
            current_timer['remaining_seconds'] -= 1
            socketio.emit('timer_update', {
                'type': 'timer_update',
                'remainingSeconds': current_timer['remaining_seconds']
            }, namespace='/ws/match_updates')
        time.sleep(1)
    if current_timer['remaining_seconds'] <= 0:
        current_timer['match_active'] = False
        socketio.emit('timer_update', {
            'type': 'timer_update',
            'remainingSeconds': 0
        }, namespace='/ws/match_updates')

# 路由：主页
@app.route('/')
def index():
    return render_template('index.html')

# 路由：学校页面
@app.route('/school')
def school():
    groups = categorize_players()
    return render_template('school.html', groups=groups)

# 路由：IBKO页面
@app.route('/ibko')
def ibok():
    return render_template('ibko.html')

# 路由：观众页面
@app.route('/spectator')
def spectator():
    return render_template('spectator_display.html')

# 路由：比赛对阵页面
@app.route('/stu_generate_matching')
def stu_generate_matching():
    global current_round

    all_players = AthStu.query.filter(AthStu.name.notin_(lost_players)).all()

    groups = categorize_players()
    matches = []
    
    for gender in groups.get('students', {}):
        for group in groups['students'][gender]:
            for program in groups['students'][gender][group]:
                for subgroup in groups['students'][gender][group][program]:
                    players = groups['students'][gender][group][program][subgroup]
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
    print("生成的对阵数据:", matches)
    return render_template('stu_generate_matching.html', matches=matches)

# 路由：选择胜利者
@app.route('/select_winner', methods=['POST'])
def select_winner():
    data = request.get_json()
    winner = data.get('winner')
    school = data.get('school')
    
    # 广播胜利信息
    socketio.emit('winner', {
        'type': 'winner',
        'winner': winner,
        'school': school
    }, namespace='/ws/match_updates')
    
    return jsonify({'status': 'success'})

# 路由：上传选手信息
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
                
                name_counts = {}
                name_counter = {}
                for name in df['姓名']:
                    name = str(name).strip()
                    name_counts[name] = name_counts.get(name, 0) + 1
                
                for _, row in df.iterrows():
                    name = str(row['姓名']).strip()
                    gender = str(row['性别']).strip()
                    if gender not in ['男', '女']:
                        raise ValueError("性别必须是'男'或'女'")
                    
                    try:
                        birth_date = pd.to_datetime(row['出生日期'], errors='coerce')
                        if pd.isna(birth_date):
                            birth_date = datetime.strptime(str(row['出生日期']).split()[0], '%Y-%m-%d')
                        birth_date = birth_date.date()
                    except:
                        birth_date = datetime.now().date()
                    
                    if name_counts[name] > 1:
                        name_counter[name] = name_counter.get(name, 0) + 1
                        name = f"{name}{name_counter[name]}"
                    
                    player = AthStu(
                        name=name,
                        gender=gender,
                        birth=birth_date,
                        group=str(row['组别']).strip(),
                        program=str(row['项目']).strip(),
                        school=str(row['所属学校']).strip() if '所属学校' in row and pd.notna(row['所属学校']) else None,
                        district=str(row['所属区']).strip() if '所属区' in row and pd.notna(row['所属区']) else None,
                        emergency_phone_call=str(row['紧急联系人']).strip(),
                        win_num=0
                    )
                    db.session.add(player)
                
                db.session.commit()
                flash('数据导入成功！', 'success')
                return redirect('/school')
                
            except Exception as e:
                db.session.rollback()
                flash(f'导入失败: {str(e)}', 'error')
                return redirect(request.url)
            finally:
                if os.path.exists(file_path):
                    os.remove(file_path)
        else:
            flash('请上传有效的Excel文件(.xlsx)', 'error')
            return redirect(request.url)
    return render_template('upload.html')

# 路由：获取选手详情
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
        'birth': player.birth.isoformat(),
        'group': player.group,
        'program': player.program,
        'school': player.school,
        'district': player.district,
        'emergency_phone_call': player.emergency_phone_call,
        'win_num': player.win_num
    })

# 路由：更新比赛信息（不启动倒计时），仅支持 POST 方法
@app.route('/update_match', methods=['POST'])
def update_match():
    data = request.get_json()
    
    # 保存比赛到数据库
    match = Match(
        player1=data['player1'],
        player2=data['player2'],
        gender=data['gender'],
        group_name=data['group'],
        program=data['program'],
        subgroup=data['subgroup'],
        school1=data['school1'],
        school2=data['school2']
    )
    db.session.add(match)
    db.session.commit()
    
    # 广播比赛信息
    socketio.emit('match_update', data, namespace='/ws/match_updates')
    
    return jsonify({'status': 'success'})

# 路由：开始比赛（启动倒计时）
@app.route('/start_match', methods=['POST'])
def start_match():
    global current_timer
    data = request.get_json()
    current_timer = {
        'remaining_seconds': data['remainingSeconds'],
        'is_paused': False,
        'match_active': True
    }
    
    # 启动倒计时
    start_timer_thread()
    return jsonify({'status': 'success'})

# 路由：更新倒计时状态
@app.route('/update_timer', methods=['POST'])
def update_timer():
    data = request.get_json()
    global current_timer
    if data['type'] == 'timer_pause':
        current_timer['is_paused'] = True
        current_timer['remaining_seconds'] = data['remainingSeconds']
        socketio.emit('timer_pause', {
            'type': 'timer_pause',
            'remainingSeconds': current_timer['remaining_seconds']
        }, namespace='/ws/match_updates')
    elif data['type'] == 'timer_resume':
        current_timer['is_paused'] = False
        socketio.emit('timer_resume', {
            'type': 'timer_resume',
            'remainingSeconds': current_timer['remaining_seconds']
        }, namespace='/ws/match_updates')
        start_timer_thread()
    elif data['type'] == 'timer_update':
        current_timer['remaining_seconds'] = data['remainingSeconds']
        socketio.emit('timer_update', {
            'type': 'timer_update',
            'remainingSeconds': current_timer['remaining_seconds']
        }, namespace='/ws/match_updates')
    return jsonify({'status': 'success'})

# 路由：暂停比赛
@app.route('/pause_match', methods=['POST'])
def pause_match():
    global current_timer
    data = request.get_json()
    current_timer['remaining_seconds'] = data['remainingSeconds']
    current_timer['is_paused'] = True
    socketio.emit('timer_pause', {
        'type': 'timer_pause',
        'remainingSeconds': current_timer['remaining_seconds']
    }, namespace='/ws/match_updates')
    return jsonify({'status': 'success'})

# 路由：恢复比赛
@app.route('/resume_match', methods=['POST'])
def resume_match():
    global current_timer
    data = request.get_json()
    current_timer['remaining_seconds'] = data['remainingSeconds']
    current_timer['is_paused'] = False
    socketio.emit('timer_resume', {
        'type': 'timer_resume',
        'remainingSeconds': current_timer['remaining_seconds']
    }, namespace='/ws/match_updates')
    start_timer_thread()
    return jsonify({'status': 'success'})

# 路由：更新获胜者
@app.route('/update_winner', methods=['POST'])
def update_winner():
    data = request.get_json()
    winner_name = data['winner']
    global current_timer
    current_timer['is_paused'] = True
    current_timer['match_active'] = False
    
    # 更新选手胜场数
    student = AthStu.query.filter_by(name=winner_name).first()
    if student:
        student.win_num += 1
    else:
        adult = AthAdult.query.filter_by(name=winner_name).first()
        if adult:
            adult.win_num += 1
    
    # 更新比赛记录
    match = Match.query.filter_by(
        player1=data['player1'],
        player2=data['player2'],
        group_name=data['group'],
        program=data['program'],
        subgroup=data['subgroup']
    ).first()
    if match:
        match.winner = winner_name
    db.session.commit()
    
    # 广播获胜者信息
    socketio.emit('winner', {
        'type': 'winner',
        'winner': winner_name,
        'school': data['school']
    }, namespace='/ws/match_updates')
    
    return jsonify({'status': 'success'})

# 分组选手
def categorize_players():
    groups = {
        'students': {'男': {}, '女': {}, '未知': {}},
        'adults': {'男': {}, '女': {}, '未知': {}}
    }
    
    students = AthStu.query.order_by(AthStu.win_num.desc()).all()
    for player in students:
        gender = player.gender or '未知'
        group = player.gender or '其它'
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

# 查找可用端口
def find_available_port(start_port=5000, max_attempts=10):
    for port in range(start_port, start_port + max_attempts):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(('127.0.0.1', port))
                return port
        except OSError:
            continue
    raise OSError(f"无法在端口{start_port}-{start_port+max_attempts-1}范围内找到可用端口")
@app.route('/test')
def test():
    return render_template('test.html')
# WebSocket 连接处理
@socketio.on('connect', namespace='/ws/match_updates')
def handle_connect():
    print('客户端已连接')


@socketio.on('disconnect', namespace='/ws/match_updates')
def handle_disconnect():
    print('客户端已断开')

if __name__ == '__main__':
    with app.app_context(): 
        init_db()
    port = find_available_port()
    socketio.run(app, host='0.0.0.0', port=port, debug=True)