from flask import Flask, render_template, request, redirect, url_for, flash
import sqlite3
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'

# 修改后的数据库配置
DATABASE = os.path.join('model', 'ath.db')

def init_db():
    # 确保model目录存在
    os.makedirs('model', exist_ok=True)
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS ath_stu (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            gender TEXT CHECK(gender IN ('m','f')) NOT NULL,
            birth DATE NOT NULL,
            "group" TEXT NOT NULL,
            program TEXT NOT NULL,
            school TEXT,
            district TEXT,
            emergency_phone_call TEXT NOT NULL,
            win_num INTEGER DEFAULT 0
        )
    ''')
    conn.commit()
    conn.close()

def map_excel_to_db(row):
    """将Excel行数据映射到数据库字段"""
    return {
        'name': row['姓名'],
        'gender': 'm' if row['性别'] == '男' else 'f',
        'birth': datetime.strptime(row['出生日期'], '%Y-%m-%d').date(),
        'group': row['组别'],
        'program': row['项目'],
        'school': row['所属学校'],
        'district': row['所属区'],
        'emergency_phone_call': row['紧急联系人']
    }

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('没有选择文件')
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '':
            flash('没有选择文件')
            return redirect(request.url)
        
        if file and file.filename.endswith('.xlsx'):
            try:
                # 读取Excel文件
                df = pd.read_excel(file)
                
                # 检查数据库表是否存在
                conn = sqlite3.connect(DATABASE)
                cursor = conn.cursor()
                cursor.execute("PRAGMA table_info(ath_stu)")
                columns = [column[1] for column in cursor.fetchall()]
                
                # 映射并筛选需要的字段
                mapped_data = [map_excel_to_db(row) for _, row in df.iterrows()]
                filtered_df = pd.DataFrame(mapped_data)
                
                # 确保只写入数据库表中存在的列
                filtered_df = filtered_df[[col for col in filtered_df.columns if col in columns]]
                
                # 将数据写入ath_stu表
                filtered_df.to_sql('ath_stu', conn, if_exists='append', index=False)
                
                conn.close()
                flash('数据导入成功！')
            except Exception as e:
                flash(f'导入失败: {str(e)}')
            return redirect(url_for('upload_file'))
        else:
            flash('请上传有效的Excel文件(.xlsx)')
            return redirect(request.url)
    
    return render_template('upload.html')

if __name__ == '__main__':
    init_db()
    app.run(debug=True)