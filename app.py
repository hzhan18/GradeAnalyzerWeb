import os
import time  # Optional delay for troubleshooting
import logging
from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify, session
from werkzeug.utils import secure_filename
from data_processing import run_report_generation
from models import db, User
from flask_sqlalchemy import SQLAlchemy
from celery import Celery
from report_generation import generate_report_task  # 确保任务导入正确


# Initialize Flask app
app = Flask(__name__)
app.secret_key = 'your_secret_key'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['GENERATED_REPORTS'] = 'generated_reports'

# Configure SQLAlchemy with PostgreSQL
uri = os.getenv('DATABASE_URL')
if uri and uri.startswith("postgres://"):
    uri = uri.replace("postgres://", "postgresql://", 1)
app.config['SQLALCHEMY_DATABASE_URI'] = uri
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Initialize Celery
def make_celery(app):
    celery = Celery(
        app.import_name,
        backend=os.getenv("REDISCLOUD_URL"),
        broker=os.getenv("REDISCLOUD_URL")
    )
    celery.conf.update(app.config)
    celery.conf.imports = ["report_generation"]  # 添加 Celery 任务的导入路径
    return celery

# Update the Celery configuration in app.py
app.config.update(
    CELERY_BROKER_URL=os.getenv("REDISCLOUD_URL"),
    CELERY_RESULT_BACKEND=os.getenv("REDISCLOUD_URL")
)

celery = make_celery(app)

db.init_app(app)

# Ensure the upload and report folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_REPORTS'], exist_ok=True)

# Set up logging
logging.basicConfig(level=logging.INFO)

# Home route (for file upload form)
@app.route('/')
def index():
    return render_template('index.html')

# Process file upload and report generation
@app.route('/process', methods=['POST'])
def process_file():
    if 'file' not in request.files:
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        return redirect(url_for('index'))

    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)

    class_name1 = request.form.get('class_name1', '')
    class_name2 = request.form.get('class_name2', '')
    report_style = session.get('report_style', 'formal')

    session_id = os.urandom(8).hex()
    session[session_id] = {"progress": 0, "status": "Initializing"}

    task = generate_report_task.delay(file_path, class_name1, class_name2, session_id, report_style)
    session[session_id]["task_id"] = task.id
    
    return jsonify({"status": "processing", "session_id": session_id})

# Check progress of report generation
@app.route('/progress/<session_id>')
def check_progress(session_id):
    task_id = session.get(session_id, {}).get("task_id")
    task = celery.AsyncResult(task_id)
    if task.state == "SUCCESS":
        return jsonify({"progress": 100, "status": "completed", "download_url": url_for('download_report', session_id=session_id)})
    elif task.state == "PENDING":
        return jsonify({"progress": 0, "status": "pending"})
    elif task.state == "PROGRESS":
        return jsonify({"progress": task.info.get('progress', 0), "status": "in_progress"})
    else:
        return jsonify({"progress": 0, "status": "failed"})

# Download the generated report
@app.route('/download/<session_id>')
def download_report(session_id):
    report_path = session.get(session_id, {}).get("report_path")
    logging.info(f"Attempting to download report from: {report_path}")
    time.sleep(0.5)

    if report_path and os.path.isfile(report_path):
        return send_file(report_path, as_attachment=True)
    else:
        logging.error("File not found at the specified path: %s", report_path)
        return "File not found", 404

@app.route('/register', methods=['POST'])
def register():
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')
    email = data.get('email')

    if User.query.filter_by(username=username).first():
        return jsonify({"status": "fail", "message": "用户名已存在"}), 400
    if User.query.filter_by(email=email).first():
        return jsonify({"status": "fail", "message": "邮箱已被注册"}), 400

    new_user = User(username=username, email=email)
    new_user.set_password(password)
    db.session.add(new_user)
    db.session.commit()

    session['user_id'] = new_user.id
    session['username'] = new_user.username
    session['email'] = new_user.email
    session['logged_in'] = True

    return jsonify({"status": "success", "message": "注册成功并已自动登录", "username": new_user.username, "email": new_user.email}), 201

@app.route('/login', methods=['POST'])
def login():
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')

    user = User.query.filter_by(username=username).first()
    if user and user.check_password(password):
        session['user_id'] = user.id
        session['username'] = user.username
        session['email'] = user.email
        session['logged_in'] = True
        return jsonify({"status": "success", "username": user.username, "email": user.email})
    else:
        return jsonify({"status": "fail", "message": "用户名或密码错误"}), 401

@app.route('/save_style', methods=['POST'])
def save_style():
    if 'user_id' not in session:
        logging.warning("用户未登录，无法保存报告风格")
        return jsonify({"status": "fail", "message": "用户未登录"}), 401

    data = request.get_json()
    report_style = data.get('report_style')
    logging.info(f"接收到的报告风格: {report_style}")

    user = User.query.get(session['user_id'])
    if user:
        user.report_style = report_style
        db.session.commit()
        logging.info(f"报告风格已更新为: {user.report_style}")
        return jsonify({"status": "success", "message": "报告风格已保存"})
    else:
        logging.error("用户不存在，无法保存报告风格")
        return jsonify({"status": "fail", "message": "用户不存在"}), 404

@app.route('/check_login')
def check_login():
    return jsonify({"logged_in": session.get('logged_in', False)})

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))


if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
