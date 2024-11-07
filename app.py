import os
import time  # Optional delay for troubleshooting
import logging
from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify, session
from werkzeug.utils import secure_filename
from data_processing import run_report_generation
from models import db, User  # 引入数据库和用户模型

# Initialize Flask app
app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Set a secret key for session management
app.config['UPLOAD_FOLDER'] = 'uploads'  # Directory for file uploads
app.config['GENERATED_REPORTS'] = 'generated_reports'  # Directory for generated reports
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///user.db'  # 设置数据库路径
db.init_app(app)  # 初始化数据库


# Ensure the upload and report folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_REPORTS'], exist_ok=True)

# Set up logging
logging.basicConfig(level=logging.INFO)

# Home route (for file upload form)
@app.route('/')
def index():
    return render_template('index.html')  # Render the home page template

# Process file upload and report generation
@app.route('/process', methods=['POST'])
def process_file():
    # Check if the file is in the request
    if 'file' not in request.files:
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        return redirect(url_for('index'))

    # Save the uploaded file
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)

    # Get class names from the form
    class_name1 = request.form.get('class_name1', '')
    class_name2 = request.form.get('class_name2', '')

    # Generate a session ID to track progress
    session_id = os.urandom(8).hex()
    session[session_id] = {"progress": 0, "status": "Initializing"}

    # Run the report generation process
    result = run_report_generation(file_path, class_name1, class_name2, session_id)

    # Error handling
    if result["status"] == "error":
        logging.error("Error during report generation.")
        return jsonify(result), 400

    # Use the path directly from result to avoid adding "uploads" twice
    report_path = result["report_path"]
    session[session_id]["report_path"] = report_path

    # Check if the file exists and list directory contents for debugging
    if os.path.isfile(report_path):
        logging.info(f"Report generated and stored at: {report_path}")
    else:
        logging.error("Report file not found immediately after generation: %s", report_path)
        logging.info("Listing contents of the directory for verification:")
        logging.info(os.listdir(os.path.dirname(report_path)))
        return jsonify({"status": "error", "message": "Report generation failed"}), 500

    return jsonify({"status": "success", "session_id": session_id})

# Check progress of report generation
@app.route('/progress/<session_id>')
def check_progress(session_id):
    progress_info = session.get(session_id, {"progress": 0, "status": "No progress data"})
    return jsonify(progress_info)

# Download the generated report
@app.route('/download/<session_id>')
def download_report(session_id):
    report_path = session.get(session_id, {}).get("report_path")
    logging.info(f"Attempting to download report from: {report_path}")

    # Optional delay for troubleshooting
    time.sleep(0.5)  # Half-second delay to ensure file is accessible

    # Ensure the path exists and is correct
    if report_path and os.path.isfile(report_path):
        return send_file(report_path, as_attachment=True)
    else:
        logging.error("File not found at the specified path: %s", report_path)
        return "File not found", 404



# 注册路由
@app.route('/register', methods=['POST'])
def register():
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')
    email = data.get('email')

    # 检查用户名和邮箱是否已存在
    if User.query.filter_by(username=username).first():
        return jsonify({"status": "fail", "message": "用户名已存在"}), 400
    if User.query.filter_by(email=email).first():
        return jsonify({"status": "fail", "message": "邮箱已被注册"}), 400

    # 创建用户并加密密码
    new_user = User(username=username, email=email)
    new_user.set_password(password)  # 加密密码
    db.session.add(new_user)
    db.session.commit()

    return jsonify({"status": "success", "message": "注册成功"}), 201
#登陆路由
@app.route('/login', methods=['POST'])
def login():
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')

    # 查询用户
    user = User.query.filter_by(username=username).first()
    if user and user.check_password(password):
        session['user_id'] = user.id  # 将用户ID存入session
        session['username'] = user.username
        session['email'] = user.email
        # 返回 JSON 响应，不跳转页面
        return jsonify({"status": "success", "username": user.username, "email": user.email})
    else:
        return jsonify({"status": "fail", "message": "用户名或密码错误"}), 401

@app.route('/save_style', methods=['POST'])
def save_style():
    if 'user_id' not in session:
        return jsonify({"status": "fail", "message": "用户未登录"}), 401

    data = request.get_json()
    report_style = data.get('report_style')

    user = User.query.get(session['user_id'])
    if user:
        user.report_style = report_style  # 保存报告风格
        db.session.commit()
        return jsonify({"status": "success", "message": "报告风格已保存"})
    else:
        return jsonify({"status": "fail", "message": "用户不存在"}), 404


# 登出路由
@app.route('/logout')
def logout():
    session.clear() 
    return redirect(url_for('index'))


if __name__ == "__main__":
    with app.app_context():
        db.create_all()  # 创建数据库表
    app.run(debug=True)