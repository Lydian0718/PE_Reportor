from flask import Flask, render_template, request, send_file, jsonify
import json
import socket
import webbrowser
import threading
import time
import sys
import os
from report_generator import generate_brief_report, generate_detailed_report, resource_path

template_dir = resource_path("templates")
app = Flask(__name__, template_folder=template_dir)

@app.route('/')
def index():
    try:
        return render_template('index.html')
    except Exception as e:
        return f"模板加载失败，请检查 templates 文件夹是否存在。错误详情: {str(e)}", 500


@app.route('/generate', methods=['POST'])



def generate():
    try:
        report_type = request.form.get('report_type')
        json_data_str = request.form.get('json_data')
        
        if not json_data_str:
            return jsonify({"error": "JSON 数据不能为空。"}), 400
            
        data = json.loads(json_data_str)
        company_name = data.get("company_short_name", "项目").replace("/", "-").replace("\\", "-")
        
        if report_type == 'brief':
            file_stream = generate_brief_report(data)
            filename = f"{company_name}项目研判报告(简要版).docx"
        else:
            file_stream = generate_detailed_report(data)
            filename = f"{company_name}项目分析报告(详细版).docx"
            
        return send_file(
            file_stream,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except json.JSONDecodeError:
        return jsonify({"error": "JSON 格式错误，请检查输入。"}), 400
    except Exception as e:
        return jsonify({"error": f"发生意外错误: {str(e)}"}), 500

def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]

if __name__ == '__main__':
    port = find_free_port()
    url = f"http://127.0.0.1:{port}"
    
    # 修复 debug=True 时打开两次浏览器的问题
    # 只有当不在 Werkzeug 重载器进程中时才打开浏览器
    if os.environ.get('WERKZEUG_RUN_MAIN') != 'true':
        threading.Thread(target=lambda: (time.sleep(1.5), webbrowser.open(url)), daemon=True).start()
    
    # 运行 Flask
    print(f" * 正在启动服务器: {url}")
    app.run(port=port, debug=True)