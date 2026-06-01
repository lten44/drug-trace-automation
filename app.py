import os
import sys
import json
import webview
from flask import Flask, render_template, request, jsonify
from threading import Thread
import time

# 添加原项目路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'drug_trace_automation'))

from modules.excel_reader import read_sales_excel
from modules.excel_writer import write_result_excel
from modules.code_query_ui import CodeQueryUI

app = Flask(__name__)

# 配置
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(__file__), 'output')
CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'config.json')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# 全局状态
processing_status = {
    'running': False,
    'current': 0,
    'total': 0,
    'message': ''
}


def load_config():
    """加载配置"""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'output_path': OUTPUT_FOLDER}


def save_config(config):
    """保存配置"""
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/get_config')
def get_config():
    """获取配置"""
    config = load_config()
    return jsonify(config)


@app.route('/select_output_path', methods=['POST'])
def select_output_path():
    """弹出文件夹选择对话框"""
    try:
        result = webview.windows[0].create_file_dialog(
            webview.FOLDER_DIALOG,
            directory=load_config().get('output_path', OUTPUT_FOLDER)
        )
        if result:
            path = result[0] if isinstance(result, list) else result
            config = load_config()
            config['output_path'] = path
            save_config(config)
            return jsonify({'success': True, 'path': path})
    except Exception as e:
        print(f"选择文件夹失败: {e}")
    return jsonify({'success': False, 'path': None})


@app.route('/select_input_file', methods=['POST'])
def select_input_file():
    """弹出文件选择对话框"""
    try:
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=('Excel Files (*.xlsx)',)
        )
        if result:
            path = result[0] if isinstance(result, list) else result
            return jsonify({'success': True, 'path': path})
    except Exception as e:
        print(f"选择文件失败: {e}")
    return jsonify({'success': False, 'path': None})


@app.route('/start_process', methods=['POST'])
def start_process():
    """开始处理"""
    global processing_status
    
    data = request.json
    input_file = data.get('input_file')
    output_path = data.get('output_path', OUTPUT_FOLDER)
    client_title = data.get('client_title', '码上放心')
    
    if not input_file or not os.path.exists(input_file):
        return jsonify({'success': False, 'error': '请选择有效的Excel文件'})
    
    if processing_status['running']:
        return jsonify({'success': False, 'error': '正在处理中，请等待'})
    
    # 在新线程中处理
    def process():
        global processing_status
        processing_status['running'] = True
        processing_status['current'] = 0
        processing_status['total'] = 0
        processing_status['message'] = '正在读取Excel文件...'
        
        try:
            # 读取Excel
            data = read_sales_excel(input_file)
            if not data.records:
                processing_status['message'] = '错误：未找到追溯码记录'
                processing_status['running'] = False
                return
            
            trace_codes = [r.trace_code for r in data.records]
            total = len(trace_codes)
            processing_status['total'] = total
            processing_status['message'] = f'共{total}条追溯码，开始查询...'
            
            # 使用UI自动化查询
            try:
                ui = CodeQueryUI()
                
                def progress_callback(idx, total, code, result):
                    processing_status['current'] = idx
                    if result:
                        if result.get('success'):
                            n = len(result.get('level_one_codes', []))
                            processing_status['message'] = f'正在处理: {idx}/{total} - 获取{n}个1级码'
                        else:
                            processing_status['message'] = f'正在处理: {idx}/{total} - 查询失败'
                    else:
                        processing_status['message'] = f'正在处理: {idx}/{total}'
                
                level_one_map = ui.batch_query(trace_codes, callback=progress_callback)
                
                # 写入结果
                processing_status['message'] = '正在生成Excel文件...'
                output_file = write_result_excel(data, level_one_map, output_path)
                
                processing_status['message'] = f'完成！文件已保存到: {os.path.basename(output_file)}'
                
            except Exception as e:
                processing_status['message'] = f'错误: {str(e)}'
                
        except Exception as e:
            processing_status['message'] = f'错误: {str(e)}'
        finally:
            processing_status['running'] = False
    
    thread = Thread(target=process)
    thread.start()
    
    return jsonify({'success': True})


@app.route('/get_status')
def get_status():
    """获取处理状态"""
    return jsonify(processing_status)


def start_server():
    """启动Flask服务器"""
    app.run(host='127.0.0.1', port=5000, debug=False, threaded=True)


def main():
    # 启动Flask服务器（在后台线程）
    server_thread = Thread(target=start_server, daemon=True)
    server_thread.start()
    
    # 等待服务器启动
    time.sleep(1)
    
    # 创建桌面窗口
    webview.create_window(
        '药品追溯码处理工具',
        'http://127.0.0.1:5000',
        width=700,
        height=600,
        resizable=False
    )
    
    webview.start()


if __name__ == '__main__':
    main()
