#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地 HTTP 服务器 - 周会数据看板实时更新
启动后访问: http://localhost:8765/周会数据看板.html

功能:
- 提供静态文件服务（HTML、JSON 等）
- /sync 端点：从腾讯文档拉取最新数据并重新生成看板
- /api/data 端点：返回最新 tdoc_data.json
"""

import http.server
import json
import os
import subprocess
import sys
import io
import urllib.parse

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PORT = 8765


class DashboardHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', '*')
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == '/sync':
            self.handle_sync()
        elif path == '/api/data':
            self.handle_api_data()
        elif path == '/api/status':
            self.handle_api_status()
        elif path == '/' or path == '':
            self.send_response(302)
            self.send_header('Location', '/index.html')
            self.end_headers()
        else:
            super().do_GET()

    def handle_sync(self):
        """从腾讯文档同步最新数据并重新生成看板"""
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()

        result = {'success': False, 'steps': []}

        # Step 1: 运行 tdoc_sync.py
        tdoc_sync = os.path.join(BASE_DIR, 'tdoc_sync.py')
        try:
            proc = subprocess.run(
                [sys.executable, tdoc_sync],
                capture_output=True, text=True, timeout=120, cwd=BASE_DIR,
                encoding='utf-8', errors='replace'
            )
            result['steps'].append({
                'step': 'tdoc_sync',
                'success': proc.returncode == 0,
                'output': proc.stdout[-500:] if proc.stdout else '',
                'error': proc.stderr[-500:] if proc.stderr else ''
            })
        except Exception as e:
            result['steps'].append({'step': 'tdoc_sync', 'success': False, 'error': str(e)})

        # Step 2: 运行 build_standalone.py
        build_script = os.path.join(BASE_DIR, 'build_standalone.py')
        try:
            proc = subprocess.run(
                [sys.executable, build_script],
                capture_output=True, text=True, timeout=60, cwd=BASE_DIR,
                encoding='utf-8', errors='replace'
            )
            result['steps'].append({
                'step': 'build_standalone',
                'success': proc.returncode == 0,
                'output': proc.stdout[-500:] if proc.stdout else '',
                'error': proc.stderr[-500:] if proc.stderr else ''
            })
        except Exception as e:
            result['steps'].append({'step': 'build_standalone', 'success': False, 'error': str(e)})

        result['success'] = all(s['success'] for s in result['steps'])

        # 更新数据时间
        data_file = os.path.join(BASE_DIR, 'tdoc_data.json')
        if os.path.exists(data_file):
            try:
                with open(data_file, 'r', encoding='utf-8') as f:
                    jd = json.load(f)
                result['data_time'] = jd.get('generated_at', '')
                result['sheets'] = list(jd.get('sheets', {}).keys())
            except:
                pass

        self.wfile.write(json.dumps(result, ensure_ascii=False).encode('utf-8'))

    def handle_api_data(self):
        """返回最新的 tdoc_data.json"""
        data_file = os.path.join(BASE_DIR, 'tdoc_data.json')
        if os.path.exists(data_file):
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            with open(data_file, 'rb') as f:
                self.wfile.write(f.read())
        else:
            self.send_response(404)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'error': 'tdoc_data.json not found'}).encode('utf-8'))

    def handle_api_status(self):
        """返回服务器和数据状态"""
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()

        status = {'server': 'running', 'port': PORT}
        data_file = os.path.join(BASE_DIR, 'tdoc_data.json')
        if os.path.exists(data_file):
            try:
                with open(data_file, 'r', encoding='utf-8') as f:
                    jd = json.load(f)
                status['data_source'] = jd.get('source', '')
                status['data_time'] = jd.get('generated_at', '')
                status['sheets_count'] = len(jd.get('sheets', {}))
                status['file_size'] = os.path.getsize(data_file)
            except Exception as e:
                status['error'] = f'Cannot parse: {e}'
        else:
            status['error'] = 'tdoc_data.json not found'

        html_file = os.path.join(BASE_DIR, 'index.html')
        if os.path.exists(html_file):
            status['html_size'] = os.path.getsize(html_file)

        self.wfile.write(json.dumps(status, ensure_ascii=False).encode('utf-8'))

    def log_message(self, format, *args):
        # 简化日志
        if '/api/' in str(args) or '/sync' in str(args):
            print(f"[{self.log_date_time_string()}] {args[0]}")


def main():
    print("=" * 60)
    print("周会数据看板 - 本地服务器")
    print("=" * 60)
    print(f"目录: {BASE_DIR}")
    print(f"地址: http://localhost:{PORT}/周会数据看板.html")
    print(f"同步: http://localhost:{PORT}/sync")
    print(f"数据: http://localhost:{PORT}/api/data")
    print(f"状态: http://localhost:{PORT}/api/status")
    print()
    print("在浏览器中打开上述地址即可使用。")
    print("点击看板中的「从腾讯文档同步」按钮可实时拉取最新数据。")
    print("按 Ctrl+C 停止服务器。")
    print("=" * 60)

    server = http.server.HTTPServer(('0.0.0.0', PORT), DashboardHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务器已停止。")
        server.server_close()


if __name__ == '__main__':
    main()