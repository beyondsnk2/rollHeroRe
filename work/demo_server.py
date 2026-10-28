# -*- coding: utf-8 -*-
"""本地 demo 服务器：
- GET /            托管 web/index.html（no-store，保证刷新拿到最新）
- GET /api/refresh 重新运行 work/export_map_config.py 导出最新配置，返回 JSON
由 start_map_demo.bat 启动；端口被占用时本实例直接退出（说明已在运行）。"""
import json
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PORT = 8765


class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/api/refresh':
            return self.refresh()
        if self.path in ('/', '/index.html'):
            data = (ROOT / 'web' / 'index.html').read_bytes()
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Cache-Control', 'no-store')
            self.end_headers()
            self.wfile.write(data)
        else:
            self.send_error(404)

    def refresh(self):
        ok, out = False, ''
        try:
            r = subprocess.run(
                [sys.executable, str(ROOT / 'work' / 'export_map_config.py')],
                capture_output=True, text=True, timeout=120)
            out = (r.stdout + r.stderr).strip()
            ok = r.returncode == 0
        except Exception as e:  # noqa: BLE001
            out = repr(e)
        body = json.dumps({'ok': ok, 'log': out}, ensure_ascii=False).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args):  # 静默访问日志
        pass


def main():
    try:
        srv = HTTPServer(('127.0.0.1', PORT), Handler)
    except OSError:
        print(f'端口 {PORT} 已被占用（服务器可能已在运行），直接退出。')
        return
    threading.Timer(0.8, webbrowser.open, [f'http://127.0.0.1:{PORT}/']).start()
    print(f'demo 服务器运行中：http://127.0.0.1:{PORT}/（Ctrl+C 退出）')
    srv.serve_forever()


if __name__ == '__main__':
    main()
