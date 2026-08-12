# -*- coding: utf-8 -*-
"""网页版后端启动入口。

用法：
    python run.py            # 默认 0.0.0.0:15618，内网其他电脑可访问
    HOST=127.0.0.1 PORT=9000 python run.py   # 改地址/端口
"""
import os
import sys

import uvicorn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


if __name__ == '__main__':
    host = os.environ.get('HOST', '0.0.0.0')
    port = int(os.environ.get('PORT', '15618'))
    print('财务内部在线工具网页版启动中')
    print('  本机访问：http://127.0.0.1:%d' % port)
    if host in ('0.0.0.0', ''):
        try:
            import socket
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(('10.255.255.255', 1))
            lan = s.getsockname()[0]
            s.close()
            print('  局域网访问：http://%s:%d' % (lan, port))
        except Exception:
            pass
    uvicorn.run('app:app', host=host, port=port, reload=False)
