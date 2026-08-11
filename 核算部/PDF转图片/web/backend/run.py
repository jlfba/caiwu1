# -*- coding: utf-8 -*-
"""网页版后端启动入口。

用法：
    python run.py            # 默认 0.0.0.0:8000，内网其他电脑可访问
    HOST=127.0.0.1 PORT=9000 python run.py   # 改地址/端口
"""
import os
import sys

import uvicorn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


if __name__ == '__main__':
    host = os.environ.get('HOST', '0.0.0.0')
    port = int(os.environ.get('PORT', '8000'))
    print('PDF 工具网页版启动中：http://%s:%d' %
          ('127.0.0.1' if host in ('0.0.0.0', '127.0.0.1') else host, port))
    uvicorn.run('app:app', host=host, port=port, reload=False)
