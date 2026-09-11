# 财务内部在线工具网页版 — 镜像构建
# 多阶段：先构建前端（Vue + vite）生成 dist，再组装 Python 后端镜像。
# 注意：内部结构与仓库 master 分支的镜像保持一致（/app/backend/run.py、端口 15618），
# 保证服务器现有部署脚本无需改动即可切换。

# ---------- 阶段 1：构建前端 ----------
FROM node:20-alpine AS frontend
WORKDIR /build
# 先装依赖再拷源码，利用 Docker 层缓存
COPY web/frontend/package.json web/frontend/package-lock.json ./
RUN npm ci
COPY web/frontend/ ./
RUN npm run build

# ---------- 阶段 2：运行后端（FastAPI + OCR） ----------
FROM python:3.12-slim
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1
WORKDIR /app

# OCR（onnxruntime / opencv）需要的底层动态库
RUN apt-get update && apt-get install -y --no-install-recommends \
        libgomp1 \
        libgl1 \
        libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# Python 依赖
COPY web/requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# pdf转图片.py 放容器根目录 /（backend/processor.py 上溯 3 级 = /）
COPY pdf转图片.py /pdf转图片.py

# 后端代码（结构对齐 master：/app/backend/）
COPY web/backend/ backend/

# 前端构建产物（阶段 1，结构对齐 master：/app/frontend/dist/）
COPY --from=frontend /build/dist frontend/dist/

# 后端临时任务目录
RUN mkdir -p backend/.tmp

# 端口对齐 master（web/backend/run.py 默认 8000，这里用环境变量覆盖）
ENV PORT=15618
EXPOSE 15618

CMD ["python", "backend/run.py"]
