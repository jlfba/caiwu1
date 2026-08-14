# 财务内部在线工具网页版 — 镜像构建
# 多阶段：先构建前端（Vue + vite）生成 dist，再组装 Python 后端镜像。

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
        libgl1 \
        libglib2.0-0 \
        libgomp1 \
    && rm -rf /var/lib/apt/lists/*

# Python 依赖
COPY web/requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# 业务代码：核心识别脚本 + 后端
COPY pdf转图片.py ./
COPY web/backend/ ./web/backend/

# 前端构建产物（阶段 1）
COPY --from=frontend /build/dist ./web/frontend/dist

WORKDIR /app/web/backend
EXPOSE 8000
# run.py 默认监听 0.0.0.0:8000（可用 HOST/PORT 环境变量覆盖）
CMD ["python", "run.py"]
