# ============================================
# Stage 1: 构建前端（Vue 3 + Vite）
# ============================================
FROM node:20-alpine AS frontend-build
WORKDIR /build
COPY frontend/package.json frontend/package-lock.json ./
RUN npm ci
COPY frontend/ ./
RUN npm run build

# ============================================
# Stage 2: 最终镜像（Python FastAPI 后端）
# ============================================
FROM python:3.11-slim

# 系统依赖：
# - libgomp1     onnxruntime/OpenMP 运行时
# - libgl1       OpenCV（rapidocr 依赖）需要 libGL.so.1
# - libglib2.0-0 OpenCV 常用依赖
RUN apt-get update && apt-get install -y --no-install-recommends \
    libgomp1 \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# pdf转图片.py 必须放在容器根目录 /
# （backend/processor.py 的 _BASE_DIR 从 d:/test/backend/processor.py 上溯 3 级 = /）
COPY pdf转图片.py /pdf转图片.py

# Python 依赖
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 后端代码
COPY backend/ backend/

# 前端构建产物
COPY --from=frontend-build /build/dist frontend/dist/

# 创建后端临时目录
RUN mkdir -p backend/.tmp

EXPOSE 15618

CMD ["python", "backend/run.py"]
