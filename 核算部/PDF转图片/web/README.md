# PDF 工具网页版

在浏览器里操作 PDF 发票：批量上传 → 选择模式 → 制作 → 直接下载 Excel。

## 快速开始

```bash
# 1. 构建前端（改过前端才需要重跑）
cd web/frontend
npm install
npm run build

# 2. 启动后端（页面 + 接口一起提供）
cd web/backend
python run.py
# 默认 0.0.0.0:8000；内网其他电脑访问 http://<本机IP>:8000

# 3. 浏览器打开
http://127.0.0.1:8000
```

## 功能

| 功能模式 | 干什么 | 下载 |
|---|---|---|
| **收款组** | PDF 逐页转图片，OCR 识别发票号码/购买方/销售方/金额并重命名 | `发票图片表.xlsx`（含图片） |
| **付款组 - canexs** | 识别 INVOICE / TRACKING NO. / 明细行 | `canexs发票明细表.xlsx` |
| **付款组 - 精准** | Accuracy Customs Brokers 清关发票 | `精准发票明细表.xlsx` |

## 常见问题

- **首次收款组较慢**：需下载一次 PaddleOCR 模型（约几百 MB），之后走缓存。
- **"没有识别到任何明细"**：发票类型选错了（精准发票选成了 canexs 等），换个类型重试。
- **前端开发调试**：`web/frontend` 下 `npm run dev`（vite 已代理 `/api` 到 8000）。

详见 [../开发文档.md](../开发文档.md)。
