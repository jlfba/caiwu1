# 财务内部在线工具网页版

在浏览器里操作 PDF 发票：批量上传 → 选择模式 → 制作 → 直接下载 Excel。

## 快速开始

**日常使用（推荐，一条命令）**——前端已构建，页面和接口都在 8000 端口：

```bash
cd web
python backend/run.py
# 或用脚本：双击 web/start.bat
```

浏览器打开 `http://127.0.0.1:8000` 即可，**不需要 npm**。

**开发前端（要热更新）**——运行 `web/dev.bat`（自动开两个窗口），或手动：

```bash
# 终端1：后端
cd web/backend && python run.py
# 终端2：前端（vite 代理 /api → 8000，改代码自动刷新）
cd web/frontend && npm run dev   # 打开 http://127.0.0.1:5173
```

> `npm run dev` 必须在 `web/frontend/` 目录下执行（package.json 在那里）。
> 改完前端要发布时：`cd web/frontend && npm run build`，之后 `python backend/run.py` 就直接提供新版页面。

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
