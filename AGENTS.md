# 项目规则

## caiwu1 局域网部署

- 部署对象：`caiwu1`，服务器为 `jl@192.168.24.29`（1lou）。
- 仅使用本地构建的镜像部署，禁止因发布而从 GitHub/GHCR 拉取镜像。
- 源码必须取自 `PDF转图片/web/frontend` 与 `PDF转图片/web/backend`；根目录的 `frontend`、`backend` 不属于 caiwu1。
- 本机 Docker 构建的完整入口是：

  ```powershell
  .\scripts\deploy-lan.ps1 -RemoteHost 192.168.24.29
  ```

- 脚本会创建临时扁平 Docker 构建上下文（`Dockerfile`、入口 Python、`requirements.txt`、`frontend/`、`backend/`），排除 `node_modules`、前端 `dist`、Python 缓存和 `backend/.tmp`，然后执行 `docker build`、`docker save`、`scp`、远端 `docker load`。
- SSH 使用本机已有的 `%USERPROFILE%\.ssh\id_ed25519`，用户为 `jl`；不得将私钥、密码或令牌写入代码、规则或日志。
- 更新时只操作 `caiwu1`：停止并重命名旧容器为带时间戳的 `caiwu1-previous-*`，新容器必须沿用端口 `15618:15618`、`59323:59323`、环境变量和日志轮转设置。
- 新容器状态不是 `running` 时，删除失败的新容器并恢复旧容器；验证成功前不得删除旧容器。
- 部署完成至少验证：容器为 `running`、`http://192.168.24.29:15618` 返回 HTTP 200、容器日志无启动错误。远端镜像 tar 与本机临时上下文、tar 必须清理。
- 不得运行 `pkill`、不明确的 `kill`、Docker daemon 重启或任何会影响其他容器的系统级命令。除非用户明确要求，禁止操作 caiwu1 以外的容器。
