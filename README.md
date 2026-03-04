# Excel 检查工具

一个可部署在云服务器上的 Excel（`.xlsx`）检查平台，支持：

- 上传 Excel 文件
- 勾选多个检查规则并执行检查
- 在页面中查看结果与按级别筛选
- 导出包含 `Summary` 与 `Issues` 工作表的 Excel 报告
- 页面刷新后自动恢复上次检查结果（`localStorage`）

## 技术栈

- 后端：FastAPI + Pydantic + openpyxl
- 前端：原生 HTML/CSS/JavaScript
- 容器：Docker + docker-compose

## 本地运行

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

访问：<http://localhost:8000>

## Docker 运行

```bash
docker compose up --build
```

访问：<http://localhost:8000>

## API 概览

- `GET /api/rules`：获取规则列表与说明
- `POST /api/check`：上传文件并执行选中规则
- `POST /api/report`：根据检查结果导出 Excel 报告

## 规则扩展

规则位于 `app/rules/` 目录。每条规则是一个独立 Python 模块，统一返回问题列表（issues），便于后续持续扩展。
