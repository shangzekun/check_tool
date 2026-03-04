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

## 与 `main` 分支冲突处理建议

当分支提示与 `main` 存在冲突时，建议按以下流程处理：

```bash
git fetch origin
git checkout <your-branch>
git merge origin/main
```

- 优先保留 `app/main.py`、`app/models.py` 与 `app/rules/` 中接口和数据模型的一致性。
- 前端冲突优先检查 `app/templates/index.html` 与 `app/static/style.css` 的布局是否仍满足：上传区在左上、规则说明在右上（1:2）、统计区在中间、结果表格在最下方。
- 冲突解决后执行基础校验：

```bash
python -m compileall app
```

## 仍然提示冲突时（排查清单）

如果你已经在本地解决过冲突，但平台（如 GitHub/GitLab）仍提示有冲突，通常是因为**目标分支有新提交**或**本地并未同步到正确的 `main`**。按下面步骤排查：

1. 确认远端与主分支存在：

```bash
git remote -v
git branch -a
```

2. 同步远端并重新合并（或 rebase）：

```bash
git fetch origin
git checkout <your-branch>
git merge origin/main
# 或者：git rebase origin/main
```

3. 若出现冲突，逐个文件解决后继续：

```bash
git add <resolved-files>
git commit
# 如果是 rebase：git rebase --continue
```

4. 本地再次确认没有冲突标记：

```bash
rg -n "^(<<<<<<<|=======|>>>>>>>)"
```

5. 推送最新分支到远端，再回到 PR 页面刷新：

```bash
git push
```

> 说明：如果 `git branch -a` 里看不到 `origin/main`，说明当前仓库副本没有配置远端主分支，需先补齐远端或在正确仓库中执行以上步骤。
