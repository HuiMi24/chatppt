# ChatPPT (React + FastAPI)

本项目已扩展为 **前后端分离**：
- 前端：React(Vite) Web UI，支持浏览/编辑 PPT 文本、聊天改 PPT、自动主题应用
- 后端：FastAPI + python-pptx，提供解析、编辑、聊天、主题 API

> 原有 `chatppt.py` / `chatppt_ui.py` 仍保留，可继续使用。

## 目录结构

```text
chatppt/
├── backend/
│   ├── app/
│   │   ├── main.py            # FastAPI 入口
│   │   ├── models.py          # 请求/响应模型
│   │   ├── ppt_service.py     # PPT解析、编辑、主题应用
│   │   └── chat_service.py    # 聊天规划（LLM + rule fallback）
│   ├── tests/test_services.py # 最小测试
│   ├── examples/demo_api.sh   # API示例脚本
│   └── requirements.txt
├── frontend/
│   ├── src/App.jsx            # 页面编辑 + 聊天窗口
│   └── package.json
├── chatppt.py
└── chatppt_ui.py
```

## 功能说明

### 1) 解析 PPT 为结构化 JSON
- `GET /api/ppt?path=/abs/path/to/file.pptx`
- 返回：`slides -> shapes(text)`，包含 `slide_index/shape_index/role/title/body/text`

### 2) 按指令编辑并保存
- `POST /api/ppt/update`
- 入参：`path + edits[]`
- 输出：新的 `output_path`

### 3) 聊天改 PPT（支持无 Key fallback）
- `POST /api/chat`
- 有 `OPENAI_API_KEY` 时优先调用 LLM 生成编辑计划
- 无 key 时使用规则解析（例如：`把第3页标题改为xxx`）

### 4) 主题生成与应用
- `POST /api/theme/apply`
- 根据 PPT 文本关键词推断主题（corporate/playful/tech/clean）
- 自动应用配色、字体、字号、背景

## 环境变量

后端可选：

```bash
export OPENAI_API_KEY=your_key        # 可选，不设则走fallback
export OPENAI_MODEL=gpt-3.5-turbo     # 可选
```

前端可选：

```bash
# 默认 http://127.0.0.1:8000
export VITE_API_BASE=http://127.0.0.1:8000
```

## 本地启动

### 后端

```bash
cd backend
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

### 前端

```bash
cd frontend
npm install
npm run dev
```

打开 `http://127.0.0.1:5173`。

## 使用流程（Web）

1. 在输入框填入本机 PPT 绝对路径
2. 点击“加载”，查看每页可编辑文本
3. 手动改标题/正文后点击“保存文本修改”
4. 在聊天框输入自然语言（如“把第3页标题改为增长战略”）并发送
5. 点击“自动应用主题”可根据内容自动套用主题

## 测试与检查

### 后端测试

```bash
cd backend
PYTHONPATH=. python -m unittest tests/test_services.py
```

### 前端构建

```bash
cd frontend
npm run build
```

### API 示例脚本

```bash
cd backend
bash examples/demo_api.sh http://127.0.0.1:8000 /abs/path/to/demo.pptx
```

## 已知限制

- 目前编辑粒度基于 shape 文本，复杂富文本(run级样式)不会完全保留
- fallback 规则解析仅覆盖常见中英文表达
- 主题策略是启发式关键词匹配，不是设计系统级智能排版
