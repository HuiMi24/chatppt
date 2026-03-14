# ChatPPT Studio (React + FastAPI)

把原先的 Web Editor 升级为 **AI 生成 PPT + 在线编辑 + 聊天改稿** 的产品体验。

- 后端：FastAPI + python-pptx（保留原有编辑接口）
- 前端：React (Vite)
- Python 依赖管理：**uv**
- LLM Key 可选：有 key 用 LLM 产出大纲；无 key 自动 fallback 模板生成

## 功能概览

1. **生成新 PPT**（不是只改已有文件）
   - `POST /api/generate`
   - 输入主题/受众/语气/页数/语言
   - 自动输出新 `.pptx`，并返回结构化大纲和主题信息

2. **在线预览并继续编辑文本**
   - `GET /api/ppt`
   - `POST /api/ppt/update`

3. **聊天修改 PPT**（兼容原能力）
   - `POST /api/chat`

4. **自动主题应用**（兼容原能力）
   - `POST /api/theme/apply`

---

## 目录结构

```text
chatppt/
├── backend/
│   ├── app/
│   │   ├── main.py               # FastAPI 入口
│   │   ├── models.py
│   │   ├── ppt_service.py        # 解析/编辑/主题
│   │   ├── chat_service.py       # 聊天改稿计划
│   │   └── generator_service.py  # 新增：PPT 生成服务（LLM + fallback）
│   ├── tests/test_services.py    # 包含 generate fallback 最小测试
│   └── requirements.txt
├── frontend/
│   ├── src/App.jsx               # 生成器 + 编辑区 + 聊天区
│   └── src/styles.css            # 现代化深色渐变 UI
└── README.md
```

---

## 环境变量

后端可选（不填也可运行）：

```bash
export OPENAI_API_KEY=your_key         # 可选
export OPENAI_MODEL=gpt-3.5-turbo      # 可选
```

前端可选：

```bash
export VITE_API_BASE=http://127.0.0.1:8000
```

---

## 启动方式（uv）

### 1) 启动后端（FastAPI）

```bash
cd backend
uv venv
source .venv/bin/activate
uv pip install -r requirements.txt
uv run uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

### 2) 启动前端（React）

```bash
cd frontend
npm install
npm run dev
```

打开：`http://127.0.0.1:5173`

---

## 接口说明

### POST /api/generate

生成全新的 PPT。

**Request JSON**

```json
{
  "topic": "AI 客服产品路线图",
  "audience": "管理层",
  "tone": "专业",
  "slide_count": 6,
  "language": "zh-CN",
  "output_path": "generated/ai-roadmap.pptx"
}
```

**Response JSON**

```json
{
  "output_path": "generated/ai-roadmap.pptx",
  "outline": [
    { "title": "背景与价值", "bullets": ["..."] }
  ],
  "theme": {
    "name": "tech",
    "font_name": "Segoe UI",
    "title_size_pt": 38,
    "body_size_pt": 20
  }
}
```

### GET /api/ppt?path=/abs/path/file.pptx

解析 PPT，返回可编辑文本结构。

### POST /api/ppt/update

按 `edits` 批量更新文本并保存新文件。

### POST /api/chat

继续对话修改 PPT（有 key 走 LLM，无 key 走规则 fallback）。

### POST /api/theme/apply

基于文本语义推断主题并应用。

---

## 质量检查

### 后端测试

```bash
cd backend
PYTHONPATH=. uv run python -m unittest tests/test_services.py
```

### 前端构建

```bash
cd frontend
npm run build
```

---

## 说明

- 保留原有 `/api/ppt`、`/api/ppt/update`、`/api/chat`、`/api/theme/apply` 兼容能力。
- `generator_service.py` 独立负责「大纲生成 + 新 PPT 创建 + 主题自动应用」。
- fallback 逻辑确保无 LLM key 也可稳定生成可编辑 PPT。
