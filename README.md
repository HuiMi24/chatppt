# ChatPPT Studio (React + FastAPI)

ChatPPT Studio generates new PowerPoint decks, supports live text editing, and applies chat-driven revisions.

- Backend: FastAPI + `python-pptx`
- Frontend: React (Vite)
- Dependency manager (backend): `uv`
- LLM usage is optional. If `OPENAI_API_KEY` is missing, deterministic fallback generation is used.

## Features

1. Generate a new PPT from topic input
- `POST /api/generate`
- Supports `preset`, `audience`, `tone`, `slide_count`, and `language`
- Returns output path, outline, and theme metadata

2. Parse and edit PPT text content
- `GET /api/ppt`
- `POST /api/ppt/update`

3. Chat-based edits (existing compatibility retained)
- `POST /api/chat`

4. Apply inferred theme to an existing PPT (existing compatibility retained)
- `POST /api/theme/apply`

## Project Layout

```text
chatppt/
├── backend/
│   ├── app/
│   │   ├── main.py
│   │   ├── models.py
│   │   ├── ppt_service.py
│   │   ├── chat_service.py
│   │   └── generator_service.py
│   ├── tests/test_services.py
│   └── requirements.txt
├── frontend/
│   ├── src/App.jsx
│   ├── src/App.test.jsx
│   └── src/styles.css
└── README.md
```

## Environment Variables

Backend (optional):

```bash
export OPENAI_API_KEY=your_key
export OPENAI_MODEL=gpt-3.5-turbo
```

Frontend (optional):

```bash
export VITE_API_BASE=http://127.0.0.1:8000
```

## Run Locally

### 1) Start backend

```bash
cd backend
uv venv
source .venv/bin/activate
uv pip install -r requirements.txt
uv run uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

### 2) Start frontend

```bash
cd frontend
npm install
npm run dev
```

Open `http://127.0.0.1:5173`.

## API Example

### `POST /api/generate`

Request:

```json
{
  "topic": "AI customer support product roadmap",
  "preset": "Tech",
  "audience": "Executive team",
  "tone": "Concise",
  "slide_count": 6,
  "language": "en-US",
  "output_path": "generated/ai-roadmap.pptx"
}
```

Response:

```json
{
  "output_path": "generated/ai-roadmap.pptx",
  "outline": [{ "title": "...", "bullets": ["..."] }],
  "theme": {
    "name": "preset-tech",
    "font_name": "Segoe UI",
    "title_size_pt": 38,
    "body_size_pt": 19
  }
}
```

Notes:
- `preset` is optional (`Business`, `Tech`, `Education`, `Marketing`).
- Explicit `audience`, `tone`, `language`, and `slide_count` still work as before and take precedence over preset defaults.

## Quality Checks

Backend tests:

```bash
cd backend
PYTHONPATH=. uv run python -m unittest tests/test_services.py
```

Frontend tests and build:

```bash
cd frontend
npm test
npm run build
```
