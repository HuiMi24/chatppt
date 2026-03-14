from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from .chat_service import ChatPlanner
from .generator_service import GeneratorService
from .models import (
    ChatRequest,
    ChatResponse,
    GenerateRequest,
    GenerateResponse,
    ThemeApplyRequest,
    UpdateRequest,
)
from .ppt_service import PPTService

app = FastAPI(title="ChatPPT API", version="0.2.0")

PREVIEW_ROOT = Path(__file__).resolve().parent.parent / "generated" / "previews"
PREVIEW_ROOT.mkdir(parents=True, exist_ok=True)
app.mount("/static/previews", StaticFiles(directory=str(PREVIEW_ROOT)), name="previews")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

ppt_service = PPTService()
chat_planner = ChatPlanner(ppt_service)
generator_service = GeneratorService(ppt_service)


@app.get("/api/health")
def health():
    return {"ok": True}


@app.post("/api/generate", response_model=GenerateResponse)
def generate_ppt(req: GenerateRequest):
    try:
        output_path, outline, theme = generator_service.generate(req)
        return GenerateResponse(
            output_path=output_path,
            outline=outline,
            theme={
                "name": theme.name,
                "font_name": theme.font_name,
                "title_size_pt": theme.title_size_pt,
                "body_size_pt": theme.body_size_pt,
            },
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/api/ppt")
def parse_ppt(path: str):
    try:
        return ppt_service.parse_ppt(path)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/api/ppt/preview")
def preview_ppt(path: str):
    try:
        rel_images = ppt_service.render_preview_images(path, str(PREVIEW_ROOT))
        return {
            "images": [f"/static/previews/{p}" for p in rel_images],
            "count": len(rel_images),
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/ppt/update")
def update_ppt(req: UpdateRequest):
    try:
        output = ppt_service.apply_edits(req.path, req.edits, req.output_path)
        return {"output_path": output}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/chat", response_model=ChatResponse)
def chat_edit(req: ChatRequest):
    try:
        plan, used_llm = chat_planner.build_plan(req.path, req.message)
        if not plan:
            raise HTTPException(status_code=400, detail="No editable action parsed from message")
        output = ppt_service.apply_edits(req.path, plan, req.output_path)
        return ChatResponse(plan=plan, output_path=output, used_llm=used_llm)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/theme/apply")
def apply_theme(req: ThemeApplyRequest):
    try:
        output, theme = ppt_service.apply_theme(req.path, req.output_path)
        return {
            "output_path": output,
            "theme": {
                "name": theme.name,
                "font_name": theme.font_name,
                "title_size_pt": theme.title_size_pt,
                "body_size_pt": theme.body_size_pt,
            },
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
