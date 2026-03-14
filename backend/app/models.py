from typing import List, Optional

from pydantic import BaseModel, Field


class ShapeText(BaseModel):
    shape_index: int
    role: str  # title/body/text
    text: str


class SlideContent(BaseModel):
    slide_index: int
    title: Optional[str] = None
    shapes: List[ShapeText] = Field(default_factory=list)


class PptDocument(BaseModel):
    path: str
    slide_count: int
    slides: List[SlideContent]


class EditInstruction(BaseModel):
    slide_index: int
    shape_index: int
    new_text: str


class UpdateRequest(BaseModel):
    path: str
    edits: List[EditInstruction]
    output_path: Optional[str] = None


class ChatRequest(BaseModel):
    path: str
    message: str
    output_path: Optional[str] = None


class ThemeApplyRequest(BaseModel):
    path: str
    output_path: Optional[str] = None


class ThemePresetApplyRequest(BaseModel):
    path: str
    preset: str
    output_path: Optional[str] = None


class ChatResponse(BaseModel):
    plan: List[EditInstruction]
    output_path: str
    used_llm: bool


class OutlineSlide(BaseModel):
    title: str
    bullets: List[str] = Field(default_factory=list)


class GenerateRequest(BaseModel):
    topic: str
    audience: Optional[str] = None
    tone: Optional[str] = None
    preset: Optional[str] = None
    slide_count: int = Field(default=6, ge=3, le=20)
    language: Optional[str] = None
    output_path: Optional[str] = None


class GenerateResponse(BaseModel):
    output_path: str
    outline: List[OutlineSlide]
    theme: dict
