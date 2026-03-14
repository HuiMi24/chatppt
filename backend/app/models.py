from pydantic import BaseModel, Field
from typing import List, Optional


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


class ChatResponse(BaseModel):
    plan: List[EditInstruction]
    output_path: str
    used_llm: bool
