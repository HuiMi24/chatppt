import os
import re
from typing import List

from .models import EditInstruction
from .ppt_service import PPTService


class ChatPlanner:
    def __init__(self, ppt_service: PPTService):
        self.ppt_service = ppt_service

    def build_plan(self, path: str, message: str) -> tuple[List[EditInstruction], bool]:
        api_key = os.getenv("OPENAI_API_KEY")
        if api_key:
            try:
                plan = self._llm_plan(path, message, api_key)
                if plan:
                    return plan, True
            except Exception:
                pass
        return self._rule_based_plan(path, message), False

    def _llm_plan(self, path: str, message: str, api_key: str) -> List[EditInstruction]:
        import json
        import openai

        openai.api_key = api_key
        doc = self.ppt_service.parse_ppt(path).model_dump()

        prompt = (
            "You are a PPT edit planner. Given user intent and PPT structure, return a JSON array only. "
            "Each item must include: slide_index(int), shape_index(int), new_text(str). "
            "Do not include explanations."
        )

        resp = openai.ChatCompletion.create(
            model=os.getenv("OPENAI_MODEL", "gpt-3.5-turbo"),
            messages=[
                {"role": "system", "content": prompt},
                {"role": "user", "content": f"request={message}\nppt={json.dumps(doc, ensure_ascii=False)}"},
            ],
            temperature=0,
        )
        content = resp.choices[0].message.content
        match = re.search(r"(\[.*\])", content, re.S)
        raw = match.group(1) if match else content
        data = json.loads(raw)
        return [EditInstruction(**x) for x in data]

    def _rule_based_plan(self, path: str, message: str) -> List[EditInstruction]:
        doc = self.ppt_service.parse_ppt(path)

        # CN pattern: 把第3页标题改为xxx
        cn = re.search(r"第\s*(\d+)\s*页\s*(标题|正文|内容)?\s*改为\s*[:：]?\s*(.+)$", message)
        if cn:
            slide_no = int(cn.group(1)) - 1
            target = cn.group(2) or "标题"
            text = cn.group(3).strip().strip('"')
            return self._pick_shape(doc, slide_no, target, text)

        # EN pattern: change slide 3 title to xxx
        en = re.search(r"slide\s*(\d+)\s*(title|body|content)?\s*(?:to|as)\s*(.+)$", message, re.I)
        if en:
            slide_no = int(en.group(1)) - 1
            target = en.group(2) or "title"
            text = en.group(3).strip().strip('"')
            return self._pick_shape(doc, slide_no, target, text)

        # fallback: edit first slide title
        if doc.slides and doc.slides[0].shapes:
            first = next((s for s in doc.slides[0].shapes if s.role == "title"), doc.slides[0].shapes[0])
            return [EditInstruction(slide_index=0, shape_index=first.shape_index, new_text=message)]
        return []

    def _pick_shape(self, doc, slide_no: int, target: str, text: str) -> List[EditInstruction]:
        if slide_no < 0 or slide_no >= len(doc.slides):
            return []
        slide = doc.slides[slide_no]

        target_role = "title" if target.lower() in ["标题", "title"] else "body"
        shape = next((s for s in slide.shapes if s.role == target_role), None)
        if not shape:
            shape = slide.shapes[0] if slide.shapes else None
        if not shape:
            return []

        return [EditInstruction(slide_index=slide_no, shape_index=shape.shape_index, new_text=text)]
