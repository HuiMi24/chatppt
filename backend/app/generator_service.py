from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import List, Tuple

from pptx import Presentation

from .models import GenerateRequest, OutlineSlide
from .ppt_service import PPTService, ThemeConfig


class GeneratorService:
    def __init__(self, ppt_service: PPTService):
        self.ppt_service = ppt_service

    def generate(self, req: GenerateRequest) -> tuple[str, List[OutlineSlide], ThemeConfig]:
        api_key = os.getenv("OPENAI_API_KEY")

        if api_key:
            try:
                outline = self._outline_from_llm(req, api_key)
            except Exception:
                outline = self._fallback_outline(req)
        else:
            outline = self._fallback_outline(req)

        output_path = req.output_path or self._default_output_path(req.topic)
        subtitle = self._build_subtitle(req)
        self.ppt_service.create_from_outline(outline, output_path=output_path, topic=req.topic, subtitle=subtitle)

        prs = Presentation(output_path)
        theme = self.ppt_service.infer_theme_from_outline(req.topic, outline)
        self.ppt_service.apply_theme_to_presentation(prs, theme)
        prs.save(output_path)

        return output_path, outline, theme

    def _outline_from_llm(self, req: GenerateRequest, api_key: str) -> List[OutlineSlide]:
        import openai

        openai.api_key = api_key
        prompt = (
            "你是专业的PPT策划师。请根据用户需求，生成结构化 JSON 数组。"
            "每个元素格式: {\"title\": string, \"bullets\": string[]}。"
            "要求: bullets 2-4 条，简洁可演讲，输出页数与 slide_count 一致。"
            "只能输出 JSON 数组，不要输出其他内容。"
        )

        payload = {
            "topic": req.topic,
            "audience": req.audience,
            "tone": req.tone,
            "slide_count": req.slide_count,
            "language": req.language or "zh-CN",
        }

        resp = openai.ChatCompletion.create(
            model=os.getenv("OPENAI_MODEL", "gpt-3.5-turbo"),
            messages=[
                {"role": "system", "content": prompt},
                {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
            ],
            temperature=0.7,
        )
        content = resp.choices[0].message.content
        match = re.search(r"(\[.*\])", content, re.S)
        raw = match.group(1) if match else content
        data = json.loads(raw)

        outline = [OutlineSlide(**item) for item in data if item.get("title")]
        if not outline:
            raise ValueError("LLM returned empty outline")
        return self._normalize_outline(outline, req.slide_count)

    def _fallback_outline(self, req: GenerateRequest) -> List[OutlineSlide]:
        tone = req.tone or "专业"
        audience = req.audience or "通用受众"
        language = (req.language or "zh").lower()

        if language.startswith("en"):
            base = [
                (f"Why {req.topic} matters", [f"Audience: {audience}", "Current landscape", "Key opportunity"]),
                ("Problem & pain points", ["Main challenges", "Business impact", "Urgency"]),
                ("Goals & success metrics", ["Primary objective", "KPIs", "Target timeline"]),
                ("Approach", [f"Tone: {tone}", "Core methods", "Execution steps"]),
                ("Plan & milestones", ["Phase 1", "Phase 2", "Phase 3"]),
                ("Summary & next actions", ["Key takeaways", "Immediate next step", "Owner & follow-up"]),
            ]
        else:
            base = [
                (f"{req.topic}背景与价值", [f"受众：{audience}", "现状与趋势", "核心机会点"]),
                ("当前问题与挑战", ["主要痛点", "影响范围", "为什么现在要做"]),
                ("目标与衡量指标", ["阶段目标", "关键KPI", "预期达成时间"]),
                ("策略与方案设计", [f"表达风格：{tone}", "关键策略", "实施路径"]),
                ("落地计划与里程碑", ["短期行动", "中期推进", "长期优化"]),
                ("结论与下一步", ["核心结论", "待决策事项", "下一步行动"]),
            ]

        outline = [OutlineSlide(title=t, bullets=b) for t, b in base[: req.slide_count]]

        while len(outline) < req.slide_count:
            idx = len(outline) + 1
            if language.startswith("en"):
                outline.append(
                    OutlineSlide(
                        title=f"Extended topic {idx}",
                        bullets=["Supporting point", "Data / example", "Action item"],
                    )
                )
            else:
                outline.append(
                    OutlineSlide(
                        title=f"补充章节 {idx}",
                        bullets=["补充观点", "案例或数据", "建议动作"],
                    )
                )
        return outline

    def _normalize_outline(self, outline: List[OutlineSlide], slide_count: int) -> List[OutlineSlide]:
        normalized = outline[:slide_count]
        while len(normalized) < slide_count:
            idx = len(normalized) + 1
            normalized.append(OutlineSlide(title=f"补充页 {idx}", bullets=["关键点", "说明", "行动建议"]))
        return normalized

    def _build_subtitle(self, req: GenerateRequest) -> str:
        items = [x for x in [req.audience, req.tone, req.language] if x]
        return " | ".join(items) if items else "Generated by ChatPPT"

    def _default_output_path(self, topic: str) -> str:
        slug = re.sub(r"[^a-zA-Z0-9\u4e00-\u9fff]+", "-", topic).strip("-") or "generated"
        output_dir = Path("generated")
        output_dir.mkdir(parents=True, exist_ok=True)
        return str(output_dir / f"{slug}.pptx")
