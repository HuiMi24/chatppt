from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Dict, List

from pptx.dml.color import RGBColor
from pptx import Presentation

from .models import GenerateRequest, OutlineSlide
from .ppt_service import PPTService, ThemeConfig


PRESET_PROFILES: Dict[str, dict] = {
    "business": {
        "label": "Business",
        "audience": "Business stakeholders",
        "tone": "Executive and data-driven",
        "language": "en-US",
    },
    "tech": {
        "label": "Tech",
        "audience": "Product and engineering teams",
        "tone": "Analytical and system-oriented",
        "language": "en-US",
    },
    "education": {
        "label": "Education",
        "audience": "Students and instructors",
        "tone": "Instructional and clear",
        "language": "en-US",
    },
    "marketing": {
        "label": "Marketing",
        "audience": "Growth and brand teams",
        "tone": "Persuasive and audience-focused",
        "language": "en-US",
    },
}

PRESET_ALIASES = {
    "business": "business",
    "tech": "tech",
    "education": "education",
    "marketing": "marketing",
}


class GeneratorService:
    def __init__(self, ppt_service: PPTService):
        self.ppt_service = ppt_service

    def generate(self, req: GenerateRequest) -> tuple[str, List[OutlineSlide], ThemeConfig]:
        if os.getenv("FAKE_LLM_RESPONSES", "0") == "1":
            outline = self._fake_llm_outline(req)
        else:
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
        theme = self._select_theme(req, outline)
        self.ppt_service.apply_theme_to_presentation(prs, theme)
        prs.save(output_path)

        return output_path, outline, theme

    def _fake_llm_outline(self, req: GenerateRequest) -> List[OutlineSlide]:
        language = req.language or "en-US"
        lang = self._language_family(language)
        topic = req.topic
        if lang == "zh":
            slides = [
                OutlineSlide(title=f"[FAKE LLM] {topic} - 项目概览", bullets=["目标与范围", "关键里程碑", "当前状态"]),
                OutlineSlide(title=f"[FAKE LLM] {topic} - 实施路径", bullets=["阶段划分", "资源与分工", "风险与对策"]),
            ]
        else:
            slides = [
                OutlineSlide(title=f"[FAKE LLM] {topic} - Executive Overview", bullets=["Goal and scope", "Key milestones", "Current status"]),
                OutlineSlide(title=f"[FAKE LLM] {topic} - Execution Plan", bullets=["Phases", "Ownership", "Risks and mitigation"]),
            ]
        return self._normalize_outline(slides, req.slide_count, language)

    def _outline_from_llm(self, req: GenerateRequest, api_key: str) -> List[OutlineSlide]:
        import openai

        openai.api_key = api_key

        preset_key = self._normalize_preset(req.preset)
        profile = PRESET_PROFILES.get(preset_key, {})

        prompt = (
            "You are a professional presentation strategist. Generate a structured JSON array only. "
            "Each item must match: {\"title\": string, \"bullets\": string[]}. "
            "Return exactly slide_count slides, 2-4 bullets each, concise and presentation-ready. "
            "Use the requested language and keep the outline tone aligned with the preset style when provided."
        )

        payload = {
            "topic": req.topic,
            "preset": profile.get("label") if profile else req.preset,
            "audience": req.audience or profile.get("audience"),
            "tone": req.tone or profile.get("tone"),
            "slide_count": req.slide_count,
            "language": req.language or profile.get("language") or "zh-CN",
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
            raise ValueError("LLM returned an empty outline")
        return self._normalize_outline(outline, req.slide_count, req.language)

    def _fallback_outline(self, req: GenerateRequest) -> List[OutlineSlide]:
        preset_key = self._normalize_preset(req.preset)
        profile = PRESET_PROFILES.get(preset_key, {})

        language = req.language or profile.get("language") or "zh-CN"
        lang = self._language_family(language)
        default_tone, default_audience = self._default_style_for_language(lang)
        tone = req.tone or profile.get("tone") or default_tone
        audience = req.audience or profile.get("audience") or default_audience

        if lang == "en":
            base = self._english_outline(req.topic, preset_key, audience, tone)
        elif lang == "ja":
            base = self._japanese_outline(req.topic, preset_key, audience, tone)
        else:
            base = self._chinese_outline(req.topic, preset_key, audience, tone)

        outline = [OutlineSlide(title=t, bullets=b) for t, b in base[: req.slide_count]]
        return self._normalize_outline(outline, req.slide_count, language)

    def _english_outline(self, topic: str, preset: str | None, audience: str, tone: str):
        if preset == "business":
            return [
                (f"Executive Summary: {topic}", [f"Audience: {audience}", "Current business context", "Decision required"]),
                ("Market Context", ["Industry shift", "Customer demand signal", "Competitive pressure"]),
                ("Strategic Objectives", ["Primary objective", "Success metrics", "Expected timeline"]),
                ("Operating Plan", [f"Narrative tone: {tone}", "Cross-team ownership", "Execution rhythm"]),
                ("Financial Impact", ["Investment required", "Expected return", "Risk controls"]),
                ("Recommendation & Next Steps", ["Key recommendation", "Immediate action", "Review cadence"]),
            ]

        if preset == "tech":
            return [
                (f"Problem Landscape: {topic}", [f"Target users: {audience}", "Current system constraints", "Opportunity statement"]),
                ("Architecture Direction", ["Core components", "Data and integration flow", "Scalability strategy"]),
                ("Product & Engineering Goals", ["Outcome targets", "Technical KPIs", "Reliability goals"]),
                ("Implementation Blueprint", [f"Delivery tone: {tone}", "Phased rollout", "Quality gates"]),
                ("Risks and Mitigations", ["Known risks", "Fallback plans", "Monitoring signals"]),
                ("Roadmap and Ownership", ["Milestones", "Team ownership", "Next sprint actions"]),
            ]

        if preset == "education":
            return [
                (f"Learning Focus: {topic}", [f"Learners: {audience}", "Why this matters", "Session outcomes"]),
                ("Foundational Concepts", ["Key definitions", "Core principles", "Common misconceptions"]),
                ("Learning Objectives", ["Knowledge objective", "Skill objective", "Assessment criteria"]),
                ("Teaching Plan", [f"Instruction style: {tone}", "Activities and examples", "Practice checkpoints"]),
                ("Assessment and Feedback", ["Formative checks", "Summative evaluation", "Feedback loop"]),
                ("Recap and Follow-up", ["Key takeaways", "Homework or extension", "Next class preparation"]),
            ]

        if preset == "marketing":
            return [
                (f"Campaign Narrative: {topic}", [f"Target audience: {audience}", "Brand challenge", "Growth objective"]),
                ("Audience Insights", ["Primary segment", "Need states", "Purchase triggers"]),
                ("Value Proposition", ["Core message", "Differentiators", "Proof points"]),
                ("Channel Strategy", [f"Communication tone: {tone}", "Channel mix", "Content pillars"]),
                ("Launch Plan", ["Timeline", "Budget focus", "Experiment plan"]),
                ("Performance and Optimization", ["Success metrics", "Reporting cadence", "Optimization actions"]),
            ]

        return [
            (f"Why {topic} matters", [f"Audience: {audience}", "Current landscape", "Key opportunity"]),
            ("Problem and Pain Points", ["Main challenges", "Business impact", "Urgency"]),
            ("Goals and Success Metrics", ["Primary objective", "KPIs", "Target timeline"]),
            ("Approach", [f"Tone: {tone}", "Core methods", "Execution steps"]),
            ("Plan and Milestones", ["Phase 1", "Phase 2", "Phase 3"]),
            ("Summary and Next Actions", ["Key takeaways", "Immediate next step", "Owner and follow-up"]),
        ]

    def _chinese_outline(self, topic: str, preset: str | None, audience: str, tone: str):
        if preset == "business":
            return [
                (f"{topic}执行摘要", [f"受众：{audience}", "业务背景", "决策目标"]),
                ("市场与竞争环境", ["行业变化", "客户需求", "竞争态势"]),
                ("战略目标", ["核心目标", "衡量指标", "时间预期"]),
                ("执行路径", [f"表达风格：{tone}", "协同机制", "推进节奏"]),
                ("投入与收益", ["资源投入", "预期收益", "风险控制"]),
                ("结论与下一步", ["关键结论", "立即行动", "复盘机制"]),
            ]

        if preset == "tech":
            return [
                (f"{topic}问题定义", [f"受众：{audience}", "现有系统瓶颈", "技术机会"]),
                ("技术架构方向", ["核心模块", "数据与接口", "扩展性方案"]),
                ("目标与指标", ["产品目标", "技术KPI", "稳定性指标"]),
                ("实施蓝图", [f"表达风格：{tone}", "分阶段交付", "质量关卡"]),
                ("风险与应对", ["主要风险", "预案", "监控信号"]),
                ("路线图与分工", ["里程碑", "团队职责", "近期动作"]),
            ]

        if preset == "education":
            return [
                (f"{topic}学习导入", [f"对象：{audience}", "学习价值", "课程产出"]),
                ("核心概念", ["关键定义", "基础原理", "易错点"]),
                ("学习目标", ["知识目标", "能力目标", "评估方式"]),
                ("教学安排", [f"表达风格：{tone}", "教学活动", "练习节点"]),
                ("评估与反馈", ["过程评估", "结果评估", "反馈改进"]),
                ("总结与延伸", ["重点回顾", "课后任务", "下一步准备"]),
            ]

        if preset == "marketing":
            return [
                (f"{topic}营销叙事", [f"目标受众：{audience}", "品牌挑战", "增长目标"]),
                ("用户洞察", ["核心人群", "需求动机", "转化触发点"]),
                ("价值主张", ["核心信息", "差异化", "信任证据"]),
                ("渠道与内容策略", [f"表达风格：{tone}", "渠道组合", "内容主线"]),
                ("投放与节奏", ["上线计划", "预算分配", "实验策略"]),
                ("效果复盘", ["关键指标", "追踪节奏", "优化动作"]),
            ]

        return [
            (f"{topic}背景与价值", [f"受众：{audience}", "现状与趋势", "核心机会点"]),
            ("当前问题与挑战", ["主要痛点", "影响范围", "为什么现在要做"]),
            ("目标与衡量指标", ["阶段目标", "关键KPI", "预期达成时间"]),
            ("策略与方案设计", [f"表达风格：{tone}", "关键策略", "实施路径"]),
            ("落地计划与里程碑", ["短期行动", "中期推进", "长期优化"]),
            ("结论与下一步", ["核心结论", "待决策事项", "下一步行动"]),
        ]

    def _japanese_outline(self, topic: str, preset: str | None, audience: str, tone: str):
        if preset == "business":
            return [
                (f"{topic} エグゼクティブサマリー", [f"対象: {audience}", "事業背景", "意思決定ポイント"]),
                ("市場環境", ["業界トレンド", "顧客ニーズ", "競争状況"]),
                ("戦略目標", ["主要目標", "KPI", "達成スケジュール"]),
                ("実行計画", [f"トーン: {tone}", "推進体制", "進行リズム"]),
                ("投資対効果", ["必要投資", "期待効果", "リスク管理"]),
                ("提案と次のアクション", ["結論", "直近アクション", "レビュー計画"]),
            ]

        if preset == "tech":
            return [
                (f"{topic} の課題整理", [f"対象: {audience}", "現行システムの制約", "技術機会"]),
                ("アーキテクチャ方針", ["主要コンポーネント", "データ連携", "拡張性"]),
                ("目標と指標", ["プロダクト目標", "技術KPI", "信頼性目標"]),
                ("実装ロードマップ", [f"トーン: {tone}", "段階導入", "品質ゲート"]),
                ("リスク管理", ["主要リスク", "回避策", "監視指標"]),
                ("体制と次工程", ["マイルストーン", "担当", "次スプリント"]),
            ]

        if preset == "education":
            return [
                (f"{topic} 学習の導入", [f"学習者: {audience}", "学ぶ意義", "到達目標"]),
                ("基礎概念", ["重要用語", "基本原理", "つまずきやすい点"]),
                ("学習目標", ["知識目標", "技能目標", "評価観点"]),
                ("授業設計", [f"トーン: {tone}", "演習と例", "理解チェック"]),
                ("評価とフィードバック", ["形成的評価", "総括評価", "改善サイクル"]),
                ("まとめと課題", ["要点整理", "復習課題", "次回準備"]),
            ]

        if preset == "marketing":
            return [
                (f"{topic} キャンペーン構想", [f"対象: {audience}", "ブランド課題", "成長目標"]),
                ("顧客インサイト", ["主要セグメント", "ニーズ", "購買トリガー"]),
                ("価値提案", ["主メッセージ", "差別化要素", "根拠"]),
                ("チャネル戦略", [f"トーン: {tone}", "チャネル配分", "コンテンツ方針"]),
                ("実行計画", ["スケジュール", "予算重点", "検証プラン"]),
                ("成果測定", ["主要指標", "報告頻度", "改善アクション"]),
            ]

        return [
            (f"{topic} の重要性", [f"対象: {audience}", "現状", "主要機会"]),
            ("課題と背景", ["主な課題", "影響", "緊急性"]),
            ("目標とKPI", ["主要目標", "測定指標", "予定期間"]),
            ("アプローチ", [f"トーン: {tone}", "主要手法", "実行ステップ"]),
            ("計画とマイルストーン", ["フェーズ1", "フェーズ2", "フェーズ3"]),
            ("まとめと次の行動", ["要点", "直近アクション", "担当とフォロー"]),
        ]

    def _select_theme(self, req: GenerateRequest, outline: List[OutlineSlide]) -> ThemeConfig:
        preset_theme = self._theme_from_preset(req.preset)
        if preset_theme:
            return preset_theme
        return self.ppt_service.infer_theme_from_outline(req.topic, outline)

    def _theme_from_preset(self, preset: str | None) -> ThemeConfig | None:
        normalized = self._normalize_preset(preset)
        if normalized == "business":
            return ThemeConfig(
                "preset-business",
                RGBColor(16, 48, 89),
                RGBColor(43, 59, 76),
                RGBColor(243, 246, 252),
                "Calibri",
                36,
                20,
            )
        if normalized == "tech":
            return ThemeConfig(
                "preset-tech",
                RGBColor(15, 64, 132),
                RGBColor(13, 110, 122),
                RGBColor(237, 248, 255),
                "Segoe UI",
                38,
                19,
            )
        if normalized == "education":
            return ThemeConfig(
                "preset-education",
                RGBColor(78, 59, 14),
                RGBColor(64, 80, 32),
                RGBColor(255, 252, 232),
                "Verdana",
                37,
                21,
            )
        if normalized == "marketing":
            return ThemeConfig(
                "preset-marketing",
                RGBColor(132, 28, 46),
                RGBColor(107, 33, 95),
                RGBColor(255, 243, 247),
                "Trebuchet MS",
                40,
                20,
            )
        return None

    def _normalize_outline(self, outline: List[OutlineSlide], slide_count: int, language: str | None = None) -> List[OutlineSlide]:
        normalized = outline[:slide_count]
        lang = self._language_family(language or "")
        while len(normalized) < slide_count:
            idx = len(normalized) + 1
            if lang == "en":
                normalized.append(
                    OutlineSlide(
                        title=f"Extended section {idx}",
                        bullets=["Supporting point", "Data or example", "Action item"],
                    )
                )
            elif lang == "ja":
                normalized.append(
                    OutlineSlide(
                        title=f"補足セクション {idx}",
                        bullets=["補足ポイント", "データまたは事例", "次アクション"],
                    )
                )
            else:
                normalized.append(
                    OutlineSlide(
                        title=f"补充章节 {idx}",
                        bullets=["补充观点", "案例或数据", "建议动作"],
                    )
                )
        return normalized

    def _build_subtitle(self, req: GenerateRequest) -> str:
        preset = self._normalize_preset(req.preset)
        profile = PRESET_PROFILES.get(preset, {})
        items = [x for x in [profile.get("label"), req.audience, req.tone, req.language] if x]
        return " | ".join(items) if items else "Generated by ChatPPT"

    def _normalize_preset(self, preset: str | None) -> str | None:
        if not preset:
            return None
        return PRESET_ALIASES.get(preset.strip().lower())

    def _language_family(self, language: str) -> str:
        normalized = (language or "").lower()
        if normalized.startswith("en"):
            return "en"
        if normalized.startswith("ja"):
            return "ja"
        return "zh"

    def _default_style_for_language(self, language_family: str) -> tuple[str, str]:
        if language_family == "en":
            return "Professional", "General audience"
        if language_family == "ja":
            return "明確で実務的", "一般向け"
        return "专业", "通用受众"

    def _default_output_path(self, topic: str) -> str:
        slug = re.sub(r"[^a-zA-Z0-9\u4e00-\u9fff]+", "-", topic).strip("-") or "generated"
        output_dir = Path("generated")
        output_dir.mkdir(parents=True, exist_ok=True)
        return str(output_dir / f"{slug}.pptx")
