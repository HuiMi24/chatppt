import os
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation

from app.chat_service import ChatPlanner
from app.generator_service import GeneratorService
from app.models import EditInstruction, GenerateRequest
from app.ppt_service import PPTService


class TestPPTServices(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.base = Path(self.tmpdir.name) / "sample.pptx"

        prs = Presentation()
        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title_slide.shapes.title.text = "Old Title"
        title_slide.placeholders[1].text = "Old Subtitle"
        prs.save(self.base)

        self.ppt = PPTService()
        self.chat = ChatPlanner(self.ppt)
        self.generator = GeneratorService(self.ppt)

    def tearDown(self):
        self.tmpdir.cleanup()

    def test_parse_and_edit(self):
        doc = self.ppt.parse_ppt(str(self.base))
        self.assertEqual(doc.slide_count, 1)
        self.assertTrue(len(doc.slides[0].shapes) >= 1)

        title_shape = next(s for s in doc.slides[0].shapes if s.role == "title")
        out = self.ppt.apply_edits(
            str(self.base),
            [
                EditInstruction(
                    slide_index=0,
                    shape_index=title_shape.shape_index,
                    new_text="New Title",
                )
            ],
        )
        out_doc = self.ppt.parse_ppt(out)
        self.assertEqual(out_doc.slides[0].title, "New Title")

    def test_rule_based_chat(self):
        plan, used_llm = self.chat.build_plan(str(self.base), "把第1页标题改为年度复盘")
        self.assertFalse(used_llm)
        self.assertEqual(len(plan), 1)

    def test_generate_fallback(self):
        os.environ.pop("OPENAI_API_KEY", None)

        req = GenerateRequest(
            topic="AI产品路线图",
            audience="管理层",
            tone="专业",
            slide_count=6,
            language="zh-CN",
            output_path=str(Path(self.tmpdir.name) / "generated.pptx"),
        )
        output_path, outline, theme = self.generator.generate(req)

        self.assertTrue(Path(output_path).exists())
        self.assertEqual(len(outline), 6)
        self.assertTrue(theme.name in {"corporate", "playful", "tech", "clean"})

        doc = self.ppt.parse_ppt(output_path)
        self.assertGreaterEqual(doc.slide_count, 7)  # cover + outline slides


if __name__ == "__main__":
    unittest.main()
