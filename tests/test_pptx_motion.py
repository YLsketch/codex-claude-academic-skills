import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
from lxml import etree

SCRIPT = Path(__file__).resolve().parents[1] / "office-academic-skill" / "references" / "office-pptx" / "scripts" / "add_pptx_motion.py"
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}

class MotionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.source = self.root / "sample.pptx"
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        title = slide.shapes.add_textbox(Inches(.5), Inches(.5), Inches(8), Inches(.8))
        title.name = "Title"
        title.text_frame.text = "Research findings"
        body = slide.shapes.add_textbox(Inches(.5), Inches(1.5), Inches(8), Inches(2))
        body.name = "Key findings"
        body.text_frame.text = "First finding"
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(.5), Inches(4), Inches(3), Inches(1))
        card.name = "Conclusion card"
        card.text_frame.text = "Conclusion"
        image_path = self.root / "pixel.png"
        Image.new("RGB", (2, 2), "white").save(image_path)
        image_slide = prs.slides.add_slide(prs.slide_layouts[6])
        image_slide.shapes.add_picture(str(image_path), 0, 0, width=prs.slide_width, height=prs.slide_height)
        prs.save(self.source)

    def tearDown(self):
        self.tmp.cleanup()

    def slide_xml(self, path, number):
        with zipfile.ZipFile(path) as archive:
            return etree.fromstring(archive.read(f"ppt/slides/slide{number}.xml"))

    def test_native_transition_and_object_timing(self):
        output = self.root / "animated.pptx"
        result = subprocess.run([sys.executable, str(SCRIPT), str(self.source), "-o", str(output)], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr + result.stdout)
        self.assertIsNotNone(self.slide_xml(output, 1).find("./p:transition", NS))
        self.assertIsNotNone(self.slide_xml(output, 1).find("./p:timing", NS))
        self.assertIsNotNone(self.slide_xml(output, 2).find("./p:transition", NS))
        self.assertIsNone(self.slide_xml(output, 2).find("./p:timing", NS))
        self.assertIn("flattened full-slide picture detected", result.stdout)
        self.assertEqual(len(Presentation(output).slides), 2)

    def test_existing_animation_is_preserved(self):
        first = self.root / "first.pptx"
        second = self.root / "second.pptx"
        subprocess.run([sys.executable, str(SCRIPT), str(self.source), "-o", str(first)], check=True)
        subprocess.run([sys.executable, str(SCRIPT), str(first), "-o", str(second)], check=True)
        slide = self.slide_xml(second, 1)
        self.assertEqual(len(slide.findall("./p:transition", NS)), 1)
        self.assertEqual(len(slide.findall("./p:timing", NS)), 1)

    def test_source_is_not_overwritten(self):
        result = subprocess.run([sys.executable, str(SCRIPT), str(self.source), "-o", str(self.source)], capture_output=True, text=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("refusing to overwrite source", result.stderr)

if __name__ == "__main__":
    unittest.main()
