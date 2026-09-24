import base64
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from md_to_docx import ConversionError, collect_markdown, convert


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
ET.register_namespace("w", W[1:-1])


class TreeTests(unittest.TestCase):
    def test_files_before_recursive_directories(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            names = ["z.md", "A.MD", "a/z.md", "a/nested/a.md", "b/a.md", "ignored.txt"]
            for name in names:
                path = root / name
                path.parent.mkdir(parents=True, exist_ok=True)
                path.touch()
            (root / "loop").symlink_to(root, target_is_directory=True)
            (root / "link.md").symlink_to(root / "z.md")
            self.assertEqual(
                [str(path.relative_to(root)) for path in collect_markdown(root)],
                ["A.MD", "z.md", "a/z.md", "a/nested/a.md", "b/a.md"],
            )


@unittest.skipUnless(shutil.which("pandoc"), "Pandoc is required")
class ConversionTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.source = self.root / "Материалы с пробелами"
        self.source.mkdir()
        self.template = self.root / "template.docx"
        self.output = self.root / "output" / "result.docx"
        self.template.write_bytes(subprocess.check_output([
            "pandoc", "--print-default-data-file=reference.docx",
        ]))
        # A distinctive page margin proves that reference formatting is applied.
        with ZipFile(self.template) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        document = ET.fromstring(entries["word/document.xml"])
        section = document.find(f".//{W}sectPr")
        margin = section.find(f"{W}pgMar")
        if margin is None:
            margin = ET.SubElement(section, f"{W}pgMar")
        for name, value in {"left": "2345", "right": "1440", "top": "1440", "bottom": "1440", "header": "720", "footer": "720", "gutter": "0"}.items():
            margin.set(f"{W}{name}", value)
        entries["word/document.xml"] = ET.tostring(document)
        with ZipFile(self.template, "w") as archive:
            for name, data in entries.items():
                archive.writestr(name, data)

    def write(self, name, content):
        path = self.source / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")
        return path

    def test_real_docx_order_formatting_and_template(self):
        self.write("z.md", "Второй текст\n")
        self.write("A.md", "\ufeff# Первый заголовок\n\n**Жирный**\n\n- Пункт\n\n| A | B |\n|---|---|\n| 1 | 2 |\n")
        self.write("a/z.md", "Третий текст\n")
        self.write("a/nested/a.md", "Четвёртый текст\n")
        self.write("b/a.md", "Пятый текст\n")
        files = convert(self.source, self.template, self.output)
        self.assertEqual(len(files), 5)
        with ZipFile(self.output) as archive:
            document = ET.fromstring(archive.read("word/document.xml"))
        text = " ".join(node.text or "" for node in document.iter(f"{W}t"))
        markers = ["Первый", "Второй", "Третий", "Четвёртый", "Пятый"]
        positions = [text.index(marker) for marker in markers]
        self.assertEqual(positions, sorted(positions))
        self.assertIsNotNone(document.find(f".//{W}tbl"))
        self.assertIsNotNone(document.find(f".//{W}b"))
        self.assertEqual(document.find(f".//{W}pgMar").get(f"{W}left"), "2345")

    def test_images_relative_to_each_file(self):
        # Two distinct, valid one-pixel GIFs with the same basename.
        gif = base64.b64decode("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7")
        other = gif[:13] + b"\xff\x00\x00" + gif[16:]
        for folder, data in [("a", gif), ("b", other)]:
            path = self.write(f"{folder}/text.md", "![Картинка](image%20one.gif)")
            (path.parent / "image one.gif").write_bytes(data)
        convert(self.source, self.template, self.output)
        with ZipFile(self.output) as archive:
            media = [archive.read(name) for name in archive.namelist() if name.startswith("word/media/")]
        self.assertCountEqual(media, [gif, other])

    def test_empty_tree_and_invalid_template(self):
        with self.assertRaisesRegex(ConversionError, "нет Markdown"):
            convert(self.source, self.template, self.output)
        self.template.write_text("not docx")
        with self.assertRaisesRegex(ConversionError, "Некорректный"):
            convert(self.source, self.template, self.output)

    def test_existing_output_and_failed_conversion(self):
        self.write("text.md", "Текст")
        convert(self.source, self.template, self.output)
        original = self.output.read_bytes()
        with self.assertRaisesRegex(ConversionError, "уже существует"):
            convert(self.source, self.template, self.output)
        self.write("text.md", "![Нет изображения](missing.png)")
        with self.assertRaises(ConversionError):
            convert(self.source, self.template, self.output, force=True)
        self.assertEqual(self.output.read_bytes(), original)
        self.write("text.md", "Новый текст")
        convert(self.source, self.template, self.output, force=True)
        with ZipFile(self.output) as archive:
            self.assertIn("Новый текст", archive.read("word/document.xml").decode())

    def test_template_cannot_be_overwritten(self):
        self.write("text.md", "Текст")
        with self.assertRaisesRegex(ConversionError, "совпадать с шаблоном"):
            convert(self.source, self.template, self.template, force=True)

    def test_cli(self):
        self.write("text.md", "Текст")
        script = Path(__file__).resolve().parents[1] / "md_to_docx.py"
        result = subprocess.run([
            "python3", str(script), str(self.source), "-t", str(self.template),
            "-o", str(self.output),
        ], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue(self.output.is_file())

    def test_cli_without_template(self):
        self.write("text.md", "# Заголовок\n\nТекст без шаблона")
        self.template.unlink()
        script = Path(__file__).resolve().parents[1] / "md_to_docx.py"
        result = subprocess.run([
            "python3", str(script), str(self.source), "-o", str(self.output),
        ], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        with ZipFile(self.output) as archive:
            document = ET.fromstring(archive.read("word/document.xml"))
            self.assertIn("word/styles.xml", archive.namelist())
        text = " ".join(node.text or "" for node in document.iter(f"{W}t"))
        self.assertIn("Заголовок", text)
        self.assertIn("Текст без шаблона", text)


if __name__ == "__main__":
    unittest.main()
