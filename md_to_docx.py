#!/usr/bin/env python3
"""Collect a Markdown directory tree into a DOCX using a reference document."""

import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from urllib.parse import unquote, urlsplit
from zipfile import BadZipFile, ZipFile


class ConversionError(Exception):
    """An actionable conversion failure."""


def collect_markdown(folder):
    """Files first, then subdirectories; never follow symbolic links."""
    def sort_key(path):
        name = unicodedata.normalize("NFC", path.name)
        return name.casefold(), name, path.name

    entries = sorted(folder.iterdir(), key=sort_key)
    for entry in entries:
        if not entry.is_symlink() and entry.is_file() and entry.suffix.lower() == ".md":
            yield entry
    for entry in entries:
        if not entry.is_symlink() and entry.is_dir():
            yield from collect_markdown(entry)


def resolve_images(node, folder):
    """Resolve images against their own Markdown file, avoiding name collisions."""
    if isinstance(node, dict):
        if node.get("t") == "Image":
            target = node["c"][2]
            url = urlsplit(target[0])
            if not url.scheme and not url.netloc and url.path:
                target[0] = str((folder / unquote(url.path)).resolve())
        for value in node.values():
            resolve_images(value, folder)
    elif isinstance(node, list):
        for value in node:
            resolve_images(value, folder)


def run_pandoc(pandoc, arguments, content=None, cwd=None):
    result = subprocess.run(
        [pandoc, *arguments], input=content, cwd=cwd,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
    )
    if result.returncode:
        detail = result.stderr.decode("utf-8", errors="replace").strip()
        raise ConversionError(f"Pandoc: {detail or 'ошибка преобразования'}")
    return result.stdout


def convert(source, template, output, force=False):
    source, output = (Path(p).resolve() for p in (source, output))
    template = Path(template).resolve() if template is not None else None
    if not source.is_dir():
        raise ConversionError(f"Исходная папка не найдена: {source}")
    if template is not None:
        if not template.is_file() or template.suffix.lower() != ".docx":
            raise ConversionError(f"Укажите существующий DOCX-шаблон: {template}")
        try:
            with ZipFile(template) as archive:
                if not {"[Content_Types].xml", "word/document.xml", "word/styles.xml"}.issubset(archive.namelist()):
                    raise ConversionError(f"Некорректный DOCX-шаблон: {template}")
        except BadZipFile as exc:
            raise ConversionError(f"Некорректный DOCX-шаблон: {template}") from exc
    if output.suffix.lower() != ".docx":
        raise ConversionError("Выходной файл должен иметь расширение .docx")
    if output == template:
        raise ConversionError("Выходной файл не должен совпадать с шаблоном")
    if output.exists() and (not force or not output.is_file()):
        raise ConversionError(f"Выходной файл уже существует: {output}. Для замены используйте --force")
    files = list(collect_markdown(source))
    if not files:
        raise ConversionError(f"В папке и подпапках нет Markdown-файлов: {source}")
    pandoc = shutil.which("pandoc")
    if pandoc is None:
        raise ConversionError("Pandoc не найден. Установите его и добавьте в PATH (см. README.md)")

    document = None
    for path in files:
        try:
            content = path.read_text(encoding="utf-8-sig").encode("utf-8")
        except UnicodeError as exc:
            raise ConversionError(f"Файл должен быть в UTF-8: {path}") from exc
        try:
            part = json.loads(run_pandoc(
                pandoc, ["--from=markdown", "--to=json", "--fail-if-warnings"],
                content, cwd=path.parent,
            ))
        except ConversionError as exc:
            raise ConversionError(f"{path}: {exc}") from exc
        resolve_images(part["blocks"], path.parent)
        if document is None:
            document = {"pandoc-api-version": part["pandoc-api-version"], "meta": {}, "blocks": []}
        document["blocks"].extend(part["blocks"])

    output.parent.mkdir(parents=True, exist_ok=True)
    # Keep an existing output intact if conversion fails.
    with tempfile.TemporaryDirectory(prefix=".md-to-docx-", dir=output.parent) as temporary:
        result = Path(temporary) / "result.docx"
        arguments = [
            "--from=json", "--to=docx", "--standalone", "--fail-if-warnings",
            f"--output={result}",
        ]
        if template is not None:
            arguments.append(f"--reference-doc={template}")
        run_pandoc(pandoc, arguments, json.dumps(document, ensure_ascii=False).encode("utf-8"))
        if force:
            os.replace(result, output)
        else:
            # Atomic creation protects a file created during conversion too.
            os.link(result, output)
    return files


def main(argv=None):
    parser = argparse.ArgumentParser(description="Собрать Markdown-файлы из дерева папок в DOCX.")
    parser.add_argument("source", type=Path, help="папка с Markdown-файлами")
    parser.add_argument("-t", "--template", type=Path, help="DOCX-шаблон (по умолчанию — оформление Pandoc)")
    parser.add_argument("-o", "--output", required=True, type=Path, help="результирующий DOCX")
    parser.add_argument("--force", action="store_true", help="заменить существующий результат")
    args = parser.parse_args(argv)
    try:
        files = convert(args.source, args.template, args.output, args.force)
    except (ConversionError, OSError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1
    print(f"Создан {args.output.resolve()} (Markdown-файлов: {len(files)})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
