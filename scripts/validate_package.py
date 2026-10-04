#!/usr/bin/env python3
"""Check package integrity, local Markdown links, and blank-template metadata."""
import hashlib
import json
from pathlib import Path
import re
import sys
from urllib.parse import unquote
import xml.etree.ElementTree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]


def validate():
    errors = []
    manifest = json.loads((ROOT / "references/source-manifest.json").read_text())
    for entry in manifest["files"]:
        path = ROOT / entry["bundled_path"]
        if not path.is_file():
            errors.append(f"缺少文件：{entry['bundled_path']}")
            continue
        raw = path.read_bytes()
        if hashlib.sha256(raw).hexdigest() != entry["sha256"]:
            errors.append(f"快照散列不符：{entry['bundled_path']}")
        if "bytes" in entry and len(raw) != entry["bytes"]:
            errors.append(f"快照大小不符：{entry['bundled_path']}")
    templates = list((ROOT / "assets/templates").glob("*.docx"))
    if len(templates) != 4:
        errors.append("应恰好包含四份空白附件。")
    for path in templates:
        with ZipFile(path) as archive:
            if archive.testzip():
                errors.append(f"损坏的DOCX：{path.name}")
            for name in archive.namelist():
                if name.endswith((".xml", ".rels")):
                    ET.fromstring(archive.read(name))
            core = ET.fromstring(archive.read("docProps/core.xml"))
            for element in core:
                if element.tag.split("}")[-1] in {"creator", "lastModifiedBy", "lastPrinted"} and element.text:
                    errors.append(f"残留个人作者元数据：{path.name}")
    links = 0
    for path in ROOT.rglob("*.md"):
        if ".git" in path.parts:
            continue
        text = path.read_text()
        for target in re.findall(r"\[[^\]\n]*\]\(([^)\n]+)\)", text):
            target = target.strip("<>")
            if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*:", target) or target.startswith("#"):
                continue
            target = unquote(target.split("#")[0])
            links += 1
            if not (path.parent / target).exists():
                errors.append(f"断开的相对链接：{path.relative_to(ROOT)} → {target}")
    if errors:
        print("\n".join(errors), file=sys.stderr)
        return 1
    print(f"通过：{len(manifest['files'])} 项快照、4份DOCX及元数据、{links}个本地链接。")
    print("这不是申报内容、当年规则或Word版面的验收。")
    return 0


if __name__ == "__main__":
    raise SystemExit(validate())
