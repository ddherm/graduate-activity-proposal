#!/usr/bin/env python3
"""Create a new activity workspace from bundled blank templates; never overwrite."""
import argparse
from pathlib import Path
import shutil
import sys

ROOT = Path(__file__).resolve().parents[1]
KINDS = {"school": (1, 3, 4), "college": (2, 3, 4), "all": (1, 2, 3, 4)}


def prepare(dest, kind):
    sources = []
    for number in KINDS[kind]:
        matches = list((ROOT / "assets/templates").glob(f"附件{number} *.docx"))
        if len(matches) != 1:
            raise ValueError(f"附件{number}缺失或存在多个版本，请检查技能目录。")
        sources.append(matches[0])
    intake = ROOT / "examples/intake.md"
    if not intake.is_file():
        raise ValueError("缺少活动信息清单。")
    # Exclusive creation protects both existing work and concurrent invocations.
    dest.mkdir(parents=True, exist_ok=False)
    (dest / "附件").mkdir()
    (dest / "资料").mkdir()
    shutil.copyfile(intake, dest / "活动信息.md")
    for source in sources:
        shutil.copyfile(source, dest / "附件" / source.name)
    return sources


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dest", type=Path, required=True, help="尚不存在的新活动目录")
    parser.add_argument("--kind", choices=KINDS, required=True)
    args = parser.parse_args()
    try:
        copied = prepare(args.dest.expanduser(), args.kind)
    except (OSError, ValueError) as exc:
        print(f"未完成初始化：{exc}", file=sys.stderr)
        return 1
    print(f"已创建 {args.dest.expanduser().resolve()}，含 {len(copied)} 份空白附件。")
    print("请补充当年通知和活动信息；这些文件尚未填写，不是申报成稿。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
