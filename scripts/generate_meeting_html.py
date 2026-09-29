#!/usr/bin/env python3
"""
Meeting/Template/MeetingTemplate.md のレイアウトに沿って、
Meeting/{year}/ 配下の会議録 Markdown を HTML に変換し、
Meeting/HTML/ 配下へ同名の .html として出力するスクリプト。

使い方:
    python3 scripts/generate_meeting_html.py --year 2026
    python3 scripts/generate_meeting_html.py --year 2026 --files 2026-09-29_検査システム開発計画打合せ.md
"""

from __future__ import annotations

import argparse
import html
import re
import sys
from pathlib import Path

import markdown

REPO_ROOT = Path(__file__).resolve().parent.parent
MEETING_DIR = REPO_ROOT / "Meeting"
TEMPLATE_CSS = MEETING_DIR / "Template" / "meeting.css"
HTML_DIR = MEETING_DIR / "HTML"

MD_EXTENSIONS = ["tables", "sane_lists", "nl2br"]

# 見出しテキスト -> (絵文字, 種別) の対応表。
# 種別によって HTML への包み方(枠のスタイル)を変える。
SECTION_STYLES = {
    "会議概要": ("📋", "summary"),
    "概要": ("📋", "summary"),
    "決定事項": ("✅", "decision"),
    "課題": ("⚠️", "plain"),
    "todo": ("🔥", "todo"),
    "スケジュール": ("📅", "plain"),
    "メモ": ("📝", "plain"),
    "次回までのアクション": ("📌", "action"),
    "担当者": ("👤", "action"),
    "期限整理": ("⏰", "action"),
    "まとめ": ("🧾", "plain"),
}
DEFAULT_STYLE = ("📄", "plain")

META_HEADINGS = {"日時", "開催日時", "日付"}
PARTICIPANT_HEADINGS = {"参加者"}
LOCATION_HEADINGS = {"場所", "会場"}

HR_RE = re.compile(r"^\s*-{3,}\s*$")
HEADING_RE = re.compile(r"^(#{1,6})\s*(.*?)\s*$")
DATE_RE = re.compile(r"^(\d{4})[-/](\d{2})[-/](\d{2})$")


class Node:
    def __init__(self, level: int, title: str):
        self.level = level
        self.title = title
        self.body_lines: list[str] = []
        self.children: list["Node"] = []

    @property
    def body(self) -> str:
        return "\n".join(self.body_lines).strip()

    def is_empty(self) -> bool:
        return not self.title.strip() and not self.body.strip() and not self.children


def parse_outline(text: str) -> Node:
    """Markdown を見出しレベルに応じた木構造へ変換する。"""
    root = Node(level=0, title="")
    stack = [root]
    for line in text.splitlines():
        if HR_RE.match(line):
            continue
        m = HEADING_RE.match(line)
        if m:
            level = len(m.group(1))
            title = m.group(2).strip()
            node = Node(level=level, title=title)
            while stack[-1].level >= level:
                stack.pop()
            stack[-1].children.append(node)
            stack.append(node)
        else:
            stack[-1].body_lines.append(line)
    return root


def md_to_html(text: str) -> str:
    text = text.strip()
    if not text:
        return ""
    return markdown.markdown(text, extensions=MD_EXTENSIONS)


def normalize_heading(title: str) -> str:
    return title.strip().lstrip("#").strip()


def format_date(raw: str) -> str:
    raw = raw.strip()
    m = DATE_RE.match(raw)
    if m:
        return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
    return raw


def parse_participants(node: Node) -> list[str]:
    names = []
    for line in node.body_lines:
        line = line.strip()
        m = re.match(r"^[-*]\s+(.*)$", line)
        if m:
            names.append(m.group(1).strip())
        elif line:
            names.append(line)
    return names


def render_meta_block(title_node: Node) -> str:
    date_text = ""
    location_text = ""
    participants: list[str] = []

    for child in title_node.children:
        key = normalize_heading(child.title)
        if key in META_HEADINGS:
            date_text = format_date(child.body)
        elif key in LOCATION_HEADINGS:
            location_text = child.body.strip()
        elif key in PARTICIPANT_HEADINGS:
            participants = parse_participants(child)

    lines = ['<blockquote class="meeting-meta">']
    parts = []
    if date_text:
        parts.append(f"<p>日時：{html.escape(date_text)}</p>")
    if location_text:
        parts.append(f"<p>場所：{html.escape(location_text)}</p>")
    if participants:
        parts.append("<p>参加者：</p>")
        parts.append("<ul>")
        for name in participants:
            parts.append(f"<li>{html.escape(name)}</li>")
        parts.append("</ul>")
    lines.extend(parts)
    lines.append("</blockquote>")
    return "\n".join(lines)


def render_section(node: Node) -> str:
    key = normalize_heading(node.title)
    lookup_key = key.lower() if key.lower() == "todo" else key
    emoji, style = SECTION_STYLES.get(lookup_key, DEFAULT_STYLE)

    out = [f"<h2>{emoji} {html.escape(node.title)}</h2>"]

    if node.children:
        inner_parts = []
        for child in node.children:
            if child.is_empty():
                continue
            child_html = md_to_html(child.body)
            inner_parts.append(
                f"<h3>{html.escape(child.title)}</h3>\n{child_html}"
            )
        inner_html = "\n".join(inner_parts)
        # セクション本文(子見出しの前にある文章)があれば、先頭に含める
        preface_html = md_to_html(node.body)
        if preface_html:
            inner_html = preface_html + "\n" + inner_html
        if style in ("summary", "decision", "todo"):
            if style == "decision":
                # 決定事項は子見出し単位で個別の枠にする
                decision_blocks = []
                if preface_html:
                    decision_blocks.append(f'<div class="decision">{preface_html}</div>')
                for child in node.children:
                    if child.is_empty():
                        continue
                    child_html = md_to_html(child.body)
                    decision_blocks.append(
                        f'<div class="decision">\n<h3>{html.escape(child.title)}</h3>\n{child_html}\n</div>'
                    )
                out.append("\n".join(decision_blocks))
            else:
                out.append(f'<div class="{style}">\n{inner_html}\n</div>')
        else:
            out.append(inner_html)
    else:
        body_html = md_to_html(node.body)
        if style in ("summary", "decision", "todo"):
            out.append(f'<div class="{style}">\n{body_html}\n</div>')
        else:
            out.append(body_html)

    return "\n".join(out)


def build_html_document(md_text: str, css_text: str, source_name: str) -> str:
    root = parse_outline(md_text)
    if not root.children:
        raise ValueError(f"見出しが見つかりませんでした: {source_name}")

    title_node = root.children[0]
    title_text = title_node.title or source_name

    body_parts = [f"<h1>{html.escape(title_text)}</h1>"]
    body_parts.append(render_meta_block(title_node))

    for section in root.children[1:]:
        if section.is_empty():
            continue
        body_parts.append(render_section(section))

    body_html = "\n".join(body_parts)

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{html.escape(title_text)}</title>
{css_text.strip()}
</head>
<body>
{body_html}
</body>
</html>
"""


def convert_file(md_path: Path, css_text: str) -> Path:
    md_text = md_path.read_text(encoding="utf-8")
    html_text = build_html_document(md_text, css_text, md_path.name)
    HTML_DIR.mkdir(parents=True, exist_ok=True)
    out_path = HTML_DIR / (md_path.stem + ".html")
    out_path.write_text(html_text, encoding="utf-8")
    return out_path


def main() -> int:
    parser = argparse.ArgumentParser(description="会議録 Markdown を HTML に変換します。")
    parser.add_argument("--year", required=True, help="Meeting/ 配下の年度フォルダ名 (例: 2026)")
    parser.add_argument(
        "--files",
        nargs="*",
        default=None,
        help="変換対象の Markdown ファイル名 (省略時は年度フォルダ内の全 .md)",
    )
    args = parser.parse_args()

    year_dir = MEETING_DIR / args.year
    if not year_dir.is_dir():
        print(f"指定した年度フォルダが見つかりません: {year_dir}", file=sys.stderr)
        return 1

    if not TEMPLATE_CSS.is_file():
        print(f"CSS テンプレートが見つかりません: {TEMPLATE_CSS}", file=sys.stderr)
        return 1
    css_text = TEMPLATE_CSS.read_text(encoding="utf-8")

    if args.files:
        md_paths = [year_dir / name for name in args.files]
    else:
        md_paths = sorted(year_dir.glob("*.md"))

    if not md_paths:
        print(f"対象の Markdown ファイルが見つかりません: {year_dir}", file=sys.stderr)
        return 0

    for md_path in md_paths:
        if not md_path.is_file():
            print(f"ファイルが見つかりません: {md_path}", file=sys.stderr)
            continue
        out_path = convert_file(md_path, css_text)
        print(f"生成しました: {out_path.relative_to(REPO_ROOT)}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
