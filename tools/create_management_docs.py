from __future__ import annotations

from pathlib import Path

import create_word_regulations
from create_word_regulations import convert_markdown


ROOT = Path(__file__).resolve().parents[1]
DOCS = [
    (
        ROOT / "00_進行管理" / "未確定事項一覧.md",
        ROOT / "00_進行管理" / "未確定事項一覧.docx",
    ),
    (
        ROOT / "00_進行管理" / "上場準備マイルストーン.md",
        ROOT / "00_進行管理" / "上場準備マイルストーン.docx",
    ),
]


def main() -> None:
    create_word_regulations.PRESET["base_font"] = "Yu Gothic"
    create_word_regulations.PRESET["jp_font"] = "Yu Gothic"

    generated: list[Path] = []
    for md_path, out_path in DOCS:
        convert_markdown(md_path, out_path)
        generated.append(out_path)

    for path in generated:
        print(path)


if __name__ == "__main__":
    main()
