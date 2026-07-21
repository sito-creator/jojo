from __future__ import annotations

import math
import shutil
import subprocess
from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
PYTHON = Path("/Users/shoito/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3")
RENDER = Path("/Users/shoito/.codex/plugins/cache/openai-primary-runtime/documents/26.715.12143/skills/documents/render_docx.py")
DOCX_DIR = ROOT / "02_Word規程ドラフト"
QA_DIR = ROOT / "03_Word規程_QAレンダー"
CONTACT_DIR = QA_DIR / "_contact_sheets"


def render_all() -> list[Path]:
    if QA_DIR.exists():
        shutil.rmtree(QA_DIR)
    QA_DIR.mkdir(parents=True, exist_ok=True)
    rendered_pages: list[Path] = []
    for docx_path in sorted(DOCX_DIR.glob("*.docx")):
        out_dir = QA_DIR / docx_path.stem
        subprocess.run(
            [str(PYTHON), str(RENDER), str(docx_path), "--output_dir", str(out_dir)],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        rendered_pages.extend(sorted(out_dir.glob("*.png")))
    return rendered_pages


def make_contact_sheets(images: list[Path], per_sheet: int = 12) -> list[Path]:
    CONTACT_DIR.mkdir(parents=True, exist_ok=True)
    sheets: list[Path] = []
    thumb_w, thumb_h = 380, 500
    label_h = 54
    cols = 3
    rows = math.ceil(per_sheet / cols)

    for sheet_index in range(math.ceil(len(images) / per_sheet)):
        batch = images[sheet_index * per_sheet : (sheet_index + 1) * per_sheet]
        canvas = Image.new("RGB", (cols * thumb_w, rows * (thumb_h + label_h)), "white")
        draw = ImageDraw.Draw(canvas)
        for idx, path in enumerate(batch):
            img = Image.open(path).convert("RGB")
            img.thumbnail((thumb_w - 24, thumb_h - 24))
            x = (idx % cols) * thumb_w
            y = (idx // cols) * (thumb_h + label_h)
            canvas.paste(img, (x + (thumb_w - img.width) // 2, y + 12))
            label = f"{path.parent.name} / {path.name}"
            draw.text((x + 10, y + thumb_h + 8), label[:58], fill=(20, 20, 20))
        out_path = CONTACT_DIR / f"contact_sheet_{sheet_index + 1:02d}.png"
        canvas.save(out_path)
        sheets.append(out_path)
    return sheets


def main() -> None:
    pages = render_all()
    sheets = make_contact_sheets(pages)
    print(f"rendered_pages={len(pages)}")
    print(f"contact_sheets={len(sheets)}")
    print(f"qa_dir={QA_DIR}")


if __name__ == "__main__":
    main()
