from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
PACKAGE = ROOT / "04_支援会社提出パッケージ"
ZIP_PATH = ROOT / "04_支援会社提出パッケージ.zip"
PYTHON_FONT = "Hiragino Sans"


def set_font(run, size=11, bold=False, color="000000") -> None:
    run.font.name = PYTHON_FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for key in ["ascii", "hAnsi", "eastAsia", "cs"]:
        r_fonts.set(qn(f"w:{key}"), PYTHON_FONT)


def configure_doc(doc: Document, title: str) -> None:
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1)
    section.right_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.header_distance = Inches(0.492)
    section.footer_distance = Inches(0.492)

    normal = doc.styles["Normal"]
    normal.font.name = PYTHON_FONT
    normal.font.size = Pt(11)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), PYTHON_FONT)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.10

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(14)
    run = p.add_run(title)
    set_font(run, size=18, bold=True, color="1F3A5F")

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = meta.add_run("株式会社MoMo / 上場準備レビュー資料 / 作成日：2026年7月18日")
    set_font(run, size=9, color="5B677A")


def heading(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    set_font(run, size=13, bold=True, color="2E74B5")


def para(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_font(run, size=11)


def bullet(doc: Document, text: str) -> None:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(-0.25)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.167
    run = p.add_run(text)
    set_font(run, size=11)


def shade_cell(cell, fill="F2F4F7") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def add_table(doc: Document, rows: list[list[str]]) -> None:
    table = doc.add_table(rows=len(rows), cols=max(len(row) for row in rows))
    table.style = "Table Grid"
    table.autofit = True
    for r_idx, row in enumerate(rows):
        for c_idx, value in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            if r_idx == 0:
                shade_cell(cell)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(value)
            set_font(run, size=9.5, bold=(r_idx == 0))
    doc.add_paragraph()


def create_review_memo(path: Path) -> None:
    doc = Document()
    configure_doc(doc, "上場準備資料 レビュー依頼メモ")

    heading(doc, "1. レビュー依頼の目的")
    para(doc, "株式会社MoMoの上場準備に向け、現時点で作成した社内規程、運用台帳、各種Word雛形について、上場支援会社の観点から不足事項、修正事項、優先順位をご確認いただきたい。")

    heading(doc, "2. 会社前提")
    add_table(doc, [
        ["項目", "内容"],
        ["会社名", "株式会社MoMo"],
        ["決算月", "6月"],
        ["想定市場", "TOKYO PRO Marketを経由し、グロース市場を目指す"],
        ["想定時期", "3年後。ただし、現在N-1期との関係は要整理"],
        ["現機関設計", "取締役2名、監査役なし"],
        ["代表取締役", "伊藤翔"],
        ["主管部署", "管理部"],
        ["主要システム", "マネーフォワード、Notion、Google Workspace、Google Sheets等"],
    ])

    heading(doc, "3. 提出物")
    add_table(doc, [
        ["区分", "内容", "格納先"],
        ["規程Word", "上場準備に必要な規程ドラフト26本", "01_regulations_word"],
        ["運用台帳", "規程管理、稟議、会議体、反社チェック、与信、契約、外注、個人情報・システム、関連当事者、固定資産、内部監査のExcel台帳", "02_operation_ledgers"],
        ["Word雛形", "議事録、稟議書、反社チェック記録票、与信審査票、契約審査票、内部監査調書等10本", "03_operation_word_templates"],
        ["進行管理", "ロードマップ、ヒアリングシート、内部監査計画、Google Sheets確認メモ", "04_project_management"],
    ])

    heading(doc, "4. 重点レビュー依頼事項")
    for item in [
        "TOKYO PRO Marketからグロース市場を目指す前提で、現時点の規程体系に不足がないか。",
        "取締役会設置会社化、監査役選任、追加取締役選任、定款変更・登記の順序と時期。",
        "現在N-1期と3年後上場予定の整合性。N期、N-1期、N-2期の定義。",
        "月次決算を翌月20日反映とする初期運用から、上場準備水準へ短縮するための目標水準。",
        "反社チェック、与信管理、契約審査、稟議、内部監査について、初年度に最低限運用すべき証跡。",
        "AI研修、AI伴走、AI受託開発、SaaS等に関する個人情報、顧客情報、AI入力情報、知的財産・営業秘密管理の不足。",
    ]:
        bullet(doc, item)

    heading(doc, "5. こちらで想定している次アクション")
    add_table(doc, [
        ["優先", "アクション", "目的"],
        ["1", "未確定論点の確認", "規程修正の前提を固める"],
        ["2", "反社チェック・稟議・会議体台帳にサンプル1件を入力", "運用開始イメージを作る"],
        ["3", "第1回内部監査の監査手続書・チェックリストを作成", "規程整備後の運用評価に備える"],
        ["4", "機関設計変更のタスク表作成", "取締役会設置、監査役選任、定款変更、登記を漏れなく進める"],
    ])

    doc.save(path)


def create_open_items(path: Path) -> None:
    doc = Document()
    configure_doc(doc, "未確定論点・確認事項リスト")
    heading(doc, "1. 優先確認事項")
    add_table(doc, [
        ["No", "論点", "現状", "確認したい内容", "優先度"],
        ["1", "上場スケジュール", "3年後上場予定、現在N-1期との情報が併存", "TOKYO PRO Market上場時期とグロース市場上場時期、N期定義", "高"],
        ["2", "機関設計", "取締役2名、監査役なし", "取締役会設置会社化、取締役追加、監査役選任の予定時期", "高"],
        ["3", "取締役会・経営会議", "経営会議はこれから設置", "開催頻度、議事録担当、月次報告資料、承認ルート", "高"],
        ["4", "月次決算", "翌月20日までに反映したい", "上場準備上の目標締め日、レビュー者、報告会議体", "高"],
        ["5", "反社チェック", "契約条項はあるが実査未実施", "利用サービス、対象範囲、頻度、既存取引先への遡及範囲", "高"],
        ["6", "与信管理", "Notionで取引先管理", "与信限度額、掛取引条件、未入金フォロー基準", "中"],
        ["7", "内部監査", "未実施", "内部監査担当、独立性、初年度監査範囲", "高"],
        ["8", "情報管理", "Google Workspace、Notion、マネーフォワード等を利用", "アカウント棚卸、退職時削除、AI入力禁止情報の運用", "中"],
        ["9", "関連当事者", "調査票・リスト未整備", "対象者範囲、年次調査、取締役会承認が必要な取引基準", "中"],
        ["10", "労務規程", "就業規則等は未確認", "就業規則、給与規程、ハラスメント防止、育児介護規程等の整備状況", "中"],
    ])

    heading(doc, "2. 支援会社へ確認したいレビュー観点")
    for item in [
        "現時点の規程26本の粒度が、TOKYO PRO Market準備段階として過不足ないか。",
        "グロース市場を見据えた場合に、今から追加すべき規程・細則・要領があるか。",
        "内部統制文書化に向け、業務記述書・フローチャート・RCMの作成順序はどこから始めるべきか。",
        "監査法人・J-Adviser・主幹事候補に提出する場合、先に整えるべき証跡は何か。",
    ]:
        bullet(doc, item)

    doc.save(path)


def copy_tree(src: Path, dst: Path) -> None:
    if dst.exists():
        shutil.rmtree(dst)
    shutil.copytree(src, dst)


def create_package() -> None:
    if PACKAGE.exists():
        shutil.rmtree(PACKAGE)
    PACKAGE.mkdir(parents=True, exist_ok=True)

    (PACKAGE / "00_review_request").mkdir()
    (PACKAGE / "01_regulations_word").mkdir()
    (PACKAGE / "02_operation_ledgers").mkdir()
    (PACKAGE / "03_operation_word_templates").mkdir()
    (PACKAGE / "04_project_management").mkdir()

    create_review_memo(PACKAGE / "00_review_request" / "review_request_memo.docx")
    create_open_items(PACKAGE / "00_review_request" / "open_items_for_review.docx")

    copy_tree(ROOT / "02_Word規程ドラフト_ASCIIファイル名", PACKAGE / "01_regulations_word")
    shutil.copy2(ROOT / "03_運用テンプレート" / "Excel台帳" / "上場準備_運用台帳セット.xlsx", PACKAGE / "02_operation_ledgers" / "ipo_operation_ledgers.xlsx")
    copy_tree(ROOT / "03_運用テンプレート" / "Word雛形", PACKAGE / "03_operation_word_templates")

    for src, name in [
        (ROOT / "00_進行管理" / "規程作成ロードマップ.md", "roadmap.md"),
        (ROOT / "00_進行管理" / "初回ヒアリングシート.md", "initial_hearing_sheet.md"),
        (ROOT / "00_進行管理" / "初年度内部監査計画_たたき台.md", "initial_internal_audit_plan.md"),
        (ROOT / "00_進行管理" / "Googleスプレッドシート確認メモ.md", "google_sheets_review_memo.md"),
    ]:
        if src.exists():
            shutil.copy2(src, PACKAGE / "04_project_management" / name)

    if ZIP_PATH.exists():
        ZIP_PATH.unlink()
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(PACKAGE.rglob("*")):
            if path.is_file():
                zf.write(path, path.relative_to(PACKAGE.parent))

    print(f"package={PACKAGE}")
    print(f"zip={ZIP_PATH}")
    print(f"files={sum(1 for p in PACKAGE.rglob('*') if p.is_file())}")


if __name__ == "__main__":
    create_package()
