from __future__ import annotations

import zipfile
from pathlib import Path

from create_ascii_zip_for_advisor import NAME_MAP


ROOT = Path(__file__).resolve().parents[1]
PACKAGE = ROOT / "04_支援会社提出パッケージ"
ZIP_PATH = ROOT / "04_advisor_review_package_ascii.zip"

WORD_TEMPLATE_NAMES = {
    "01_取締役会議事録_テンプレート.docx": "01_board_meeting_minutes_template.docx",
    "02_経営会議議事録_テンプレート.docx": "02_management_meeting_minutes_template.docx",
    "03_稟議書_テンプレート.docx": "03_approval_request_form_template.docx",
    "04_反社チェック記録票_テンプレート.docx": "04_anti_social_forces_check_form_template.docx",
    "05_与信審査票_テンプレート.docx": "05_credit_review_form_template.docx",
    "06_契約審査記録票_テンプレート.docx": "06_contract_review_form_template.docx",
    "07_内部監査調書_テンプレート.docx": "07_internal_audit_workpaper_template.docx",
    "08_関連当事者調査票_テンプレート.docx": "08_related_party_questionnaire_template.docx",
    "09_システムアカウント申請書_テンプレート.docx": "09_system_account_request_form_template.docx",
    "10_クレーム報告書_テンプレート.docx": "10_complaint_report_template.docx",
}

FOLDER_NAMES = {
    "00_レビュー依頼": "00_review_request",
    "01_規程Word": "01_regulations_word",
    "02_運用台帳": "02_operation_ledgers",
    "03_Word雛形": "03_operation_word_templates",
    "04_進行管理": "04_project_management",
}

REVIEW_NAMES = {
    "レビュー依頼メモ.docx": "review_request_memo.docx",
    "未確定論点・確認事項リスト.docx": "open_items_for_review.docx",
}

LEDGER_NAMES = {
    "上場準備_運用台帳セット.xlsx": "ipo_operation_ledgers.xlsx",
}

PROJECT_NAMES = {
    "規程作成ロードマップ.md": "roadmap.md",
    "初回ヒアリングシート.md": "initial_hearing_sheet.md",
    "初年度内部監査計画_たたき台.md": "initial_internal_audit_plan.md",
    "Googleスプレッドシート確認メモ.md": "google_sheets_review_memo.md",
    "未確定事項一覧.md": "open_items_current.md",
    "未確定事項一覧.docx": "open_items_current.docx",
    "規程ひな形準拠チェックリスト.md": "template_conformance_checklist.md",
    "規程ひな形読み取り_ヒアリング整理.md": "template_reading_and_hearing_notes.md",
}

REGULATION_NAMES = dict(NAME_MAP)


def arcname(path: Path) -> Path:
    rel = path.relative_to(PACKAGE)
    parts = list(rel.parts)
    source_dir = parts[0]
    parts[0] = FOLDER_NAMES.get(parts[0], parts[0])
    if source_dir == "00_レビュー依頼":
        parts[-1] = REVIEW_NAMES.get(parts[-1], parts[-1])
    elif source_dir == "01_規程Word":
        parts[-1] = REGULATION_NAMES.get(parts[-1], parts[-1])
    elif source_dir == "02_運用台帳":
        parts[-1] = LEDGER_NAMES.get(parts[-1], parts[-1])
    elif source_dir == "03_Word雛形":
        parts[-1] = WORD_TEMPLATE_NAMES.get(parts[-1], parts[-1])
    elif source_dir == "04_進行管理":
        parts[-1] = PROJECT_NAMES.get(parts[-1], parts[-1])
    return Path("advisor_review_package") / Path(*parts)


def main() -> None:
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(PACKAGE.rglob("*")):
            if path.is_file():
                zf.write(path, arcname(path))
    print(ZIP_PATH)


if __name__ == "__main__":
    main()
