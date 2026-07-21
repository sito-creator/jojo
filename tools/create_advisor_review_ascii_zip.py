from __future__ import annotations

import zipfile
from pathlib import Path


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


def arcname(path: Path) -> Path:
    rel = path.relative_to(PACKAGE)
    parts = list(rel.parts)
    if parts[0] == "03_operation_word_templates":
        parts[-1] = WORD_TEMPLATE_NAMES.get(parts[-1], parts[-1])
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
