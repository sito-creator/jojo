from __future__ import annotations

import shutil
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT / "02_Word規程ドラフト"
OUT_DIR = ROOT / "02_Word規程ドラフト_ASCIIファイル名"
ZIP_PATH = ROOT / "02_Word規程ドラフト_ASCIIファイル名.zip"

NAME_MAP = [
    ("取締役会規程_ドラフト.docx", "01_board_meeting_rules_draft.docx"),
    ("監査役規程_ドラフト.docx", "02_corporate_auditor_rules_draft.docx"),
    ("経営会議規程_ドラフト.docx", "03_management_meeting_rules_draft.docx"),
    ("規程管理規程_ドラフト.docx", "04_regulation_control_rules_draft.docx"),
    ("組織規程_ドラフト.docx", "05_organization_rules_draft.docx"),
    ("業務分掌規程_ドラフト.docx", "06_responsibility_assignment_rules_draft.docx"),
    ("職務権限規程_ドラフト.docx", "07_authority_rules_draft.docx"),
    ("稟議規程_ドラフト.docx", "08_approval_request_rules_draft.docx"),
    ("予算管理規程_ドラフト.docx", "09_budget_control_rules_draft.docx"),
    ("経理規程_ドラフト.docx", "10_accounting_rules_draft.docx"),
    ("資金管理規程_ドラフト.docx", "11_cash_management_rules_draft.docx"),
    ("販売管理規程_ドラフト.docx", "12_sales_management_rules_draft.docx"),
    ("購買管理規程_ドラフト.docx", "13_purchasing_management_rules_draft.docx"),
    ("外注管理規程_ドラフト.docx", "14_outsourcing_management_rules_draft.docx"),
    ("与信管理規程_ドラフト.docx", "15_credit_management_rules_draft.docx"),
    ("反社会的勢力対応規程_ドラフト.docx", "16_anti_social_forces_rules_draft.docx"),
    ("内部監査規程_ドラフト.docx", "17_internal_audit_rules_draft.docx"),
    ("情報セキュリティ規程_ドラフト.docx", "18_information_security_rules_draft.docx"),
    ("システム管理規程_ドラフト.docx", "19_system_management_rules_draft.docx"),
    ("個人情報保護基本規程_ドラフト.docx", "20_personal_information_protection_rules_draft.docx"),
    ("リスク・コンプライアンス管理規程_ドラフト.docx", "21_risk_compliance_management_rules_draft.docx"),
    ("内部通報ポリシー_ドラフト.docx", "22_whistleblowing_policy_draft.docx"),
    ("クレーム管理規程_ドラフト.docx", "23_complaint_management_rules_draft.docx"),
    ("知的財産管理規程_ドラフト.docx", "24_intellectual_property_management_rules_draft.docx"),
    ("関連当事者取引管理規程_ドラフト.docx", "25_related_party_transaction_rules_draft.docx"),
    ("固定資産管理規程_ドラフト.docx", "26_fixed_asset_management_rules_draft.docx"),
    ("棚卸資産管理規程_ドラフト.docx", "27_inventory_asset_management_rules_draft.docx"),
    ("棚卸実施要領_ドラフト.docx", "28_inventory_count_procedures_draft.docx"),
    ("賃金規程_ドラフト.docx", "29_wage_rules_draft.docx"),
    ("人事評価規程_ドラフト.docx", "30_performance_evaluation_rules_draft.docx"),
    ("出張・旅費規程_ドラフト.docx", "31_travel_expense_rules_draft.docx"),
    ("執行役員規程_ドラフト.docx", "32_executive_officer_rules_draft.docx"),
    ("監査役監査規程_ドラフト.docx", "33_corporate_auditor_audit_rules_draft.docx"),
    ("監査役監査基準_ドラフト.docx", "34_corporate_auditor_audit_standards_draft.docx"),
    ("文書管理規程_ドラフト.docx", "35_document_management_rules_draft.docx"),
]


def main() -> None:
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    for source_name, target_name in NAME_MAP:
        shutil.copy2(SRC_DIR / source_name, OUT_DIR / target_name)

    if ZIP_PATH.exists():
        ZIP_PATH.unlink()
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(OUT_DIR.glob("*.docx")):
            zf.write(path, Path("advisor_word_regulations_ascii") / path.name)

    print(f"copied={len(NAME_MAP)}")
    print(f"out_dir={OUT_DIR}")
    print(f"zip={ZIP_PATH}")


if __name__ == "__main__":
    main()
