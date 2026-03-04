from openpyxl.workbook import Workbook

from app.models import Issue


RULE_ID = "format_check"
RULE_NAME = "格式校验"
RULE_DESC = "检查第二列是否仅包含数字（占位规则）。"


def run(workbook: Workbook) -> list[Issue]:
    issues: list[Issue] = []
    for sheet in workbook.worksheets:
        max_row = min(sheet.max_row, 300)
        for row in range(2, max_row + 1):
            value = sheet.cell(row=row, column=2).value
            if value is None:
                continue
            as_text = str(value).strip()
            if as_text and not as_text.isdigit():
                issues.append(
                    Issue(
                        rule_id=RULE_ID,
                        rule_name=RULE_NAME,
                        severity="warning",
                        sheet=sheet.title,
                        row=row,
                        column="B",
                        message="第二列预期为数字格式",
                    )
                )
    return issues
