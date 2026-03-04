from openpyxl.workbook import Workbook

from app.models import Issue


RULE_ID = "required_check"
RULE_NAME = "必填校验"
RULE_DESC = "检查每个工作表的前 3 列是否存在空值（占位规则）。"


def run(workbook: Workbook) -> list[Issue]:
    issues: list[Issue] = []
    for sheet in workbook.worksheets:
        max_row = min(sheet.max_row, 200)
        for row in range(2, max_row + 1):
            for col_index in range(1, 4):
                value = sheet.cell(row=row, column=col_index).value
                if value is None or (isinstance(value, str) and value.strip() == ""):
                    issues.append(
                        Issue(
                            rule_id=RULE_ID,
                            rule_name=RULE_NAME,
                            severity="error",
                            sheet=sheet.title,
                            row=row,
                            column=chr(64 + col_index),
                            message=f"第 {col_index} 列存在空值",
                        )
                    )
    return issues
