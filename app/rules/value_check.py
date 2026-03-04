from openpyxl.workbook import Workbook

from app.models import Issue


RULE_ID = "value_check"
RULE_NAME = "取值范围校验"
RULE_DESC = "检查第三列数字是否在 0~100 区间（占位规则）。"


def run(workbook: Workbook) -> list[Issue]:
    issues: list[Issue] = []
    for sheet in workbook.worksheets:
        max_row = min(sheet.max_row, 300)
        for row in range(2, max_row + 1):
            value = sheet.cell(row=row, column=3).value
            if value is None:
                continue
            try:
                number = float(value)
            except (TypeError, ValueError):
                issues.append(
                    Issue(
                        rule_id=RULE_ID,
                        rule_name=RULE_NAME,
                        severity="info",
                        sheet=sheet.title,
                        row=row,
                        column="C",
                        message="第三列非数值，跳过范围检查",
                    )
                )
                continue

            if number < 0 or number > 100:
                issues.append(
                    Issue(
                        rule_id=RULE_ID,
                        rule_name=RULE_NAME,
                        severity="error",
                        sheet=sheet.title,
                        row=row,
                        column="C",
                        message="第三列数值超出 0~100 范围",
                    )
                )
    return issues
