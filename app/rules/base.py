from dataclasses import dataclass
from typing import Callable

from openpyxl.workbook import Workbook

from app.models import Issue


@dataclass
class Rule:
    id: str
    name: str
    description: str
    checker: Callable[[Workbook], list[Issue]]
