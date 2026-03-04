from typing import Literal

from pydantic import BaseModel, Field


Severity = Literal["error", "warning", "info"]


class RuleMeta(BaseModel):
    id: str
    name: str
    description: str


class Issue(BaseModel):
    rule_id: str
    rule_name: str
    severity: Severity
    sheet: str = ""
    row: int | None = None
    column: str | None = None
    message: str


class RuleResult(BaseModel):
    rule_id: str
    rule_name: str
    error_count: int = 0
    warning_count: int = 0
    info_count: int = 0
    issues: list[Issue] = Field(default_factory=list)


class CheckResponse(BaseModel):
    summary: list[RuleResult]
    issues: list[Issue]
