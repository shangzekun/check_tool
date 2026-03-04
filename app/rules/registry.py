from app.models import RuleMeta
from app.rules.base import Rule
from app.rules.format_check import RULE_DESC as FORMAT_DESC
from app.rules.format_check import RULE_ID as FORMAT_ID
from app.rules.format_check import RULE_NAME as FORMAT_NAME
from app.rules.format_check import run as run_format
from app.rules.required_check import RULE_DESC as REQUIRED_DESC
from app.rules.required_check import RULE_ID as REQUIRED_ID
from app.rules.required_check import RULE_NAME as REQUIRED_NAME
from app.rules.required_check import run as run_required
from app.rules.value_check import RULE_DESC as VALUE_DESC
from app.rules.value_check import RULE_ID as VALUE_ID
from app.rules.value_check import RULE_NAME as VALUE_NAME
from app.rules.value_check import run as run_value

RULES: dict[str, Rule] = {
    REQUIRED_ID: Rule(
        id=REQUIRED_ID,
        name=REQUIRED_NAME,
        description=REQUIRED_DESC,
        checker=run_required,
    ),
    FORMAT_ID: Rule(
        id=FORMAT_ID,
        name=FORMAT_NAME,
        description=FORMAT_DESC,
        checker=run_format,
    ),
    VALUE_ID: Rule(
        id=VALUE_ID,
        name=VALUE_NAME,
        description=VALUE_DESC,
        checker=run_value,
    ),
}


def list_rule_meta() -> list[RuleMeta]:
    return [
        RuleMeta(id=rule.id, name=rule.name, description=rule.description)
        for rule in RULES.values()
    ]
