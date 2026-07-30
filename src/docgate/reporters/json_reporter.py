from __future__ import annotations

import json

from ..core.model import CheckResult


def render_json(result: CheckResult) -> str:
    return json.dumps(result.to_dict(), ensure_ascii=False, indent=2, default=str)
