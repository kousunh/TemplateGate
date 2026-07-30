"""DocGate — policy-as-code acceptance gate for AI-edited Office documents."""

__version__ = "0.1.0"

from .api import check, detect_target, diff, snapshot
from .core.model import Change, CheckResult, SemanticFinding, Violation
from .core.policy import Policy, PolicyError, load_policy

__all__ = [
    "__version__",
    "check",
    "diff",
    "snapshot",
    "detect_target",
    "Change",
    "CheckResult",
    "SemanticFinding",
    "Violation",
    "Policy",
    "PolicyError",
    "load_policy",
]
