"""TemplateGate — policy-as-code acceptance gate for AI-edited Office documents."""

from importlib.metadata import PackageNotFoundError, version

try:
    # Single source of truth is pyproject.toml; a hardcoded copy here is what
    # let 0.1.1 and 0.1.2 both ship reporting the wrong number.
    __version__ = version("templategate")
except PackageNotFoundError:  # running from a source checkout, not installed
    __version__ = "0.0.0+unknown"

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
