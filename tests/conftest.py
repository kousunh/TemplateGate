import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "fixtures"))
from generate import generate_all  # noqa: E402

EXCEL_POLICY = """\
version: 1
target: excel
mode: normal_input
allow:
  - selector: "計画表!B4:B6"
    attributes: [value]
protect:
  - selector: "*"
    attributes: [formula, vba]
structural:
  sheets: strict
  images: strict
semantic:
  mode: "off"
"""

WORD_POLICY = """\
version: 1
target: word
mode: normal_input
allow:
  - selector: "body"
    attributes: [text]
protect:
  - selector: "*"
    attributes: [section, header_footer]
semantic:
  mode: "off"
"""


@pytest.fixture(scope="session")
def fixtures(tmp_path_factory) -> dict[str, Path]:
    out = tmp_path_factory.mktemp("docs")
    paths = generate_all(out)
    excel_policy = out / "excel.policy.yaml"
    excel_policy.write_text(EXCEL_POLICY, encoding="utf-8")
    word_policy = out / "word.policy.yaml"
    word_policy.write_text(WORD_POLICY, encoding="utf-8")
    paths["excel_policy"] = excel_policy
    paths["word_policy"] = word_policy
    return paths
