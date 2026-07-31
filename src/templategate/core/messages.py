"""Human-facing strings, in the two languages the gate speaks.

Everything a person reads goes through here.  Everything a machine reads —
the JSON report — does not: agents and CI parse that, so it stays English and
stable whatever ``--lang`` says.  The split is the whole design.  A report is
for the person deciding whether to accept an edit, and for the target
audience that person is often reading Japanese; the JSON is a contract.

Two deliberate non-choices.  There is no gettext machinery: with exactly two
locales, a dict keeps every string greppable and puts the English and the
Japanese on adjacent lines, where a translation drifting apart from its
original is obvious.  And nothing here is assembled by concatenation —
Japanese puts its particles and verbs where English puts neither, so a
sentence stitched from fragments is only ever a sentence in one of them.
Each message is a whole template with named holes.

The Japanese register is that of README.ja.md: plain 事務文書 です・ます,
addressed to somebody who has to decide whether a document may ship.  Not
literal translation — 「数式がベタ書きの値に置き換えられています」 is what the
situation is called in a Japanese office, and that is what it should say.
"""

from __future__ import annotations

DEFAULT_LANG = "en"
LANGUAGES = ("en", "ja")

# Names of things, kept out of the sentence templates so a term is spelled
# one way everywhere.  Attribute and category names are policy vocabulary —
# a reader types them into a YAML file — so they are NOT translated; only the
# prose around them is.
CATALOG: dict[str, dict[str, str]] = {
    "en": {
        # --- report frame ---
        "report.title": "TemplateGate: {status}",
        "report.pass": "PASS",
        "report.fail": "FAIL",
        "report.baseline": "baseline",
        "report.candidate": "candidate",
        "report.summary": ("changes: {total} total, {allowed} allowed, "
                           "{violations} violations"),
        "report.md.title": "## TemplateGate: {status}",
        "report.md.pass": "✅ PASS",
        "report.md.fail": "❌ FAIL",
        "report.md.baseline": "Baseline",
        "report.md.candidate": "Candidate",
        "report.md.summary": ("Changes: {total} total / {allowed} allowed / "
                              "{violations} violations"),
        "report.change_line": "{old} -> {new}",
        "report.violations": "Violations:",
        "report.violations.review_only":
            "Violations (mode: review_only — reported, not blocking):",
        "report.md.review_only":
            "> `mode: review_only` — reported for review, not blocking.",
        "report.warnings": "Policy warnings:",
        "report.md.warnings": "### Policy warnings",
        "report.semantic": "Semantic checks ({mode}):",
        "report.md.semantic": "### Semantic checks ({mode})",
        "report.md.columns": "| Severity | Location | Attribute | Old | New | Rule |",
        "report.cascade": ("content shifted because {reason} — {count} knock-on "
                           "changes collapsed; the JSON report lists each one"),
        "report.md.cascade": "{count} knock-on changes",

        # --- damage and degradation ---
        "damaged.line": "the {role} document is damaged: {reason}",
        "damaged.parts_only": "only its package parts could be compared",
        "damaged.parts_only.long": ("only its package parts could be compared — "
                                    "cell, paragraph and formatting contents "
                                    "were not"),
        "damaged.md.line": "> **The {role} document is damaged:** {reason}  ",
        "damaged.md.parts_only": "> Only its package parts could be compared.",
        "role.baseline": "baseline",
        "role.candidate": "candidate",
        "severity.error": "error",
        "severity.warning": "warning",

        # --- evaluation verdicts ---
        "verdict.protected": "protected attribute '{attribute}' changed",
        "verdict.not_allowed": ("change to '{attribute}' is not allowed by the "
                                "policy (default deny)"),

        # --- policy problems ---
        "policy.did_you_mean": " — did you mean {suggestion!r}?",
        "policy.valid_are": " (valid: {options})",
        "policy.unknown": "{where}unknown {what} {word!r}{hint}",
        "policy.rule_shape": "{context}: each rule must be a string or a mapping",
        "policy.attributes_shape": "{context}: 'attributes' must be a list of strings",
        "policy.not_found": "policy file not found: {path}",
        "policy.bad_yaml": "invalid YAML in {path}: {reason}",
        "policy.root_shape": "{path}: policy root must be a mapping",
        "policy.bad_target": "target: {value!r} is not a target{hint}",
        "policy.bad_mode": "mode: {value!r} is not a mode{hint}",
        "policy.structural_shape": "'structural' must be a mapping",
        "policy.bad_structural":
            "structural.{key}: {value!r} is not a setting{hint}",
        "policy.semantic_shape": "'semantic' must be a mapping",
        "policy.bad_semantic_mode":
            "semantic.mode: {value!r} is not a mode{hint}",
        "policy.bad_check":
            "semantic.checks entries must be strings or {{instruction: ...}}",
        "policy.must_be_list": "'{name}' must be a list",
        "policy.word.key": "key",
        "policy.word.policy_key": "policy key",
        "policy.word.attribute": "attribute",
        "policy.word.structural_key": "structural key",
        "cli.policy_error": "templategate: policy error: {reason}",
        "cli.error": "templategate: error: {reason}",
        # --- CLI ---
        "cli.warning": "templategate: warning: {message}",
        "cli.no_changes": "no changes",
        "cli.no_changes.degraded":
            "no differences among the parts that could be read",
        "cli.report_written": "report written to {path}",
        "cli.snapshot_written": "snapshot written to {path}",

        # --- liveness ---
        "liveness.dead_rule": ("{kind} rule {number} selector {selector!r} "
                               "matches nothing in the baseline: {complaint}"),
        "liveness.no_sheet": "no sheet named {sheet!r}; {listed}",
        "liveness.listed_none": "the document has none",
        "liveness.listed": "document has {names}",
        "liveness.listed_more": "document has {names}, ...",
        "liveness.no_paragraph.one":
            "the document has 1 body paragraph, so p{wanted} does not exist",
        "liveness.no_paragraph.many":
            "the document has {available} body paragraphs, so p{wanted} does "
            "not exist",
        "liveness.no_indexed.one":
            "the document has 1 {kind}, so {kind}{wanted} does not exist",
        "liveness.no_indexed.many":
            "the document has {available} {kind}s, so {kind}{wanted} does not "
            "exist",

        # --- change details ---
        "detail.formula_hardcoded":
            "formula replaced by a hardcoded value ({old} -> {new})",
        "detail.formula_changed": "formula changed ({old} -> {new})",
        "detail.formula_added": "formula added ({new})",
        "detail.cell_format": "cell format changed",
        "detail.merged": "cells merged (only the top-left value survives)",
        "detail.unmerged": "cells unmerged (the block splits back apart)",
        "detail.image_added": "image added at {where}",
        "detail.image_removed": "image removed from {where}",
        "detail.image_replaced": "a different image is here now",
        "detail.image_moved": "same image, moved from {old} to {new}",
        "detail.image_resized": "same image, resized from {old} to {new}",
        "detail.image_moved_resized":
            "same image, moved from {old_at} to {new_at} and resized from "
            "{old} to {new}",
        "detail.sheet_visibility": "sheet visibility changed",
        "detail.sheet_protection": "sheet protection changed",
        "detail.sheet_settings": "sheet settings changed",
        "detail.sheet_kind": "sheet kind changed",
        "detail.workbook_settings": "workbook settings changed",
        "detail.print_settings": "print settings changed",
        "detail.defined_name": "defined name changed",
        "detail.defined_name_sheet": "sheet-scoped defined name changed",
        "detail.header_footer": "header/footer changed",
        "detail.row_column": "row or column changed",
        "detail.still_hidden": "{body} (still hidden)",
        "detail.part_removed": "{category} part removed: {part}",
        "detail.part_added": "{category} part added: {part}",
        "detail.part_modified": "{category} part modified: {part}",
        "detail.text_changed": "text changed",
        "detail.paragraph_added": "paragraph added",
        "detail.paragraph_removed": "paragraph removed",
        "detail.paragraph_moved": "paragraph moved from {old} to {new}",
        "detail.run_format": "run formatting changed",
        "detail.paragraph_format": "paragraph formatting changed",
        "detail.style_changed": "style changed",
        "detail.section_changed": "section setup changed",
        "detail.markup_changed":
            "markup changed that no other attribute accounts for",
        "detail.revision": "an edit is recorded as a tracked change",
        "detail.field_changed": "field changed",
        "detail.content_control": "content control changed",
        "detail.and_more": "and {count} more",
        "detail.delta": "{subject}: {fields}",
        "detail.field_pair": "{name} {old} -> {new}",
        "detail.join": ", ",
        "detail.format_changed":
            "workbook saved as .{new} instead of .{old}",
        "detail.format_changed_vba":
            "workbook saved as .{new} instead of .{old}; a macro-enabled "
            "workbook saved as .xlsx loses its VBA project",
        "detail.sheet_removed": "sheet removed",
        "detail.sheet_added": "sheet added",
        "detail.sheet_renamed_to": "sheet renamed to {name!r}",
        "detail.sheet_renamed_from": "sheet renamed from {name!r}",
        "detail.sheet_removed_paired":
            "sheet removed (contents compared against {name!r}, which is not "
            "a rename)",
        "detail.sheet_added_paired":
            "sheet added (contents compared against {name!r}, which is not "
            "a rename)",
        "detail.join_clauses": "; ",
        "detail.join_passthrough": "{clauses}",
        "detail.shifted_here": "content shifted here by an earlier edit ({summary})",
        "shift.paragraphs.one": "1 paragraph",
        "shift.paragraphs.many": "{count} paragraphs",
        "shift.removed_at": "{what} removed at p{position}",
        "shift.added_at": "{what} added at p{position}",
        "shift.rewritten_at": "{what} rewritten at p{position}",
        "detail.image_rewritten": "same image, rewritten",
        "detail.sheet_moved": "sheet moved",
        "detail.print_settings_fields": "print settings changed: {fields}",
        "detail.field_code": "field code changed",
        "detail.bookmarks": "bookmarks changed",
        "detail.tracked_markup": "tracked-change markup changed",
        "detail.markup_other": "markup changed that no other check covers",
        "detail.content_control_props": "content control properties changed",
        "detail.table_style": "table style changed",
        "detail.table_geometry": "table widths, borders or shading changed",
        "detail.table_dimensions": "table dimensions changed",
        "detail.table_cell": "table cell changed",
        "detail.table_added": "table added",
        "detail.table_removed": "table removed",
        "detail.block_added": "{kind} added",
        "detail.block_removed": "{kind} removed",
        "detail.section_setup": "section/page setup changed",
        "detail.image_removed_replaced": "image removed or replaced",
        "detail.image_added_replaced": "image added or replaced",
        "detail.part_missing": "{name} is missing from the candidate",
        "detail.part_arrived": "{name} was added to the candidate",
        "size.unknown": "an unknown size",
        "size.unreadable": "an unreadable size",
        "size.cell_span": "a different cell span",
        "value.none": "none",
        "value.empty": "(empty)",
        "value.yes": "yes",
        "value.no": "no",
        "value.pixels": "{width}x{height} pixels",
        "value.cells": "{width}x{height}",
    },
    "ja": {
        # --- report frame ---
        "report.title": "TemplateGate: {status}",
        "report.pass": "合格",
        "report.fail": "不合格",
        "report.baseline": "変更前",
        "report.candidate": "変更後",
        "report.summary": ("変更点: 全{total}件 / 許可{allowed}件 / "
                           "違反{violations}件"),
        "report.md.title": "## TemplateGate: {status}",
        "report.md.pass": "✅ 合格",
        "report.md.fail": "❌ 不合格",
        "report.md.baseline": "変更前",
        "report.md.candidate": "変更後",
        "report.md.summary": ("変更点: 全{total}件 / 許可{allowed}件 / "
                              "違反{violations}件"),
        "report.change_line": "{old} → {new}",
        "report.violations": "ポリシー違反:",
        "report.violations.review_only":
            "ポリシー違反（mode: review_only のため、報告のみで停止はしません）:",
        "report.md.review_only":
            "> `mode: review_only` のため、報告のみで停止はしません。",
        "report.warnings": "ポリシーに関する警告:",
        "report.md.warnings": "### ポリシーに関する警告",
        "report.semantic": "意味チェック（{mode}）:",
        "report.md.semantic": "### 意味チェック（{mode}）",
        "report.md.columns": "| 重大度 | 位置 | 属性 | 変更前 | 変更後 | 規則 |",
        "report.cascade": ("{reason}により内容が繰り下がりました。"
                           "連鎖した変更{count}件をまとめています"
                           "（個々の変更は JSON レポートに記載）"),
        "report.md.cascade": "連鎖した変更{count}件",

        # --- damage and degradation ---
        "damaged.line": "{role}のファイルが破損しています: {reason}",
        "damaged.parts_only": "パッケージ内のパーツのみ比較しました",
        "damaged.parts_only.long": ("パッケージ内のパーツのみ比較しました。"
                                    "セル・段落・書式の内容は比較できていません"),
        "damaged.md.line": "> **{role}のファイルが破損しています:** {reason}  ",
        "damaged.md.parts_only": "> パッケージ内のパーツのみ比較しました。",
        "role.baseline": "変更前",
        "role.candidate": "変更後",
        "severity.error": "エラー",
        "severity.warning": "警告",

        # --- evaluation verdicts ---
        "verdict.protected": "保護属性 '{attribute}' が変更されています",
        "verdict.not_allowed": ("'{attribute}' の変更はポリシーで許可されていません"
                                "（既定で拒否）"),

        # --- policy problems ---
        "policy.did_you_mean": "（もしかして {suggestion!r} ですか？）",
        "policy.valid_are": "（指定できるのは {options} です）",
        "policy.unknown": "{where}{what} {word!r} は認識できません{hint}",
        "policy.rule_shape": "{context}: 各ルールは文字列かマッピングで書いてください",
        "policy.attributes_shape":
            "{context}: 'attributes' は文字列のリストで書いてください",
        "policy.not_found": "ポリシーファイルが見つかりません: {path}",
        "policy.bad_yaml": "{path} の YAML が不正です: {reason}",
        "policy.root_shape": "{path}: ポリシーの最上位はマッピングで書いてください",
        "policy.bad_target": "target: {value!r} は対象として指定できません{hint}",
        "policy.bad_mode": "mode: {value!r} はモードとして指定できません{hint}",
        "policy.structural_shape": "'structural' はマッピングで書いてください",
        "policy.bad_structural":
            "structural.{key}: {value!r} は設定値として指定できません{hint}",
        "policy.semantic_shape": "'semantic' はマッピングで書いてください",
        "policy.bad_semantic_mode":
            "semantic.mode: {value!r} はモードとして指定できません{hint}",
        "policy.bad_check":
            "semantic.checks の各項目は文字列か {{instruction: ...}} で書いてください",
        "policy.must_be_list": "'{name}' はリストで書いてください",
        "policy.word.key": "キー",
        "policy.word.policy_key": "ポリシーキー",
        "policy.word.attribute": "属性",
        "policy.word.structural_key": "structural のキー",
        "cli.policy_error": "templategate: ポリシーエラー: {reason}",
        "cli.error": "templategate: エラー: {reason}",
        # --- CLI ---
        "cli.warning": "templategate: 警告: {message}",
        "cli.no_changes": "変更はありません",
        "cli.no_changes.degraded":
            "読み取れたパーツの範囲では差分はありません",
        "cli.report_written": "レポートを {path} に書き出しました",
        "cli.snapshot_written": "スナップショットを {path} に書き出しました",

        # --- liveness ---
        "liveness.dead_rule": ("{kind} ルール {number} 番のセレクタ {selector!r} は"
                               "変更前のファイルのどこにも一致しません: {complaint}"),
        "liveness.no_sheet": "{sheet!r} というシートはありません。{listed}",
        "liveness.listed_none": "この文書には1つもありません",
        "liveness.listed": "この文書にあるのは {names} です",
        "liveness.listed_more": "この文書にあるのは {names} などです",
        "liveness.no_paragraph.one":
            "この文書の本文は1段落しかないため、p{wanted} は存在しません",
        "liveness.no_paragraph.many":
            "この文書の本文は{available}段落しかないため、p{wanted} は存在しません",
        "liveness.no_indexed.one":
            "この文書に {kind} は1つしかないため、{kind}{wanted} は存在しません",
        "liveness.no_indexed.many":
            "この文書に {kind} は{available}つしかないため、{kind}{wanted} は"
            "存在しません",

        # --- change details ---
        "detail.formula_hardcoded":
            "数式がベタ書きの値に置き換えられています（{old} → {new}）",
        "detail.formula_changed": "数式が変更されています（{old} → {new}）",
        "detail.formula_added": "数式が追加されています（{new}）",
        "detail.cell_format": "セルの書式が変更されています",
        "detail.merged": "セルが結合されています（左上のセルの値だけが残ります）",
        "detail.unmerged": "セルの結合が解除されています（結合前の状態に戻ります）",
        "detail.image_added": "{where} に画像が追加されています",
        "detail.image_removed": "{where} から画像が削除されています",
        "detail.image_replaced": "別の画像に差し替えられています",
        "detail.image_moved": "同じ画像が {old} から {new} に移動しています",
        "detail.image_resized": "同じ画像が {old} から {new} に変寸されています",
        "detail.image_moved_resized":
            "同じ画像が {old_at} から {new_at} に移動し、"
            "{old} から {new} に変寸されています",
        "detail.sheet_visibility": "シートの表示・非表示が変更されています",
        "detail.sheet_protection": "シートの保護が変更されています",
        "detail.sheet_settings": "シートの設定が変更されています",
        "detail.sheet_kind": "シートの種類が変更されています",
        "detail.workbook_settings": "ブックの設定が変更されています",
        "detail.print_settings": "印刷設定が変更されています",
        "detail.defined_name": "定義された名前が変更されています",
        "detail.defined_name_sheet": "シート単位の定義された名前が変更されています",
        "detail.header_footer": "ヘッダー・フッターが変更されています",
        "detail.row_column": "行または列が変更されています",
        "detail.still_hidden": "{body}（非表示のままです）",
        "detail.part_removed": "{category} のパーツが削除されています: {part}",
        "detail.part_added": "{category} のパーツが追加されています: {part}",
        "detail.part_modified": "{category} のパーツが変更されています: {part}",
        "detail.text_changed": "本文が変更されています",
        "detail.paragraph_added": "段落が追加されています",
        "detail.paragraph_removed": "段落が削除されています",
        "detail.paragraph_moved": "段落が {old} から {new} に移動しています",
        "detail.run_format": "文字書式が変更されています",
        "detail.paragraph_format": "段落書式が変更されています",
        "detail.style_changed": "スタイルが変更されています",
        "detail.section_changed": "セクション設定が変更されています",
        "detail.markup_changed":
            "他のどの属性にも該当しないマークアップが変更されています",
        "detail.revision": "変更履歴として記録された編集があります",
        "detail.field_changed": "フィールドが変更されています",
        "detail.content_control": "コンテンツコントロールが変更されています",
        "detail.and_more": "ほか{count}件",
        "detail.delta": "{subject}: {fields}",
        "detail.field_pair": "{name} {old} → {new}",
        "detail.join": "、",
        "detail.format_changed":
            "ブックが .{old} ではなく .{new} として保存されています",
        "detail.format_changed_vba":
            "ブックが .{old} ではなく .{new} として保存されています。"
            "マクロ有効ブックを .xlsx で保存すると VBA プロジェクトが失われます",
        "detail.sheet_removed": "シートが削除されています",
        "detail.sheet_added": "シートが追加されています",
        "detail.sheet_renamed_to": "シート名が {name!r} に変更されています",
        "detail.sheet_renamed_from": "シート名が {name!r} から変更されています",
        "detail.sheet_removed_paired":
            "シートが削除されています（内容は {name!r} と比較しましたが、"
            "名前の変更ではありません）",
        "detail.sheet_added_paired":
            "シートが追加されています（内容は {name!r} と比較しましたが、"
            "名前の変更ではありません）",
        "detail.join_clauses": "、",
        "detail.join_passthrough": "{clauses}",
        "detail.shifted_here": "前方の編集により内容が繰り下がっています（{summary}）",
        "shift.paragraphs.one": "段落1つ",
        "shift.paragraphs.many": "段落{count}つ",
        "shift.removed_at": "p{position} で{what}が削除",
        "shift.added_at": "p{position} に{what}が追加",
        "shift.rewritten_at": "p{position} で{what}が書き換え",
        "detail.image_rewritten": "同じ画像が書き直されています",
        "detail.sheet_moved": "シートの並び順が変更されています",
        "detail.print_settings_fields": "印刷設定が変更されています: {fields}",
        "detail.field_code": "フィールドコードが変更されています",
        "detail.bookmarks": "ブックマークが変更されています",
        "detail.tracked_markup": "変更履歴のマークアップが変更されています",
        "detail.markup_other": "他のどのチェックにも該当しないマークアップが変更されています",
        "detail.content_control_props": "コンテンツコントロールの設定が変更されています",
        "detail.table_style": "表のスタイルが変更されています",
        "detail.table_geometry": "表の列幅・罫線・網かけが変更されています",
        "detail.table_dimensions": "表の行数・列数が変更されています",
        "detail.table_cell": "表のセルが変更されています",
        "detail.table_added": "表が追加されています",
        "detail.table_removed": "表が削除されています",
        "detail.block_added": "{kind} が追加されています",
        "detail.block_removed": "{kind} が削除されています",
        "detail.section_setup": "セクション・ページ設定が変更されています",
        "detail.image_removed_replaced": "画像が削除または差し替えられています",
        "detail.image_added_replaced": "画像が追加または差し替えられています",
        "detail.part_missing": "{name} が変更後のファイルにありません",
        "detail.part_arrived": "{name} が変更後のファイルに追加されています",
        "size.unknown": "不明なサイズ",
        "size.unreadable": "読み取れないサイズ",
        "size.cell_span": "別のセル範囲",
        "value.none": "なし",
        "value.empty": "(空)",
        "value.yes": "はい",
        "value.no": "いいえ",
        "value.pixels": "{width}x{height} ピクセル",
        "value.cells": "{width}x{height}",
    },
}


def normalise(lang: str | None) -> str:
    """The locale we will actually speak.

    Anything unrecognised falls back to English rather than raising: a report
    that comes out in the wrong language is a much smaller problem than a gate
    that refuses to run because of a environment variable somebody exported
    years ago.
    """
    if not lang:
        return DEFAULT_LANG
    tag = str(lang).strip().replace("_", "-").lower()
    if not tag:
        return DEFAULT_LANG
    primary = tag.split("-")[0]
    return primary if primary in LANGUAGES else DEFAULT_LANG


class Translator:
    """Renders one message id in one language.

    Missing keys fall back to English and then to the id itself.  The
    completeness test is what keeps that from being load-bearing.
    """

    __slots__ = ("lang",)

    def __init__(self, lang: str | None = DEFAULT_LANG) -> None:
        self.lang = normalise(lang)

    def __call__(self, key: str, /, **args) -> str:
        template = CATALOG.get(self.lang, {}).get(key)
        if template is None:
            template = CATALOG[DEFAULT_LANG].get(key)
        if template is None:
            return key
        try:
            return template.format(**args)
        except (KeyError, IndexError):
            # A template asking for something the caller did not pass is a
            # bug, but printing the raw template beats crashing the report
            # somebody is waiting on.
            return template

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return f"Translator({self.lang!r})"


ENGLISH = Translator(DEFAULT_LANG)


class Detail(str):
    """A message that already reads as its own English text.

    Subclassing ``str`` is the point.  Diffing runs once and its output is
    read twice: as English, by the JSON contract and by every caller that
    treats a detail as the string it has always been, and as whatever
    ``--lang`` asked for, by the report a person sees.  Being a string means
    nothing downstream has to know this class exists — ``asdict``, ``json``,
    substring tests and equality all behave exactly as before — while the
    message id and its arguments ride along for anyone who does.

    That is what keeps the comparison layer ignorant of locale: it has no
    business knowing one, and it never has to.
    """

    __slots__ = ("key", "args")

    def __new__(cls, key: str, /, **args):
        obj = super().__new__(cls, ENGLISH(key, **{
            name: _render_arg(value, ENGLISH) for name, value in args.items()}))
        obj.key = key
        obj.args = args
        return obj

    def render(self, t: Translator) -> str:
        # An argument may itself be a message — an image's size inside the
        # sentence about the image — and it has to be put into words in the
        # same language as the sentence around it.
        return t(self.key, **{name: _render_arg(value, t)
                              for name, value in self.args.items()})

    def __deepcopy__(self, memo):
        return self

    def __reduce__(self):
        return (_rebuild_detail, (type(self), self.key, self.args))

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return f"Detail({self.key!r}, {self.args!r})"


def _rebuild_detail(cls, key, args):  # pragma: no cover - pickle support
    return cls(key, **args)


def _render_arg(value, t: Translator):
    if isinstance(value, Detail):
        return value.render(t)
    if isinstance(value, (list, tuple)) and any(isinstance(v, Detail) for v in value):
        return t("detail.join_clauses").join(str(_render_arg(v, t)) for v in value)
    return value


def message(key: str, /, **args) -> Detail:
    """One message id plus its arguments, reading as English until asked."""
    return Detail(key, **args)


def say(value, t: Translator) -> str:
    """Whatever this is, in the reader's language."""
    return value.render(t) if isinstance(value, Detail) else value


class DeltaDetail(Detail):
    """A named-field difference: "cell format changed: numfmt A -> B".

    Holds the fields as raw triples rather than a finished sentence, because
    the pieces around them all move between languages: the separator between
    items, the word for a value that is not set, and the "and 3 more" tail.
    """

    __slots__ = ("subject_key", "fields", "extra")

    def __new__(cls, subject_key: str, fields, extra: int = 0):
        fields = tuple(fields)
        obj = str.__new__(cls, _render_delta(ENGLISH, subject_key, fields, extra))
        obj.key = subject_key
        obj.args = {}
        obj.subject_key = subject_key
        obj.fields = fields
        obj.extra = extra
        return obj

    def render(self, t: Translator) -> str:
        return _render_delta(t, self.subject_key, self.fields, self.extra)

    def __reduce__(self):  # pragma: no cover - pickle support
        return (_rebuild_delta, (type(self), self.subject_key, self.fields,
                                 self.extra))

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return f"DeltaDetail({self.subject_key!r}, {self.fields!r}, {self.extra!r})"


def _rebuild_delta(cls, subject_key, fields, extra):  # pragma: no cover
    return cls(subject_key, fields, extra)


def _render_delta(t: Translator, subject_key: str, fields, extra: int) -> str:
    subject = t(subject_key)
    if not fields and not extra:
        return subject
    parts = [t("detail.field_pair", name=name,
               old=_value(before, t), new=_value(after, t))
             for name, before, after in fields]
    if extra:
        parts.append(t("detail.and_more", count=extra))
    return t("detail.delta", subject=subject, fields=t("detail.join").join(parts))


_MAX_VALUE = 60


def _value(value, t: Translator) -> str:
    """One side of a field difference, shortened and put into words."""
    if value is None:
        return t("value.none")
    text = value if isinstance(value, str) else repr(value)
    return text if len(text) <= _MAX_VALUE else text[:_MAX_VALUE] + "..."


def missing_keys() -> dict[str, set[str]]:
    """Which ids one locale has and another does not."""
    everything: set[str] = set()
    for table in CATALOG.values():
        everything |= set(table)
    return {lang: everything - set(table) for lang, table in CATALOG.items()}
