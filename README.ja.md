# OfficeCheck

**AIが編集したOffice文書のための、Policy-as-code 受入ゲート。**

English version: [README.md](README.md)

Claude Code・Codex・ChatGPT などのAIエージェントは Excel / Word の編集が得意ですが、
ときどき数式・結合セル・印刷レイアウトを壊します。OfficeCheck は
`.xlsx` / `.xlsm` / `.docx` のための**読み取り専用の回帰テスト**です。
編集前(Baseline)と編集後(Candidate)を比較し、
**ポリシーで許可した変更だけが行われたこと**を検証します。

- **Default Deny / Fail Closed** — 明示的に許可されていない変更はすべて違反。
- **決定的な構造検査** — セル値・数式・書式・結合セル・条件付き書式・
  データ検証・シート構造・画像(数/内容ハッシュ/位置/サイズ)・
  ヘッダーフッター・印刷設定・VBA・Wordの段落/表/セクション。
- **意味解析は任意** — `off`(既定。文書内容は一切外部送信されない)/
  `review`(警告のみ)/ `gate`(合否判定に含める)。モデル・ベンダーは利用者が選択。
- **編集ツールではない** — OfficeCheck は文書を一切変更せず、自動修復もしません。
  FAILした候補は破棄し、最終判断は人間が行います。

## 信頼境界

**文書を編集したエージェント自身に、許可範囲を決めさせてはいけません。**
エージェントはポリシー案を「提案」できますが、実際の検査は人間(またはCI)が
レビューして固定した **Trusted Policy** に対して実行されます。

## インストール

```bash
pip install officecheck
```

Python 3.10+。依存は openpyxl / python-docx / PyYAML のみ。

## クイックスタート

```bash
# 1. ポリシーの雛形を生成して編集
officecheck init --target excel

# 2. AIには文書の「コピー」を編集させる

# 3. 結果を検査
officecheck check \
  --baseline 計画書2026.xlsx \
  --candidate 計画書2026_編集後.xlsx \
  --policy officecheck.policy.yaml \
  --report text
```

終了コード: `0`=PASS / `1`=FAIL / `2`=実行エラー

ポリシーの例:

```yaml
version: 1
target: excel
allow:
  - selector: "計画表!B4:B6"      # この範囲の値の変更だけ許可
    attributes: [value]
protect:
  - selector: "*"                 # 数式・印刷設定・VBAは常に保護
    attributes: [formula, print_settings, vba]
structural:
  sheets: strict                  # シートの追加/削除/非表示化はFAIL
  images: strict
semantic:
  mode: "off"                     # off | review | gate
```

その他のコマンド:

```bash
officecheck diff --baseline a.xlsx --candidate b.xlsx   # ポリシーなしで全変更を一覧
officecheck snapshot file.docx                           # 構造スナップショットをJSON出力
```

## Python API

```python
import officecheck

result = officecheck.check("baseline.xlsx", "candidate.xlsx", "policy.yaml")
if not result.passed:
    for v in result.violations:
        print(v.change.location, v.change.attribute, v.message)
```

## AIエージェント向け

`skills/office-document-regression/` に Agent Skill(Claude Code / Codex /
ChatGPT 互換)が入っています。「コピーを編集 → `officecheck check` を実行 →
JSONレポートを解釈 → **FAILを通すためにポリシーを書き換えることは絶対にしない**」
という安全な作業手順をエージェントに教えます。

## OfficeCheck がやらないこと

- 編集・変換・自動修復
- 別のOfficeアプリで保存し直した際の差分の正規化(Round-trip)
- PDF・見た目のリグレッション検査
- クラウドへのアップロード(semantic `off` なら通信ゼロ)

## ライセンス

MIT。[LICENSE](LICENSE) を参照。

OfficeCheck は独立したオープンソースプロジェクトであり、Microsoft とは
無関係です。「Microsoft」「Office」「Excel」「Word」は Microsoft Corporation の
商標であり、本プロジェクトではファイル形式の識別のためにのみ言及しています。
