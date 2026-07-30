# TemplateGate

[English](README.md) | **日本語**

[![CI](https://github.com/kousunh/TemplateGate/actions/workflows/ci.yml/badge.svg)](https://github.com/kousunh/TemplateGate/actions/workflows/ci.yml)

**AIが編集した Excel・Word 文書のための、Policy-as-code 受入ゲート。**

Claude Code・Codex・ChatGPT などのAIエージェントは Excel / Word の編集が得意ですが、
ときどき数式・結合セル・印刷レイアウトを壊します。TemplateGate は
`.xlsx` / `.xlsm` / `.docx` のための**読み取り専用の回帰テスト**です。
編集前(Baseline)と編集後(Candidate)を比較し、
**ポリシーで許可した変更だけが行われたこと**を検証します。

- **Default Deny / Fail Closed** — 明示的に許可されていない変更はすべて違反。
- **決定的な構造検査** — セル値・数式・書式・結合セル・条件付き書式・
  データ検証・シート構造・画像(数/内容ハッシュ/位置/サイズ)・
  ヘッダーフッター・印刷設定・VBA・Wordの段落/表/セクション。
- **編集ツールが黙って捨てたものを検出** — グラフ・ピボットテーブル・
  コメント・埋め込みオブジェクト・カスタムXML、Excel の図形・
  テキストボックス、そして**それ以外のあらゆるパート**を、文書内部の
  パッケージから直接読み取ります。そのため、**編集に使ったツールが
  それらを扱えなかった場合でも**破損を検出できます。「ファイルは正常に
  開くし見た目も問題ない、ただしグラフが消えている」という壊れ方こそ、
  この検査が捉えるものです。未知のパートも例外ではありません
  (「知らない」は「許可」ではない)。
- **目に見えないものを検出** — 静かに非表示にされた行、解除された
  シート保護、手動計算に切り替えられたブック(全数式が古い値のまま表示
  される)、表示文字はそのままで飛び先だけ変えられたハイパーリンク、
  4pt 白文字にされた本文。変更履歴(トラックチェンジ)として行われた
  編集も見えます — 挿入はページ上のテキストなので、テキスト変更として
  報告されます。
- **意味解析は任意** — `off`(既定。文書内容は一切外部送信されない)/
  `review`(警告のみ)/ `gate`(合否判定に含める)。モデル・ベンダーは利用者が選択。
- **編集ツールではない** — TemplateGate は文書を一切変更せず、自動修復もしません。
  FAILした候補は破棄し、最終判断は人間が行います。

## 信頼境界

**文書を編集したエージェント自身に、許可範囲を決めさせてはいけません。**
エージェントはポリシー案を「提案」できますが、実際の検査は人間(またはCI)が
レビューして固定した **Trusted Policy** に対して実行されます。

これは仮定の話ではありません。Microsoft Research の測定では、最先端の LLM は
長い編集ワークフローの終わりまでに**文書内容の平均25%を破損**させ
(「まばらだが深刻な、静かに文書を壊すエラー」)、エージェントにツールを
持たせても改善しませんでした
([Laban, Schnabel & Neville, 2026](https://arxiv.org/abs/2604.15597))。
検証は、エージェントの外側に置く必要があります。

## インストール

```bash
pip install templategate
```

Python 3.10+。依存は openpyxl / python-docx / PyYAML のみ。

## クイックスタート

```bash
# 1. ポリシーの雛形を生成して編集
templategate init --target excel

# 2. AIには文書の「コピー」を編集させる

# 3. 結果を検査
templategate check \
  --baseline 計画書2026.xlsx \
  --candidate 計画書2026_編集後.xlsx \
  --policy templategate.policy.yaml \
  --report text
```

終了コード: `0`=PASS / `1`=FAIL / `2`=実行エラー

パッケージとしては開けるが参照しているパートを失っている候補は、
**「読めない」のではなく「破損している」**と扱われます。これはツールの
エラーではなく、欠落したパート名を挙げた上での FAIL です。exit 2 は
「そもそもファイルが開けない場合」と設定ミスのためだけに使われます。
`review_only` モードも `structural: ignore` も、**開けない文書を
通すことはできません**。

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
  charts: strict                  # グラフ・ピボットテーブル・コメント・
  pivot_tables: strict            # 埋め込みオブジェクト・カスタムXML・
  comments: strict                # Excelの図形/テキストボックスの消失もFAIL
  parts: strict                   # その他すべてのパートの消失も
  links: strict                   # ハイパーリンクの飛び先変更も
semantic:
  mode: "off"                     # off | review | gate
```

structural の各カテゴリは、書かなくても既定で `strict` です(行を消しても
無効にはなりません)。除外したい場合は `ignore` を指定します。

### ポリシーで指定できるもの

`templategate init` が生成する雛形には、全属性・全 structural キーが
コメント付きで書き出されます。陳腐化しない参照先として、まずそちらを
見てください。特に知っておく価値があるもの:

- **パッケージパート(両形式共通)** — `charts` / `pivot_tables` /
  `drawings` / `comments` / `embedded` / `custom_xml` に加えて
  `parts` と `links`。`parts` はそれ以外すべてを受け止める包括カテゴリで、
  TemplateGate が名前を知らないパートも含みます(消失は既定で破損扱い)。
  `links` はリレーションシップの外部参照先を比較するので、
  **表示文字はそのままで飛び先URLだけ差し替えられたハイパーリンク**を
  捕まえられます。
- **Excel** — `layout`(行・列の非表示とサイズ)、`protection`
  (シート/ブックの保護)、`sheet_settings`(手動計算への切り替えを含む)。
- **Word** — `paragraph_format` / `field` / `bookmark` /
  `content_control` / `revision`(変更履歴)、および内容は残ったまま
  位置だけ変わったブロックを示す `moved`。

VBAプロジェクトは `package#` 配下ではなく、従来どおり `vba` セレクタ・
`vba` 属性で指定します。

指定できるロケーション:

```yaml
protect:
  - selector: "package#*"                              # すべてのパッケージパート
  - selector: "package#charts:*"                       # カテゴリ単位
  - selector: "package#charts:xl/charts/chart1.xml"    # 特定のパートのみ
  - selector: "package#links:https://example.com/x"    # 特定の外部参照先
  - selector: "Sheet1!1:10"                            # 行(layout用)
  - selector: "'Q1!Q4'!A1"                             # 下記参照
```

Word のコンテンツコントロールとテキストボックスは `sdt1` / `textbox1`。
セル内の入れ子表は親のロケーションを引き継ぐため、`r1c1` セルに入れ子に
なった表の最初のセルは `table1!r1c1!table1!r1c1` となります。

シート名に `'` `!` `#` が含まれる場合は、Excel と同じ流儀でクォートします。
`'Q1!Q4'!A1` は「Q1!Q4」というシートのセル A1 を指し、クォートなしの
`Q1!Q4` は「Q1」というシートのセル Q4 を意味します。クォートが必要なのは
これらの文字を含む名前のときだけです。

### レポートの読み方

レポートは「何が変わったか」を名指しします。書式の違反は「書式が変わった」
ではなく `font.bold True -> False` のように表示されます。1箇所の編集で
以降がすべてずれた場合、テキスト/Markdown レポートは連鎖分を1行に
まとめます(例: `p4..p11: content shifted because 1 paragraph removed at p3`)。
JSON レポートは常に個々の変更をすべて保持し、`group` フィールドで
関連付けます。

その他のコマンド:

```bash
templategate diff --baseline a.xlsx --candidate b.xlsx   # ポリシーなしで全変更を一覧
templategate snapshot file.docx                           # 構造スナップショットをJSON出力
```

## FAILしたとき

TemplateGate は文書に一切書き込みません。自動修復モードは存在せず、
今後も追加しません — **検査対象を編集できるツールは、二重に信用しなければ
ならない**からです。復旧はあえてツールの外側に置かれており、必要なものは
すでに手元にあります。

**Baseline がバックアップそのものです。** Baseline は手つかずの原本であり、
FAIL した候補は使い捨てのファイルにすぎません。手順:

1. 違反を読む。何がどこで壊れたかが書かれています。
2. 候補は捨てる。修正しようとしない、そして何より**保存し直さない** —
   1回目の保存でパートを落としたものは、2回目でもっと落とします。
3. Baseline をもう一度コピーしてやり直す。その際、違反レポートを
   そのままエージェントへフィードバックとして渡します
   (JSONレポートはそのまま入力できるように作られています)。
4. 同じ違反が2回続くなら、そのツールではそのパートを保持できません。
   **ポリシーではなく手段を変えてください。**

CI では Baseline が Git 履歴に残るため、コミットごとに世代バックアップが
積み上がり、最後に PASS したバージョンといつでも比較できます。

**FAIL を通すためにポリシーを緩めてはいけません。** それは検出できた
エラーを、黙って通るエラーに変える行為であり、このツールが防ごうとして
いるものそのものです。

## Python API

```python
import templategate

result = templategate.check("baseline.xlsx", "candidate.xlsx", "policy.yaml")
if not result.passed:
    for v in result.violations:
        print(v.change.location, v.change.attribute, v.message)
```

## AIエージェント向け

`skills/office-document-regression/` に Agent Skill(Claude Code / Codex /
ChatGPT 互換)が入っています。「コピーを編集 → `templategate check` を実行 →
JSONレポートを解釈 → **FAILを通すためにポリシーを書き換えることは絶対にしない**」
という安全な作業手順をエージェントに教えます。

## GitHub Action

`action.yml` は `action/` サブディレクトリにあるため、`uses:` のパスにも
`/action` を含めます。TemplateGate 本体はランナーが自動で取得するので、
利用者側は**自分のリポジトリ**をチェックアウトするだけで構いません
(検査対象の文書とポリシーをディスク上に置くため)。

```yaml
- uses: actions/checkout@v4
- uses: actions/setup-python@v5
  with:
    python-version: "3.12"
- uses: kousunh/TemplateGate/action@v1
  with:
    baseline: docs/plan_baseline.xlsx
    candidate: docs/plan.xlsx
    policy: .templategate/plan.policy.yaml
```

検査がFAILするとジョブも失敗します。レポートは成否にかかわらず
ジョブサマリーに出力されます。出力は `passed`(`true` / `false`)と
`report-path` の2つ。ジョブを失敗させずに出力だけ読みたい場合は、
ステップに `continue-on-error: true` を付けてください。

## TemplateGate がやらないこと

- 編集・変換・自動修復([FAILしたとき](#failしたとき)を参照)
- 個別にモデル化されていないパート内部の変更は、「そのパートが変わった」
  としか報告されません(`word/header1.xml` が変更されたことは分かりますが、
  どのランが白文字になったかまでは分かりません)
- 別のOfficeアプリで保存し直した際の差分の正規化(Round-trip)
- PDF・見た目のリグレッション検査
- クラウドへのアップロード(semantic `off` なら通信ゼロ)

## ライセンス

MIT。[LICENSE](LICENSE) を参照。

TemplateGate は独立したオープンソースプロジェクトであり、Microsoft とは
無関係です。「Microsoft」「Office」「Excel」「Word」は Microsoft Corporation の
商標であり、本プロジェクトではファイル形式の識別のためにのみ言及しています。
