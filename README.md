# sample-feedback-excel

BカートCSV（`bcart_order.csv`）から、メーカー別の「サンプルフィードバックExcel」を自動生成するツール／Claude Skill。

## これは何か

サンプル品を配布した後、どの会社が実際の注文につながったか（サンプル→注文の転換率）を
まとめたExcelレポートを、Bカートの受注CSV1本から自動生成します。

- サンプル提供社数 / 注文転換社数 / 注文転換率を自動集計
- 会社・納品日ごとに商品を束ねて見やすい表を作成
- 注文あり／なしを色分け表示

## 使い方（Claude Skillとして）

このリポジトリを `SKILL.md` ごとClaudeにアップロードすると、
「Bカートの受注CSVからサンプルフィードバックExcelを作って」といった依頼で自動的に使われます。

## 使い方（スクリプト単体）

```bash
pip install -r requirements.txt
python scripts/generate_feedback_excel.py \
  --csv bcart_order.csv \
  --maker-name "○○" \
  --maker-code "121" \
  --output "○○様サンプルフィードバック.xlsx"
```

### オプション

| オプション | 必須 | 説明 |
|---|---|---|
| `--csv` | ✅ | Bカート受注CSVのパス |
| `--maker-name` | ✅ | タイトル・ファイル名に使うメーカー名 |
| `--maker-code` | - | 商品名冒頭 `[121]` のようなメーカーコード。カンマ区切りで複数可。省略時はCSV全行が対象 |
| `--output` | ✅ | 出力するxlsxファイルのパス |

## CSVの仕様（前提）

- BR列（69列目）が「セット名」。値に「サンプル」を含む行をサンプル注文として扱う
- 会社名に「ナゴミヤ」「なごみや」を含む行は自社テストとして除外
- 文字コードは cp932 / utf-8 を自動判定

## ディレクトリ構成

```
sample-feedback-excel/
├── SKILL.md                        # Claude Skill定義
├── README.md
├── requirements.txt
└── scripts/
    └── generate_feedback_excel.py  # 生成本体
```

## 既知の制約 / TODO

- CSVの実際のヘッダー名がBカート標準と異なる場合、列の自動検出に失敗することがあります
  （その場合はエラーメッセージに実際のヘッダー一覧が表示されます）
- 「サンプル用途」「サンプル評価結果」列はCSVに含まれないため空欄で出力し、手入力する想定です
- 実データでの動作確認はまだ行っていません。実際のCSVで一度テストしてから本番運用してください
