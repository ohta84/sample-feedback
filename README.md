# sample-feedback-excel

Bカートの「サンプル評価回収CSV」と「受注CSV」から、メーカー別の「サンプルフィードバックExcel」を自動生成するツール／Claude Skill。

## これは何か

サンプル品を配布した後、どの会社が実際の注文につながったか（サンプル→注文の転換率）をまとめたExcelレポートを、
Bカートの2種類のCSVから自動生成します。

- サンプル提供社数 / 注文転換社数 / 注文転換率を自動集計
- 会社・納品日ごとに商品を束ねて見やすい表を作成
- 注文あり／なしを色分け表示
- 業態別の提供社数・転換率・評価傾向（OK/NG/検討中の概算件数）をシート2にまとめる
- メーカーコード（商品名冒頭の `[141]` 等）から会社名を自動解決（`data/maker_codes.csv`）

## 使い方（Webツール・GitHub Pages）

Claudeなしで、ブラウザだけで使える版が `index.html` です。CSVはどこにも送信されず、
ブラウザの中だけで処理してExcelをダウンロードします。

**公開する手順**
1. GitHubのリポジトリ画面で「Settings」→左メニューの「Pages」を開く
2. 「Build and deployment」の「Source」で「Deploy from a branch」を選択
3. 「Branch」を `main` / `/ (root)` にして「Save」
4. 数分待つと、ページ上部に公開URL（`https://ohta84.github.io/sample-feedback/` のような形）が表示される
5. そのURLを開けば、CSVをアップロードするだけでExcelがダウンロードできる

## 使い方（Claude Skillとして）

このリポジトリを `SKILL.md` ごとClaudeにアップロードすると、
「Bカートのサンプル評価回収CSVと受注CSVからサンプルフィードバックExcelを作って」といった依頼で自動的に使われます。

## 使い方（スクリプト単体）

```bash
pip install -r requirements.txt
python scripts/generate_feedback_excel.py \
  --eval-csv "サンプル評価回収.csv" \
  --order-csv "bcart_order.csv" \
  --maker-code "141" \
  --output "141東京堂様サンプルフィードバック.xlsx"
```

### オプション

| オプション | 必須 | 説明 |
|---|---|---|
| `--eval-csv` | ✅ | Bカートの「サンプル評価回収CSV」のパス |
| `--order-csv` | ✅ | Bカートの受注CSV（全項目版）のパス |
| `--maker-name` | - | タイトル・ファイル名に使うメーカー名。省略時は `--maker-code` と `data/maker_codes.csv` から自動解決 |
| `--maker-code` | - | 商品名冒頭 `[141]` のようなメーカーコード。カンマ区切りで複数可 |
| `--output` | ✅ | 出力するxlsxファイルのパス |

## CSVの仕様（前提）

- 受注CSVの「セット名」列に「サンプル」を含む行をサンプル注文として扱う
  （評価回収CSVには通常注文も混在しているため、受注CSVと突き合わせて判定する）
- 会社名に「ナゴミヤ」「なごみや」を含む行は自社テストとして除外
- 文字コードは cp932 / utf-8 を自動判定
- 2つのCSVは「受注番号＋商品名」をキーに結合する

## メーカーコード対応表

`data/maker_codes.csv` に「コード, 会社名」の対応表を持っている。
新しいメーカーコードが出てきたら随時追記していく（Webツール・スクリプトの両方から参照される）。

## ディレクトリ構成

```
sample-feedback-excel/
├── SKILL.md                        # Claude Skill定義
├── README.md
├── requirements.txt
├── index.html                      # ブラウザだけで動くWebツール（GitHub Pages用）
├── data/
│   └── maker_codes.csv             # メーカーコード→会社名の対応表
└── scripts/
    └── generate_feedback_excel.py  # 生成本体（Python版）
```

## 既知の制約 / TODO

- CSVの実際のヘッダー名がBカート標準と異なる場合、列の自動検出に失敗することがあります
  （その場合はエラーメッセージに詳細が表示されます）
- 業態別分析シートのOK/NG/検討中件数は、評価コメント内の文字列一致による概算値です
  （1コメントに複数商品分の評価が含まれる場合、正確な件数と一致しないことがあります）
- Webツール（index.html）とPythonスクリプトはロジックを別々に実装しているため、
  仕様変更時は両方を更新する必要があります

