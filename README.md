# callcenter-report-automation

クレジットカード会社のコールセンター業務を想定した、架空データによる日報集計・グラフ化の自動化ポートフォリオです。

## 背景

クレジットカード会社で約20年、不正利用防止・オーソリ部門を中心にコールセンター業務に携わってきた実務経験をもとに、現場でよくある「日々の入電データを集計してレポート化する」作業を、Python（データ生成）とExcel VBA（集計・グラフ化）で自動化しました。
実データは一切使用せず、すべて架空のデータで構築しています。

## 構成
callcenter-report-automation/
├── generate_fake_data.py # 架空データ生成（Python）
├── data/
│ └── callcenter_report.csv # 生成されたCSVデータ
├── callcenter-report-automation.xlsm # 集計・グラフ化（Excel VBA）
└── README.md


処理は2段構えです。

1. **generate_fake_data.py**（Python）で、日勤/夜勤シフトごとの架空コールセンターデータをCSV出力
2. **callcenter-report-automation.xlsm**（Excel VBA）で、そのCSVを取り込み、集計・グラフ化

## データ項目

| 列 | 内容 |
|---|---|
| date | 日付 |
| shift | シフト区分（day＝日勤／night＝夜勤） |
| auth_request_count | 加盟店オーソリ取得入電数 |
| cardholder_inquiry_count | カード会員問い合わせ入電数 |
| lost_stolen_count | 紛失・盗難入電数 |

## 使い方

1. `generate_fake_data.py` を実行し、`data/callcenter_report.csv` を生成
2. `callcenter-report-automation.xlsm` を開き、マクロを有効化
3. `ImportCallCenterData` マクロを実行し、CSVをDataシートに取り込み
4. `CreateSummary` マクロを実行し、Summaryシートに集計表とグラフを自動生成

## できること（Summaryシート）

- **月別集計**：月ごとの各項目の合計・平均を折れ線グラフで可視化
- **曜日別集計**：曜日ごとの各項目の平均入電数を積み上げ棒グラフで可視化
- **項目別統計**：各項目の合計・平均・最大・最小を一覧化
- **シフト別集計**：日勤/夜勤ごとの各項目の合計・平均を、比較しやすいクラスター棒グラフで可視化

## 実装で詰まった点と対応

実務経験だけでは気づけなかった、VBA特有のハマりどころです。

- **CSV取り込み時の文字化け**：`Workbooks.Open`や`OpenText`では日本語が文字化けしたため、`QueryTables`経由で`TextFilePlatform=65001`（UTF-8）を明示して解決
- **Dictionary.Keysの型不一致**：`Scripting.Dictionary`の`.Keys`メソッドは文字列配列ではなくVariant配列を返す仕様のため、受け取り側の型宣言を`Variant`に修正
- **グラフの余分な系列**：`.SetSourceData`と`.SeriesCollection.NewSeries`を併用すると、自動生成される既定系列とインデックスがズレて空の系列が残るバグが発生。`SetSourceData`を使わず、`NewSeries`の戻り値を直接受け取る形に修正して解決

## 今後の拡張アイデア

- 時間帯別（hourly）集計：日次データより実態に即すが、複雑化を避けるため見送り中
- 複数月データへの対応：現状は1ヶ月分のみのため、月別推移グラフを実質的に機能させるには複数月分のデータ生成が必要

## 技術スタック

- Python（架空データ生成）
- Excel VBA（CSV取り込み・集計・グラフ化）