# callcenter-report-automation

クレジットカード会社のコールセンター業務を想定した、架空データによる日報集計・グラフ化の自動化ポートフォリオです。

## 背景

コールセンターのSVやクレジットカード会社の不正利用防止業務など、約20年にわたり現場で働いてきました。その経験をもとに、コールセンターでよくある「日々の入電データを集計してレポート化する」作業を、Python（データ生成）とExcel VBA（集計・グラフ化）で自動化しました。

実データは一切使用せず、すべて架空のデータで構築しています。

## 構成
```
callcenter-report-automation/
├── generate_fake_data.py              # 架空データ生成（Python）
├── data/
│   ├── callcenter_report.csv          # 生成された日次・シフト別CSVデータ
│   └── callcenter_hourly_report.csv   # 生成された時間帯別CSVデータ
├── callcenter-report-automation.xlsm  # 集計・グラフ化（Excel VBA）
└── README.md
```


処理は2段構えです。

1. **generate_fake_data.py**（Python）で、日勤/夜勤シフトごとの架空コールセンターデータ（日次）と、0〜23時の時間帯別データをそれぞれCSV出力（直近1年分＝365日）
2. **callcenter-report-automation.xlsm**（Excel VBA）で、その2つのCSVを取り込み、集計・グラフ化

## データ項目

**日次・シフト別（callcenter_report.csv）**

| 列 | 内容 |
|---|---|
| date | 日付 |
| shift | シフト区分（日勤／夜勤） |
| auth_request_count | 加盟店オーソリ取得入電数 |
| cardholder_inquiry_count | カード会員問合せ入電数 |
| lost_stolen_count | 紛失・盗難入電数 |

**時間帯別（callcenter_hourly_report.csv）**

コールセンターの日報でよく見られる「着信数・応答数・応答率・計画値・計画差異」という指標構成を参考にしています。

| 列 | 内容 |
|---|---|
| date | 日付 |
| hour | 時刻（0〜23） |
| shift | シフト区分（日勤／夜勤、時刻から自動判定） |
| auth_request_count | 加盟店オーソリ取得入電数（時間帯別） |
| cardholder_inquiry_count | カード会員問合せ入電数（時間帯別） |
| lost_stolen_count | 紛失・盗難入電数（時間帯別） |
| call_count | 着信数（実績） |
| planned_call_count | 着信計画 |
| call_variance | 着信計画差異（実績－計画） |
| answer_rate | 応答率（実績） |
| answered_count | 応答数（実績） |
| planned_answered_count | 応答数計画（目標応答率95%固定） |
| answered_variance | 応答数計画差異（実績－計画） |

## 使い方

1. `generate_fake_data.py` を実行し、`data/callcenter_report.csv` と `data/callcenter_hourly_report.csv` を生成（直近1年分・365日）
2. `callcenter-report-automation.xlsm` を開き、マクロを有効化
3. `ImportCallCenterData` マクロを実行し、日次CSVをDataシートに取り込み
4. `ImportHourlyData` マクロを実行し、時間帯別CSVをHourlyDataシートに取り込み
5. `CreateSummary` マクロを実行し、Summaryシートに集計表とグラフを自動生成（日次系・時間帯別系の集計をまとめて実行）

## Settingsシートについて

SLA目標や各種閾値は、Settingsシートに名前付き範囲として登録しており、コードを直接書き換えずに調整できます。

| 名前付き範囲 | 内容 | 初期値 |
|---|---|---|
| SLA_Target | SLA目標の応答率 | 90% |
| Inbound_Upper_Threshold | 入電上振れとみなす増加率の閾値 | +20% |
| Min_Diff_Hourly | 入電上振れ判定に使う、時間帯あたりの最小件数差 | 50件 |
| Min_Diff_Daily | （将来の日次判定用）最小件数差 | 200件 |
| Min_Missed_Hourly | SLA未達のうち「要確認」とする、取りこぼし件数の最小値 | 2件 |

## できること（Summaryシート）

- **分析コメント自動生成**：SLA未達時間帯数（応答率がSettingsシートの目標値を下回った時間帯の件数、取りこぼし件数に応じて「要確認」「軽微」に分類）と、入電上振れ時間帯数（計画に対して一定割合・一定件数以上入電が増えた時間帯の件数）を自動集計し、Summaryシート上部に表示
- **時間帯別集計**：0〜23時それぞれの1年平均を算出し、着信数実績（棒グラフ）と着信計画（折れ線）を並べた複合グラフで、1日の呼量カーブと計画とのズレを可視化
- **曜日別集計**：曜日ごとの各項目の平均入電数を積み上げ棒グラフで可視化
- **月別集計**：月ごとの各項目の合計・平均を折れ線グラフで可視化（1年分のため12ヶ月分の推移が見える）
- **シフト別集計**：日勤/夜勤ごとの各項目の合計・平均を、比較しやすいクラスター棒グラフで可視化
- 全グラフで色分けを統一（オーソリ＝青、会員問合せ＝赤、紛失盗難＝緑）し、縦軸の目盛りも実データの範囲から自動計算（3〜5本程度に抑えて読みやすく、3桁ごとにカンマ表示）

## 実装で詰まった点と対応

実装中に気づいた、VBA特有のハマりどころです。

- **CSV取り込み時の文字化け**：`Workbooks.Open`や`OpenText`では日本語が文字化けしたため、`QueryTables`経由で`TextFilePlatform=65001`（UTF-8）を明示して解決
- **Dictionary.Keysの型不一致**：`Scripting.Dictionary`の`.Keys`メソッドは文字列配列ではなくVariant配列を返す仕様のため、受け取り側の型宣言を`Variant`に修正
- **グラフの余分な系列**：`.SetSourceData`と`.SeriesCollection.NewSeries`を併用すると、自動生成される既定系列とインデックスがズレて空の系列が残るバグが発生。`SetSourceData`を使わず、`NewSeries`の戻り値を直接受け取る形に修正して解決

## 既知の制約

- **入電上振れ検出が0件になることがある**：障害イベント（入電急増・応答率低下を伴う突発事象）は年2〜3回とまれに発生する設計です。今回生成した1年分のサンプルデータでは、障害の発生シフトが夜勤に偏り、夜勤はもともとの入電数が少ないため、上振れ判定の最小件数差の基準に届きませんでした。日勤での障害発生は0回でした。このように、まれにしか起きない事象を閾値で検出する仕組みは、サンプル期間や乱数の巡り合わせによって検出件数が0になることがあります。実務でも起こり得ることであり、サンプル数が十分であるかを確認する重要性を示す一例として、あえてそのまま残しています。

## 今後の拡張アイデア

- 着信計画差異（call_variance）だけを取り出した専用グラフ：時間帯別の実績・計画の折れ線グラフは値が近く差が見えにくいため、差異のみを棒グラフで見せる形も検討候補

## 技術スタック

- Python（架空データ生成）
- Excel VBA（CSV取り込み・集計・グラフ化）
