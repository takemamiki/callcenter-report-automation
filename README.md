# callcenter-report-automation

クレジットカード会社のコールセンター業務を想定した、架空データによる日報集計・グラフ化の自動化ポートフォリオです。

## 背景

クレジットカード会社で約20年、不正利用防止・オーソリ部門を中心にコールセンター業務に携わってきた実務経験をもとに、現場でよくある「日々の入電データを集計してレポート化する」作業を、Python（データ生成）とExcel VBA（集計・グラフ化）で自動化しました。
実データは一切使用せず、すべて架空のデータで構築しています。

## 構成
callcenter-report-automation/
├── generate_fake_data.py # 架空データ生成（Python）
├── data/
│ ├── callcenter_report.csv # 生成された日次・シフト別CSVデータ
│ └── callcenter_hourly_report.csv # 生成された時間帯別CSVデータ
├── callcenter-report-automation.xlsm # 集計・グラフ化（Excel VBA）
└── README.md


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

実務のコールセンター運営で日々使われている「着信数・応答数・応答率・計画値・計画差異」という指標構成を模しています。着信計画は時間帯ごとの重み配分から算出した固定値（乱数なし）、実績はそこに±20%程度のブレを加えて生成しています。

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

## できること（Summaryシート）

- **月別集計**：月ごとの各項目の合計・平均を折れ線グラフで可視化（1年分のため12ヶ月分の推移が見える）
- **曜日別集計**：曜日ごとの各項目の平均入電数を積み上げ棒グラフで可視化
- **項目別統計**：各項目の合計・平均・最大・最小を一覧化
- **シフト別集計**：日勤/夜勤ごとの各項目の合計・平均を、比較しやすいクラスター棒グラフで可視化
- **時間帯別集計**：0〜23時それぞれの1年平均を算出し、着信数実績と着信計画を並べた折れ線グラフで、1日の呼量カーブと計画とのズレを可視化
- 全グラフで色分けを統一（オーソリ＝青、会員問合せ＝赤、紛失盗難＝緑）し、目盛り線もデータの実際の範囲に合わせて自動調整（3〜4本程度に抑えて読みやすく）

## 実装で詰まった点と対応

実務経験だけでは気づけなかった、VBA特有のハマりどころです。

- **CSV取り込み時の文字化け**：`Workbooks.Open`や`OpenText`では日本語が文字化けしたため、`QueryTables`経由で`TextFilePlatform=65001`（UTF-8）を明示して解決
- **Dictionary.Keysの型不一致**：`Scripting.Dictionary`の`.Keys`メソッドは文字列配列ではなくVariant配列を返す仕様のため、受け取り側の型宣言を`Variant`に修正
- **グラフの余分な系列**：`.SetSourceData`と`.SeriesCollection.NewSeries`を併用すると、自動生成される既定系列とインデックスがズレて空の系列が残るバグが発生。`SetSourceData`を使わず、`NewSeries`の戻り値を直接受け取る形に修正して解決
- **`Option Explicit`の重複記述**：モジュール内に`Option Explicit`が2回記述されていたことが原因で、モジュールレベルで宣言したはずの`Const`が別プロシージャから「変数が定義されていません」エラーになる現象が発生。重複を削除し、宣言セクションを一本化して解決
- **棒グラフに線用のスタイル関数を適用**：配色統一の実装時、折れ線グラフ用の`ApplyLineStyle`（線の色を変更）を誤って棒グラフに適用してしまい、色が反映されない現象が発生。棒グラフの塗りつぶし色を変えるには`ApplyBarStyle`（Fill.ForeColor）が必要で、グラフの種類ごとに適切な関数を使い分ける必要があると再確認

## 今後の拡張アイデア

- 時間帯×シフトのクロス集計：現状は時間帯別（全体平均）とシフト別（日次合計）を別々に集計しているが、shiftはhourに従属する値のため情報量が薄く、優先度は低いと判断
- 着信計画差異（call_variance）だけを取り出した専用グラフ：時間帯別の実績・計画の折れ線グラフは値が近く差が見えにくいため、差異のみを棒グラフで見せる形も検討候補

## 技術スタック

- Python（架空データ生成）
- Excel VBA（CSV取り込み・集計・グラフ化）
