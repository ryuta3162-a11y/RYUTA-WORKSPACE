# RYUTA Workspace 集約ハブ — 引き継ぎメモ

他PC・別担当でも同じ構成を復元・理解するための共有データです。  
機械可読版: [`workspace-sources.json`](./workspace-sources.json)

## 方針

- **集約用**: RYUTA Workspace（このリポジトリの GAS が紐づくスプシ）
- **入力元**: 各業務スプシは独立のまま（EAST全店・フォーム等）。Workspace 側は基本 **読み取りのみ**
- **同期手段**: `IMPORTRANGE` / `QUERY(IMPORTRANGE(...))`（ライブ）。口コミ・入会者フラグは TRUE/FALSE をチェック風表示
- **元ブックを勝手に編集しない**（特に EAST 口コミは全店利用）

## スプレッドシート一覧

| 役割 | 名前 | ID | URL |
|------|------|-----|-----|
| **集約ハブ** | RYUTA Workspace | `1deuG2zYdIMegMnCCT7lVl4AD7J75K8KisEsH2NVH10Q` | https://docs.google.com/spreadsheets/d/1deuG2zYdIMegMnCCT7lVl4AD7J75K8KisEsH2NVH10Q/edit |
| 受付状況表／日報（マスタKPI元） | 経堂　受付状況表 | `14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w` | https://docs.google.com/spreadsheets/d/14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w/edit |
| 見学・体験フォーム | 経堂　見学・体験フォーム | `1RPUw0slNCit9ZwJgINGfv89oc2Hxw8zzAZyMt6g_QuY` | https://docs.google.com/spreadsheets/d/1RPUw0slNCit9ZwJgINGfv89oc2Hxw8zzAZyMt6g_QuY/edit |
| 追加販促（スタッフ入力） | JOYFIT24経堂追加販促 | `1w7ExndmZn7t2_z55CvxRDMZy4QAcuEyNhIuj-6sUy3E` | https://docs.google.com/spreadsheets/d/1w7ExndmZn7t2_z55CvxRDMZy4QAcuEyNhIuj-6sUy3E/edit |
| EAST口コミ回答 | EAST口コミ回答者 | `13_E8m3vQa_61hcoMAPb7XZTyVDVtQ9O7rkVDNtHQvRM` | https://docs.google.com/spreadsheets/d/13_E8m3vQa_61hcoMAPb7XZTyVDVtQ9O7rkVDNtHQvRM/edit |
| マシンレクチャー／自動メール | 20分マシンレクチャー・自動送信メール | `1wntzhyPGcz9hW4saswppYmVG-zHINbjAibu9VkCyEQ8` | https://docs.google.com/spreadsheets/d/1wntzhyPGcz9hW4saswppYmVG-zHINbjAibu9VkCyEQ8/edit |
| 未納管理（オーナー別・自分は編集者） | 26年度未納管理ドライブ【経堂】 | `10vpQRDfTdwx_Wb7JaSm3lZCkTk8msLyf8ggAHhI1shI` | https://docs.google.com/spreadsheets/d/10vpQRDfTdwx_Wb7JaSm3lZCkTk8msLyf8ggAHhI1shI/edit |
| 未納の対応後☑（リンクのみ） | 【JOYFIT】26年度未納回収フロー管理表 | `1NvIIRTXC9XCAuib5USFouigkvmM8H2WDBOTWLjfN8oM` | https://docs.google.com/spreadsheets/d/1NvIIRTXC9XCAuib5USFouigkvmM8H2WDBOTWLjfN8oM/edit |
| 請求報告（EAST全店・経堂行を相互連携） | 26年度未納一覧【EAST運営本部】 | `1qFF8HGOlSOczshMI5Vg5iTAgN_iLQ2aemJLp35V3rbA` | https://docs.google.com/spreadsheets/d/1qFF8HGOlSOczshMI5Vg5iTAgN_iLQ2aemJLp35V3rbA/edit |
| FIT365 入会者一覧 | FIT365 入会者一覧＋自動メール管理 | `1BbExBUCfyq1cfNqw4TvlwUriL-AfvghU9XT6McdzGTQ` | https://docs.google.com/spreadsheets/d/1BbExBUCfyq1cfNqw4TvlwUriL-AfvghU9XT6McdzGTQ/edit |

## 制限付き共有 + IMPORTRANGE 再許可（推奨運用）

方針: **元スプシは制限付きのまま**。`IMPORTRANGE` は残す。GAS 値コピーへの切り替えはしない。

### 誰に共有するか

Web アプリ実行ユーザー（`appsscript.json` の `USER_DEPLOYING`）:

- **`r-kusaka@okamoto-group.co.jp`** … 閲覧者以上（編集者でも可）

このアカウントが元スプシを開けないと、Workspace 上の `IMPORTRANGE` は `#REF!` のままになる。

### 共有チェックリスト（元スプシごと）

各ファイル → **共有** → 上記メールが入っているか確認（なければ閲覧者で追加）。

1. [経堂　受付状況表](https://docs.google.com/spreadsheets/d/14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w/edit)（日報・入会・退会・OP）
2. [経堂　見学・体験フォーム](https://docs.google.com/spreadsheets/d/1RPUw0slNCit9ZwJgINGfv89oc2Hxw8zzAZyMt6g_QuY/edit)
3. [JOYFIT24経堂追加販促](https://docs.google.com/spreadsheets/d/1w7ExndmZn7t2_z55CvxRDMZy4QAcuEyNhIuj-6sUy3E/edit)
4. [EAST口コミ回答者](https://docs.google.com/spreadsheets/d/13_E8m3vQa_61hcoMAPb7XZTyVDVtQ9O7rkVDNtHQvRM/edit)
5. [マシンレクチャー](https://docs.google.com/spreadsheets/d/1wntzhyPGcz9hW4saswppYmVG-zHINbjAibu9VkCyEQ8/edit)
6. 会員動向など、マスタがさらに参照している元があれば同様

### Workspace 側の再許可（1回／元ブック）

1. **`r-kusaka@okamoto-group.co.jp` でログインしたブラウザ**で [RYUTA Workspace](https://docs.google.com/spreadsheets/d/1deuG2zYdIMegMnCCT7lVl4AD7J75K8KisEsH2NVH10Q/edit) を開く
2. `#REF!` が出ているセルを選ぶ（よくある場所）
   - `経堂マスタ` の **AB4 / AB5 / AB6** 付近（受付・月・見学の土台）
   - `経堂_入会` / `経堂_退会` / `経堂_OP` の A1
   - `見学体験申請` など
3. セル上または数式バー近くの **「アクセスを許可」** を押す（元ブックごとに1回）
4. 数秒〜数十秒で値が戻るか確認。`経堂マスタ` の **F17 / F18 / F19** など KPI が空や `#REF!` でないこと

共有設定を「リンクを知っている全員」→「制限付き」に変えた直後は、許可が切れることがある。そのときは **共有は維持したまま、上記の再許可だけ** やり直せばよい。

### 診断（任意）

```
https://script.google.com/macros/s/AKfycbzMELimQThNdPUShwo2_KBzJd8kGy9BNdRyOYNgu_sg41t2SleVRiWXFztZJ48e2l9L/exec?api=diagnoseImports
```

- `executedAs` が `r-kusaka@...` であること
- `summary.brokenSheets` / `blockedSources` が空なら健全
- `sources` の受付・見学が `ok: true` なのにシートが `#REF!` なら → **再許可だけ不足**

## Workspace 内シート（ミラー／ハブ）

| Workspace シート名 | 元 | 方式 |
|--------------------|----|------|
| `販促_乗り換え` ほか `販促_*` | 追加販促の各シート | `IMPORTRANGE`（見た目整形のみ元側可） |
| `口コミ_経堂` | `回答シート_JOYFIT` の `storeId=kyodo` | `QUERY(IMPORTRANGE(...))` |
| `見学体験申請` | 見学・体験フォーム（`1RPUw0…`）の `見学体験申請` | A2 `IMPORTRANGE`。K列＝入会日（K1 の式。申込の前日〜180日以内で `経堂_入会` の氏名・メール、または入会者一覧のメールが一致した最初の入会日）。J列＝K列があれば「入会」、なければ「未入会」（見学・体験の行のみ）。経堂マスタ R1 に今月の入会率、R:W の一覧に入会／未入会 |
| `経堂_入会` / `経堂_退会` | 受付状況表 `入会・退会_データ` の A:F / G:L | `IMPORTRANGE`（白黒整形のみ）。E列「メールID」は Gmail の通知メールID（重複防止用）。F列「メールアドレス」は入会メールの宛先（受付状況表 GAS が `入会_メールアドレス` 対応表から毎回再生成）。見学体験申請 J列はこのF列でもメール一致を見る。退会キャンセル列は6ヶ月継続の途中退会がほぼ無くなったため Workspace では扱わない（受付状況表側では今後使う可能性があるので残す） |
| `マシンレクチャー申込` | 同名 | `IMPORTRANGE` |
| `未納管理` | 26年度未納管理ドライブ【経堂】（`10vpQRDf…`）の月タブ | B1 で年月（25年8月〜27年12月）を選ぶと A5 の `IMPORTRANGE("'"&B1&"'!A1:AK")` がその月のタブを表示（☑/☐・¥表示）。会員名の右に 入会日・入会区分・未納開始・入会→未納 を差し込み（`経堂_入会` を氏名照合、未納開始月末までの最新入会。2ヶ月以内は赤）。D1:M3 は選択月の集計（支払額ベース。回収は回収金額、空なら右隣3列の「〇〇入金」で支払額を回収扱い）。元ファイルは触らない。API `setupUnpaidView` |
| `未納管理_推移` | 同上の全月タブ | 1行＝1ヶ月の集計（件数・未納総額・回収額・回収率・カテゴリ別回収率）。数値のまま。`setupUnpaidView` で一緒に作成 |
| `入会者一覧＋自動メール管理` | 同名 | `IMPORTRANGE`（アンケート等はチェック風表示） |
| `URL一覧` | リンク索引（先頭に口コミ付与アプリ） | 値 |
| `Tasks` / `WorkspaceSync` | 本 GAS 用 | 自動作成 |

## 外部アプリ URL

- エンジョイポイント付与画面（口コミの付与確認後に使う・経堂 clubCode=1304）: https://main.d5z4bnw4wyrxn.amplifyapp.com/store-settings/basic/points?clubCode=1304 — トップの引用元リンク（H列）末尾に自動で追加（`ensureEnjoyPointLink_`）

| 用途 | URL |
|------|-----|
| Workspace GAS Web App（本番） | https://script.google.com/macros/s/AKfycbzMELimQThNdPUShwo2_KBzJd8kGy9BNdRyOYNgu_sg41t2SleVRiWXFztZJ48e2l9L/exec |
| EAST 口コミ付与アプリ | https://script.google.com/a/macros/okamoto-group.co.jp/s/AKfycbwu1eUxJzePa494p-343axfwgUcnHATf-db7FKw806rXZQsHn_ea0uHc6415yw-RZ80/exec |

## GAS プロジェクト

| 項目 | 値 |
|------|-----|
| scriptId | `1YjNLFjfLFNYM2Pyt248fGd90QW9wOKvAYJu-CKxHHTRobtZQNAxlobjp` |
| ローカル | `gas-remote/`（`clasp push` の root） |
| サーバ本体 | `gas-remote/r.js`（同期コピー: `Code.gs`, `gas/Code.gs`） |

### セットアップ用 API（再構築）

ベース: `.../exec?api=`

| api | 内容 |
|-----|------|
| `setupPromoImport` | 追加販促 → `販促_*` |
| `setupReviewImport` | 口コミ経堂 → `口コミ_経堂` |
| `setupMachineImport` | マシンレクチャー2シート |
| `listSheets` | Workspace シート一覧 |
| `diagnoseImports` | IMPORTRANGE／元スプシ疎通の健全性 |
| `formatJoinList` | 入会者一覧：枠線削除＋チェック列を下まで適用（IMPORTRANGE維持） |
| `setupUnpaidView` | `未納管理`（ダッシュボード＋月表示＋入会照合4列）と `未納管理_推移` を再作成。B1 の選択月は保持。IMPORTRANGE 許可もスクリプトで付与 |
| `setupKengakuJoinLive&mode=` | 見学体験申請の入会判定（`joinDate`=K列入会日／`joinLabel`=J列／`emailJoin`=経堂_入会・退会の列拡張／`leaveList`=経堂マスタ D9 今日の退会者） |
| `readRange&name=&range=` | Workspace の任意範囲の表示値・数式を読む（確認用、書き込みなし） |
| `peekExternal&id=&name=&range=` | 外部ブックを読み取りのみで確認（name 省略でシート一覧） |
| `inspectBook&id=` | 任意ブックのシート／ヘッダー確認 |
| `rebuildUrlIndex` | URL一覧再生成（deta がある場合） |

**シート内容の一括書き換え API（`repairRestrictedImports` 等）は使わない。** ラベル消失の原因になりうる。権限切れは上の共有＋再許可で直す。

## 他PCでの再開手順（最短）

1. このリポジトリを clone（既にあれば `git pull`）
2. Node + `@google/clasp`、`clasp login`（できれば `r-kusaka@okamoto-group.co.jp`）
3. **編集を始める前に** `clasp pull` して `git diff` で差分確認（編集途中に pull すると手元の変更が消える。途中で確認したい時は別フォルダに `clasp clone`）（もう片方のPCが push した HEAD を消さないため）。問題なければ編集→`clasp push --force`→ git commit/push
4. 必要なら `clasp deploy -i AKfycbzMELimQThNdPUShwo2_KBzJd8kGy9BNdRyOYNgu_sg41t2SleVRiWXFztZJ48e2l9L -d "..."`
5. 上記「制限付き共有 + IMPORTRANGE 再許可」を実施

## 未納管理（2026-09-27 作成）— 引き継ぎメモ

- **元**: 未納管理ドライブ【経堂】。月ごとのタブ（`25年8月`〜）。オーナーは別の人、自分は編集者。IMPORTRANGE なので **元ファイルへの通知・編集履歴は発生しない**（読み取りのみ）。Workspace の共有相手には未納者の氏名・電話・金額が見える点だけ注意
- **`未納管理` シートの構成**
  - A1:B1 年月選択（ドロップダウン 25年8月〜27年12月）、A2 元ファイルへのリンク、A3 推移シートへのリンク
  - D1:M3 選択月の集計カード（未納件数／未納総額／回収額／回収率（黒反転）／未回収額／回収済み／1ヶ月・2ヶ月・貸倒候補・JACCS の回収率）
  - A5〜 元タブの表示用変換（TRUE/FALSE→☑/☐、金額→¥、率→%、レジDL→m/d、A列区分を下へ埋める、エラーは空欄）
  - 会員名の右に 4列（入会日・入会区分・未納開始・入会→未納）。`経堂_入会` を氏名（空白除去）で照合し、未納開始月末までの最新入会を採用。入会→未納が2ヶ月以内は赤字、見つからない人は「該当なし」（2023/10 以前の入会・他店番号など）
  - A〜D 列と 1〜6 行を固定。配色は白黒グレー＋赤1色のみ（ユーザー希望：カラフルにしない）
- **`未納管理_推移`**: 1行＝1ヶ月（25年8月〜27年12月）。数値のまま（計算に使える）。元にタブが増えれば自動で埋まる。累計の回収額は各月の単純合計（繰越で二重計上の可能性あり）
- **集計ロジック**（`unpaidStatsLet_`）: 新旧レイアウトが混在（26年3月以前は会員番号がB列・区分列なし）するため、列は見出し文字（会員番号／支払額／総額／回収金額|入金金額|レジ打ち金額）で探す。対象行＝会員番号あり・支払額>0。回収＝回収金額、空なら右隣3列に「〇〇入金」があれば支払額を回収扱い。元ファイル右上の独自集計欄とは定義が違うので数字は一致しない
- **すべて数式**なのでトリガー不要
- A4 に「対応後☑用シートを開く ↗」（【JOYFIT】26年度未納回収フロー管理表へのリンクのみ・IMPORTRANGE なし）。`onOpen` で消えていれば自動で付け直す

## 入会者一覧＋自動メール管理（2026-09-30）

- 経堂 A:E ＝ `1wntzhy…` / FIT365 F:J ＝ `1BbExBU…` の同名シートを **IMPORTRANGE で常時連動**（`linkJoinListLive_`）。チェック列は ☑/☐ 表示
- 値コピー同期（`syncJoinListWithCheckboxes_`）は使わない。このプロジェクトは `script.scriptapp` 権限がなく時間トリガーが作れないため、コピーは一度きりで古くなる
- A2 が IMPORTRANGE でなくなると `onOpen` が自動で張り直す。手動はメニュー「数値更新 → 入会者一覧を元シートと連動し直す」
- **clasp push の前に必ず `clasp pull`（別PCのHEADを消さない）**。2026-09-29 に別PCの v192〜200 を上書きしたため v200 に統合し直した

## 請求・回収実績（2026-09-29 作成・旧名 未納_請求報告）

- **元**: 26年度未納一覧【EAST運営本部】（全店入力・自分はEAST店舗分の編集者）。月度タブ `26年4月度`〜。**経堂は B列「経堂」の行（現在72行目）**。C:AH を使用（クレカ／ジャックス／合計 × 当月振替結果・翌月振替結果後）
- **Workspace 側**: 縦＝月度（26年4月度〜27年3月度）、横＝項目（30列）。見出し1〜3行目は元と同じ
  - 白セル＝手入力項目。入力すると **installable onEdit（`billingOnEdit`）で元シートの経堂行へ即書き戻し**
  - 赤＝未入力（該当ゼロは「0」を入れる）。グレー＝元シートの数式（自動計算）→ Workspace で入力しても元には書かず元の値に戻す
  - 元→Workspace は `billingPullTriggered`（開いた時＋5分ごと）で取り込み
  - 列は位置ではなく「グループ｜当月/翌月/他｜見出し名」で照合（4・5月度は列構成が違うため）。基準タブ `26年6月度`。4・5月度だけにある SMSレジ入金・代弁入金・総入金は未表示
  - 経堂の行は毎回店舗名で検索。**元シートの数式・書式・他店舗・構成には一切触れない**（書くのは経堂行の非数式セルだけ）。元の変更履歴には自分の名前で残る
- **19行目〜「売上・回収の分析」（9/30 整理）**: 月度・売上・請求件数・客単価・回収金額・回収率・未納額・未納率・売上前月比・翌月振替の不納額（参考・灰色）。表の下に「用語」（客単価＝売上÷請求件数 など）。グラフ2つ（売上と客単価／回収率）はデータのある月だけ表示し、月が増えると自動で伸びる。**最終回収率は削除**（元の「翌月振替結果後」の件数が当月より多い月があり意味が未確認）。作り直したい時は `BILL_ANALYSIS_VER_` を上げて push
- 初回セットアップ（済）: メニュー「数値更新 → 請求報告の自動連携を有効にする（初回のみ）」。トリガー3つ（onEdit／5分／onOpen）を作成。manifest に `script.scriptapp` を追加済み
- 実績（9/29時点）: 売上 月1,120〜1,220万円、客単価 ¥8,100〜8,700（低下傾向）、当月回収率 93.3%→95.0%（改善）

## 会員分析（2026-09-29 作成）

- **元**: Workspace 内の `累計退会データ`（F:S＝会員番号〜理由、P＝在籍期間、R＝退会年月）と `累計入会データ`（F:O＝会員番号〜手続、N＝利用開始年月）。**手動貼り付け**
- 貼り替えで行数が変わると 5分以内に自動で作り直し（`memberAnalysisIfChanged_`）。手動はメニュー「数値更新 → 会員分析を作り直す」
- **9/30 刷新**: 最上段（2〜5行）は受付状況表 `日報` を IMPORTRANGE で直接参照（C12 月初会員・F12/H12 男女、C13 当月入会・D14/F14/H14 移籍/復会/紹介、C15 当月末退会、C16 月末安定、C17 翌月月初、C18 休会）。月間退会率は「直近12ヶ月の月平均退会 ÷ 月初会員(A4)」。11〜12行に項目メニュー（タップでジャンプ）、各項目は行グループで折りたたみ（左の＋で開く、お金への影響だけ最初から開く）。方針: セル関数で完結できるものは関数優先
- 旧中身: KPI（在籍推定・累計入会/退会・月間退会率・退会者平均在籍・平均客単価〔請求・回収実績から連動〕・生涯売上2種）、入会年別／入会月別（直近36ヶ月）の 3・6・12・24ヶ月継続率、年代・性別・契約プラン別（2023年以降入会）、継続率グラフ
- 9/29 時点の要点: 在籍約1,480人／月間退会率3.61%／生涯売上 ¥126,247（退会者平均在籍15ヶ月）〜¥232,463（1÷退会率）。**3ヶ月継続率が 2017〜22年入会 96〜99% → 2024・25年入会 83.2% に悪化**（2026年は97%に回復）。20代ほど早期退会が多い
- 注意: 年齢は現在年齢。契約名称はほぼ空欄（＝通常）。退会理由は半数以上空欄＋記号（A・M 等）で意味不明
- 9/30 追加（クロス分析）: 
  - **お金への影響**: 直近月の売上、退会で毎月なくなる売上（月平均退会×客単価 ≒ ¥44万）、入会1人の生涯売上（≒ ¥23万＝獲得費の上限目安）、3ヶ月継続率を2017〜22年水準（98.1%）に戻した時の年間増収（≒ ¥1,118万・試算）
  - **販促・レクチャー別の継続率**: 販促シート（乗り換え／紹介・ペア／学割／ラグビー割／6ヶ月継続／マシンレクチャー申込）の申請者を累計入会データと **電話番号→氏名（漢字・カナ）** で照合（申請月の前2〜後3ヶ月に入会した人だけ候補）。比較用に「販促なし（同時期）」。対象10人未満は空欄
  - **季節別**: 入会した月（1〜12月）ごとの継続率＋「今月の一言」（赤字）。9月入会は12ヶ月継続 37.5%（平均44.2%）
  - 経堂マスタ「今月の読み」（F7:K14）は別の仕組みが自動で書いている（Workspace GAS 内には無い）ので触っていない
  - 分析内容を変えたら `MEMBER_ANALYSIS_VER_` を上げて push → 5分以内に作り直される

- **9/30 見やすさ改善**: 見出しを短く、各項目の見出し行の右に赤字で「結論1行」（閉じたままでも要点が分かる）。A列210・他108幅。6ヶ月継続特典は結論の比較から除外（生存バイアス）

## 経堂マスタ「当月の申請」（2026-09-30 追加分）

- 表は 1行目に「当月の申請」の式がある列から始まる（**9/30 に左へ会員番号列を挿入したので Q列〜**。タブ選択は Q2 の赤いプルダウン、月は AG5、受付ID は AG1）。スクリプトは列を固定せず `masterApplyCol_` で探す
- **会員番号（P列）**: P4 の関数で、表示中の名前（紹介は被紹介者）を `累計入会データ` G列（空白無視）と照合し F列の番号を出す。同名が複数（再入会など754人）なら電話番号下9桁で絞り、**届出日（M列）が最新の番号**。口コミのように会員番号列がある表はそれをそのまま表示。未入会・未貼り付けは空欄（`ensureMasterMemberNo_`）
- **☑/☐**: 表の QUERY を `LET(src_,…,MAP(…))` で包み TRUE/FALSE を ☑/☐ 表示（`ensureMasterApplyCheckmarks_`）。表示だけなので付与の記録は元シート側で
- **年月（B2）**: `=TEXT(TODAY(),"yyyy年m月")` で毎月自動切替（2026-10-01〜。以前は手入力で9月のまま止まっていた）。AG5・5ヶ月データベース・当月の申請はすべて B2 から計算。`ensureMasterMonthAuto_` が式を守る
- **入会 同日比較（5ヶ月データベース・入会実績の下の行）**: 各月の1日〜今日と同じ日までの入会数。`経堂_入会` の C列（月度＝利用開始月）と A列（メール受信日時）で COUNTIFS。A〜O列だけ挿入したので右の申請一覧は動いていない（`ensureMasterSameDayRow_`）
- どちらも5分トリガー（＋開いた時の installable onOpen → `billingPullTriggered`）で自動復旧
- 口コミの「来店日」が 46294 のような数値表示（未対応・ユーザー確認待ち）
- トップ H10 に「エンジョイポイント付与（口コミ確認後）」リンク（`ensureEnjoyPointLink_`）

## 整理整頓（2026-10-01）

- 削除したシート: 受付状況表 `OP契約一覧`（約3.5万行・参照なし）／ワークスペース `見学体験申請_backup_0927_1814`・`経堂マスタ_backup_0927_1814`（数式333個が裏で再計算されていた）・`未納_同期`（9/28 の古いログ）
- Workspace GAS（r.js）: 一度きりの初期設定・装飾・修復用の関数 69 個と、それを呼ぶ Web API 分岐（setupHubHome / setupUnpaidView / setupKengakuJoinLive / fixMasterActuals など）を削除。約236KB→約152KB。必要になったら git 履歴（2026-10-01 以前）から戻す
- 空の `onEdit` を削除。simple `onOpen` はメニュー作成だけにし、自動復旧（入会者一覧の連動・未納リンク・エンジョイ・会員番号・☑）は `billingPullTriggered`（5分＋開いた時）に一本化
- 受付状況表 GAS: 読み取り専用の `api=audit`（シートごとのサイズ・数式数・参照）と `api=triggers` を追加。Web アプリは @34
- 受付状況表のトリガー（r-kusaka 分）は 10時・17時の数値更新の2本。20時/21時の更新＋送信は店舗アカウント側のトリガー（送信直前に数値更新が走る。`api=triggers` の sentDays で毎日送信を確認できる）
- 日報 D21:E（契約・解約）は **受付状況表の数値更新だけ**が書く（10/17/20/21時の自動＋ワークスペースのメニュー「受付状況表の数値を更新」）。旧 `日報オプションをB1連動`（関数を書き込む版）は削除

## 日報の男女（2026-10-01）

- 月初（12行）: 男 F12 は3ファイル取込で「年齢表の男 − 特例の男 − 当月開始の男」。女 H12 は `=C12-F12`
- 月末安定（16行）・翌月月初（17行）: `F16=F12+K13`、`F17=F16-K15`、女は合計−男
- K13 当月入会の男 = 累計入会データで性別が分かる当月開始者の男（M13/N13）＋ 載っていない人数 × 月初の男比率
- K15 当月末退会の男 = 当月末退会 × 月初の男比率（退会メールに性別がない）
- J〜N列はメールの範囲（B〜I）の外。`ensureNippoGenderFlow_`（5分トリガー）が式を守る
- 受付状況表 GAS は **nippo/gas から push**（一時コピー %TEMP%\nippoclone は古いので使わない）

## 課題・未対応（2026-09-29 時点）

- 【2026-10-01 解消】GAS バージョン上限 200 で Web アプリが更新できなくなっていた → プロジェクト履歴で古い版を一括削除し、本番デプロイを **@201**（2026/10/02 0:14）に更新済み。URL／deploymentId は変更なし。また上限に当たったら同じく一括削除 → デプロイ更新。メニュー・トリガーは HEAD なので `clasp push` だけで反映される
- 【2026-10-01 時点の進捗】会員動向 10月のオプション月初（I31:I46）は未入力。ユーザーが「月初３ファイル」を再取込すれば `gessho3WriteHq_` が当月末契約数で自動記入（日報の月初 1515／男 1074 は変わらない）
- 【2026-10-01】日報 10月月初を 1515（男1074・女441）、会員動向 9月解除を 50 に修正済み。規約退会リスト 10名（4〜9月分）
- 【2026-10-01】受付状況表 Web アプリ（`AKfycbyQzrG0…` @37）：競合分析.gs の doGet を `competitorPage_` に改名し、ワークスペース連携.gs の doGet から振り分け（token/api 無し → 競合分析画面）。`inspect` API（id・name・range・bg）で他ブックの値・数式・背景を読める
- 退会理由の記号表が手に入れば理由別集計を追加（累計退会データの理由 X＝規約退会）
- 請求・回収実績：4・5月度の追加列（SMSレジ入金 等）は `billCanon_` が全月タブから自動で拾う（2026-10-01）。不要なら非表示にするだけ
- 以前からの保留: SHIFT-DETA-NEW の stash（local-before-pull-2026-09-29）破棄/保持、SHIFT の .env.local キー、未納で「該当なし」の人を会員番号で照合できる名簿、F71 書き込み、2609 曜日列、backup シート整理

## 見た目の統一（2026-10-02）

- 基準は **未納管理**（白・黒・グレー＋赤1色）。`hubTheme_()` に集約。旧クリーム紙（デスノート風）はやめ、ミラーシートも白地に揃える
- **会員分析**の円グラフ：虹色を廃止 → 墨〜グレー＋赤（`MEMBER_ANALYSIS_VER_=17`）。メニュー「ワークスペース → 会員分析を作り直す」か、5分トリガーで反映
- **請求・回収実績**のグラフ：棒＝墨、線＝赤、背景白（`BILL_ANALYSIS_VER_=6`）。列は全月タブから集約済み
- **口コミ管理**シート新設：件数／付与済み／未付与／付与率＋手順リンク（EAST付与・エンジョイ・一覧）。トップ「現場」の先頭付近に追加（`TOP_LAYOUT_VERSION_=simple-v4`）
- **口コミ_経堂**：来店日・付与日時の日付書式を修正（46294 問題）。1行目は黒帯

## 口コミ管理（2026-10-02）

- Workspace 内で「確認 → 付与 → エンジョイ」を1枚にまとめた。**付与そのものの書き込みは EAST／エンジョイ側のまま**（全店共通のため）
- 自動：5分トリガー（`billingPullTriggered`）でシートが無ければ作成。手動はメニュー「ワークスペース → 口コミ管理を作り直す」

## 規約退会＝過去の強制退会のあぶり出し（稼働中）

- **できている。** `syncKiyakuList_` が未納管理ドライブの各月シートで **会員名（D列）背景 `#ff0000`** の人を集め、「規約退会リスト」に出す（1時間ごと＋メニューで即時）
- 2026-10-02 時点で **10名**（例: 落合 悠野 2026/10/1 強制退会）。会員動向の退会月・累計入会の退会年月・未納管理の赤字表示まで連動
- 累計退会データに未反映の人は J列が「未反映（会員数から外す）」→ 会員分析から外れる

## あなた側でやること（設定・初回だけ）

1. このブランチを `clasp pull` せず上書きしないよう注意したうえで **`clasp push -f`**（またはエディタに `gas-remote/r.js` を反映）→ 必要なら Web アプリを新バージョンにデプロイ
2. スプシを開き直す（メニュー「ワークスペース」が出る）
3. 「ワークスペース → 口コミ管理を作り直す」「会員分析を作り直す」「トップを作り直す」「請求・回収を今すぐ取得」を1回ずつ
4. 累計入会／退会の貼り付けは従来どおり **3日おき手動**（自動取得手段なし）

## 今後やりたいこと（案）

- 会員分析：退会理由別、入会月（7月キャンペーン等）ごとの継続率比較、早期退会者の特徴（年代×プラン）

## 規約退会（2026-10）

- 未納管理ドライブの月シート（`26年4月` など）で D列の背景が `#ff0000` の人＝2ヶ月未納で強制退会。各シート備考の「M/D … 強制退会」が実施日（翌月1〜2日）
- `syncKiyakuList_`（billingPullTriggered・1時間ごと）が Workspace「規約退会リスト」を作り直す。元シートは読むだけ
- 会員動向では「赤くなったシートの月」の退会に数える（例：9月シートの落合 悠野 → 9月の解除 50）。CASIO の退会年月は翌月（リスト K列）
- 累計入会データ U列（退会年月）は 累計退会データに無ければ 規約退会リストの K列を使う → 会員分析からも外れる
- 未納管理の会員名は条件付き書式で規約退会なら赤
- 月初３ファイル取込：規約退会リストで退会年月が当月の人を月初会員・男から引く（`gessho3KiyakuMinus_`）。取込時に会員動向（本部 `1LOOUG97…` 経堂 31〜46行）の当月列へオプションの当月末契約数を書く（空のセルだけ・`gessho3WriteHq_`）
- 会員動向の入会・解除の実績は 累計入会（利用開始年月）・累計退会（退会年月の前月）と 4〜9月一致を確認済み。9月の解除だけ 49→50 に修正

## トップ（2026-10 シンプル化）

- クリックで開閉する仕組み（`onSelectionChange` / `handleHubHomeSelect_` / キー列 AX〜）は廃止。トップはただのリンク集
- `ensureTopSimple_`（billingPullTriggered 内）が `TOP_LAYOUT_VERSION_` の版ごとに1回だけ作り直す。並びは `topSections_()`、右の外部リンクは `topExternalLinks_()`。変えたら版を上げて push
- 区分：数字（経堂マスタ・日報・会員分析・会員動向）／未納（未納管理・推移・請求・回収実績）／現場（5枚）／販促（4枚）／データ（累計入会・累計退会）
- `#gid=` リンクは非表示シートを開けないので、トップに載せたシートは表示のまま。経堂_* ・Tasks・WorkspaceSync・月初系は非表示

## ハマりどころ（GAS から数式を入れるとき）

- LET の変数名に **セル参照っぽい名前（`d1` `k1` など）は使えない** → `#NAME?`。`kone` `djoin` のように付ける
- LET/MAP 内で範囲演算するときは `ARRAYFORMULA` で包む
- IMPORTRANGE の「アクセスを許可」は GAS から付与できる: `POST https://docs.google.com/spreadsheets/d/<WorkspaceID>/externaldata/addimportrangepermissions?donorDocId=<元ID>`（`Authorization: Bearer ScriptApp.getOAuthToken()`）
- `clasp push -f` は2回、`clasp deploy -i <deploymentId>` 直後は旧コードが返ることがあるので30秒待ってリトライ
- Windows PowerShell でインライン python / node を書くとクォートが壊れる → `.py` / `.js` ファイルにして実行。`Set-Content -Encoding utf8` は BOM が付くので GAS ファイルには使わない

## 関連リポジトリ

- 受付状況表 GAS（入会・退会の取り込み、メールアドレス抽出、タイムゾーン）: `Documents/GitHub/nippo/gas/`（scriptId `1JvaBDxH580M-WCRa1bk5QZOAZd8veLJDbQccODMDni5Iv9_DHtZVRw7d`）。受付状況表は 2026-09 に Asia/Tokyo へ変更済み。入会・退会シート E/J＝メールアドレス、K/L（メールID・退会キャンセル）は非表示
- 法人会員の退会は経堂マスタ「今日の退会者」に出さない。月の法人退会合計（D71 など）はスタッフが決算時に手入力するので上書きしない

## 注意

- IMPORTRANGE の権限は「許可した人」のアクセスに依存する。引き継ぐ人／実行ユーザーは元スプシへの権限が必要
- Workspace を共有すると、共有相手は元スプシ権限がなくてもミラー内容を見られる
- 口コミのポイント付与・入会メールフラグの**操作は元側／EASTアプリ**。Workspace は表示用
- エージェント／GAS から Workspace のラベル列を `setFormulas` 等で一括上書きしない
