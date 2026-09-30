# RYUTA Workspace 集約ハブ — 引き継ぎメモ

他PC・別担当でも同じ構成を復元・理解するための共有データです。  
機械可読版: [`workspace-sources.json`](./workspace-sources.json)

## 方針

- **集約用**: RYUTA Workspace（このリポジトリの GAS が紐づくスプシ）
- **入力元**: 各業務スプシは独立のまま（EAST全店・フォーム等）。Workspace 側は基本 **読み取りのみ**
- **同期手段**: `IMPORTRANGE` / `QUERY(IMPORTRANGE(...))`（ライブ）。口コミ・入会者フラグは TRUE/FALSE をチェック風表示
- **元ブックを勝手に編集しない**（特に EAST 口コミは全店利用）。例外: 受付状況表 `日報` の移籍/復会/紹介セル（D10/F10/H10・D14/F14/H14）。**オプションの契約/解約は受付状況表 GAS が書く**（Workspace は上書きしない。月初列は触らない）
- **日報の月シート参照は使わない**: `INDIRECT($B$1&"!H36")` のように B1（2609）の月タブを見る式は、月シートへの手入力が回っておらず機能していない。削除してよい（2026-10-01 確認）

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

## GAS プロジェクトは2つ（混ぜない）

| 項目 | ワークスペース | 受付状況表（NIPPO） |
|------|----------------|---------------------|
| スプシ | 経堂　ワークスペース `1deuG2zYd…` | 経堂　受付状況表 `14hxiLBzv…` |
| scriptId | `1YjNLFjfLFNYM2Pyt248fGd90QW9wOKvAYJu-CKxHHTRobtZQNAxlobjp` | `1JvaBDxH580M-WCRa1bk5QZOAZd8veLJDbQccODMDni5Iv9_DHtZVRw7d` |
| ローカル | `gas-remote/r.js`（同期コピー: `Code.gs`, `gas/Code.gs`） | 会社PC `Documents/GitHub/nippo/gas/`。OPの正本は `kyodo-master-deta/option/Code.gs` |
| Web App | `AKfycbzMELimQTh…` | `AKfycbyQzrG0awDL…`（`refreshNumbers`） |
| 役割 | 販促ミラー、会員分析、移籍/復会/紹介と契約の日報書き戻し。メニューは「数値更新」だけ | Gmail 入会・退会、OP取込、日報メール、数値更新 |

スクリプトIDは「どのプロジェクトか」を特定できる。貼ってあるソース全文をこの環境からダウンロードするには Google ログインが要る。シートの数字は `peekExternal` で読める。日報の契約はワークスペース GAS が受付状況表を開いて書く（移籍と同じ）。この環境に `CLASPRC_JSON` が無いと本番へ push できない。

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
| `syncJoinBreakdown` | 移籍・復会・紹介を数えて日報 D14/F14/H14（と当日 D10/F10/H10）へ書き戻す。**契約/解約は触らない** |
| `ensureNippoOpByB1` | 残っているが、5分トリガーからは呼ばない。契約/解約は受付状況表側 |

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
- **9/30 移籍・復会・紹介**: 日報 D14 はもともと `INDIRECT($B$1&"!H36")` で月シート（2609）の合計を見ていたが、月シートへの入力が回っておらず機能していない → **その INDIRECT は削除してよい**。
  - **移籍・紹介**＝Workspace 内の販促シートを数えて、日報の **数字マスだけ** 書く（D14/H14 当月、D10/H10 当日）。日報から追加販促への IMPORTRANGE は許可ができず 0 になったので使わない。C/E/G の「移籍」「/復会」「/紹介」は触らない
  - **復会**＝今月の `経堂_入会` が過去の `経堂_入会` とメール／氏名一致。F14 当月、F10 当日
  - **当日（10行）**＝日報 B1 の月のうち「その日」（今日が翌月ならその月末日）。**当月（14行）**＝B1 の月全体
  - 月シート 2609 の INDIRECT は運用されていないので使わない
- **10/1 オプション契約・解約**: 「数値更新 → 受付状況表の数値を更新」が、メール取り込みのあと移籍・復会・紹介と契約・解約も日報へ書く。契約は診断の入会月。月初列は触らない。5分更新も同じ。Google の中のプログラムが古いままだと、ボタンはメール取り込みだけ。
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
- どちらも onOpen と5分トリガーで自動復旧
- 口コミの「来店日」が 46294 のような数値表示（未対応・ユーザー確認待ち）
- トップ H10 に「エンジョイポイント付与（口コミ確認後）」リンク（`ensureEnjoyPointLink_`）

## 課題・未対応（2026-09-29 時点）

- **GAS バージョン上限 200 に到達** → Web アプリ（`AKfycbzMELim…` @200）を更新できない。新しい API を使うには Apps Script エディタ「プロジェクト履歴」で古いバージョンを一括削除（@200 は残す）してから `clasp deploy -i …`。メニュー・トリガーは HEAD で動くので push だけで反映される
- 退会理由の記号表が手に入れば理由別集計を追加
- 請求・回収実績：4・5月度の追加列（SMSレジ入金 等）を出すか要確認
- 以前からの保留: SHIFT-DETA-NEW の stash（local-before-pull-2026-09-29）破棄/保持、SHIFT の .env.local キー、未納で「該当なし」の人を会員番号で照合できる名簿、F71 書き込み、2609 曜日列、backup シート整理

## 今後やりたいこと（案）

- 会員分析：退会理由別、入会月（7月キャンペーン等）ごとの継続率比較、早期退会者の特徴（年代×プラン）
- 請求・回収実績と会員分析をトップのタイルに出す（`hubCatalog_` には追加済み。`setupHubHome` の再構築で表示）

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
