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
| 学校関係者割（全店フォーム・読み取りのみ） | 学校関係者割フォーム　JOYFIT24（回答） | `1mmG_xM1WoWFKgpmOl5obsKXo_0GjWLnAnLY9hTGanVg` | https://docs.google.com/spreadsheets/d/1mmG_xM1WoWFKgpmOl5obsKXo_0GjWLnAnLY9hTGanVg/edit |
| EAST口コミ回答 | EAST口コミ回答者 | `13_E8m3vQa_61hcoMAPb7XZTyVDVtQ9O7rkVDNtHQvRM` | https://docs.google.com/spreadsheets/d/13_E8m3vQa_61hcoMAPb7XZTyVDVtQ9O7rkVDNtHQvRM/edit |
| マシンレクチャー／自動メール | 20分マシンレクチャー・自動送信メール | `1wntzhyPGcz9hW4saswppYmVG-zHINbjAibu9VkCyEQ8` | https://docs.google.com/spreadsheets/d/1wntzhyPGcz9hW4saswppYmVG-zHINbjAibu9VkCyEQ8/edit |
| 未納管理（オーナー別・自分は編集者） | 26年度未納管理ドライブ【経堂】 | `10vpQRDfTdwx_Wb7JaSm3lZCkTk8msLyf8ggAHhI1shI` | https://docs.google.com/spreadsheets/d/10vpQRDfTdwx_Wb7JaSm3lZCkTk8msLyf8ggAHhI1shI/edit |

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
4. [学校関係者割フォーム　JOYFIT24（回答）](https://docs.google.com/spreadsheets/d/1mmG_xM1WoWFKgpmOl5obsKXo_0GjWLnAnLY9hTGanVg/edit)（閲覧者以上。**元ブックは編集しない**）
5. [EAST口コミ回答者](https://docs.google.com/spreadsheets/d/13_E8m3vQa_61hcoMAPb7XZTyVDVtQ9O7rkVDNtHQvRM/edit)
6. [マシンレクチャー](https://docs.google.com/spreadsheets/d/1wntzhyPGcz9hW4saswppYmVG-zHINbjAibu9VkCyEQ8/edit)
7. 会員動向など、マスタがさらに参照している元があれば同様

### Workspace 側の再許可（1回／元ブック）

1. **`r-kusaka@okamoto-group.co.jp` でログインしたブラウザ**で [RYUTA Workspace](https://docs.google.com/spreadsheets/d/1deuG2zYdIMegMnCCT7lVl4AD7J75K8KisEsH2NVH10Q/edit) を開く
2. `#REF!` が出ているセルを選ぶ（よくある場所）
   - `経堂マスタ` の **AB4 / AB5 / AB6** 付近（受付・月・見学の土台。口コミ右に学校関係者割を足すと土台は AT 付近へずれる）
   - `学割` の A1（全店フォーム → 経堂だけ・新しい順）
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
| `販促_乗り換え` ほか `販促_*` | 追加販促の各シート（`学校関係者` スタブは使わない） | `IMPORTRANGE`（見た目整形のみ元側可） |
| `学割` | 学校関係者割フォーム `フォームの回答 1` の A:P。B列=`JOYFIT24 経堂`。全件・新しい順 | `QUERY(IMPORTRANGE(...))`。シート名は `学割`（`販促_学校関係者` から改名）。経堂マスタの P列以降は **当月の申請**（2行目のボタンで見学体験／紹介／学割／ラグビー割／乗り換え／6ヶ月／レクチャー／口コミ／入会メール／入会／退会を切替）。詳細の全列は各名称シート。API `setupSchoolDiscountImport` |
| `口コミ_経堂` | `回答シート_JOYFIT` の `storeId=kyodo` | `QUERY(IMPORTRANGE(...))` |
| `見学体験申請` | 見学・体験フォーム（`1RPUw0…`）の `見学体験申請` | A2 `IMPORTRANGE`。K列＝入会日。J列＝入会／未入会。経堂マスタは当月移籍の右に **当月紹介**、その右に **当月学割**。P列以降は当月の申請（分野ボタンで切替） |
| `経堂_入会` / `経堂_退会` | 受付状況表 `入会・退会_データ` の A:F / G:L | `IMPORTRANGE`（白黒整形のみ）。E列「メールID」は Gmail の通知メールID（重複防止用）。F列「メールアドレス」は入会メールの宛先（受付状況表 GAS が `入会_メールアドレス` 対応表から毎回再生成）。見学体験申請 J列はこのF列でもメール一致を見る。退会キャンセル列は6ヶ月継続の途中退会がほぼ無くなったため Workspace では扱わない（受付状況表側では今後使う可能性があるので残す） |
| `マシンレクチャー申込` | 同名 | `IMPORTRANGE` |
| `未納管理` | 26年度未納管理ドライブ【経堂】（`10vpQRDf…`）の月タブ | B1 で年月（25年8月〜27年12月）を選ぶと A5 の `IMPORTRANGE("'"&B1&"'!A1:AK")` がその月のタブを表示（☑/☐・¥表示）。会員名の右に 入会日・入会区分・未納開始・入会→未納 を差し込み（`経堂_入会` を氏名照合、未納開始月末までの最新入会。2ヶ月以内は赤）。D1:M3 は選択月の集計（支払額ベース。回収は回収金額、空なら右隣3列の「〇〇入金」で支払額を回収扱い）。元ファイルは触らない。API `setupUnpaidView` |
| `未納管理_推移` | 同上の全月タブ | 1行＝1ヶ月の集計（件数・未納総額・回収額・回収率・カテゴリ別回収率）。数値のまま。`setupUnpaidView` で一緒に作成 |
| `入会者一覧＋自動メール管理` | 同名 | `IMPORTRANGE`（アンケート等はチェック風表示） |
| `トップ` | 作業の目次。上段＝シート名（`#gid=` でそのタブへ）。下段＝`元のシートを開く ↗`（元ブック）。墨の細い行が陰影。閉じるときはタブを右クリック→非表示 | HYPERLINK |
| `URL一覧` | リンク索引（先頭に口コミ付与アプリ） | 値 |
| `Tasks` / `WorkspaceSync` | 本 GAS 用 | 自動作成 |

## 外部アプリ URL

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
| `setupPromoImport` | 追加販促 → `販促_*`（元ブックの「学校関係者」はスキップ。フォーム側を使う） |
| `setupSchoolDiscountImport` | 学校関係者割フォーム → `学割`（経堂だけ A:P・新しい順）＋経堂マスタ P列以降の当月パネル。元フォームは触らない |
| `setupReviewImport` | 口コミ経堂 → `口コミ_経堂` |
| `setupMachineImport` | マシンレクチャー2シート |
| `listSheets` | Workspace シート一覧 |
| `diagnoseImports` | IMPORTRANGE／元スプシ疎通の健全性 |
| `formatJoinList` | 入会者一覧：枠線削除＋チェック列を下まで適用（IMPORTRANGE維持） |
| `setupUnpaidView` | `未納管理`（ダッシュボード＋月表示＋入会照合4列）と `未納管理_推移` を再作成。B1 の選択月は保持。IMPORTRANGE 許可もスクリプトで付与 |
| `restyleHubLook` | 見た目再適用。経堂マスタでは当月紹介・当月学割と P列以降の当月パネルも冪等に載せる |
| `setupMasterIntroKpi` | 経堂マスタの当月移籍の右に「当月紹介」（`販促_紹介・ペア入会` の当月申請件数）を追加 |
| `hubOpen&name=` | 指定シートだけ表示して開く |
| `hubClose` | トップ以外を隠す |
| `setupKengakuJoinLive&mode=` | 見学体験申請の入会判定（`joinDate`=K列入会日／`joinLabel`=J列／`emailJoin`=経堂_入会・退会の列拡張／`leaveList`=経堂マスタ D9 今日の退会者） |
| `readRange&name=&range=` | Workspace の任意範囲の表示値・数式を読む（確認用、書き込みなし） |
| `peekExternal&id=&name=&range=` | 外部ブックを読み取りのみで確認（name 省略でシート一覧） |
| `inspectBook&id=` | 任意ブックのシート／ヘッダー確認 |
| `rebuildUrlIndex` | URL一覧再生成（deta がある場合） |

**シート内容の一括書き換え API（`repairRestrictedImports` 等）は使わない。** ラベル消失の原因になりうる。権限切れは上の共有＋再許可で直す。

## GAS を本番へ出す（clasp）

以前の手動はこれだけ:
`clasp login`（初回だけ、`r-kusaka@okamoto-group.co.jp`）→ リポジトリ直下で `clasp push --force` を2回 → `clasp deploy -i AKfycbzMELimQThNdPUShwo2_KBzJd8kGy9BNdRyOYNgu_sg41t2SleVRiWXFztZJ48e2l9L`

今後は **main へマージすると GitHub Actions が同じことを自動でやる**（`.github/workflows/gas-deploy.yml`）。必要なのは GitHub secret `CLASPRC_JSON` を1回入れること。中身は手元の `%USERPROFILE%\.clasprc.json`（Mac/Linux は `~/.clasprc.json`）。

## 他PCでの再開手順（最短）

1. このリポジトリを clone
2. Node + `@google/clasp`、`clasp login`（できれば `r-kusaka@okamoto-group.co.jp`）
3. リポジトリ直下で `clasp push --force`（`.clasp.json` の scriptId を使用）
4. 必要なら `clasp deploy -i AKfycbzMELimQThNdPUShwo2_KBzJd8kGy9BNdRyOYNgu_sg41t2SleVRiWXFztZJ48e2l9L -d "..."`
5. 上記「制限付き共有 + IMPORTRANGE 再許可」を実施

## 未納管理（2026-09-27 作成）— 引き継ぎメモ

- **元**: 未納管理ドライブ【経堂】。月ごとのタブ（`25年8月`〜）。オーナーは別の人、自分は編集者。IMPORTRANGE なので **元ファイルへの通知・編集履歴は発生しない**（読み取りのみ）。Workspace の共有相手には未納者の氏名・電話・金額が見える点だけ注意
- **`未納管理` シートの構成**
  - A1:B1 年月選択（いまの月がドロップダウン先頭。25年8月〜27年12月）、A2 元ファイルへのリンク、A3 推移シートへのリンク。見出しは右クリック→メモで説明
  - D1:M3 選択月の集計カード（未納件数／未納総額／回収額／回収率（黒反転）／未回収額／回収済み／1ヶ月・2ヶ月・貸倒候補・JACCS の回収率）
  - A5〜 元タブの表示用変換（TRUE/FALSE→☑/☐、金額→¥、率→%、レジDL→m/d、A列区分を 1ヶ月／2ヶ月／貸倒／JACCS／過年度 に短縮）。☑は表示のみ。レジ送信〜SMS・元ファイル右端の集計ブロックは非表示（列見出しの「+」で出せる）
  - 会員名の右に 4列（入会日・入会区分・未納開始・入会から未納）。`経堂_入会` を氏名（空白除去）で照合し、未納開始月末までの最新入会を採用。入会から未納が2ヶ月以内は赤字、見つからない人は「該当なし」（2023/10 以前の入会・他店番号など）
  - 見方: 左の区分色 → 氏名 → 赤い支払額＝未回収。回収済みは灰。合計より下の備考はそのまま残す。配色は `#EDEDED` / `#171717` / `#444444` / `#DA0037`
  - タブ色は4色のみ（通常＝墨 `#171717`、未納＝アクセント `#DA0037`、バックアップ等＝補助 `#444444`）。見た目再適用は `restyleHubLook`（数式は触らない）。各作業シートに `元のシートを開く ↗`（未納は左上 A2 の「元の未納管理ドライブを開く ↗」、見学・会員動向は行1の右端、他は行1左上）
- **`未納管理_推移`**: 1行＝1ヶ月（25年8月〜27年12月）。`全体回収率バー` は E列の回収率（回収額÷未納総額）の横棒。貸倒候補は A列の「貸倒」「貸し倒れ」両方。該当者がいなければ「対象なし」。累計の回収額は各月の単純合計（繰越で二重計上の可能性あり）
- **集計ロジック**（`unpaidStatsLet_`）: 新旧レイアウトが混在（26年3月以前は会員番号がB列・区分列なし）するため、列は見出し文字（会員番号／支払額／総額／回収金額|入金金額|レジ打ち金額）で探す。対象行＝会員番号あり・支払額>0。回収＝回収金額、空なら右隣3列に「〇〇入金」があれば支払額を回収扱い。元ファイル右上の独自集計欄とは定義が違うので数字は一致しない。25年11月の区分表記は「貸し倒れ」
- **すべて数式**なのでトリガー不要（Workspace GAS には script.scriptapp スコープが無くトリガーを作れない）

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
