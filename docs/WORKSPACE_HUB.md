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
| 追加販促（スタッフ入力） | JOYFIT24経堂追加販促 | `1w7ExndmZn7t2_z55CvxRDMZy4QAcuEyNhIuj-6sUy3E` | https://docs.google.com/spreadsheets/d/1w7ExndmZn7t2_z55CvxRDMZy4QAcuEyNhIuj-6sUy3E/edit |
| EAST口コミ回答 | EAST口コミ回答者 | `13_E8m3vQa_61hcoMAPb7XZTyVDVtQ9O7rkVDNtHQvRM` | https://docs.google.com/spreadsheets/d/13_E8m3vQa_61hcoMAPb7XZTyVDVtQ9O7rkVDNtHQvRM/edit |
| マシンレクチャー／自動メール | 20分マシンレクチャー・自動送信メール | `1wntzhyPGcz9hW4saswppYmVG-zHINbjAibu9VkCyEQ8` | https://docs.google.com/spreadsheets/d/1wntzhyPGcz9hW4saswppYmVG-zHINbjAibu9VkCyEQ8/edit |

## Workspace 内シート（ミラー／ハブ）

| Workspace シート名 | 元 | 方式 |
|--------------------|----|------|
| `販促_乗り換え` ほか `販促_*` | 追加販促の各シート | `IMPORTRANGE`（見た目整形のみ元側可） |
| `口コミ_経堂` | `回答シート_JOYFIT` の `storeId=kyodo` | `QUERY(IMPORTRANGE(...))` |
| `マシンレクチャー申込` | 同名 | `IMPORTRANGE` |
| `入会者一覧＋自動メール管理` | 同名 | `IMPORTRANGE`（アンケート等はチェック風表示） |
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
| `setupPromoImport` | 追加販促 → `販促_*` |
| `setupReviewImport` | 口コミ経堂 → `口コミ_経堂` |
| `setupMachineImport` | マシンレクチャー2シート |
| `listSheets` | Workspace シート一覧 |
| `inspectBook&id=` | 任意ブックのシート／ヘッダー確認 |
| `rebuildUrlIndex` | URL一覧再生成（deta がある場合） |

## 他PCでの再開手順（最短）

1. このリポジトリを clone
2. Node + `@google/clasp`、`clasp login`
3. リポジトリ直下で `clasp push --force`（`.clasp.json` の scriptId を使用）
4. 必要なら `clasp deploy -i AKfycbzMELimQThNdPUShwo2_KBzJd8kGy9BNdRyOYNgu_sg41t2SleVRiWXFztZJ48e2l9L -d "..."`
5. Workspace を開き、`#REF!` のシートがあれば **アクセスを許可**（自分の Google アカウントが元スプシの編集者であること）

## 注意

- IMPORTRANGE の権限は「許可した人」のアクセスに依存する。引き継ぐ人は元スプシへの権限が必要
- Workspace を共有すると、共有相手は元スプシ権限がなくてもミラー内容を見られる
- 口コミのポイント付与・入会メールフラグの**操作は元側／EASTアプリ**。Workspace は表示用
