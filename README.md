# RS-LOG — 稟議・修繕・日報 PWA

PC 向けの黒背景アプリ。Chrome の「アプリをインストール」で追加できます。

## 構成

```
  GitHub ──push──► Vercel (web/)  = RS-LOG PWA
                      │
                      │ 日報下書き API
                      ▼
                 GAS Web App (Gmail / カレンダー / スプシ)
```

| URL | 画面 |
|-----|------|
| `vercel.app/` | **RS-LOG**（稟議申請・修繕備品・日報） |
| `vercel.app/nippo` | 日報だけの簡易版 |
| GAS `/exec` | 従来の 4 分割 Workspace |

| 役割 | 場所 |
|------|------|
| メイン UI | `web/`（PWA） |
| 日報・スプシ・Gmail | `gas/Code.gs` / 本番は `gas-remote/` |
| **集約ハブ（スプシID・同期マップ）** | [`docs/WORKSPACE_HUB.md`](docs/WORKSPACE_HUB.md) / [`docs/workspace-sources.json`](docs/workspace-sources.json) |

## セットアップ

## セットアップ

### 1. Google Apps Script

1. [script.google.com](https://script.google.com) でプロジェクトを開く（または新規）
2. `gas/Code.gs` を **コード** に、`gas/index.html` を **index**（HTML）にコピー
3. **プロジェクトの設定 → スクリプト プロパティ** に追加:
   - `WS_API_TOKEN` … ランダムな長い文字列（Vercel と共有）
   - `GEMINI_API_KEY` … （任意）GAS 上の所感校閲用。Vercel 側にも同じキーを設定
4. **デプロイ → ウェブアプリ** → `/exec` URL を控える
5. 初回: カレンダー・Gmail・スプレッドシートの権限を許可

### 2. Vercel

```bash
cd web
npm install
cp .env.example .env.local
# .env.local を編集
npm run dev
```

環境変数（Vercel ダッシュボードでも可）:

| 変数 | 内容 |
|------|------|
| `GAS_WEB_APP_URL` | GAS の `/exec` URL |
| `NEXT_PUBLIC_GAS_WEB_APP_URL` | 上と同じ（`/workspace` で 4 分割 UI を表示） |
| `GAS_API_TOKEN` | `WS_API_TOKEN` と同じ |
| `GEMINI_API_KEY` | （任意）使わないなら不要 |

GitHub リポジトリを Vercel に接続 → Root Directory を **`web`** に指定。

### 3. 動作確認

1. `https://ryuta-workspace.vercel.app/` を開く
2. Chrome で「アプリをインストール」→ デスクトップに RS-LOG
3. 稟議 / 修繕を登録。日報タブから Gmail 下書き

API 単体テスト:

- `GET {GAS_URL}?api=dayContext&token=...`
- `GET {GAS_URL}?api=status`

## リポジトリ内のファイル

| パス | 説明 |
|------|------|
| `gas/Code.gs` | GAS 本番用（`dayContext` API 含む） |
| `gas/index.html` | 従来 Workspace UI |
| `web/` | RS-LOG PWA（稟議・修繕・日報） |
| `Code.gs` / `index.html`（ルート） | 開発用コピー（`gas/` と同期推奨） |

## 今後の拡張

- Vercel 上で Gmail API まで持つ（GAS はスプシ専用にさらに薄くする）
- Chat / Gmail 件名の取り込み
- プロンプトを `web/prompts/` で版管理
