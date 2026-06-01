# RYUTA Workspace — GAS × Vercel × GitHub

業務日報を **AI で自動生成**し、**Gmail 下書き**（経堂数値つき）まで届ける構成です。

## 構成

```
  GitHub ──push──► Vercel (web/)
                      │
                      │ トップ / = iframe で GAS Workspace 表示
                      ▼
                 GAS Web App (gas/index.html + Code.gs)
                      │
                      ├── 4分割 UI・日報ボタン（いつもの画面）
                      └── API（/nippo 用・dayContext など）
```

| URL | 画面 |
|-----|------|
| `vercel.app/` | **GAS Workspace へリダイレクト**（4分割・iframe 不可のため） |
| `vercel.app/nippo` | 日報だけ作る簡易版（補助） |
| GAS `/exec` 直リンク | 上と同じ Workspace |

| 役割 | 場所 |
|------|------|
| メイン UI | `gas/index.html`（Vercel から iframe 表示） |
| 日報・スプシ・Gmail | `gas/Code.gs` |
| Vercel | ドメイン・`/nippo`・API プロキシ |

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

1. `https://あなたの-vercel.app/` を開く
2. **AI で業務・所感を生成** → カレンダー・Workspace をもとに文章化
3. 編集 → **Gmail に下書き作成** → GAS が経堂数値表を付けて下書き

API 単体テスト:

- `GET {GAS_URL}?api=dayContext&token=...`
- `GET {GAS_URL}?api=status`

## リポジトリ内のファイル

| パス | 説明 |
|------|------|
| `gas/Code.gs` | GAS 本番用（`dayContext` API 含む） |
| `gas/index.html` | 従来 Workspace UI |
| `web/` | Next.js 日報スタジオ |
| `Code.gs` / `index.html`（ルート） | 開発用コピー（`gas/` と同期推奨） |

## 今後の拡張

- Vercel 上で Gmail API まで持つ（GAS はスプシ専用にさらに薄くする）
- Chat / Gmail 件名の取り込み
- プロンプトを `web/prompts/` で版管理
