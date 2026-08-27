'use client';

import { FormEvent, useCallback, useEffect, useMemo, useState } from 'react';

type MainTab = 'ringi' | 'repair' | 'nippo';
type OpsPage = 'view' | 'add';

type OpsItem = {
  id: number;
  kind: 'ringi' | 'repair';
  title: string;
  company: string;
  amount: string;
  replyDate: string;
  mailUrl: string;
  productUrl: string;
  stuckAt: string;
  status: string;
  createdAt?: string;
};

type DayContext = {
  date: string;
  calendarEvents: Array<{ title: string; start: string; end: string }>;
  workspace: { done: string[]; kansou: string };
};

const STORAGE_KEY = 'ryuta_ops_board_v2';
const STATUSES = ['未着手', 'メール送信済', '返答待ち', '稟議中', '発注済', '完了'];
const COMPANIES = [
  'SEKAI',
  'Lifefitness',
  'LIXIL',
  'BICS',
  'amazon',
  '鳳商事',
  'アイリスオーヤマ',
  'KH',
  '藤ビル',
  'メトス',
  'プロアバンセ',
  'technogym',
];

const EMPTY_FORM = {
  title: '',
  company: '',
  amount: '',
  replyDate: '',
  mailUrl: '',
  productUrl: '',
  stuckAt: '',
  status: '未着手',
};

function loadItems(): OpsItem[] {
  try {
    const raw = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    return Array.isArray(raw) ? raw : [];
  } catch {
    return [];
  }
}

export default function RsApp() {
  const [tab, setTab] = useState<MainTab>('ringi');
  const [page, setPage] = useState<OpsPage>('view');
  const [items, setItems] = useState<OpsItem[]>([]);
  const [editId, setEditId] = useState<number | null>(null);
  const [form, setForm] = useState(EMPTY_FORM);
  const [ready, setReady] = useState(false);
  const [installEvent, setInstallEvent] = useState<BeforeInstallPromptEvent | null>(null);
  const [standalone, setStandalone] = useState(false);

  const [ctx, setCtx] = useState<DayContext | null>(null);
  const [gyomuText, setGyomuText] = useState('');
  const [kansou, setKansou] = useState('');
  const [nippoStatus, setNippoStatus] = useState('');
  const [nippoLoading, setNippoLoading] = useState(false);

  useEffect(() => {
    setItems(loadItems());
    setReady(true);
    setStandalone(window.matchMedia('(display-mode: standalone)').matches);
    const onPrompt = (event: Event) => {
      event.preventDefault();
      setInstallEvent(event as BeforeInstallPromptEvent);
    };
    window.addEventListener('beforeinstallprompt', onPrompt);
    return () => window.removeEventListener('beforeinstallprompt', onPrompt);
  }, []);

  useEffect(() => {
    if (!ready) return;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(items));
  }, [items, ready]);

  const loadContext = useCallback(async () => {
    setNippoStatus('GAS からデータ取得中…');
    const res = await fetch('/api/context');
    const data = await res.json();
    if (!data.ok) {
      setNippoStatus(data.message ?? '取得に失敗。GAS_WEB_APP_URL と GAS_API_TOKEN を確認してください。');
      return;
    }
    setCtx({
      date: data.date,
      calendarEvents: data.calendarEvents ?? [],
      workspace: data.workspace ?? { done: [], kansou: '' },
    });
    const done = (data.workspace?.done as string[]) ?? [];
    if (done.length) setGyomuText(done.map((t: string) => t.replace(/^[・\-]\s*/, '')).join('\n'));
    if (data.workspace?.kansou) setKansou(String(data.workspace.kansou));
    setNippoStatus('');
  }, []);

  useEffect(() => {
    if (tab === 'nippo' && !ctx) loadContext();
  }, [tab, ctx, loadContext]);

  const filtered = useMemo(
    () => items.filter((item) => item.kind === tab),
    [items, tab]
  );

  function openAdd(reset = true) {
    if (reset) {
      setEditId(null);
      setForm(EMPTY_FORM);
    }
    setPage('add');
  }

  function openView() {
    setEditId(null);
    setPage('view');
  }

  function switchTab(next: MainTab) {
    setTab(next);
    setEditId(null);
    setForm(EMPTY_FORM);
    setPage('view');
  }

  function editItem(item: OpsItem) {
    setEditId(item.id);
    setForm({
      title: item.title,
      company: item.company,
      amount: item.amount,
      replyDate: item.replyDate,
      mailUrl: item.mailUrl,
      productUrl: item.productUrl,
      stuckAt: item.stuckAt,
      status: item.status || '未着手',
    });
    setPage('add');
  }

  function saveItem(event: FormEvent) {
    event.preventDefault();
    if (!form.title.trim() || tab === 'nippo') return;
    const payload = {
      kind: tab,
      title: form.title.trim(),
      company: form.company.trim(),
      amount: form.amount.trim(),
      replyDate: form.replyDate.trim(),
      mailUrl: form.mailUrl.trim(),
      productUrl: form.productUrl.trim(),
      stuckAt: form.stuckAt.trim(),
      status: form.status || '未着手',
    };
    setItems((prev) => {
      if (editId) {
        return prev.map((item) => (item.id === editId ? { ...item, ...payload } : item));
      }
      return [
        {
          ...payload,
          id: Date.now(),
          createdAt: new Date().toISOString().slice(0, 10),
        },
        ...prev,
      ];
    });
    setEditId(null);
    setForm(EMPTY_FORM);
    setPage('view');
  }

  function cycleStatus(id: number) {
    setItems((prev) =>
      prev.map((item) => {
        if (item.id !== id) return item;
        const idx = STATUSES.indexOf(item.status);
        return { ...item, status: STATUSES[(idx + 1) % STATUSES.length] };
      })
    );
  }

  function removeItem(id: number) {
    setItems((prev) => prev.filter((item) => item.id !== id));
  }

  async function installApp() {
    if (!installEvent) return;
    await installEvent.prompt();
    setInstallEvent(null);
  }

  function fillGyomuFromCalendar() {
    if (!ctx?.calendarEvents.length) {
      setNippoStatus('今日のカレンダー予定がありません');
      return;
    }
    setGyomuText(
      ctx.calendarEvents
        .map((ev) => `${ev.start && ev.end ? `${ev.start}–${ev.end} ` : ''}${ev.title}`.trim())
        .join('\n')
    );
    setNippoStatus('カレンダー予定を業務欄に反映しました');
  }

  async function createDraft() {
    const gyomu = gyomuText
      .split('\n')
      .map((line) => line.replace(/^[・\-•\s]+/, '').trim())
      .filter(Boolean);
    if (!gyomu.length) {
      setNippoStatus('業務内容を 1 行以上入力してください');
      return;
    }
    if (!window.confirm('Gmail に日報下書きを作成します。よろしいですか？')) return;
    setNippoLoading(true);
    setNippoStatus('Gmail 下書き作成中…');
    try {
      const res = await fetch('/api/draft', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ gyomu, kansou }),
      });
      const data = await res.json();
      if (!data.ok) throw new Error(data.message ?? '下書き作成に失敗');
      setNippoStatus(`下書きを作成しました。\n件名: ${data.subject ?? ''}`);
    } catch (error) {
      setNippoStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setNippoLoading(false);
    }
  }

  return (
    <div className="rs-app">
      <header className="rs-header">
        <div className="rs-brand">
          <div className="rs-mark">RS</div>
          <div>
            <div className="rs-name">RS-LOG</div>
            <div className="rs-sub">稟議 / 修繕 / 日報</div>
          </div>
        </div>
        <nav className="rs-nav">
          <button type="button" className={tab === 'ringi' ? 'active' : ''} onClick={() => switchTab('ringi')}>
            稟議申請
          </button>
          <button type="button" className={tab === 'repair' ? 'active' : ''} onClick={() => switchTab('repair')}>
            修繕・備品
          </button>
          <button type="button" className={tab === 'nippo' ? 'active' : ''} onClick={() => switchTab('nippo')}>
            日報
          </button>
        </nav>
      </header>

      <main className="rs-main">
        {!standalone && (
          <div className="rs-install">
            <span>Chrome のアドレス欄「アプリをインストール」で PC に追加できます。</span>
            {installEvent && (
              <button type="button" className="rs-primary" onClick={installApp}>
                このPCに追加
              </button>
            )}
          </div>
        )}

        {tab !== 'nippo' ? (
          <>
            <div className="rs-toolbar">
              <div className="rs-page-tabs">
                <button type="button" className={page === 'view' ? 'active' : ''} onClick={openView}>
                  閲覧
                </button>
                <button type="button" className={page === 'add' ? 'active' : ''} onClick={() => openAdd(!editId)}>
                  {editId ? '修正' : '追加'}
                </button>
              </div>
              <div className="rs-hint">{tab === 'ringi' ? '稟議の進捗を管理' : '修繕・備品の発注を管理'}</div>
            </div>

            {page === 'view' ? (
              filtered.length === 0 ? (
                <div className="rs-empty">
                  {tab === 'ringi' ? '稟議申請はまだありません' : '修繕・備品はまだありません'}
                  <div style={{ marginTop: 16 }}>
                    <button type="button" className="rs-primary" onClick={() => openAdd(true)}>
                      追加する
                    </button>
                  </div>
                </div>
              ) : (
                <div className="rs-grid">
                  {filtered.map((item) => (
                    <article key={item.id} className="rs-card" onClick={() => editItem(item)} role="button">
                      <div className="rs-card-top">
                        <div className="rs-title">{item.title}</div>
                        <button
                          type="button"
                          className="rs-status"
                          onClick={(event) => {
                            event.stopPropagation();
                            cycleStatus(item.id);
                          }}
                        >
                          {item.status || '未着手'}
                        </button>
                      </div>
                      {item.company && (
                        <div className="rs-row">
                          <b>会社</b> {item.company}
                        </div>
                      )}
                      {item.amount && (
                        <div className="rs-row">
                          <b>金額</b> ¥{item.amount}
                        </div>
                      )}
                      {item.replyDate && (
                        <div className="rs-row">
                          <b>返答</b> {item.replyDate}
                        </div>
                      )}
                      {item.stuckAt && (
                        <div className="rs-row">
                          <b>停止</b> {item.stuckAt}
                        </div>
                      )}
                      {(item.mailUrl || item.productUrl) && (
                        <div className="rs-links" onClick={(event) => event.stopPropagation()}>
                          {item.mailUrl && (
                            <a href={item.mailUrl} target="_blank" rel="noopener noreferrer">
                              メール
                            </a>
                          )}
                          {item.productUrl && (
                            <a href={item.productUrl} target="_blank" rel="noopener noreferrer">
                              商品
                            </a>
                          )}
                        </div>
                      )}
                      <button
                        type="button"
                        className="rs-delete"
                        onClick={(event) => {
                          event.stopPropagation();
                          removeItem(item.id);
                        }}
                      >
                        削除
                      </button>
                    </article>
                  ))}
                </div>
              )
            ) : (
              <form className="rs-form" onSubmit={saveItem}>
                <div className="rs-span-2">
                  <label className="rs-label">件名</label>
                  <input
                    className="rs-input"
                    value={form.title}
                    onChange={(e) => setForm((prev) => ({ ...prev, title: e.target.value }))}
                    placeholder="例: ケーブル修繕"
                    required
                  />
                </div>
                <div className="rs-span-2">
                  <label className="rs-label">会社（タップで選択）</label>
                  <div className="rs-chips">
                    {COMPANIES.map((name) => (
                      <button
                        key={name}
                        type="button"
                        className={`rs-chip${form.company === name ? ' active' : ''}`}
                        onClick={() => setForm((prev) => ({ ...prev, company: name }))}
                      >
                        {name}
                      </button>
                    ))}
                  </div>
                  <input
                    className="rs-input"
                    style={{ marginTop: 8 }}
                    value={form.company}
                    onChange={(e) => setForm((prev) => ({ ...prev, company: e.target.value }))}
                    placeholder="一覧にない会社はここに入力"
                  />
                </div>
                <div className="rs-span-2">
                  <label className="rs-label">金額</label>
                  <div className="rs-amount">
                    <span>¥</span>
                    <input
                      className="rs-input"
                      type="number"
                      min="0"
                      step="1"
                      value={form.amount}
                      onChange={(e) => setForm((prev) => ({ ...prev, amount: e.target.value }))}
                      placeholder="数字のみ"
                    />
                  </div>
                </div>
                <div>
                  <label className="rs-label">返答日</label>
                  <input
                    className="rs-input"
                    value={form.replyDate}
                    onChange={(e) => setForm((prev) => ({ ...prev, replyDate: e.target.value }))}
                    placeholder="例: 8/20"
                  />
                </div>
                <div>
                  <label className="rs-label">進捗</label>
                  <select
                    className="rs-input"
                    value={form.status}
                    onChange={(e) => setForm((prev) => ({ ...prev, status: e.target.value }))}
                  >
                    {STATUSES.map((status) => (
                      <option key={status} value={status}>
                        {status}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="rs-span-2">
                  <label className="rs-label">メールURL</label>
                  <input
                    className="rs-input"
                    type="url"
                    value={form.mailUrl}
                    onChange={(e) => setForm((prev) => ({ ...prev, mailUrl: e.target.value }))}
                    placeholder="https://"
                  />
                </div>
                <div className="rs-span-2">
                  <label className="rs-label">商品URL</label>
                  <input
                    className="rs-input"
                    type="url"
                    value={form.productUrl}
                    onChange={(e) => setForm((prev) => ({ ...prev, productUrl: e.target.value }))}
                    placeholder="https://"
                  />
                </div>
                <div className="rs-span-2">
                  <label className="rs-label">止まっているところ</label>
                  <input
                    className="rs-input"
                    value={form.stuckAt}
                    onChange={(e) => setForm((prev) => ({ ...prev, stuckAt: e.target.value }))}
                    placeholder="例: メーカー返答待ち"
                  />
                </div>
                <button type="submit" className="rs-primary rs-span-2">
                  {editId ? '更新して閲覧へ' : '保存して閲覧へ'}
                </button>
              </form>
            )}
          </>
        ) : (
          <>
            {ctx && (
              <section className="rs-cal">
                <div style={{ fontWeight: 700 }}>今日（{ctx.date}）</div>
                <div className="rs-label" style={{ marginTop: 10 }}>
                  カレンダー
                </div>
                {ctx.calendarEvents.length === 0 ? (
                  <div className="rs-hint">予定なし</div>
                ) : (
                  <ul>
                    {ctx.calendarEvents.map((ev, i) => (
                      <li key={i}>
                        {ev.start}–{ev.end} {ev.title}
                      </li>
                    ))}
                  </ul>
                )}
              </section>
            )}
            <div className="rs-nippo-actions">
              <button type="button" className="rs-ghost" onClick={loadContext} disabled={nippoLoading}>
                データ再取得
              </button>
              <button type="button" className="rs-ghost" onClick={fillGyomuFromCalendar} disabled={nippoLoading}>
                カレンダー → 業務欄
              </button>
            </div>
            <div className="rs-form">
              <div>
                <label className="rs-label">業務内容（1行1項目）</label>
                <textarea value={gyomuText} onChange={(e) => setGyomuText(e.target.value)} />
              </div>
              <div>
                <label className="rs-label">所感</label>
                <textarea value={kansou} onChange={(e) => setKansou(e.target.value)} />
              </div>
              <button type="button" className="rs-primary rs-span-2" onClick={createDraft} disabled={nippoLoading}>
                Gmail に日報下書きを作成
              </button>
            </div>
            {nippoStatus && <p className="rs-status-text">{nippoStatus}</p>}
          </>
        )}
      </main>
    </div>
  );
}

interface BeforeInstallPromptEvent extends Event {
  prompt: () => Promise<void>;
}
