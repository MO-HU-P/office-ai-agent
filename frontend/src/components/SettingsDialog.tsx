import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { deleteModel, fetchModels, fetchSettings, pullModel, updateSettings } from '../api'
import type { LLMMode, ModelInfo, SettingsInfo } from '../types'

// 追加入力欄のサジェスト(datalist)。自由入力も可能
const SUGGESTED: Record<LLMMode, string[]> = {
  local: ['gpt-oss:20b', 'qwen3.5:9b', 'gemma4:12b'],
  cloud: ['gpt-oss:120b', 'gemma4:31b', 'qwen3.5:397b', 'nemotron-3-ultra'],
}


// gpt-oss系は思考の深さ(low/medium/high)指定、それ以外はオン/オフ指定
const isLevelModel = (model: string) => model.startsWith('gpt-oss')

const REASONING_OPTIONS = {
  level: [
    { value: 'auto', label: '自動（おすすめ）' },
    { value: 'low', label: '短く考える（すばやい返答）' },
    { value: 'medium', label: 'ふつうに考える' },
    { value: 'high', label: 'じっくり考える（高品質・低速）' },
  ],
  toggle: [
    { value: 'auto', label: '自動（おすすめ）' },
    { value: 'false', label: '考えずにすぐ答える（すばやい返答）' },
    { value: 'true', label: 'じっくり考える（高品質・低速）' },
  ],
}

function formatSize(bytes?: number | null): string {
  if (!bytes) return ''
  return `${(bytes / 1e9).toFixed(1)} GB`
}

interface Props {
  onClose: () => void
  onSaved: () => void
}

interface PullState {
  name: string
  percent: number | null
  status: string
}

export function SettingsDialog({ onClose, onSaved }: Props) {
  const [settings, setSettings] = useState<SettingsInfo | null>(null)
  const [mode, setMode] = useState<LLMMode>('cloud')
  const [modelLocal, setModelLocal] = useState('')
  const [modelCloud, setModelCloud] = useState('')
  const [reasoning, setReasoning] = useState('auto')
  const [models, setModels] = useState<Partial<Record<LLMMode, ModelInfo[]>>>({})
  const [unavailable, setUnavailable] = useState<Partial<Record<LLMMode, string>>>({})
  const [addName, setAddName] = useState('')
  const [pull, setPull] = useState<PullState | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [saving, setSaving] = useState(false)
  const abortRef = useRef<AbortController | null>(null)

  const loadModels = useCallback((m: LLMMode) => {
    fetchModels(m)
      .then((res) => {
        setModels((prev) => ({ ...prev, [m]: res.models }))
        setUnavailable((prev) => ({ ...prev, [m]: res.unavailable }))
      })
      .catch(() => setUnavailable((prev) => ({ ...prev, [m]: 'モデル一覧を取得できませんでした' })))
  }, [])

  useEffect(() => {
    fetchSettings()
      .then((s) => {
        setSettings(s)
        setMode(s.mode)
        setModelLocal(s.model_local)
        setModelCloud(s.model_cloud)
        setReasoning(s.reasoning)
        loadModels(s.mode)
      })
      .catch(() => setError('設定を読み込めませんでした'))
  }, [loadModels])

  // モード切替時: そのモードのモデル一覧を未取得なら取得
  useEffect(() => {
    if (settings && models[mode] === undefined && !unavailable[mode]) loadModels(mode)
  }, [mode, settings, models, unavailable, loadModels])

  // ダイアログを閉じたらダウンロードも中断する(部分ダウンロードは再開可能)
  useEffect(() => () => abortRef.current?.abort(), [])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  const currentModel = mode === 'cloud' ? modelCloud : modelLocal
  const setCurrentModel = mode === 'cloud' ? setModelCloud : setModelLocal
  const reasoningOptions = isLevelModel(currentModel) ? REASONING_OPTIONS.level : REASONING_OPTIONS.toggle

  // モデル種別が変わって選択中のreasoningが選択肢から消えたら「自動」へ戻す
  useEffect(() => {
    if (!reasoningOptions.some((o) => o.value === reasoning)) setReasoning('auto')
  }, [reasoningOptions, reasoning])

  const listed = models[mode] ?? []
  // 選択中モデルを先頭に出す(長い一覧でも現在の選択がすぐ見えるように)。
  // 一覧に無い場合(未ダウンロード等)も行として表示する
  const rows = useMemo(() => {
    const rest = listed.filter((m) => m.name !== currentModel).map((m) => ({ ...m, missing: false }))
    if (!currentModel) return rest
    const selected = listed.find((m) => m.name === currentModel)
    return selected
      ? [{ ...selected, missing: false }, ...rest]
      : [{ name: currentModel, size: null, vision: false, missing: mode === 'local' }, ...rest]
  }, [listed, currentModel, mode])

  const startPull = useCallback(
    async (name: string) => {
      name = name.trim()
      if (!name || pull) return
      setError(null)
      setPull({ name, percent: null, status: '開始しています…' })
      const ctrl = new AbortController()
      abortRef.current = ctrl
      try {
        await pullModel(
          name,
          (p) => {
            setPull({
              name,
              percent: p.total && p.completed !== undefined ? Math.round((p.completed / p.total) * 100) : null,
              status: p.status ?? '',
            })
          },
          ctrl.signal,
        )
        setCurrentModel(name)
        setAddName('')
        loadModels('local')
      } catch (e) {
        if (!ctrl.signal.aborted) setError(e instanceof Error ? e.message : 'ダウンロードに失敗しました')
      } finally {
        setPull(null)
        abortRef.current = null
      }
    },
    [pull, setCurrentModel, loadModels],
  )

  const handleAdd = useCallback(() => {
    const name = addName.trim()
    if (!name) return
    if (mode === 'local') {
      void startPull(name)
    } else {
      // クラウドはダウンロード不要だが、提供終了・入力ミスのモデル名を
      // そのまま保存してしまわないよう、取得済みの一覧と照合する
      const cloudList = models.cloud ?? []
      if (cloudList.length > 0) {
        const exists = cloudList.some((m) => m.name === name || m.name.startsWith(`${name}:`))
        if (!exists) {
          setError(`「${name}」は Ollama Cloud のモデル一覧に見つかりません。上の一覧から選ぶか、名前を確認してください。`)
          return
        }
      }
      setError(null)
      setModelCloud(name)
      setAddName('')
    }
  }, [addName, mode, models, startPull])

  const handleDelete = useCallback(
    async (name: string) => {
      if (!window.confirm(`モデル「${name}」を削除しますか？\n(必要になったら再ダウンロードできます)`)) return
      setError(null)
      try {
        await deleteModel(name)
        loadModels('local')
      } catch (e) {
        setError(e instanceof Error ? e.message : '削除に失敗しました')
      }
    },
    [loadModels],
  )

  const handleSave = useCallback(async () => {
    setSaving(true)
    setError(null)
    try {
      await updateSettings({ mode, model_local: modelLocal, model_cloud: modelCloud, reasoning })
      onSaved()
      onClose()
    } catch (e) {
      setError(e instanceof Error ? e.message : '保存に失敗しました')
      setSaving(false)
    }
  }, [mode, modelLocal, modelCloud, reasoning, onSaved, onClose])

  return (
    <div className="dialog-overlay" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="dialog" role="dialog" aria-modal="true" aria-label="設定">
        <div className="dialog-head">
          <h2>設定</h2>
          <button className="icon-btn" onClick={onClose} aria-label="閉じる" title="閉じる">
            <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor">
              <path d="M19 6.41 17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z" />
            </svg>
          </button>
        </div>

        {!settings && !error && <div className="dialog-loading">読み込み中…</div>}

        {settings && (
          <div className="dialog-body">
            <section className="settings-section">
              <h3>AIの実行場所</h3>
              <div className="mode-cards">
                <button
                  className={`mode-card ${mode === 'cloud' ? 'selected' : ''}`}
                  onClick={() => setMode('cloud')}
                >
                  <span className="mode-card-title">クラウド</span>
                  <span className="mode-card-desc">高性能なAIをインターネット経由で利用</span>
                </button>
                <button
                  className={`mode-card ${mode === 'local' ? 'selected' : ''}`}
                  onClick={() => setMode('local')}
                >
                  <span className="mode-card-title">このパソコン</span>
                  <span className="mode-card-desc">データを外部に送らず手元で実行</span>
                </button>
              </div>
              {mode === 'cloud' && !settings.cloud_key_configured && (
                <p className="settings-warning">
                  クラウド利用には APIキーが必要です。<code>.env</code> ファイルに{' '}
                  <code>OLLAMA_API_KEY</code> を設定してください（安全のため、この画面からは設定できません）。
                </p>
              )}
            </section>

            <section className="settings-section">
              <h3>使用するAIモデル</h3>
              {unavailable[mode] ? (
                <p className="settings-warning">{unavailable[mode]}</p>
              ) : (
                <div className="model-list">
                  {rows.map((m) => (
                    <label key={m.name} className={`model-row ${currentModel === m.name ? 'selected' : ''}`}>
                      <input
                        type="radio"
                        name="model"
                        checked={currentModel === m.name}
                        onChange={() => setCurrentModel(m.name)}
                      />
                      <span className="model-name">{m.name}</span>
                      {m.vision && <span className="model-badge">画像対応</span>}
                      {m.missing ? (
                        <span className="model-note">未ダウンロード</span>
                      ) : (
                        // サイズはディスク消費の目安としてローカルのみ表示
                        mode === 'local' && <span className="model-size">{formatSize(m.size)}</span>
                      )}
                      {mode === 'local' && !m.missing && currentModel !== m.name && (
                        <button
                          className="icon-btn small"
                          onClick={(e) => {
                            e.preventDefault()
                            void handleDelete(m.name)
                          }}
                          disabled={!!pull}
                          aria-label={`${m.name} を削除`}
                          title="このモデルを削除してディスクを空ける"
                        >
                          <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor">
                            <path d="M6 19c0 1.1.9 2 2 2h8a2 2 0 0 0 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z" />
                          </svg>
                        </button>
                      )}
                    </label>
                  ))}
                  {rows.length === 0 && <p className="model-note">モデルがありません。下の欄から追加してください。</p>}
                </div>
              )}
              {rows.some((m) => m.vision) && (
                <p className="model-note">
                  「画像対応」のモデルを選ぶと、AIがスライドや文書の見た目を画像で確認しながら修正できます。
                </p>
              )}

              {pull ? (
                <div className="pull-progress">
                  <div className="pull-progress-label">
                    <span className="spinner" /> {pull.name} をダウンロード中… {pull.percent !== null ? `${pull.percent}%` : ''}
                  </div>
                  <div className="progress-track">
                    <div
                      className={`progress-fill ${pull.percent === null ? 'indeterminate' : ''}`}
                      style={pull.percent !== null ? { width: `${pull.percent}%` } : undefined}
                    />
                  </div>
                  <p className="model-note">この画面を閉じると中断されます（あとから再開できます）</p>
                </div>
              ) : (
                <div className="model-add">
                  <input
                    list={`model-suggest-${mode}`}
                    placeholder={mode === 'local' ? 'モデル名を入力してダウンロード (例: qwen3:8b)' : 'モデル名を入力して追加 (例: gpt-oss:20b)'}
                    value={addName}
                    onChange={(e) => setAddName(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleAdd()}
                  />
                  <datalist id={`model-suggest-${mode}`}>
                    {SUGGESTED[mode].map((n) => (
                      <option key={n} value={n} />
                    ))}
                  </datalist>
                  <button className="text-btn" onClick={handleAdd} disabled={!addName.trim()}>
                    {mode === 'local' ? 'ダウンロード' : '追加'}
                  </button>
                </div>
              )}
            </section>

            <section className="settings-section">
              <h3>回答の考え方</h3>
              <select className="settings-select" value={reasoning} onChange={(e) => setReasoning(e.target.value)}>
                {reasoningOptions.map((o) => (
                  <option key={o.value} value={o.value}>
                    {o.label}
                  </option>
                ))}
              </select>
              <p className="model-note">「じっくり考える」ほど回答の質が上がりますが、返答に時間がかかります。</p>
            </section>

            {error && <p className="settings-error">{error}</p>}
          </div>
        )}

        <div className="dialog-actions">
          <button className="text-btn" onClick={onClose}>
            キャンセル
          </button>
          <button className="primary-btn" onClick={handleSave} disabled={!settings || saving || !!pull}>
            {saving ? '保存中…' : '保存'}
          </button>
        </div>
      </div>
    </div>
  )
}
