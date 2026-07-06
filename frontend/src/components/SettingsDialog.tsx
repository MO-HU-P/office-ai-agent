import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { deleteModel, fetchModels, fetchSettings, pullModel, updateSettings } from '../api'
import type { ExternalProvider, LLMProvider, ModelInfo, ModelSource, SettingsInfo } from '../types'

// 追加入力欄のサジェスト(datalist)。自由入力も可能。
// 外部プロバイダーの候補は config.toml(llm.*_models)由来の一覧から動的に出すため、ここは空。
const SUGGESTED: Record<ModelSource, string[]> = {
  local: ['gpt-oss:20b', 'qwen3.5:9b', 'gemma4:12b'],
  cloud: ['gpt-oss:120b', 'gemma4:31b', 'qwen3.5:397b', 'nemotron-3-ultra'],
  openai: [],
  gemini: [],
}

// 追加入力欄のプレースホルダーに出すモデル名の例
const ADD_EXAMPLE: Record<ModelSource, string> = {
  local: 'qwen3:8b',
  cloud: 'gpt-oss:120b',
  openai: 'gpt-4o-mini',
  gemini: 'gemini-2.5-flash',
}

// 外部プロバイダー(=自由入力・候補保存ができるソース)かの判定。TSの型絞り込みにも使う
const isExternalSource = (src: ModelSource): src is ExternalProvider => src !== 'local' && src !== 'cloud'

// 上段「AIの実行場所」= このパソコン(手元) / クラウド(ネット越しの高性能AI)の2択。
// プロバイダーが何個増えてもこの2枚は固定で、増えるのは下段「利用するサービス」側。
type Location = 'local' | 'cloud'
const LOCATION_CARDS: { key: Location; title: string; desc: string }[] = [
  { key: 'cloud', title: 'クラウド', desc: '高性能なAIをインターネット経由で利用' },
  { key: 'local', title: 'このパソコン', desc: 'データを外部に送らず手元で実行' },
]

// 下段「利用するサービス」= クラウド選択時に縦リストで並ぶ。プロバイダー追加は原則ここに1行
// 足すだけ(backend providers.py に _build_* を実装し provider を返す)。requiresKey を持つ
// ものは、その鍵が .env にある(または現在選択中の)ときだけ表示する opt-in。縦リストなので
// 何個増えてもレイアウトが崩れない。
interface CloudService {
  key: string
  provider: LLMProvider
  source: ModelSource
  title: string
  keyEnv: string // キー未設定時の警告に出す .env の変数名
  requiresKey?: 'openai_key_configured' | 'gemini_key_configured' // 表示条件に使う settings のフラグ名
}
const CLOUD_SERVICES: CloudService[] = [
  { key: 'ollama-cloud', provider: 'ollama', source: 'cloud', title: 'Ollama Cloud', keyEnv: 'OLLAMA_API_KEY' },
  { key: 'openai', provider: 'openai', source: 'openai', title: 'OpenAI', keyEnv: 'OPENAI_API_KEY', requiresKey: 'openai_key_configured' },
  { key: 'gemini', provider: 'gemini', source: 'gemini', title: 'Google Gemini', keyEnv: 'GEMINI_API_KEY', requiresKey: 'gemini_key_configured' },
  // 追加例: { key: 'anthropic', provider: 'anthropic', source: 'anthropic', title: 'Anthropic (Claude)', keyEnv: 'ANTHROPIC_API_KEY', requiresKey: 'anthropic_key_configured' },
]

// gpt-oss系は思考の深さ(low/medium/high)指定、それ以外のOllamaはオン/オフ指定
const isLevelModel = (model: string) => model.startsWith('gpt-oss')
// OpenAIの推論モデル(gpt-5系/o系)。これらだけが reasoning_effort を受け付ける。
// backend providers.py の _OPENAI_REASONING_RE と対応させること(gpt-4o系/gpt-4.1系は
// temperature固定で思考の深さを持たないため、設定を出しても効かない)。
const isOpenAIReasoningModel = (model: string) => /^(?:o\d|gpt-5)/.test(model)
// Gemini 3以降は思考の深さ(low/medium/high)指定、2.5系はオン/オフ指定。
// backend providers.py の _GEMINI_LEVEL_RE と対応させること。
const isGeminiLevelModel = (model: string) => /^gemini-(?:[3-9]|\d{2})/.test(model)

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
  const [provider, setProvider] = useState<LLMProvider>('ollama')
  const [mode, setMode] = useState<'local' | 'cloud'>('cloud')
  const [modelLocal, setModelLocal] = useState('')
  const [modelCloud, setModelCloud] = useState('')
  // 外部プロバイダーごとの選択モデル
  const [modelExternal, setModelExternal] = useState<Record<ExternalProvider, string>>({ openai: '', gemini: '' })
  // 自由入力し保存した外部プロバイダーのモデル候補(プリセットとは別・削除可能)。保存はsettings.json
  const [customExternal, setCustomExternal] = useState<Record<ExternalProvider, string[]>>({ openai: [], gemini: [] })
  const [reasoning, setReasoning] = useState('auto')
  const [models, setModels] = useState<Partial<Record<ModelSource, ModelInfo[]>>>({})
  const [unavailable, setUnavailable] = useState<Partial<Record<ModelSource, string>>>({})
  const [addName, setAddName] = useState('')
  const [pull, setPull] = useState<PullState | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [saving, setSaving] = useState(false)
  const abortRef = useRef<AbortController | null>(null)

  // 現在選ばれている取得元(Ollamaは mode、外部プロバイダーはその名前)
  const source: ModelSource = provider === 'ollama' ? mode : provider

  const loadModels = useCallback((src: ModelSource) => {
    fetchModels(src)
      .then((res) => {
        setModels((prev) => ({ ...prev, [src]: res.models }))
        setUnavailable((prev) => ({ ...prev, [src]: res.unavailable }))
      })
      .catch(() => setUnavailable((prev) => ({ ...prev, [src]: 'モデル一覧を取得できませんでした' })))
  }, [])

  useEffect(() => {
    fetchSettings()
      .then((s) => {
        setSettings(s)
        setProvider(s.provider)
        setMode(s.mode)
        setModelLocal(s.model_local)
        setModelCloud(s.model_cloud)
        setModelExternal({ openai: s.model_openai, gemini: s.model_gemini })
        setCustomExternal({ openai: s.openai_custom_models ?? [], gemini: s.gemini_custom_models ?? [] })
        setReasoning(s.reasoning)
        loadModels(s.provider === 'ollama' ? s.mode : s.provider)
      })
      .catch(() => setError('設定を読み込めませんでした'))
  }, [loadModels])

  // 取得元の切替時: その一覧を未取得なら取得
  useEffect(() => {
    if (settings && models[source] === undefined && !unavailable[source]) loadModels(source)
  }, [source, settings, models, unavailable, loadModels])

  // ダイアログを閉じたらダウンロードも中断する(部分ダウンロードは再開可能)
  useEffect(() => () => abortRef.current?.abort(), [])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  const currentModel =
    provider === 'ollama' ? (mode === 'cloud' ? modelCloud : modelLocal) : modelExternal[provider]
  const setCurrentModel = useCallback(
    (name: string) => {
      if (provider === 'ollama') (mode === 'cloud' ? setModelCloud : setModelLocal)(name)
      else setModelExternal((prev) => ({ ...prev, [provider]: name }))
    },
    [provider, mode],
  )
  // 「回答の考え方」を出すか。Ollama/Geminiは常に。OpenAIは推論モデル(gpt-5系/o系)のときだけ
  // (gpt-4o系/gpt-4.1系はreasoning_effort非対応でtemperature固定のため、選ばせても効かない)
  const showReasoning = provider !== 'openai' || isOpenAIReasoningModel(currentModel)
  // OpenAIの推論モデルと Gemini 3以降は段階指定。Gemini 2.5系はオン/オフ。
  // Ollamaは gpt-oss系のみ段階、他はオン/オフ
  const reasoningOptions =
    provider === 'openai' || (provider === 'gemini' ? isGeminiLevelModel(currentModel) : isLevelModel(currentModel))
      ? REASONING_OPTIONS.level
      : REASONING_OPTIONS.toggle

  // 非表示のとき、およびモデル種別が変わって選択中のreasoningが選択肢から消えたら「自動」へ戻す
  useEffect(() => {
    if (!showReasoning || !reasoningOptions.some((o) => o.value === reasoning)) {
      if (reasoning !== 'auto') setReasoning('auto')
    }
  }, [showReasoning, reasoningOptions, reasoning])

  // 上段(場所)と下段(サービス)の選択状態を provider/mode から導出する
  const location: Location = provider === 'ollama' && mode === 'local' ? 'local' : 'cloud'
  const cloudKey = provider === 'ollama' ? 'ollama-cloud' : provider // openai/gemini 等は provider 名がキー
  // 鍵未設定のサービスは隠す(現在選択中は残す)。Ollama Cloud は鍵不要扱いで常に表示
  const visibleCloudServices = CLOUD_SERVICES.filter((s) => {
    if (cloudKey === s.key) return true
    if (!s.requiresKey) return true
    return Boolean(settings?.[s.requiresKey])
  })
  const selectedService = CLOUD_SERVICES.find((s) => s.key === cloudKey)
  const keyConfigured = selectedService?.requiresKey
    ? Boolean(settings?.[selectedService.requiresKey])
    : Boolean(settings?.cloud_key_configured)

  // 下段のサービスを選ぶ(Ollamaはcloud、OpenAI等もcloud固定で統一)
  const selectCloudService = (svc: CloudService) => {
    setProvider(svc.provider)
    setMode('cloud')
  }
  // 上段の場所を選ぶ。クラウドへ切替時は、表示中の先頭サービス(通常 Ollama Cloud)を既定に
  const selectLocation = (loc: Location) => {
    if (loc === 'local') {
      setProvider('ollama')
      setMode('local')
    } else if (location !== 'cloud') {
      const first = visibleCloudServices[0]
      if (first) selectCloudService(first)
    }
  }

  // 外部プロバイダーはプリセット(config.toml由来)＋自由入力の追加分(削除可能)を並べる。
  // 現行の外部プロバイダーのチャットモデルはすべて画像対応なので追加分は vision=true 扱い。
  const listed = useMemo<(ModelInfo & { custom?: boolean })[]>(() => {
    const base = models[source] ?? []
    if (!isExternalSource(source)) return base
    const presetNames = new Set(base.map((m) => m.name))
    const extras = customExternal[source]
      .filter((n) => !presetNames.has(n))
      .map((n) => ({ name: n, size: null, vision: true, custom: true }))
    return [...base, ...extras]
  }, [models, source, customExternal])
  // 選択中モデルを先頭に出す(長い一覧でも現在の選択がすぐ見えるように)。
  // 一覧に無い場合(未ダウンロード等)も行として表示する
  const rows = useMemo(() => {
    const rest = listed.filter((m) => m.name !== currentModel).map((m) => ({ ...m, missing: false }))
    if (!currentModel) return rest
    const selected = listed.find((m) => m.name === currentModel)
    const external = isExternalSource(source)
    return selected
      ? [{ ...selected, missing: false }, ...rest]
      : [
          // 外部プロバイダーで一覧に無い選択中モデル＝自由入力分。画像対応・削除可能として扱う
          { name: currentModel, size: null, vision: external, custom: external, missing: source === 'local' },
          ...rest,
        ]
  }, [listed, currentModel, source])

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
        setModelLocal(name)
        setAddName('')
        loadModels('local')
      } catch (e) {
        if (!ctrl.signal.aborted) setError(e instanceof Error ? e.message : 'ダウンロードに失敗しました')
      } finally {
        setPull(null)
        abortRef.current = null
      }
    },
    [pull, loadModels],
  )

  const handleAdd = useCallback(() => {
    const name = addName.trim()
    if (!name) return
    // APIキー(sk-... / AIza...)の貼り付け事故は、全プロバイダーの全欄で弾く
    // (秘密情報を settings.json に保存しないため。サーバー側でも二重に検証する)
    if (/^(sk-|AIza)/i.test(name)) {
      setError('APIキーのような値は入力できません。モデル名を入力してください。')
      return
    }
    if (source === 'local') {
      void startPull(name)
      return
    }
    if (source === 'cloud') {
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
    }
    setError(null)
    // 追加分は候補として保存(プリセットと重複しないもののみ)。保存は「保存」ボタン押下時
    if (isExternalSource(source)) {
      const presetNames = new Set((models[source] ?? []).map((m) => m.name))
      if (!presetNames.has(name)) {
        setCustomExternal((prev) =>
          prev[source].includes(name) ? prev : { ...prev, [source]: [...prev[source], name] },
        )
      }
    }
    setCurrentModel(name)
    setAddName('')
  }, [addName, source, models, startPull, setCurrentModel])

  // 自由入力した外部プロバイダーのモデル候補を一覧から削除(プリセットは対象外)。反映は「保存」押下時。
  const handleDeleteCustom = useCallback(
    (name: string) => {
      if (!isExternalSource(source)) return
      setError(null)
      setCustomExternal((prev) => ({ ...prev, [source]: prev[source].filter((n) => n !== name) }))
      // 削除したモデルを選択中だった場合は既定(プリセット先頭)へ戻す
      if (modelExternal[source] === name) {
        setModelExternal((prev) => ({ ...prev, [source]: models[source]?.[0]?.name ?? '' }))
      }
    },
    [source, modelExternal, models],
  )

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
      await updateSettings({
        provider,
        mode,
        model_local: modelLocal,
        model_cloud: modelCloud,
        // 空文字は「未選択」なので送らない(サーバー側の検証エラーを避ける)
        model_openai: modelExternal.openai || undefined,
        model_gemini: modelExternal.gemini || undefined,
        openai_custom_models: customExternal.openai,
        gemini_custom_models: customExternal.gemini,
        reasoning,
      })
      onSaved()
      onClose()
    } catch (e) {
      setError(e instanceof Error ? e.message : '保存に失敗しました')
      setSaving(false)
    }
  }, [provider, mode, modelLocal, modelCloud, modelExternal, customExternal, reasoning, onSaved, onClose])

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
                {LOCATION_CARDS.map((c) => (
                  <button
                    key={c.key}
                    className={`mode-card ${location === c.key ? 'selected' : ''}`}
                    onClick={() => selectLocation(c.key)}
                  >
                    <span className="mode-card-title">{c.title}</span>
                    <span className="mode-card-desc">{c.desc}</span>
                  </button>
                ))}
              </div>

              {/* クラウド選択時のみ、利用するサービスを縦リストで選ぶ(プロバイダーが増えても崩れない) */}
              {location === 'cloud' && (
                <div className="cloud-services">
                  <p className="settings-sublabel">利用するサービス</p>
                  <div className="model-list">
                    {visibleCloudServices.map((svc) => (
                      <label
                        key={svc.key}
                        className={`model-row ${cloudKey === svc.key ? 'selected' : ''}`}
                      >
                        <input
                          type="radio"
                          name="cloud-service"
                          checked={cloudKey === svc.key}
                          onChange={() => selectCloudService(svc)}
                        />
                        <span className="model-name">{svc.title}</span>
                      </label>
                    ))}
                  </div>
                </div>
              )}

              {location === 'cloud' && !keyConfigured && (
                <p className="settings-warning">
                  {selectedService ? `${selectedService.title}の利用` : 'クラウド利用'}には APIキーが必要です。
                  <code>.env</code> ファイルに <code>{selectedService?.keyEnv ?? 'OLLAMA_API_KEY'}</code>{' '}
                  を設定してください（安全のため、この画面からは設定できません）。
                </p>
              )}
            </section>

            <section className="settings-section">
              <h3>使用するAIモデル</h3>
              {unavailable[source] ? (
                <p className="settings-warning">{unavailable[source]}</p>
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
                        source === 'local' && <span className="model-size">{formatSize(m.size)}</span>
                      )}
                      {source === 'local' && !m.missing && currentModel !== m.name && (
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
                      {/* 自由入力で追加した外部プロバイダーのモデルは一覧から削除できる(プリセットは削除不可) */}
                      {isExternalSource(source) && m.custom && (
                        <button
                          className="icon-btn small"
                          onClick={(e) => {
                            e.preventDefault()
                            handleDeleteCustom(m.name)
                          }}
                          aria-label={`${m.name} を一覧から削除`}
                          title="このモデルを一覧から削除"
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
                    list={`model-suggest-${source}`}
                    placeholder={
                      source === 'local'
                        ? `モデル名を入力してダウンロード (例: ${ADD_EXAMPLE.local})`
                        : `モデル名を入力して追加 (例: ${ADD_EXAMPLE[source]})`
                    }
                    value={addName}
                    onChange={(e) => setAddName(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleAdd()}
                  />
                  <datalist id={`model-suggest-${source}`}>
                    {/* 外部プロバイダーの候補は config.toml 由来の一覧から出す(静的リストを陳腐化させない) */}
                    {(isExternalSource(source) ? listed.map((m) => m.name) : SUGGESTED[source]).map((n) => (
                      <option key={n} value={n} />
                    ))}
                  </datalist>
                  <button className="text-btn" onClick={handleAdd} disabled={!addName.trim()}>
                    {source === 'local' ? 'ダウンロード' : '追加'}
                  </button>
                </div>
              )}
            </section>

            {showReasoning && (
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
            )}

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
