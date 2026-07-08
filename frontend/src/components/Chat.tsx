import { useEffect, useRef, useState } from 'react'
import ReactMarkdown from 'react-markdown'
import { useSpeechInput } from '../hooks/useSpeechInput'
import type { ChatMessage, TargetSelection, ToolCallPart } from '../types'

const TOOL_LABELS: Record<string, string> = {
  list_files: 'ファイル一覧',
  copy_file: 'ファイルコピー',
  rename_file: '名前の変更',
  delete_file: 'ファイル削除',
  run_python: 'Python実行',
  excel_create: 'Excel作成',
  excel_read: 'Excel読み取り',
  excel_query: '行の検索',
  excel_write_cells: 'セル書き込み',
  excel_write_rows: 'データ書き込み',
  excel_format: '書式設定',
  excel_add_sheet: 'シート追加',
  word_create: 'Word作成',
  word_read: 'Word読み取り',
  word_find: '段落の検索',
  word_append: '段落追加',
  word_edit_paragraph: '段落編集',
  word_batch_edit: '段落一括編集',
  word_add_table: '表追加',
  word_dump_style: '体裁の読み取り',
  word_apply_style: '体裁の適用',
  ppt_create: 'スライド作成',
  ppt_add_slide: 'スライド追加',
  ppt_read: 'スライド読み取り',
  ppt_edit_slide: 'スライド編集',
  ppt_batch_edit: 'スライド一括編集',
  ppt_add_shape: '図形追加',
  ppt_add_image: '画像挿入',
  ppt_add_table: '表追加',
  ppt_add_chart: 'グラフ追加',
  ppt_delete_shape: '図形削除',
  ppt_move_shape: '図形の移動',
  render_page: '見た目の確認',
  merge_template: 'テンプレート差し込み',
  check_document_issues: '文書チェック',
  doc_diff: 'ファイルの比較',
  word_suggest_edits: '変更履歴で提案',
  word_add_comments: 'コメント追加',
  anonymize_file: '個人情報のマスク',
  list_file_versions: 'バックアップ一覧',
  restore_file: '編集の巻き戻し',
}

const SUGGESTIONS = [
  '月別売上のサンプルExcelを作って、合計行と書式も設定して',
  '「生成AIの業務活用」について5枚構成のプレゼン資料を作って',
  '会議の議事録テンプレートをWordで作って',
]

// 表示中のファイルに対して1クリックで実行できる校閲プリセット。
// promptは対象ファイル名を差し込んだ指示文で、押すとそのままチャットに送る。
const PRESETS: { label: string; icon: string; prompt: (file: string) => string }[] = [
  { label: '校正', icon: '✍️', prompt: (f) => `「${f}」の誤字脱字・文法の誤り・不自然な言い回しを校正して、直した箇所を教えてください。` },
  { label: '要約', icon: '📝', prompt: (f) => `「${f}」の内容を、要点がわかるように日本語で要約してください。` },
  { label: '匿名化', icon: '🕶️', prompt: (f) => `「${f}」に含まれる個人情報を、元のファイルは残したまま匿名化してください。` },
]

function ToolChip({ part }: { part: ToolCallPart }) {
  const [open, setOpen] = useState(false)
  const label = TOOL_LABELS[part.name] ?? part.name
  return (
    <div className={`tool-chip ${part.status}`}>
      <button className="tool-chip-head" onClick={() => setOpen(!open)}>
        <span className="tool-chip-icon">
          {part.status === 'running' ? <span className="spinner" /> : part.status === 'ok' ? '✓' : '!'}
        </span>
        <span className="tool-chip-label">{label}</span>
        <code className="tool-chip-name">{part.name}</code>
      </button>
      {open && (
        <div className="tool-chip-detail">
          <div><b>引数:</b> <code>{part.args}</code></div>
          {part.result && <div><b>結果:</b> {part.result}</div>}
        </div>
      )}
    </div>
  )
}

function Message({ msg }: { msg: ChatMessage }) {
  if (msg.role === 'user') {
    const text = msg.parts.map((p) => (p.kind === 'text' ? p.content : '')).join('')
    return <div className="msg user"><div className="bubble">{text}</div></div>
  }
  return (
    <div className="msg assistant">
      <div className="avatar">AI</div>
      <div className="assistant-body">
        {msg.parts.map((part, i) =>
          part.kind === 'text' ? (
            <div className="markdown" key={i}>
              <ReactMarkdown>{part.content}</ReactMarkdown>
            </div>
          ) : (
            <ToolChip key={i} part={part} />
          ),
        )}
        {msg.parts.length === 0 && <div className="typing"><span /><span /><span /></div>}
      </div>
    </div>
  )
}

interface ChatProps {
  messages: ChatMessage[]
  busy: boolean
  connected: boolean
  statusWarning: string | null
  modelName: string
  width: number
  activeFile: string | null
  /** プレビュー上でマウス選択された「対象箇所」。次のメッセージ冒頭に差し込む */
  target: TargetSelection | null
  onClearTarget: () => void
  onSend: (text: string) => void
  onReset: () => void
}

export function Chat({ messages, busy, connected, statusWarning, modelName, width, activeFile, target, onClearTarget, onSend, onReset }: ChatProps) {
  const [input, setInput] = useState('')
  const scrollRef = useRef<HTMLDivElement>(null)
  const textareaRef = useRef<HTMLTextAreaElement>(null)
  const speech = useSpeechInput()
  const speechBaseRef = useRef('')  // 音声入力開始時点で入力欄にあったテキスト

  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight })
  }, [messages])

  // 音声認識などプログラムからの入力更新でも高さを追従させる
  useEffect(() => {
    const ta = textareaRef.current
    if (!ta) return
    ta.style.height = 'auto'
    ta.style.height = Math.min(ta.scrollHeight, 160) + 'px'
  }, [input])

  const toggleSpeech = () => {
    if (speech.listening) {
      speech.stop()
      return
    }
    speechBaseRef.current = input ? input.trimEnd() + ' ' : ''
    speech.start((text) => setInput(speechBaseRef.current + text))
  }

  const send = () => {
    const text = input.trim()
    if (!text || busy || !connected) return
    speech.stop()
    // プレビューで選択した対象箇所があれば、メッセージ冒頭に付けてAIに場所を伝える
    onSend(target ? `（対象箇所: ${target.file} の ${target.label}）\n${text}` : text)
    onClearTarget()
    setInput('')
    if (textareaRef.current) textareaRef.current.style.height = 'auto'
  }

  return (
    <div className="chat-pane" style={{ width }}>
      <div className="chat-scroll" ref={scrollRef}>
        {messages.length === 0 ? (
          <div className="empty-chat">
            <div className="empty-logo">
              <span style={{ color: '#4285f4' }}>O</span>
              <span style={{ color: '#ea4335' }}>f</span>
              <span style={{ color: '#fbbc04' }}>f</span>
              <span style={{ color: '#4285f4' }}>i</span>
              <span style={{ color: '#34a853' }}>c</span>
              <span style={{ color: '#ea4335' }}>e</span>
              <span className="empty-logo-sub"> AI Agent</span>
            </div>
            <p className="empty-hint">
              Word・Excel・PowerPointのファイルを、チャットで指示するだけで作成・編集します。
            </p>
            {statusWarning && <p className="model-warning">⏳ {statusWarning}</p>}
            <div className="suggestions">
              {SUGGESTIONS.map((s) => (
                <button key={s} className="suggestion" onClick={() => onSend(s)} disabled={busy || !connected}>
                  {s}
                </button>
              ))}
            </div>
          </div>
        ) : (
          messages.map((m, i) => <Message key={i} msg={m} />)
        )}
      </div>
      <div className="chat-input-area">
        {target && (
          <div className="target-chip" title="次のメッセージは、この場所を対象として送られます">
            <span className="target-chip-label">📍 {target.file} の {target.label}</span>
            <button onClick={onClearTarget} aria-label="対象の選択を解除" title="選択を解除">✕</button>
          </div>
        )}
        {activeFile && (
          <div className="quick-presets" title={`表示中の「${activeFile}」に対して実行します`}>
            {PRESETS.map((p) => (
              <button
                key={p.label}
                className="preset-chip"
                disabled={busy || !connected}
                onClick={() => onSend(p.prompt(activeFile))}
              >
                <span aria-hidden>{p.icon}</span> {p.label}
              </button>
            ))}
          </div>
        )}
        <div className="chat-input-box">
          <textarea
            ref={textareaRef}
            value={input}
            placeholder={
              speech.listening ? 'どうぞお話しください…' : connected ? 'AIエージェントに指示する…' : 'サーバーに接続中…'
            }
            rows={1}
            disabled={!connected}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) {
                e.preventDefault()
                send()
              }
            }}
          />
          {speech.supported && (
            <button
              className={`mic-btn ${speech.listening ? 'listening' : ''}`}
              onClick={toggleSpeech}
              disabled={!connected}
              title={speech.listening ? '音声入力を停止' : '音声で入力'}
              aria-label={speech.listening ? '音声入力を停止' : '音声で入力'}
            >
              <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor">
                <path d="M12 14a3 3 0 0 0 3-3V5a3 3 0 0 0-6 0v6a3 3 0 0 0 3 3zm5-3a5 5 0 0 1-10 0H5a7 7 0 0 0 6 6.92V21h2v-3.08A7 7 0 0 0 19 11h-2z" />
              </svg>
            </button>
          )}
          <button className="send-btn" onClick={send} disabled={busy || !connected || !input.trim()} title="送信">
            <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor">
              <path d="M2 21l21-9L2 3v7l15 2-15 2v7z" />
            </svg>
          </button>
        </div>
        {speech.error && <p className="speech-error">{speech.error}</p>}
        <div className="chat-footer">
          <span className={`status-dot ${connected ? 'on' : 'off'}`} />
          {speech.listening ? '🎤 音声を聞き取り中…' : connected ? `接続中 · ${modelName}` : '再接続中…'}
          {messages.length > 0 && (
            <button className="reset-btn" onClick={onReset} disabled={busy}>会話をクリア</button>
          )}
        </div>
      </div>
    </div>
  )
}
