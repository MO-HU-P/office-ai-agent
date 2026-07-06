import { useCallback, useEffect, useRef, useState } from 'react'
import { fetchFiles, fetchHealth } from './api'
import { Chat } from './components/Chat'
import { PreviewPane } from './components/PreviewPane'
import { SettingsDialog } from './components/SettingsDialog'
import { useAgentSocket } from './hooks/useAgentSocket'
import type { FileInfo, HealthInfo } from './types'

export default function App() {
  const [files, setFiles] = useState<FileInfo[]>([])
  const [activeFile, setActiveFile] = useState<string | null>(null)
  const [refreshKey, setRefreshKey] = useState(0)
  const [health, setHealth] = useState<HealthInfo | null>(null)
  const [settingsOpen, setSettingsOpen] = useState(false)
  const activeFileRef = useRef(activeFile)
  activeFileRef.current = activeFile

  // 左右ペインの境界をドラッグして幅を変える
  const CHAT_MIN = 320
  const [chatWidth, setChatWidth] = useState(() => {
    const saved = Number(localStorage.getItem('chatWidth'))
    return saved >= CHAT_MIN && saved <= 1000 ? saved : 420
  })
  const draggingRef = useRef(false)

  const startResize = useCallback((e: React.MouseEvent) => {
    e.preventDefault()
    draggingRef.current = true
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
  }, [])

  useEffect(() => {
    const onMove = (e: MouseEvent) => {
      if (!draggingRef.current) return
      const max = Math.min(1000, window.innerWidth - 360) // プレビュー側に最低360px残す
      setChatWidth(Math.max(CHAT_MIN, Math.min(max, e.clientX)))
    }
    const onUp = () => {
      if (!draggingRef.current) return
      draggingRef.current = false
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
  }, [])

  useEffect(() => {
    localStorage.setItem('chatWidth', String(chatWidth))
  }, [chatWidth])

  const reloadFiles = useCallback(() => {
    fetchFiles().then(setFiles).catch(() => {})
  }, [])

  const handleDocUpdated = useCallback((filename: string) => {
    reloadFiles()
    // エージェントが触ったファイルを自動で開き、開いていれば再描画する
    if (activeFileRef.current !== filename) {
      setActiveFile(filename)
    }
    setRefreshKey((k) => k + 1)
  }, [reloadFiles])

  const handleDocDeleted = useCallback((filename: string) => {
    reloadFiles()
    if (activeFileRef.current === filename) {
      setActiveFile(null)
    }
  }, [reloadFiles])

  const { messages, busy, connected, sendMessage, resetChat } = useAgentSocket({
    onDocUpdated: handleDocUpdated,
    onDocDeleted: handleDocDeleted,
  })

  const checkHealth = useCallback(() => {
    fetchHealth().then(setHealth).catch(() => setHealth(null))
  }, [])

  useEffect(() => {
    reloadFiles()
    checkHealth()
    const timer = setInterval(checkHealth, 15000)
    return () => clearInterval(timer)
  }, [reloadFiles, checkHealth])

  return (
    <div className="app">
      <header className="app-header">
        <div className="app-title">
          <svg viewBox="0 0 24 24" width="26" height="26">
            <rect x="3" y="3" width="8" height="8" rx="1.5" fill="#4285f4" />
            <rect x="13" y="3" width="8" height="8" rx="1.5" fill="#ea4335" />
            <rect x="3" y="13" width="8" height="8" rx="1.5" fill="#34a853" />
            <rect x="13" y="13" width="8" height="8" rx="1.5" fill="#fbbc04" />
          </svg>
          <h1>Office AI Agent</h1>
        </div>
        <div className="header-right">
          <div className="header-status">
            <span className={`status-dot ${health?.backend_ok ? (health.model_ready ? 'on' : 'warn') : 'off'}`} />
            {health?.key_missing
              ? `${health.provider === 'openai' ? 'OPENAI_API_KEY' : 'OLLAMA_API_KEY'} が未設定です (.env)`
              : health?.provider === 'openai'
                ? `OpenAI · ${health.model}`
                : health?.backend_ok
                  ? health.model_ready
                    ? `${health.mode === 'cloud' ? 'Ollama Cloud' : 'Ollama'} · ${health.model}`
                    : health.mode === 'cloud'
                      ? `モデル ${health.model} は利用できません`
                      : `モデル ${health.model} をダウンロード中…`
                  : 'Ollama 未接続'}
          </div>
          <button
            className="icon-btn"
            onClick={() => setSettingsOpen(true)}
            aria-label="設定"
            title="設定 (AIモデルの切り替え・追加・削除)"
          >
            <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor">
              <path d="M19.14 12.94c.04-.3.06-.61.06-.94 0-.32-.02-.64-.07-.94l2.03-1.58a.49.49 0 0 0 .12-.61l-1.92-3.32a.488.488 0 0 0-.59-.22l-2.39.96c-.5-.38-1.03-.7-1.62-.94l-.36-2.54a.484.484 0 0 0-.48-.41h-3.84c-.24 0-.43.17-.47.41l-.36 2.54c-.59.24-1.13.57-1.62.94l-2.39-.96c-.22-.08-.47 0-.59.22L2.74 8.87c-.12.21-.08.47.12.61l2.03 1.58c-.05.3-.09.63-.09.94s.02.64.07.94l-2.03 1.58a.49.49 0 0 0-.12.61l1.92 3.32c.12.22.37.29.59.22l2.39-.96c.5.38 1.03.7 1.62.94l.36 2.54c.05.24.24.41.48.41h3.84c.24 0 .44-.17.47-.41l.36-2.54c.59-.24 1.13-.56 1.62-.94l2.39.96c.22.08.47 0 .59-.22l1.92-3.32c.12-.22.07-.47-.12-.61l-2.01-1.58zM12 15.6A3.61 3.61 0 0 1 8.4 12c0-1.98 1.62-3.6 3.6-3.6s3.6 1.62 3.6 3.6-1.62 3.6-3.6 3.6z" />
            </svg>
          </button>
        </div>
      </header>
      {settingsOpen && (
        <SettingsDialog onClose={() => setSettingsOpen(false)} onSaved={checkHealth} />
      )}
      <main className="app-main">
        <Chat
          messages={messages}
          busy={busy}
          connected={connected}
          statusWarning={
            health?.key_missing
              ? `.env に ${health.provider === 'openai' ? 'OPENAI_API_KEY' : 'OLLAMA_API_KEY'} を設定して docker compose up -d を再実行してください。`
              : health && !health.model_ready && health.provider === 'ollama'
                ? health.mode === 'cloud'
                  ? health.backend_ok
                    ? `選択中のAIモデル「${health.model}」は Ollama Cloud で見つかりません。提供終了した可能性があります。右上の設定（歯車アイコン）から別のモデルを選んでください。`
                    : 'Ollama Cloud に接続できていません。APIキーとネットワークを確認してください。'
                  : `モデル ${health.model} をダウンロード中です。初回は数分かかります。`
                : null
          }
          modelName={health?.model ?? '…'}
          width={chatWidth}
          onSend={sendMessage}
          onReset={resetChat}
        />
        <div
          className="pane-resizer"
          onMouseDown={startResize}
          role="separator"
          aria-orientation="vertical"
          title="ドラッグして幅を変更"
        />
        <PreviewPane
          files={files}
          activeFile={activeFile}
          refreshKey={refreshKey}
          onSelect={setActiveFile}
          onFilesChanged={reloadFiles}
        />
      </main>
    </div>
  )
}
