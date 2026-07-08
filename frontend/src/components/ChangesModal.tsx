import { useEffect, useState } from 'react'
import { fetchChanges, restoreFile } from '../api'
import type { ChangesResult } from '../types'

interface ChangesModalProps {
  filename: string
  onClose: () => void
  /** 巻き戻し完了後に呼ぶ(プレビューとファイル一覧の再読み込み用) */
  onRestored: (filename: string) => void
}

function formatTime(iso: string): string {
  const d = new Date(iso)
  return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, '0')}:${String(
    d.getMinutes(),
  ).padStart(2, '0')}`
}

/** 「最後の変更でどこが変わったか」を色分け表示し、その変更前へ巻き戻せるモーダル。 */
export function ChangesModal({ filename, onClose, onRestored }: ChangesModalProps) {
  const [changes, setChanges] = useState<ChangesResult | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [restoring, setRestoring] = useState(false)

  useEffect(() => {
    let cancelled = false
    setChanges(null)
    setError(null)
    fetchChanges(filename)
      .then((data) => {
        if (!cancelled) setChanges(data)
      })
      .catch((e) => {
        if (!cancelled) setError(String(e.message ?? e))
      })
    return () => {
      cancelled = true
    }
  }, [filename])

  const handleRestore = async () => {
    if (!changes?.available) return
    const when = formatTime(changes.base_time)
    if (!confirm(`「${filename}」を ${when} 時点(この変更前)の状態に戻しますか？\n今の状態も自動バックアップされるため、巻き戻しはあとから取り消せます。`)) {
      return
    }
    setRestoring(true)
    try {
      await restoreFile(filename, changes.base_version)
      onRestored(filename)
      onClose()
    } catch (e) {
      alert(e instanceof Error ? e.message : String(e))
      setRestoring(false)
    }
  }

  return (
    <div className="dialog-overlay" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="dialog changes-dialog" role="dialog" aria-modal="true" aria-label="変更箇所">
        <div className="dialog-head">
          <h2>「{filename}」の変更箇所</h2>
          <button className="icon-btn" onClick={onClose} aria-label="閉じる" title="閉じる">
            <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor">
              <path d="M19 6.41 17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z" />
            </svg>
          </button>
        </div>
        <div className="dialog-body changes-body">
          {!changes && !error && <div className="dialog-loading"><span className="spinner dark" /> 読み込み中…</div>}
          {error && <p className="changes-empty">{error}</p>}
          {changes && !changes.available && <p className="changes-empty">{changes.reason}</p>}
          {changes?.available && (
            <>
              <p className="changes-meta">
                {formatTime(changes.base_time)} 時点との比較 ·{' '}
                <span className="diff-count add">＋{changes.added}</span>{' '}
                <span className="diff-count del">－{changes.removed}</span>
                （文章・セルの値の変更のみ。色や書式の変更は表示されません）
              </p>
              <div className="diff-lines">
                {changes.lines.map((line, i) => (
                  <div key={i} className={`diff-line ${line.op}`}>
                    <span className="diff-sign">{line.op === 'add' ? '＋' : line.op === 'del' ? '－' : ' '}</span>
                    {line.text || ' '}
                  </div>
                ))}
                {changes.truncated && <div className="diff-line skip">…変更が多いため一部のみ表示しています…</div>}
              </div>
            </>
          )}
        </div>
        <div className="dialog-actions">
          <button className="text-btn" onClick={onClose}>閉じる</button>
          {changes?.available && (
            <button className="primary-btn" onClick={handleRestore} disabled={restoring}>
              {restoring ? '戻しています…' : 'この変更前に戻す'}
            </button>
          )}
        </div>
      </div>
    </div>
  )
}
