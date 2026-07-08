import { useEffect, useRef, useState } from 'react'
import { deleteFile, downloadUrl, fetchPreview, uploadFile } from '../api'
import type { FileInfo, PreviewData } from '../types'
import { ChangesModal } from './ChangesModal'
import { ExcelPreview } from './ExcelPreview'
import { PptPreview } from './PptPreview'
import { WordPreview } from './WordPreview'

const TYPE_ICONS: Record<FileInfo['type'], { label: string; color: string }> = {
  word: { label: 'W', color: '#185abd' },
  excel: { label: 'X', color: '#107c41' },
  powerpoint: { label: 'P', color: '#c43e1c' },
  csv: { label: 'C', color: '#5f6368' },
  other: { label: 'F', color: '#5f6368' },
}

interface PreviewPaneProps {
  files: FileInfo[]
  activeFile: string | null
  refreshKey: number
  onSelect: (name: string | null) => void
  onFilesChanged: () => void
  /** 巻き戻し・上書きアップロードでファイルの中身が変わったとき(一覧とプレビューの再読み込み用) */
  onFileChanged: (name: string) => void
  /** 「対象箇所」の選択解除の合図(増えるたびにプレビュー側のハイライトを消す) */
  targetEpoch: number
  /** プレビュー上でマウス選択した対象箇所の通知(nullで解除) */
  onTarget: (label: string | null) => void
}

export function PreviewPane({ files, activeFile, refreshKey, onSelect, onFilesChanged, onFileChanged, targetEpoch, onTarget }: PreviewPaneProps) {
  const [preview, setPreview] = useState<PreviewData | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [changesOpen, setChangesOpen] = useState(false)
  const uploadRef = useRef<HTMLInputElement>(null)

  const active = files.find((f) => f.name === activeFile) ?? null

  useEffect(() => {
    if (!activeFile) {
      setPreview(null)
      return
    }
    let cancelled = false
    setLoading(true)
    setError(null)
    fetchPreview(activeFile)
      .then((data) => {
        if (!cancelled) setPreview(data)
      })
      .catch((e) => {
        if (!cancelled) setError(String(e.message ?? e))
      })
      .finally(() => {
        if (!cancelled) setLoading(false)
      })
    return () => {
      cancelled = true
    }
  }, [activeFile, refreshKey])

  const handleUpload = async (fileList: FileList | null) => {
    if (!fileList?.length) return
    for (const f of Array.from(fileList)) {
      // 同名ファイルへの上書きアップロードでもプレビューが新しい中身に更新されるよう、
      // アップロード成功ごとに「中身が変わった」通知を出す(ファイルも自動で開かれる)
      await uploadFile(f).then(() => onFileChanged(f.name)).catch((e) => alert(e.message))
    }
    onFilesChanged()
  }

  const handleDelete = async () => {
    if (!activeFile) return
    if (!confirm(`「${activeFile}」を削除しますか？`)) return
    await deleteFile(activeFile).catch((e) => alert(e.message))
    onSelect(null)
    onFilesChanged()
  }

  return (
    <div className="preview-pane">
      <div className="preview-toolbar">
        <div className="file-tabs">
          {files.map((f) => {
            const icon = TYPE_ICONS[f.type]
            return (
              <button
                key={f.name}
                className={`file-tab ${f.name === activeFile ? 'active' : ''}`}
                onClick={() => onSelect(f.name)}
                title={f.name}
              >
                <span className="file-icon" style={{ backgroundColor: icon.color }}>{icon.label}</span>
                <span className="file-tab-name">{f.name}</span>
              </button>
            )
          })}
          {files.length === 0 && <span className="no-files">ファイルはまだありません</span>}
        </div>
        <div className="preview-actions">
          <input
            ref={uploadRef}
            type="file"
            multiple
            hidden
            accept=".docx,.xlsx,.pptx,.csv"
            onChange={(e) => {
              handleUpload(e.target.files)
              e.target.value = ''
            }}
          />
          <button title="ファイルをアップロード" onClick={() => uploadRef.current?.click()}>
            <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M9 16h6v-6h4l-7-7-7 7h4zm-4 2h14v2H5z"/></svg>
          </button>
          {active && (
            <>
              <button title="変更箇所を確認・元に戻す" onClick={() => setChangesOpen(true)}>
                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M13 3a9 9 0 0 0-9 9H1l3.89 3.89.07.14L9 12H6c0-3.87 3.13-7 7-7s7 3.13 7 7-3.13 7-7 7c-1.93 0-3.68-.79-4.94-2.06l-1.42 1.42A8.954 8.954 0 0 0 13 21a9 9 0 0 0 0-18zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8H12z"/></svg>
              </button>
              <a title="ダウンロード" href={downloadUrl(active.name)} download>
                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M19 9h-4V3H9v6H5l7 7zM5 18v2h14v-2z"/></svg>
              </a>
              <button title="削除" onClick={handleDelete}>
                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6zM19 4h-3.5l-1-1h-5l-1 1H5v2h14z"/></svg>
              </button>
            </>
          )}
        </div>
      </div>
      {changesOpen && activeFile && (
        <ChangesModal
          filename={activeFile}
          onClose={() => setChangesOpen(false)}
          onRestored={onFileChanged}
        />
      )}
      <div className="preview-body">
        {!activeFile && (
          <div className="preview-empty">
            <svg viewBox="0 0 24 24" width="56" height="56" fill="#dadce0">
              <path d="M14 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V8zm4 18H6V4h7v5h5z"/>
            </svg>
            <p>左のチャットでAIに指示すると、<br />作成・編集されたファイルがここに表示されます</p>
          </div>
        )}
        {activeFile && loading && !preview && <div className="preview-message"><span className="spinner dark" /> 読み込み中…</div>}
        {activeFile && error && <div className="preview-message error">{error}</div>}
        {activeFile && !error && preview?.type === 'excel' && (
          <ExcelPreview sheets={preview.sheets} targetEpoch={targetEpoch} onTarget={onTarget} />
        )}
        {activeFile && !error && preview?.type === 'pptx' && (
          <PptPreview slides={preview.slides} targetEpoch={targetEpoch} onTarget={onTarget} />
        )}
        {activeFile && !error && preview?.type === 'docx' && (
          <WordPreview filename={activeFile} refreshKey={refreshKey} onTarget={onTarget} />
        )}
        {activeFile && !error && preview?.type === 'csv' && <pre className="csv-preview">{preview.content}</pre>}
        {activeFile && !error && preview?.type === 'unsupported' && (
          <div className="preview-message">このファイル形式のプレビューには対応していません</div>
        )}
      </div>
    </div>
  )
}
