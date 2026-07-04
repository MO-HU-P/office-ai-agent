import { useEffect, useRef, useState } from 'react'
import { deleteFile, downloadUrl, fetchPreview, uploadFile } from '../api'
import type { FileInfo, PreviewData } from '../types'
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
}

export function PreviewPane({ files, activeFile, refreshKey, onSelect, onFilesChanged }: PreviewPaneProps) {
  const [preview, setPreview] = useState<PreviewData | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
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
      await uploadFile(f).catch((e) => alert(e.message))
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
        {activeFile && !error && preview?.type === 'excel' && <ExcelPreview sheets={preview.sheets} />}
        {activeFile && !error && preview?.type === 'pptx' && <PptPreview slides={preview.slides} />}
        {activeFile && !error && preview?.type === 'docx' && <WordPreview filename={activeFile} refreshKey={refreshKey} />}
        {activeFile && !error && preview?.type === 'csv' && <pre className="csv-preview">{preview.content}</pre>}
        {activeFile && !error && preview?.type === 'unsupported' && (
          <div className="preview-message">このファイル形式のプレビューには対応していません</div>
        )}
      </div>
    </div>
  )
}
