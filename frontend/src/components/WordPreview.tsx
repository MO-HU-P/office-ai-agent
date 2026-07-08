import { useEffect, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import { fetchRawBlob } from '../api'

interface WordPreviewProps {
  filename: string
  refreshKey: number
  /** 本文の文字列をマウス選択したとき(その文言の箇所を質問の対象にする) */
  onTarget: (label: string | null) => void
}

export function WordPreview({ filename, refreshKey, onTarget }: WordPreviewProps) {
  const outerRef = useRef<HTMLDivElement>(null)
  const containerRef = useRef<HTMLDivElement>(null)
  const [error, setError] = useState<string | null>(null)

  // 本文をマウスで選択したら、その文言を「対象箇所」として親へ通知する
  // (AI側は word_find で該当段落を特定できる)。空選択では何もしない
  const handleMouseUp = () => {
    const text = window.getSelection()?.toString().replace(/\s+/g, ' ').trim()
    if (text && text.length >= 2) {
      const snippet = text.length > 80 ? text.slice(0, 80) : text
      onTarget(`「${snippet}」と書かれている箇所`)
    }
  }

  // ページ(A4固定幅)がペインより広い場合は縮小表示する
  const fitZoom = () => {
    const outer = outerRef.current
    const page = containerRef.current?.querySelector<HTMLElement>('section.docx')
    if (!outer || !page) return
    const available = outer.clientWidth - 48
    const zoom = Math.min(1, available / page.offsetWidth)
    ;(containerRef.current as HTMLElement & { style: CSSStyleDeclaration }).style.zoom = String(zoom)
  }

  useEffect(() => {
    let cancelled = false
    setError(null)
    fetchRawBlob(filename)
      .then((blob) => {
        if (cancelled || !containerRef.current) return
        return renderAsync(blob, containerRef.current, undefined, {
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: true,
          experimental: true,
          renderChanges: true,   // 変更履歴(見え消し)を赤字で表示
          renderComments: true,  // レビューコメント(吹き出し)を表示
        }).then(() => fitZoom())
      })
      .catch((e) => {
        if (!cancelled) setError(String(e))
      })
    return () => {
      cancelled = true
    }
  }, [filename, refreshKey])

  useEffect(() => {
    if (!outerRef.current) return
    const observer = new ResizeObserver(() => fitZoom())
    observer.observe(outerRef.current)
    return () => observer.disconnect()
  }, [])

  if (error) return <div className="preview-message error">Word文書の描画に失敗しました: {error}</div>
  return (
    <div className="word-preview-outer" ref={outerRef} onMouseUp={handleMouseUp}>
      <div className="word-preview" ref={containerRef} />
    </div>
  )
}
