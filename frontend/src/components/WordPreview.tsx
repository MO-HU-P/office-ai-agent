import { useEffect, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import { fetchRawBlob } from '../api'

export function WordPreview({ filename, refreshKey }: { filename: string; refreshKey: number }) {
  const outerRef = useRef<HTMLDivElement>(null)
  const containerRef = useRef<HTMLDivElement>(null)
  const [error, setError] = useState<string | null>(null)

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
    <div className="word-preview-outer" ref={outerRef}>
      <div className="word-preview" ref={containerRef} />
    </div>
  )
}
