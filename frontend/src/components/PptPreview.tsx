import { useEffect, useState } from 'react'

export function PptPreview({ slides }: { slides: string[] }) {
  const [current, setCurrent] = useState(0)

  useEffect(() => {
    setCurrent((c) => Math.min(c, slides.length - 1))
  }, [slides])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'ArrowRight') setCurrent((c) => Math.min(c + 1, slides.length - 1))
      if (e.key === 'ArrowLeft') setCurrent((c) => Math.max(c - 1, 0))
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [slides.length])

  if (slides.length === 0) return <div className="preview-message">スライドがありません</div>

  return (
    <div className="ppt-preview">
      <div className="ppt-main">
        <img src={slides[current]} alt={`スライド ${current + 1}`} />
        <div className="ppt-nav">
          <button onClick={() => setCurrent((c) => Math.max(c - 1, 0))} disabled={current === 0}>‹</button>
          <span>{current + 1} / {slides.length}</span>
          <button onClick={() => setCurrent((c) => Math.min(c + 1, slides.length - 1))} disabled={current === slides.length - 1}>›</button>
        </div>
      </div>
      <div className="ppt-thumbs">
        {slides.map((src, i) => (
          <button key={src} className={i === current ? 'active' : ''} onClick={() => setCurrent(i)}>
            <img src={src} alt={`サムネイル ${i + 1}`} loading="lazy" />
            <span>{i + 1}</span>
          </button>
        ))}
      </div>
    </div>
  )
}
