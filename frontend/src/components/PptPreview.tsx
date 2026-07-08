import { useEffect, useState } from 'react'

interface PptPreviewProps {
  slides: string[]
  /** 対象箇所の選択が解除されたときに増える(選択ハイライトのリセット用) */
  targetEpoch: number
  /** スライドをクリックで対象に選んだとき(もう一度クリックで解除) */
  onTarget: (label: string | null) => void
}

export function PptPreview({ slides, targetEpoch, onTarget }: PptPreviewProps) {
  const [current, setCurrent] = useState(0)
  // 質問の対象として選択中のスライド(番号は0始まり、表示は1始まり)
  const [selected, setSelected] = useState<number | null>(null)

  useEffect(() => {
    setCurrent((c) => Math.min(c, slides.length - 1))
  }, [slides])

  // 親側で選択が解除されたらハイライトも消す
  useEffect(() => {
    setSelected(null)
  }, [targetEpoch])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'ArrowRight') setCurrent((c) => Math.min(c + 1, slides.length - 1))
      if (e.key === 'ArrowLeft') setCurrent((c) => Math.max(c - 1, 0))
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [slides.length])

  if (slides.length === 0) return <div className="preview-message">スライドがありません</div>

  const toggleTarget = () => {
    if (selected === current) {
      setSelected(null)
      onTarget(null)
    } else {
      setSelected(current)
      onTarget(`スライド${current + 1}`)
    }
  }

  return (
    <div className="ppt-preview">
      <div className="ppt-main">
        <img
          src={slides[current]}
          alt={`スライド ${current + 1}`}
          className={selected === current ? 'target-selected' : undefined}
          title="クリックすると、このスライドを質問の対象にできます(もう一度クリックで解除)"
          onClick={toggleTarget}
        />
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
            <span>{i + 1}{i === selected ? ' 📍' : ''}</span>
          </button>
        ))}
      </div>
    </div>
  )
}
