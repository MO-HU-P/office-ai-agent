import { useEffect, useMemo, useRef, useState } from 'react'
import type { ExcelCell, ExcelSheet } from '../types'

function colLabel(i: number): string {
  let label = ''
  i++
  while (i > 0) {
    const rem = (i - 1) % 26
    label = String.fromCharCode(65 + rem) + label
    i = Math.floor((i - 1) / 26)
  }
  return label
}

function cellStyle(cell: ExcelCell): React.CSSProperties {
  const s = cell.s
  if (!s) return {}
  return {
    fontWeight: s.b ? 600 : undefined,
    fontStyle: s.i ? 'italic' : undefined,
    fontSize: s.fs ? `${s.fs}pt` : undefined,
    color: s.fc,
    backgroundColor: s.bg,
    textAlign: s.ha,
  }
}

interface Sel {
  r1: number
  c1: number
  r2: number
  c2: number
}

function rangeLabel(sheetName: string, sel: Sel): string {
  const rMin = Math.min(sel.r1, sel.r2)
  const rMax = Math.max(sel.r1, sel.r2)
  const cMin = Math.min(sel.c1, sel.c2)
  const cMax = Math.max(sel.c1, sel.c2)
  const start = `${colLabel(cMin)}${rMin + 1}`
  const end = `${colLabel(cMax)}${rMax + 1}`
  const range = start === end ? `セル ${start}` : `セル範囲 ${start}:${end}`
  return `シート「${sheetName}」の${range}`
}

interface SheetGridProps {
  sheet: ExcelSheet
  /** セルをクリック/ドラッグで選択したとき(対象箇所として親へ通知)。nullで解除 */
  onTarget: (label: string | null) => void
}

function SheetGrid({ sheet, onTarget }: SheetGridProps) {
  const { skip, spans } = useMemo(() => {
    const skip = new Set<string>()
    const spans = new Map<string, { rs: number; cs: number }>()
    for (const m of sheet.merges) {
      spans.set(`${m.r}:${m.c}`, { rs: m.rs, cs: m.cs })
      for (let r = m.r; r < m.r + m.rs; r++)
        for (let c = m.c; c < m.c + m.cs; c++)
          if (r !== m.r || c !== m.c) skip.add(`${r}:${c}`)
    }
    return { skip, spans }
  }, [sheet])

  // マウスでのセル選択(クリック=1セル、ドラッグ=範囲)。マウスを離した時点で親へ通知する
  const [sel, setSel] = useState<Sel | null>(null)
  const selRef = useRef<Sel | null>(null)
  selRef.current = sel
  const draggingRef = useRef(false)

  useEffect(() => {
    const onUp = () => {
      if (!draggingRef.current) return
      draggingRef.current = false
      if (selRef.current) onTarget(rangeLabel(sheet.name, selRef.current))
    }
    window.addEventListener('mouseup', onUp)
    return () => window.removeEventListener('mouseup', onUp)
  }, [sheet.name, onTarget])

  const startSel = (r: number, c: number, e: React.MouseEvent) => {
    if (e.button !== 0) return
    e.preventDefault() // ドラッグ中の文字選択を防ぐ
    draggingRef.current = true
    setSel({ r1: r, c1: c, r2: r, c2: c })
  }

  const extendSel = (r: number, c: number) => {
    if (!draggingRef.current) return
    setSel((s) => (s ? { ...s, r2: r, c2: c } : s))
  }

  const isSel = (r: number, c: number) =>
    !!sel &&
    r >= Math.min(sel.r1, sel.r2) && r <= Math.max(sel.r1, sel.r2) &&
    c >= Math.min(sel.c1, sel.c2) && c <= Math.max(sel.c1, sel.c2)

  const nCols = sheet.rows[0]?.length ?? 0

  return (
    <div className="excel-grid-wrap">
      <table className="excel-grid">
        <thead>
          <tr>
            <th className="corner" />
            {Array.from({ length: nCols }, (_, c) => (
              <th key={c} style={{ minWidth: sheet.colWidths[c] ?? 68 }}>{colLabel(c)}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {sheet.rows.map((row, r) => (
            <tr key={r}>
              <th className="rowhead">{r + 1}</th>
              {row.map((cell, c) => {
                const key = `${r}:${c}`
                if (skip.has(key)) return null
                const span = spans.get(key)
                return (
                  <td
                    key={c}
                    rowSpan={span?.rs}
                    colSpan={span?.cs}
                    className={isSel(r, c) ? 'sel' : undefined}
                    style={cellStyle(cell)}
                    title={cell.f ?? 'クリック/ドラッグでこのセルを質問の対象にできます'}
                    onMouseDown={(e) => startSel(r, c, e)}
                    onMouseEnter={() => extendSel(r, c)}
                  >
                    {cell.v ?? ''}
                  </td>
                )
              })}
            </tr>
          ))}
        </tbody>
      </table>
      {sheet.truncated && <div className="truncated-note">※ 表示は300行×60列までに制限されています</div>}
    </div>
  )
}

interface ExcelPreviewProps {
  sheets: ExcelSheet[]
  /** 対象箇所の選択が解除されたときに増える(選択ハイライトのリセット用) */
  targetEpoch: number
  onTarget: (label: string | null) => void
}

export function ExcelPreview({ sheets, targetEpoch, onTarget }: ExcelPreviewProps) {
  const [active, setActive] = useState(0)
  const sheet = sheets[Math.min(active, sheets.length - 1)]
  if (!sheet) return <div className="preview-message">シートがありません</div>
  return (
    <div className="excel-preview">
      {/* keyにtargetEpochを含め、選択解除時はグリッドを作り直してハイライトを消す */}
      <SheetGrid key={`${sheet.name}:${targetEpoch}`} sheet={sheet} onTarget={onTarget} />
      <div className="sheet-tabs">
        {sheets.map((s, i) => (
          <button
            key={s.name}
            className={i === active ? 'active' : ''}
            onClick={() => {
              setActive(i)
              if (i !== active) onTarget(null) // 別シートに移ったら選択は無効
            }}
          >
            {s.name}
          </button>
        ))}
      </div>
    </div>
  )
}
