import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { ExcelCell, ExcelChart as ChartSpec, ExcelSheet } from '../types'
import { ExcelChart } from './ExcelChart'

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

/** グリッドに重ねて表示するもの(貼られた画像 / ネイティブグラフ)と、その実測位置。
 *  ChartSpec自身が種別の`kind`を持つので、判別子は`overlay`という別名にしている */
type Overlay =
  | ({ overlay: 'image' } & ExcelSheet['images'][number])
  | ({ overlay: 'chart' } & ChartSpec)

interface Placed {
  item: Overlay
  left: number
  top: number
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

  // シートに貼られた画像・ネイティブグラフを、アンカーのセルの位置に重ねる。
  // セルの実座標はレイアウト後にしか分からないので、DOMを測ってから配置する
  const wrapRef = useRef<HTMLDivElement>(null)
  const tableRef = useRef<HTMLTableElement>(null)
  const cellRefs = useRef(new Map<string, HTMLTableCellElement>())
  const [placed, setPlaced] = useState<Placed[]>([])

  const overlays = useMemo(
    () => [
      ...sheet.images.map((im) => ({ ...im, overlay: 'image' as const })),
      ...sheet.charts.map((ch) => ({ ...ch, overlay: 'chart' as const })),
    ],
    [sheet],
  )

  useLayoutEffect(() => {
    if (!overlays.length) {
      setPlaced([])
      return
    }
    const measure = () => {
      const wrap = wrapRef.current
      if (!wrap) return
      // td.offsetLeft はtable基準になるため、ラッパー基準の座標を実測して求める
      const base = wrap.getBoundingClientRect()
      const out: Placed[] = []
      for (const ov of overlays) {
        // 表示範囲(300行×60列)やデータ範囲の外に置かれたものは、いちばん近い端のセルに寄せる
        const r = Math.min(ov.r, sheet.rows.length - 1)
        const c = Math.min(ov.c, nCols - 1)
        const td = cellRefs.current.get(`${r}:${c}`)
        if (!td) continue // 結合セルの内側などで対応するセルが無いときは諦める
        const rect = td.getBoundingClientRect()
        out.push({
          item: ov,
          left: rect.left - base.left + wrap.scrollLeft - wrap.clientLeft + ov.dx,
          top: rect.top - base.top + wrap.scrollTop - wrap.clientTop + ov.dy,
        })
      }
      setPlaced(out)
    }
    measure()
    // 列幅の確定やフォント読み込みでレイアウトがずれるので、表のサイズ変化に追従する
    const table = tableRef.current
    if (!table) return
    const ro = new ResizeObserver(measure)
    ro.observe(table)
    return () => ro.disconnect()
  }, [overlays, sheet.rows.length, nCols])

  return (
    <div className="excel-grid-wrap" ref={wrapRef}>
      <table className="excel-grid" ref={tableRef}>
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
                    ref={(el) => {
                      if (el) cellRefs.current.set(key, el)
                      else cellRefs.current.delete(key)
                    }}
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
      {placed.map((p, i) => (
        <div key={i} className="excel-overlay" style={{ left: p.left, top: p.top }}>
          {p.item.overlay === 'image' ? (
            <img src={p.item.url} alt="シートに貼られた画像" style={{ width: p.item.w, height: p.item.h }} />
          ) : p.item.kind === null ? (
            // 描画に未対応の種類(散布図など)。存在だけは伝える
            <div className="excel-chart-unsupported" style={{ width: p.item.w, height: p.item.h }}>
              {p.item.title || 'グラフ'}
              <span>この種類のグラフはプレビューできません。ダウンロードしてExcelで開くと表示されます</span>
            </div>
          ) : (
            <ExcelChart chart={p.item} />
          )}
        </div>
      ))}
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
