import { useMemo, useState } from 'react'
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

function SheetGrid({ sheet }: { sheet: ExcelSheet }) {
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
                    style={cellStyle(cell)}
                    title={cell.f ?? undefined}
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

export function ExcelPreview({ sheets }: { sheets: ExcelSheet[] }) {
  const [active, setActive] = useState(0)
  const sheet = sheets[Math.min(active, sheets.length - 1)]
  if (!sheet) return <div className="preview-message">シートがありません</div>
  return (
    <div className="excel-preview">
      <SheetGrid sheet={sheet} />
      <div className="sheet-tabs">
        {sheets.map((s, i) => (
          <button key={s.name} className={i === active ? 'active' : ''} onClick={() => setActive(i)}>
            {s.name}
          </button>
        ))}
      </div>
    </div>
  )
}
