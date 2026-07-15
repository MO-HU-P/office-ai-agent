import { useMemo } from 'react'
import type { ExcelChart as ChartSpec } from '../types'

/** 系列に色が指定されていないときの既定色。検証済みの並び順なので、順番は変えないこと
 *  (隣り合う色が色覚特性のある人にも見分けられるように選ばれている) */
const SERIES_COLORS = [
  '#2a78d6', '#008300', '#e87ba4', '#eda100',
  '#1baf7a', '#eb6834', '#4a3aa7', '#e34948',
]

const TEXT = '#202124'
const TEXT_2 = '#5f6368'
const GRID = '#e0e0e0'
const SURFACE = '#ffffff'

const seriesColor = (i: number, explicit: string | null) => explicit ?? SERIES_COLORS[i % SERIES_COLORS.length]

/** 目盛りをキリのいい数値に丸める */
function niceTicks(min: number, max: number, count = 5): number[] {
  if (min === max) {
    if (min === 0) return [0, 1]
    min = Math.min(0, min)
    max = Math.max(0, max)
  }
  const span = max - min || 1
  const raw = span / count
  const mag = Math.pow(10, Math.floor(Math.log10(raw)))
  const step = [1, 2, 2.5, 5, 10].find((m) => raw <= m * mag)! * mag
  const start = Math.floor(min / step) * step
  const ticks: number[] = []
  for (let v = start; v <= max + step / 2; v += step) ticks.push(Math.round(v * 1e6) / 1e6)
  return ticks
}

const fmt = (v: number) => (Math.abs(v) >= 1000 ? v.toLocaleString('ja-JP') : String(Math.round(v * 100) / 100))

interface Props {
  chart: ChartSpec
}

/** Excelのネイティブグラフをプレビュー用にSVGで再現する。
 *  値はセルから解決済みなので、セルを直せばこのグラフも変わる(Excelと同じ挙動)。 */
export function ExcelChart({ chart }: Props) {
  const { kind, series = [], categories = [], title, xTitle, yTitle, stacked, w, h } = chart

  const legendH = series.length >= 2 ? 22 : 0
  const titleH = title ? 22 : 0

  const layout = useMemo(() => {
    const values = series.flatMap((s) => s.values.filter((v): v is number => v !== null))
    if (!values.length) return null
    // エラーバーのひげが軸からはみ出さないように、±の分まで含めて範囲を取る
    const withErrors = series.flatMap((s) =>
      s.values.flatMap((v, i) => {
        if (v === null) return []
        const e = s.errors?.[i]
        return e == null ? [v] : [v + e, v - e]
      }))
    let min = Math.min(0, ...withErrors)
    let max = Math.max(0, ...withErrors)
    if (stacked) {
      // 積み上げは各項目の合計が最大になる
      const totals = categories.map((_, i) =>
        series.reduce((sum, s) => sum + Math.max(0, s.values[i] ?? 0), 0))
      max = Math.max(0, ...totals)
      min = 0
    }
    return { min, max }
  }, [series, categories, stacked])

  if (!layout || !series.length) return null

  if (kind === 'pie') {
    return <PieChart chart={chart} titleH={titleH} />
  }

  const horizontal = kind === 'bar_horizontal'
  const ticks = niceTicks(layout.min, layout.max)
  const vMin = Math.min(layout.min, ticks[0])
  const vMax = Math.max(layout.max, ticks[ticks.length - 1])

  // 目盛りラベルの幅を実測できないので、桁数からおおよそで確保する
  const tickLabelW = Math.max(...ticks.map((t) => fmt(t).length)) * 7 + 8
  const padL = (horizontal ? Math.max(...categories.map((c) => c.length)) * 12 + 8 : tickLabelW) + (yTitle ? 16 : 0)
  const padR = 12
  const padT = titleH + 8
  const padB = 26 + (xTitle ? 16 : 0) + legendH
  const plotW = Math.max(w - padL - padR, 10)
  const plotH = Math.max(h - padT - padB, 10)

  // 値 → ピクセル位置
  const scale = (v: number) => (v - vMin) / (vMax - vMin || 1)
  const zero = scale(0)

  const bandCount = categories.length || 1
  const band = (horizontal ? plotH : plotW) / bandCount
  // 棒は太くしすぎない(24px上限)。帯の余りは余白として残す
  const groupW = Math.min(band * 0.7, 24 * (stacked ? 1 : series.length))
  const barW = stacked ? groupW : groupW / series.length

  return (
    <svg width={w} height={h} className="excel-chart-svg" role="img" aria-label={title || 'グラフ'}>
      <rect width={w} height={h} fill={SURFACE} />
      {title && (
        <text x={w / 2} y={16} textAnchor="middle" fontSize="13" fontWeight="500" fill={TEXT}>{title}</text>
      )}

      {/* 目盛り線と目盛りラベル */}
      {ticks.map((t) => {
        const p = scale(t)
        return horizontal ? (
          <g key={t}>
            <line x1={padL + p * plotW} y1={padT} x2={padL + p * plotW} y2={padT + plotH} stroke={GRID} strokeWidth="1" />
            <text x={padL + p * plotW} y={padT + plotH + 14} textAnchor="middle" fontSize="10" fill={TEXT_2}>{fmt(t)}</text>
          </g>
        ) : (
          <g key={t}>
            <line x1={padL} y1={padT + (1 - p) * plotH} x2={padL + plotW} y2={padT + (1 - p) * plotH} stroke={GRID} strokeWidth="1" />
            <text x={padL - 6} y={padT + (1 - p) * plotH + 3.5} textAnchor="end" fontSize="10" fill={TEXT_2}>{fmt(t)}</text>
          </g>
        )
      })}

      {/* 項目名 */}
      {categories.map((cat, i) => {
        const center = (i + 0.5) * band
        return horizontal ? (
          <text key={i} x={padL - 6} y={padT + center + 3.5} textAnchor="end" fontSize="10" fill={TEXT_2}>{cat}</text>
        ) : (
          <text key={i} x={padL + center} y={padT + plotH + 14} textAnchor="middle" fontSize="10" fill={TEXT_2}>{cat}</text>
        )
      })}

      {/* データのマーク */}
      {kind === 'line'
        ? series.map((s, si) => {
            const color = seriesColor(si, s.color)
            const pts = s.values
              .map((v, i) => (v === null ? null : [padL + (i + 0.5) * band, padT + (1 - scale(v)) * plotH] as const))
              .filter((p): p is readonly [number, number] => p !== null)
            return (
              <g key={si}>
                <polyline
                  points={pts.map(([x, y]) => `${x},${y}`).join(' ')}
                  fill="none" stroke={color} strokeWidth="2" strokeLinejoin="round" strokeLinecap="round"
                />
                {pts.map(([x, y], i) => (
                  // 線が重なっても見分けられるよう、点は背景色のリングで縁取る
                  <circle key={i} cx={x} cy={y} r="4" fill={color} stroke={SURFACE} strokeWidth="2" />
                ))}
              </g>
            )
          })
        : categories.map((_, ci) => {
            let posStack = 0
            return (
              <g key={ci}>
                {series.map((s, si) => {
                  const v = s.values[ci]
                  if (v === null || v === undefined) return null
                  const color = seriesColor(si, s.color)
                  const from = stacked ? posStack : 0
                  const to = stacked ? posStack + v : v
                  if (stacked) posStack = to
                  const a = scale(from)
                  const b = scale(to)
                  const lo = Math.min(a, b)
                  const hi = Math.max(a, b)
                  // 積み上げの各段・隣り合う棒は2pxの隙間で分ける(枠線は引かない)
                  const gap = stacked && from !== 0 ? 2 : 0
                  const offset = stacked ? (band - groupW) / 2 : (band - groupW) / 2 + si * barW
                  const thick = stacked ? groupW : Math.max(barW - 2, 1)
                  const err = s.errors?.[ci]
                  if (horizontal) {
                    return (
                      <g key={si}>
                        <rect x={padL + lo * plotW + gap} y={padT + ci * band + offset}
                          width={Math.max((hi - lo) * plotW - gap, 0)} height={thick}
                          fill={color} rx="2" />
                        {err != null && (
                          <ErrorBar horizontal x1={padL + scale(v - err) * plotW} x2={padL + scale(v + err) * plotW}
                            at={padT + ci * band + offset + thick / 2} />
                        )}
                      </g>
                    )
                  }
                  return (
                    <g key={si}>
                      <rect x={padL + ci * band + offset} y={padT + (1 - hi) * plotH}
                        width={thick} height={Math.max((hi - lo) * plotH - gap, 0)}
                        fill={color} rx="2" />
                      {err != null && (
                        <ErrorBar y1={padT + (1 - scale(v - err)) * plotH} y2={padT + (1 - scale(v + err)) * plotH}
                          at={padL + ci * band + offset + thick / 2} />
                      )}
                    </g>
                  )
                })}
              </g>
            )
          })}

      {/* 基準線(0の位置) */}
      {horizontal ? (
        <line x1={padL + zero * plotW} y1={padT} x2={padL + zero * plotW} y2={padT + plotH} stroke={TEXT_2} strokeWidth="1" />
      ) : (
        <line x1={padL} y1={padT + (1 - zero) * plotH} x2={padL + plotW} y2={padT + (1 - zero) * plotH} stroke={TEXT_2} strokeWidth="1" />
      )}

      {/* 軸タイトル */}
      {yTitle && (
        <text transform={`translate(10 ${padT + plotH / 2}) rotate(-90)`} textAnchor="middle" fontSize="10" fill={TEXT_2}>{yTitle}</text>
      )}
      {xTitle && (
        <text x={padL + plotW / 2} y={h - legendH - 4} textAnchor="middle" fontSize="10" fill={TEXT_2}>{xTitle}</text>
      )}

      {/* 凡例(2系列以上のときだけ。色だけに頼らせない) */}
      {legendH > 0 && <Legend series={series} y={h - 6} w={w} />}
    </svg>
  )
}

/** エラーバー(誤差範囲)のひげ。両端に短いキャップを付ける */
function ErrorBar(props: { horizontal?: boolean; x1?: number; x2?: number; y1?: number; y2?: number; at: number }) {
  const { horizontal, x1 = 0, x2 = 0, y1 = 0, y2 = 0, at } = props
  const cap = 3
  return (
    <g stroke={TEXT} strokeWidth="1.2" fill="none">
      {horizontal ? (
        <>
          <line x1={x1} y1={at} x2={x2} y2={at} />
          <line x1={x1} y1={at - cap} x2={x1} y2={at + cap} />
          <line x1={x2} y1={at - cap} x2={x2} y2={at + cap} />
        </>
      ) : (
        <>
          <line x1={at} y1={y1} x2={at} y2={y2} />
          <line x1={at - cap} y1={y1} x2={at + cap} y2={y1} />
          <line x1={at - cap} y1={y2} x2={at + cap} y2={y2} />
        </>
      )}
    </g>
  )
}

function Legend({ series, y, w }: { series: ChartSpec['series'] & {}; y: number; w: number }) {
  // だいたいの文字幅で中央寄せの位置を決める
  const widths = series.map((s) => 14 + s.name.length * 11 + 12)
  const total = widths.reduce((a, b) => a + b, 0)
  let x = Math.max((w - total) / 2, 4)
  return (
    <g>
      {series.map((s, i) => {
        const at = x
        x += widths[i]
        return (
          <g key={i}>
            <rect x={at} y={y - 8} width="8" height="8" rx="2" fill={seriesColor(i, s.color)} />
            <text x={at + 12} y={y} fontSize="10" fill={TEXT_2}>{s.name || `系列${i + 1}`}</text>
          </g>
        )
      })}
    </g>
  )
}

function PieChart({ chart, titleH }: { chart: ChartSpec; titleH: number }) {
  const { series = [], categories = [], pointColors = [], title, w, h } = chart
  // 円グラフは最初の1系列だけを使う(Excelと同じ)
  const values = (series[0]?.values ?? []).map((v) => (v === null ? 0 : Math.max(v, 0)))
  const total = values.reduce((a, b) => a + b, 0)
  if (!total) return null

  const legendW = 110
  const cx = (w - legendW) / 2
  const cy = titleH + (h - titleH) / 2
  const r = Math.max(Math.min(cx, (h - titleH) / 2) - 8, 4)

  let angle = -Math.PI / 2
  const slices = values.map((v, i) => {
    const start = angle
    const sweep = (v / total) * Math.PI * 2
    angle += sweep
    const end = angle
    const large = sweep > Math.PI ? 1 : 0
    const p = (a: number) => `${cx + r * Math.cos(a)},${cy + r * Math.sin(a)}`
    // 1項目だけで全周のときはパスが閉じないので円で描く
    const d = values.filter((x) => x > 0).length === 1
      ? null
      : `M ${cx},${cy} L ${p(start)} A ${r},${r} 0 ${large} 1 ${p(end)} Z`
    // 円は系列ではなく項目ごとに色を変える(系列の色は使わない)
    return { d, color: seriesColor(i, pointColors[i] ?? null), pct: v / total, label: categories[i] ?? `項目${i + 1}` }
  })

  return (
    <svg width={w} height={h} className="excel-chart-svg" role="img" aria-label={title || '円グラフ'}>
      <rect width={w} height={h} fill={SURFACE} />
      {title && <text x={w / 2} y={16} textAnchor="middle" fontSize="13" fontWeight="500" fill={TEXT}>{title}</text>}
      {slices.map((s, i) =>
        s.d === null
          ? <circle key={i} cx={cx} cy={cy} r={r} fill={s.color} />
          // 隣り合う扇形は背景色の隙間で分ける
          : <path key={i} d={s.d} fill={s.color} stroke={SURFACE} strokeWidth="2" />,
      )}
      {/* 円は面積で量を読ませにくいので、凡例に割合を併記する */}
      {slices.map((s, i) => (
        <g key={i}>
          <rect x={w - legendW + 4} y={titleH + 8 + i * 16} width="8" height="8" rx="2" fill={s.color} />
          <text x={w - legendW + 16} y={titleH + 16 + i * 16} fontSize="10" fill={TEXT_2}>
            {s.label} {Math.round(s.pct * 100)}%
          </text>
        </g>
      ))}
    </svg>
  )
}
