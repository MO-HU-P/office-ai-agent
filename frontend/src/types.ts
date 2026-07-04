export interface FileInfo {
  name: string
  size: number
  mtime: string
  type: 'word' | 'excel' | 'powerpoint' | 'csv' | 'other'
}

export interface ToolCallPart {
  kind: 'tool'
  name: string
  args: string
  status: 'running' | 'ok' | 'error'
  result?: string
}

export interface TextPart {
  kind: 'text'
  content: string
}

export type MessagePart = TextPart | ToolCallPart

export interface ChatMessage {
  role: 'user' | 'assistant'
  parts: MessagePart[]
  error?: string
}

export interface ExcelCellStyle {
  b?: number
  i?: number
  fs?: number
  fc?: string
  bg?: string
  ha?: 'left' | 'center' | 'right'
}

export interface ExcelCell {
  v?: string
  f?: string
  s?: ExcelCellStyle
}

export interface ExcelSheet {
  name: string
  rows: ExcelCell[][]
  merges: { r: number; c: number; rs: number; cs: number }[]
  colWidths: number[]
  truncated: boolean
}

export type PreviewData =
  | { type: 'excel'; sheets: ExcelSheet[] }
  | { type: 'pptx'; slides: string[] }
  | { type: 'docx' }
  | { type: 'csv'; content: string }
  | { type: 'unsupported' }

export interface HealthInfo {
  mode: 'local' | 'cloud'
  ollama: boolean
  key_missing: boolean
  model: string
  model_ready: boolean
}

export type LLMMode = 'local' | 'cloud'

export interface SettingsInfo {
  mode: LLMMode
  model: string
  model_local: string
  model_cloud: string
  reasoning: string
  cloud_key_configured: boolean
}

export interface ModelInfo {
  name: string
  size?: number | null
}

export interface ModelListResult {
  mode: LLMMode
  models: ModelInfo[]
  unavailable?: string
}

export interface PullProgress {
  status?: string
  total?: number
  completed?: number
  error?: string
}
