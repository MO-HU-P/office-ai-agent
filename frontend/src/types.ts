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

// 「最後の変更でどこが変わったか」(GET /api/files/{name}/changes)
export interface ChangeLine {
  op: 'add' | 'del' | 'ctx' | 'skip'
  text: string
}

export type ChangesResult =
  | { available: false; reason: string }
  | {
      available: true
      filename: string
      base_version: string
      base_time: string
      base_label: string
      added: number
      removed: number
      truncated: boolean
      lines: ChangeLine[]
    }

export type LLMProvider = 'ollama' | 'openai' | 'gemini'
export type LLMMode = 'local' | 'cloud'
// モデル一覧の取得元。Ollamaは local/cloud、外部プロバイダーはその名前
export type ModelSource = 'local' | 'cloud' | 'openai' | 'gemini'
// 外部プロバイダー(=モデルを自由入力でき、候補をsettings.jsonに保存できるもの)
export type ExternalProvider = 'openai' | 'gemini'

export interface HealthInfo {
  provider: LLMProvider
  mode: LLMMode
  backend_ok: boolean
  key_missing: boolean
  model: string
  model_ready: boolean
}

export interface SettingsInfo {
  provider: LLMProvider
  mode: LLMMode
  model: string
  model_local: string
  model_cloud: string
  model_openai: string
  model_gemini: string
  // 設定画面で自由入力し保存した外部プロバイダーのモデル候補(settings.json)。プリセットとは別で削除可能
  openai_custom_models: string[]
  gemini_custom_models: string[]
  reasoning: string
  cloud_key_configured: boolean
  openai_key_configured: boolean
  gemini_key_configured: boolean
}

export interface ModelInfo {
  name: string
  size?: number | null
  vision?: boolean
}

export interface ModelListResult {
  source: ModelSource
  models: ModelInfo[]
  unavailable?: string
}

export interface PullProgress {
  status?: string
  total?: number
  completed?: number
  error?: string
}
