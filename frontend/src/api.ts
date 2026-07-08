import type {
  ChangesResult,
  FileInfo,
  HealthInfo,
  ModelListResult,
  ModelSource,
  PreviewData,
  PullProgress,
  SettingsInfo,
} from './types'

export async function fetchFiles(): Promise<FileInfo[]> {
  const res = await fetch('/api/files')
  if (!res.ok) throw new Error('ファイル一覧の取得に失敗しました')
  return res.json()
}

export async function fetchPreview(name: string): Promise<PreviewData> {
  const res = await fetch(`/api/files/${encodeURIComponent(name)}/preview`)
  if (!res.ok) {
    const body = await res.json().catch(() => null)
    throw new Error(body?.detail ?? `プレビューの取得に失敗しました (${res.status})`)
  }
  return res.json()
}

export async function fetchRawBlob(name: string): Promise<Blob> {
  const res = await fetch(`/api/files/${encodeURIComponent(name)}/raw`)
  if (!res.ok) throw new Error('ファイルの取得に失敗しました')
  return res.blob()
}

export async function uploadFile(file: File): Promise<void> {
  const form = new FormData()
  form.append('file', file)
  const res = await fetch('/api/files/upload', { method: 'POST', body: form })
  if (!res.ok) throw new Error('アップロードに失敗しました')
}

export async function deleteFile(name: string): Promise<void> {
  const res = await fetch(`/api/files/${encodeURIComponent(name)}`, { method: 'DELETE' })
  if (!res.ok) throw new Error('削除に失敗しました')
}

/** 「最後の変更でどこが変わったか」の差分を取得する。 */
export async function fetchChanges(name: string): Promise<ChangesResult> {
  const res = await fetch(`/api/files/${encodeURIComponent(name)}/changes`)
  if (!res.ok) {
    const body = await res.json().catch(() => null)
    throw new Error(body?.detail ?? `変更箇所の取得に失敗しました (${res.status})`)
  }
  return res.json()
}

/** ファイルを自動バックアップの状態に巻き戻す(version省略時は最後の変更前)。 */
export async function restoreFile(name: string, version?: string): Promise<void> {
  const res = await fetch(`/api/files/${encodeURIComponent(name)}/restore`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(version ? { version } : {}),
  })
  if (!res.ok) {
    const body = await res.json().catch(() => null)
    throw new Error(body?.detail ?? '巻き戻しに失敗しました')
  }
}

export async function fetchHealth(): Promise<HealthInfo> {
  const res = await fetch('/api/health')
  if (!res.ok) throw new Error('health check failed')
  return res.json()
}

export function downloadUrl(name: string): string {
  return `/api/files/${encodeURIComponent(name)}/raw`
}

// ---------- 設定・モデル管理 ----------

export async function fetchSettings(): Promise<SettingsInfo> {
  const res = await fetch('/api/settings')
  if (!res.ok) throw new Error('設定の取得に失敗しました')
  return res.json()
}

export async function updateSettings(patch: Partial<SettingsInfo>): Promise<SettingsInfo> {
  const res = await fetch('/api/settings', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(patch),
  })
  if (!res.ok) {
    const body = await res.json().catch(() => null)
    throw new Error(body?.detail ?? '設定の保存に失敗しました')
  }
  return res.json()
}

export async function fetchModels(source: ModelSource): Promise<ModelListResult> {
  const res = await fetch(`/api/models?source=${source}`)
  if (!res.ok) throw new Error('モデル一覧の取得に失敗しました')
  return res.json()
}

export async function deleteModel(name: string): Promise<void> {
  const res = await fetch(`/api/models/${encodeURIComponent(name)}`, { method: 'DELETE' })
  if (!res.ok) {
    const body = await res.json().catch(() => null)
    throw new Error(body?.detail ?? 'モデルの削除に失敗しました')
  }
}

/** モデルをダウンロードし、NDJSONの進捗を逐次コールバックする。エラー行は例外にする。 */
export async function pullModel(
  name: string,
  onProgress: (p: PullProgress) => void,
  signal?: AbortSignal,
): Promise<void> {
  const res = await fetch('/api/models/pull', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name }),
    signal,
  })
  if (!res.ok || !res.body) {
    const body = await res.json().catch(() => null)
    throw new Error(body?.detail ?? 'ダウンロードを開始できませんでした')
  }
  const reader = res.body.getReader()
  const decoder = new TextDecoder()
  let buf = ''
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    buf += decoder.decode(value, { stream: true })
    const lines = buf.split('\n')
    buf = lines.pop() ?? ''
    for (const line of lines) {
      if (!line.trim()) continue
      let progress: PullProgress
      try {
        progress = JSON.parse(line)
      } catch {
        continue
      }
      if (progress.error) throw new Error(progress.error)
      onProgress(progress)
    }
  }
}
