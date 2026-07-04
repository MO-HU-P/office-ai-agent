import { useCallback, useEffect, useRef, useState } from 'react'

/** ブラウザ内蔵の音声認識(Web Speech API)でテキスト入力を補助するフック。
 *
 * - Chrome/Edge等でのみ利用可能。非対応ブラウザでは supported=false になり、
 *   呼び出し側はマイクボタン自体を表示しない。
 * - 認識はブラウザベンダーの音声認識サービスで処理される(完全ローカルではない)。
 * - 認識テキストはコールバックでのみ渡し、consoleやログには一切出さない。
 */

// Web Speech APIの型はTypeScriptの標準libに無いため、必要最小限を宣言する
interface SpeechResultEvent {
  results: ArrayLike<{ isFinal: boolean; 0: { transcript: string } }>
}

interface SpeechRecognitionLike {
  lang: string
  continuous: boolean
  interimResults: boolean
  onresult: ((e: SpeechResultEvent) => void) | null
  onerror: ((e: { error: string }) => void) | null
  onend: (() => void) | null
  start(): void
  stop(): void
  abort(): void
}

const SpeechRecognitionCtor: (new () => SpeechRecognitionLike) | undefined =
  (window as { SpeechRecognition?: new () => SpeechRecognitionLike }).SpeechRecognition ??
  (window as { webkitSpeechRecognition?: new () => SpeechRecognitionLike }).webkitSpeechRecognition

const ERROR_MESSAGES: Record<string, string> = {
  'not-allowed': 'マイクの使用が許可されていません。ブラウザのアドレスバーからマイクを許可してください。',
  'service-not-allowed': 'このブラウザでは音声認識サービスを利用できません。',
  'audio-capture': 'マイクが見つかりません。接続を確認してください。',
  network: '音声認識サービスに接続できませんでした。ネットワークを確認してください。',
}

export function useSpeechInput() {
  const [listening, setListening] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const recognitionRef = useRef<SpeechRecognitionLike | null>(null)

  // アンマウント時は即座に打ち切る
  useEffect(() => () => recognitionRef.current?.abort(), [])

  /** 認識を開始する。認識が進むたびに onUpdate(セッション全体の認識テキスト) を呼ぶ。 */
  const start = useCallback((onUpdate: (text: string) => void) => {
    if (!SpeechRecognitionCtor || recognitionRef.current) return
    setError(null)
    const rec = new SpeechRecognitionCtor()
    rec.lang = 'ja-JP'
    rec.continuous = true
    rec.interimResults = true
    rec.onresult = (e) => {
      // resultsはセッション全体を保持しているので、毎回すべて結合し直す(取りこぼし防止)
      let text = ''
      for (let i = 0; i < e.results.length; i++) {
        text += e.results[i][0].transcript
      }
      onUpdate(text)
    }
    rec.onerror = (e) => {
      // no-speech(無音で終了)やaborted(手動停止)はエラー表示しない
      const message = ERROR_MESSAGES[e.error]
      if (message) setError(message)
    }
    rec.onend = () => {
      recognitionRef.current = null
      setListening(false)
    }
    recognitionRef.current = rec
    setListening(true)
    rec.start()
  }, [])

  /** 認識を停止する(確定済みテキストは残る)。 */
  const stop = useCallback(() => {
    recognitionRef.current?.stop()
  }, [])

  return { supported: !!SpeechRecognitionCtor, listening, error, start, stop }
}
