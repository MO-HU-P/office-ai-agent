import { useCallback, useEffect, useRef, useState } from 'react'
import type { ChatMessage, ToolCallPart } from '../types'

interface AgentSocketOptions {
  onDocUpdated: (filename: string) => void
  onDocDeleted: (filename: string) => void
}

export function useAgentSocket({ onDocUpdated, onDocDeleted }: AgentSocketOptions) {
  const [messages, setMessages] = useState<ChatMessage[]>([])
  const [busy, setBusy] = useState(false)
  const [connected, setConnected] = useState(false)
  const wsRef = useRef<WebSocket | null>(null)
  const onDocUpdatedRef = useRef(onDocUpdated)
  onDocUpdatedRef.current = onDocUpdated
  const onDocDeletedRef = useRef(onDocDeleted)
  onDocDeletedRef.current = onDocDeleted

  const appendToAssistant = useCallback((fn: (parts: ChatMessage['parts']) => ChatMessage['parts']) => {
    setMessages((prev) => {
      const next = [...prev]
      let last = next[next.length - 1]
      if (!last || last.role !== 'assistant') {
        last = { role: 'assistant', parts: [] }
        next.push(last)
      } else {
        last = { ...last }
        next[next.length - 1] = last
      }
      last.parts = fn(last.parts)
      return next
    })
  }, [])

  useEffect(() => {
    let disposed = false
    let retryTimer: number | undefined

    const connect = () => {
      const proto = location.protocol === 'https:' ? 'wss' : 'ws'
      const ws = new WebSocket(`${proto}://${location.host}/ws`)
      wsRef.current = ws

      ws.onopen = () => setConnected(true)
      ws.onclose = () => {
        setConnected(false)
        setBusy(false)
        if (!disposed) retryTimer = window.setTimeout(connect, 2000)
      }
      ws.onmessage = (ev) => {
        const msg = JSON.parse(ev.data)
        switch (msg.type) {
          case 'start':
            setBusy(true)
            setMessages((prev) => [...prev, { role: 'assistant', parts: [] }])
            break
          case 'token':
            appendToAssistant((parts) => {
              const next = [...parts]
              const last = next[next.length - 1]
              if (last && last.kind === 'text') {
                next[next.length - 1] = { ...last, content: last.content + msg.content }
              } else {
                next.push({ kind: 'text', content: msg.content })
              }
              return next
            })
            break
          case 'tool_start':
            appendToAssistant((parts) => [
              ...parts,
              { kind: 'tool', name: msg.name, args: msg.args, status: 'running' } as ToolCallPart,
            ])
            break
          case 'tool_end':
            appendToAssistant((parts) => {
              const next = [...parts]
              for (let i = next.length - 1; i >= 0; i--) {
                const p = next[i]
                if (p.kind === 'tool' && p.status === 'running' && p.name === msg.name) {
                  next[i] = { ...p, status: msg.ok ? 'ok' : 'error', result: msg.result }
                  break
                }
              }
              return next
            })
            break
          case 'doc_updated':
            onDocUpdatedRef.current(msg.filename)
            break
          case 'doc_deleted':
            onDocDeletedRef.current(msg.filename)
            break
          case 'error':
            appendToAssistant((parts) => [
              ...parts,
              { kind: 'text', content: `\n\n⚠️ ${msg.message}` },
            ])
            break
          case 'done':
            setBusy(false)
            break
        }
      }
    }

    connect()
    return () => {
      disposed = true
      if (retryTimer) clearTimeout(retryTimer)
      wsRef.current?.close()
    }
  }, [appendToAssistant])

  const sendMessage = useCallback((content: string) => {
    const ws = wsRef.current
    if (!ws || ws.readyState !== WebSocket.OPEN) return false
    setMessages((prev) => [...prev, { role: 'user', parts: [{ kind: 'text', content }] }])
    ws.send(JSON.stringify({ type: 'chat', content }))
    return true
  }, [])

  const resetChat = useCallback(() => {
    setMessages([])
    wsRef.current?.send(JSON.stringify({ type: 'reset' }))
  }, [])

  return { messages, busy, connected, sendMessage, resetChat }
}
