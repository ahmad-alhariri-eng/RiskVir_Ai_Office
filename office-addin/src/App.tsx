import { useEffect, useState, useRef } from 'react'
import './App.css'
import { getSelectedWordText, type WordContext } from './utils/wordReader'
import { getSelectedExcelData, type ExcelContext } from './utils/excelReader'
import { getCurrentSlideText, type PptContext } from './utils/pptReader'
import { executeWordAction } from './utils/wordActions'
import { executeExcelAction } from './utils/excelActions'
import { executePptAction } from './utils/pptActions'
import { MessageContent } from './components/MessageContent'
import type { OfficeAction } from './types/actions'

// ─── Backend URL — HTTPS first, HTTP fallback ────────────────────
const BACKEND_HOST = 'lectures-rv.alwaysdata.net'
let BACKEND_URL = `https://${BACKEND_HOST}`

// ─── Types ───────────────────────────────────────────────────────
type Message = {
  role: 'user' | 'assistant'
  content: string
}

type OfficeContext = {
  app: string
  selection: string | unknown[][] | ExcelContext | WordContext | PptContext | null
}

// ─── Helpers ─────────────────────────────────────────────────────

/** Dispatches a parsed action to the correct Office host handler. */
async function dispatchAction(action: OfficeAction, host: string) {
  if (host === String(Office.HostType.Word)) {
    await executeWordAction(action as any)
  } else if (host === String(Office.HostType.Excel)) {
    await executeExcelAction(action as any)
  } else if (host === String(Office.HostType.PowerPoint)) {
    await executePptAction(action as any)
  }
}

/**
 * Robust streaming parser for SSE chunks.
 * Handles <office_action> blocks that may span multiple chunks,
 * and multiple action blocks in a single response.
 *
 * Returns updated { displayText, pendingActionBuf, inAction }.
 */
function processStreamChunk(
  chunk: string,
  displayText: string,
  pendingActionBuf: string,
  inAction: boolean
): {
  displayText: string
  pendingActionBuf: string
  inAction: boolean
  completedActions: string[]
  done: boolean
} {
  const completedActions: string[] = []
  let done = false

  const lines = chunk.split('\n')

  for (const line of lines) {
    if (!line.startsWith('data: ')) continue

    const dataStr = line.slice(6)
    if (dataStr.trim() === '[DONE]') { done = true; break }

    let token: string
    try {
      const parsed = JSON.parse(dataStr)
      token = parsed.token ?? ''
    } catch {
      continue
    }

    if (!inAction) {
      // Check if this token starts (or completes) an action tag
      const combined = displayText + token
      const openIdx = combined.indexOf('<office_action>')
      if (openIdx !== -1) {
        // Everything before the tag is visible text
        displayText = combined.slice(0, openIdx)
        inAction = true
        pendingActionBuf = combined.slice(openIdx + '<office_action>'.length)
      } else {
        displayText = combined
      }
    } else {
      pendingActionBuf += token
    }

    // While in action mode, check if closing tag has arrived
    while (inAction) {
      const closeIdx = pendingActionBuf.indexOf('</office_action>')
      if (closeIdx === -1) break

      // Extract the JSON payload
      const jsonPayload = pendingActionBuf.slice(0, closeIdx)
      completedActions.push(jsonPayload)

      // What remains after the closing tag
      const remainder = pendingActionBuf.slice(closeIdx + '</office_action>'.length)

      // Check if another action starts immediately in the remainder
      const nextOpen = remainder.indexOf('<office_action>')
      if (nextOpen !== -1) {
        displayText += remainder.slice(0, nextOpen)
        pendingActionBuf = remainder.slice(nextOpen + '<office_action>'.length)
        inAction = true
      } else {
        displayText += remainder
        pendingActionBuf = ''
        inAction = false
      }
    }
  }

  return { displayText, pendingActionBuf, inAction, completedActions, done }
}

// ─── Component ───────────────────────────────────────────────────
function App() {
  const [officeHost, setOfficeHost]       = useState<string>('Initializing...')
  const [backendStatus, setBackendStatus] = useState<string>('Connecting...')
  const [messages, setMessages]           = useState<Message[]>([])
  const [inputValue, setInputValue]       = useState('')
  const [isTyping, setIsTyping]           = useState(false)
  const messagesEndRef = useRef<HTMLDivElement>(null)

  // ── Office & backend initialization ──────────────────────────
  useEffect(() => {
    Office.onReady((info) => {
      setOfficeHost(info.host ? String(info.host) : 'Outside Office')
    })

    const checkHealth = (url: string) =>
      fetch(`${url}/health`, { headers: { 'Bypass-Tunnel-Reminder': 'true' } })
        .then((r) => r.json())
        .then((data) => {
          if (data.status === 'ok') {
            BACKEND_URL = url
            setBackendStatus(
              data.model_loaded ? '✅ Connected (Model Loaded)' : '⚠️ Connected (No Model)'
            )
          } else {
            setBackendStatus('⚠️ Backend Error')
          }
        })

    checkHealth(`https://${BACKEND_HOST}`).catch(() =>
      checkHealth(`http://${BACKEND_HOST}`).catch(() =>
        setBackendStatus('❌ Disconnected')
      )
    )
  }, [])

  // ── Auto-scroll ───────────────────────────────────────────────
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages])

  // ── Get current document context ─────────────────────────────
  const getContext = async (): Promise<OfficeContext> => {
    let selection: OfficeContext['selection'] = null
    try {
      if (officeHost === String(Office.HostType.Word)) {
        selection = await getSelectedWordText()
      } else if (officeHost === String(Office.HostType.Excel)) {
        selection = await getSelectedExcelData()
      } else if (officeHost === String(Office.HostType.PowerPoint)) {
        selection = await getCurrentSlideText()
      }
    } catch (e) {
      console.warn('Could not get document selection:', e)
    }
    return { app: officeHost, selection }
  }

  // ── Send message ──────────────────────────────────────────────
  const handleSend = async (overrideMessage?: string) => {
    const textToSend = (overrideMessage || inputValue).trim()
    if (!textToSend) return

    const userMsg: Message = { role: 'user', content: textToSend }
    setMessages((prev) => [...prev, userMsg])
    if (!overrideMessage) setInputValue('')
    setIsTyping(true)

    // Add an empty assistant bubble that we'll fill progressively
    setMessages((prev) => [...prev, { role: 'assistant', content: '' }])

    try {
      const context = await getContext()

      const response = await fetch(`${BACKEND_URL}/chat`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Bypass-Tunnel-Reminder': 'true' },
        body: JSON.stringify({
          message: userMsg.content,
          context,
          history: messages,
        }),
      })

      if (!response.ok) {
        const err = await response.json().catch(() => ({}))
        throw new Error(err.error || `HTTP ${response.status}`)
      }

      if (!response.body) throw new Error('No response body')

      const reader  = response.body.getReader()
      const decoder = new TextDecoder('utf-8')

      let displayText     = ''
      let pendingActionBuf = ''
      let inAction        = false
      let isDone          = false

      while (!isDone) {
        const { value, done } = await reader.read()
        if (done) break

        const chunk = decoder.decode(value, { stream: true })
        const result = processStreamChunk(chunk, displayText, pendingActionBuf, inAction)

        displayText      = result.displayText
        pendingActionBuf = result.pendingActionBuf
        inAction         = result.inAction
        isDone           = result.done

        // Execute any completed action blocks
        for (const jsonPayload of result.completedActions) {
          try {
            const action = JSON.parse(jsonPayload) as OfficeAction
            await dispatchAction(action, officeHost)
            displayText += '\n✅ Action executed.'
          } catch (e) {
            console.error('Action execution failed:', e)
            displayText += '\n⚠️ Action failed: ' + String(e)
          }
        }

        // Update the last assistant message reactively
        const snapshot = displayText
        setMessages((prev) => {
          const updated = [...prev]
          updated[updated.length - 1] = { role: 'assistant', content: snapshot }
          return updated
        })
      }
    } catch (error) {
      console.error('Chat error:', error)
      setMessages((prev) => {
        const updated = [...prev]
        updated[updated.length - 1] = {
          role: 'assistant',
          content: `⚠️ Error: ${String(error)}`,
        }
        return updated
      })
    } finally {
      setIsTyping(false)
    }
  }

  // ── Quick action prompts per host ─────────────────────────────
  const quickActions: Record<string, { label: string; prompt: string }[]> = {
    [String(Office.HostType?.Word)]: [
      { label: '✍️ Add Greeting',    prompt: 'Insert a professional greeting at the end of the document' },
      { label: '📋 Summarize',       prompt: 'Summarize the selected text in 3 bullet points' },
      { label: '🔍 Find & Replace',  prompt: 'Replace all occurrences of "Company" with "Acme Corp"' },
    ],
    [String(Office.HostType?.Excel)]: [
      { label: '📊 Sample Data',     prompt: 'Insert sample employee sales data with 5 rows' },
      { label: '📈 Add Chart',       prompt: 'Create a column chart from the selected data range' },
      { label: '🔢 SUM Formula',     prompt: 'Insert a SUM formula in the next empty cell below the selection' },
    ],
    [String(Office.HostType?.PowerPoint)]: [
      { label: '➕ Add Slide',       prompt: 'Add a slide about Artificial Intelligence trends' },
      { label: '📊 Chart Slide',     prompt: 'Add a slide with a revenue growth chart for Q1–Q4' },
      { label: '📋 Table Slide',     prompt: 'Add a slide with a comparison table of 3 products' },
    ],
  }

  const currentQuickActions = quickActions[officeHost] ?? []

  // ── Render ────────────────────────────────────────────────────
  return (
    <div className="app-shell">
      {/* Header */}
      <header className="app-header">
        <div className="app-title">RiskVir AI Office</div>
        <div className="app-meta">
          <span className="host-badge">{officeHost}</span>
          <span className={`status-badge ${backendStatus.startsWith('✅') ? 'ok' : backendStatus.startsWith('⚠️') ? 'warn' : 'err'}`}>
            {backendStatus}
          </span>
        </div>
      </header>

      {/* Messages */}
      <main className="messages-area">
        {messages.length === 0 ? (
          <div className="empty-state">
            <p className="empty-title">How can I help?</p>
            <p className="empty-sub">Select text in your document, then ask a question.</p>
            {currentQuickActions.length > 0 && (
              <div className="quick-actions">
                {currentQuickActions.map(({ label, prompt }) => (
                  <button key={label} className="quick-btn" onClick={() => handleSend(prompt)}>
                    {label}
                  </button>
                ))}
              </div>
            )}
          </div>
        ) : (
          messages.map((msg, idx) => (
            <div key={idx} className={`message-row ${msg.role}`}>
              <div className="bubble">
                <MessageContent content={msg.content} isUser={msg.role === 'user'} />
              </div>
            </div>
          ))
        )}
        {isTyping && (
          <div className="message-row assistant">
            <div className="bubble typing-indicator">
              <span /><span /><span />
            </div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </main>

      {/* Input bar */}
      <footer className="input-bar">
        <input
          type="text"
          className="chat-input"
          value={inputValue}
          onChange={(e) => setInputValue(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && !e.shiftKey && handleSend()}
          placeholder="Ask a question…"
          disabled={isTyping}
        />
        <button
          className="send-btn"
          onClick={() => handleSend()}
          disabled={isTyping || !inputValue.trim()}
        >
          Send
        </button>
      </footer>
    </div>
  )
}

export default App
