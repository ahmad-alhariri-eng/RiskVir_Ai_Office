import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'

interface MessageContentProps {
  content: string
  isUser: boolean
}

/**
 * Renders assistant messages as formatted Markdown.
 * User messages are shown as plain text (they rarely contain markdown).
 */
export function MessageContent({ content, isUser }: MessageContentProps) {
  if (isUser) {
    return <span className="bubble-plain">{content}</span>
  }

  return (
    <div className="bubble-markdown">
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        components={{
          // ── Code blocks ──────────────────────────────────────
          code({ node, className, children, ...props }: any) {
            const isBlock = className?.startsWith('language-')
            const code = String(children).replace(/\n$/, '')

            if (isBlock) {
              const lang = className?.replace('language-', '') ?? ''
              return (
                <div className="code-block">
                  {lang && <span className="code-lang">{lang}</span>}
                  <pre><code {...props}>{code}</code></pre>
                </div>
              )
            }
            return <code className="code-inline" {...props}>{code}</code>
          },

          // ── Tables ───────────────────────────────────────────
          table({ children }: any) {
            return (
              <div className="table-wrapper">
                <table>{children}</table>
              </div>
            )
          },

          // ── Links — open in external browser ─────────────────
          a({ href, children }: any) {
            return (
              <a href={href} target="_blank" rel="noopener noreferrer">
                {children}
              </a>
            )
          },

          // ── Blockquotes ──────────────────────────────────────
          blockquote({ children }: any) {
            return <blockquote className="md-blockquote">{children}</blockquote>
          },
        }}
      >
        {content}
      </ReactMarkdown>
    </div>
  )
}
