import { useMemo, useState } from 'react'

const API_BASE =
  import.meta.env.VITE_API_BASE ||
  `${window.location.protocol}//${window.location.hostname}:8000`

const PRESET_OPTIONS = ['Business', 'Tech', 'Education', 'Marketing']

const LANGUAGE_OPTIONS = [
  { label: 'English', value: 'en-US' },
  { label: 'Chinese', value: 'zh-CN' },
  { label: 'Japanese', value: 'ja-JP' },
]

async function api(path, options = {}) {
  const res = await fetch(`${API_BASE}${path}`, {
    headers: { 'Content-Type': 'application/json', ...(options.headers || {}) },
    ...options,
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: res.statusText }))
    throw new Error(err.detail || 'Request failed')
  }
  return res.json()
}

const initialForm = {
  topic: '',
  audience: '',
  tone: '',
  preset: 'Business',
  slide_count: 6,
  language: 'en-US',
}

export default function App() {
  const [form, setForm] = useState(initialForm)
  const [pptPath, setPptPath] = useState('')
  const [doc, setDoc] = useState(null)
  const [loading, setLoading] = useState(false)
  const [notice, setNotice] = useState('')
  const [chatInput, setChatInput] = useState('')
  const [messages, setMessages] = useState([
    { role: 'assistant', text: 'Enter a topic and click "Generate PPT". Then continue editing through chat.' },
  ])

  const titleAndBody = useMemo(() => {
    if (!doc) return []
    return doc.slides.map((s) => ({
      ...s,
      editableShapes: s.shapes.filter((x) => ['title', 'body', 'text'].includes(x.role)),
    }))
  }, [doc])

  const loadPPT = async (path = pptPath) => {
    setLoading(true)
    try {
      const data = await api(`/api/ppt?path=${encodeURIComponent(path)}`)
      setDoc(data)
      setPptPath(path)
    } finally {
      setLoading(false)
    }
  }

  const generatePPT = async () => {
    if (!form.topic.trim()) return
    setLoading(true)
    setNotice('')
    try {
      const payload = {
        ...form,
        topic: form.topic.trim(),
        audience: form.audience || undefined,
        tone: form.tone || undefined,
        preset: form.preset || undefined,
        language: form.language || undefined,
      }
      const data = await api('/api/generate', {
        method: 'POST',
        body: JSON.stringify(payload),
      })
      await loadPPT(data.output_path)
      setNotice(`Generated and loaded: ${data.output_path} (Theme: ${data.theme.name})`)
      setMessages((m) => [
        ...m,
        {
          role: 'assistant',
          text: `Generated ${data.outline.length} slides using theme ${data.theme.name}. You can keep editing manually or by chat.`,
        },
      ])
    } catch (e) {
      setNotice(e.message)
    } finally {
      setLoading(false)
    }
  }

  const updateText = (slideIndex, shapeIndex, value) => {
    setDoc((prev) => {
      const cloned = structuredClone(prev)
      const slide = cloned.slides.find((s) => s.slide_index === slideIndex)
      const shape = slide?.shapes.find((sh) => sh.shape_index === shapeIndex)
      if (shape) shape.text = value
      return cloned
    })
  }

  const saveManualEdits = async () => {
    if (!doc || !pptPath) return
    setLoading(true)
    setNotice('')
    try {
      const edits = []
      for (const s of doc.slides) {
        for (const sh of s.shapes) {
          edits.push({ slide_index: s.slide_index, shape_index: sh.shape_index, new_text: sh.text })
        }
      }
      const data = await api('/api/ppt/update', {
        method: 'POST',
        body: JSON.stringify({ path: pptPath, edits }),
      })
      await loadPPT(data.output_path)
      setNotice(`Text edits saved: ${data.output_path}`)
    } catch (e) {
      setNotice(e.message)
    } finally {
      setLoading(false)
    }
  }

  const sendChat = async () => {
    if (!chatInput.trim() || !pptPath) return
    const message = chatInput.trim()
    setLoading(true)
    setMessages((m) => [...m, { role: 'user', text: message }])
    setChatInput('')
    try {
      const data = await api('/api/chat', {
        method: 'POST',
        body: JSON.stringify({ path: pptPath, message }),
      })
      setMessages((m) => [
        ...m,
        {
          role: 'assistant',
          text: `Applied ${data.plan.length} edits (${data.used_llm ? 'LLM' : 'fallback'}) and saved to ${data.output_path}.`,
        },
      ])
      await loadPPT(data.output_path)
    } catch (e) {
      setMessages((m) => [...m, { role: 'assistant', text: `Failed: ${e.message}` }])
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div>
          <h1>ChatPPT Studio</h1>
          <p>AI presentation generation, live preview editing, and chat-based revisions</p>
        </div>
        <span className="badge">FastAPI + React</span>
      </header>

      {notice && <div className="notice">{notice}</div>}

      <main className="main-grid">
        <aside className="panel generator-panel">
          <h2>Generator</h2>
          <label>
            Topic / Prompt *
            <textarea
              value={form.topic}
              placeholder="Example: Q4 product strategy for AI customer support"
              onChange={(e) => setForm((f) => ({ ...f, topic: e.target.value }))}
            />
          </label>

          <div className="field-grid">
            <label>
              Preset
              <select
                data-testid="preset-select"
                value={form.preset}
                onChange={(e) => setForm((f) => ({ ...f, preset: e.target.value }))}
              >
                {PRESET_OPTIONS.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
            </label>
            <label>
              Audience
              <input
                value={form.audience}
                onChange={(e) => setForm((f) => ({ ...f, audience: e.target.value }))}
                placeholder="e.g. leadership team, students"
              />
            </label>
            <label>
              Tone
              <input
                value={form.tone}
                onChange={(e) => setForm((f) => ({ ...f, tone: e.target.value }))}
                placeholder="e.g. concise, formal"
              />
            </label>
            <label>
              Slides
              <input
                type="number"
                min={3}
                max={20}
                value={form.slide_count}
                onChange={(e) => setForm((f) => ({ ...f, slide_count: Number(e.target.value || 6) }))}
              />
            </label>
            <label>
              Language
              <select
                data-testid="language-select"
                value={form.language}
                onChange={(e) => setForm((f) => ({ ...f, language: e.target.value }))}
              >
                {LANGUAGE_OPTIONS.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                  </option>
                ))}
              </select>
            </label>
          </div>

          <button className="primary-btn" onClick={generatePPT} disabled={loading || !form.topic.trim()}>
            {loading ? 'Generating...' : 'Generate PPT'}
          </button>
          <p className="hint">The generated deck will load in the editor automatically.</p>
        </aside>

        <section className="panel editor-panel">
          <div className="panel-head">
            <h2>Preview and Text Editing</h2>
            <button className="ghost-btn" onClick={saveManualEdits} disabled={loading || !doc}>Save Text Changes</button>
          </div>

          {!doc && <div className="skeleton">Start by entering a topic and generating a PPT from the left panel.</div>}
          {doc && (
            <div className="slides-list">
              {titleAndBody.map((slide) => (
                <article className="slide-card" key={slide.slide_index}>
                  <h3>Slide {slide.slide_index + 1}</h3>
                  {slide.editableShapes.map((sh) => (
                    <label key={sh.shape_index}>
                      <span>{sh.role} · shape #{sh.shape_index}</span>
                      <textarea value={sh.text} onChange={(e) => updateText(slide.slide_index, sh.shape_index, e.target.value)} />
                    </label>
                  ))}
                </article>
              ))}
            </div>
          )}
        </section>

        <section className="panel chat-panel">
          <h2>Chat Edits</h2>
          <div className="chat-box">
            {messages.map((m, i) => (
              <div key={i} className={`msg ${m.role}`}>{m.text}</div>
            ))}
            {loading && <div className="msg assistant">Processing...</div>}
          </div>
          <div className="chat-input-row">
            <input
              value={chatInput}
              onChange={(e) => setChatInput(e.target.value)}
              placeholder="Example: change slide 3 title to Growth Loop"
            />
            <button onClick={sendChat} disabled={loading || !doc}>Send</button>
          </div>
        </section>
      </main>
    </div>
  )
}
