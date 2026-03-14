import { useEffect, useMemo, useRef, useState } from 'react'

const API_BASE =
  import.meta.env.VITE_API_BASE ||
  `${window.location.protocol}//${window.location.hostname}:8000`

const PRESET_OPTIONS = [
  { name: 'Business', desc: 'Executive, clean, data-focused', colors: ['#3b82f6', '#1e3a8a'] },
  { name: 'Tech', desc: 'Modern, product & engineering', colors: ['#22c55e', '#0f766e'] },
  { name: 'Education', desc: 'Clear, instructional, calm', colors: ['#eab308', '#a16207'] },
  { name: 'Marketing', desc: 'Bold, storytelling, conversion', colors: ['#f43f5e', '#be123c'] },
]

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

const toAbsolute = (url) => (url?.startsWith('http') ? url : `${API_BASE}${url}`)

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
  const [leftWidth, setLeftWidth] = useState(360)
  const [isResizing, setIsResizing] = useState(false)

  const [pptPath, setPptPath] = useState('')
  const [doc, setDoc] = useState(null)
  const [loading, setLoading] = useState(false)
  const [notice, setNotice] = useState('')

  const [previewImages, setPreviewImages] = useState([])
  const [selectedSlideIndex, setSelectedSlideIndex] = useState(0)
  const [dirtySlides, setDirtySlides] = useState(() => new Set())

  const [undoStack, setUndoStack] = useState([])
  const [redoStack, setRedoStack] = useState([])
  const [isDirty, setIsDirty] = useState(false)

  const [chatOpen, setChatOpen] = useState(true)
  const [chatInput, setChatInput] = useState('')
  const [messages, setMessages] = useState([
    { role: 'assistant', text: 'Enter a topic and click "Generate PPT". Then continue editing via chat.' },
  ])

  const autosaveTimerRef = useRef(null)

  const titleAndBody = useMemo(() => {
    if (!doc) return []
    return doc.slides.map((s) => ({
      ...s,
      editableShapes: s.shapes.filter((x) => ['title', 'body', 'text'].includes(x.role)),
    }))
  }, [doc])

  const selectedSlide = useMemo(
    () => titleAndBody.find((s) => s.slide_index === selectedSlideIndex) || titleAndBody[0] || null,
    [titleAndBody, selectedSlideIndex],
  )

  useEffect(() => {
    if (!isResizing) return

    const onMove = (e) => {
      const min = 300
      const max = Math.min(520, window.innerWidth * 0.45)
      setLeftWidth(Math.max(min, Math.min(max, e.clientX - 32)))
    }

    const onUp = () => setIsResizing(false)

    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
  }, [isResizing])

  useEffect(() => {
    if (!isDirty || !doc || !pptPath) return
    clearTimeout(autosaveTimerRef.current)
    autosaveTimerRef.current = setTimeout(() => {
      saveManualEdits({ silent: true })
    }, 1500)
    return () => clearTimeout(autosaveTimerRef.current)
  }, [doc, isDirty, pptPath])

  const loadPPT = async (path = pptPath) => {
    setLoading(true)
    try {
      const data = await api(`/api/ppt?path=${encodeURIComponent(path)}`)
      setDoc(data)
      setSelectedSlideIndex(0)
      setPptPath(path)
      setDirtySlides(new Set())
      setUndoStack([])
      setRedoStack([])
      setIsDirty(false)
      const preview = await api(`/api/ppt/preview?path=${encodeURIComponent(path)}`)
      setPreviewImages(preview.images || [])
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
      if (!prev) return prev
      setUndoStack((s) => [...s, structuredClone(prev)])
      setRedoStack([])
      const cloned = structuredClone(prev)
      const slide = cloned.slides.find((s) => s.slide_index === slideIndex)
      const shape = slide?.shapes.find((sh) => sh.shape_index === shapeIndex)
      if (shape) {
        shape.text = value
        setDirtySlides((set0) => new Set([...set0, slideIndex]))
        setIsDirty(true)
      }
      return cloned
    })
  }

  const undo = () => {
    if (!undoStack.length || !doc) return
    const prev = undoStack[undoStack.length - 1]
    setUndoStack((s) => s.slice(0, -1))
    setRedoStack((s) => [...s, structuredClone(doc)])
    setDoc(prev)
    setIsDirty(true)
  }

  const redo = () => {
    if (!redoStack.length || !doc) return
    const next = redoStack[redoStack.length - 1]
    setRedoStack((s) => s.slice(0, -1))
    setUndoStack((s) => [...s, structuredClone(doc)])
    setDoc(next)
    setIsDirty(true)
  }

  const saveManualEdits = async ({ silent = false } = {}) => {
    if (!doc || !pptPath) return
    setLoading(true)
    if (!silent) setNotice('')
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
      setNotice(silent ? `Auto-saved at ${new Date().toLocaleTimeString()}` : `Text edits saved: ${data.output_path}`)
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

  const copySlideText = async () => {
    if (!selectedSlide) return
    const text = selectedSlide.editableShapes.map((s) => `[${s.role}] ${s.text}`).join('\n')
    await navigator.clipboard.writeText(text)
    setNotice(`Slide ${selectedSlide.slide_index + 1} text copied`)
  }

  const rewriteTitle = () => {
    if (!selectedSlide) return
    const titleShape = selectedSlide.editableShapes.find((s) => s.role === 'title') || selectedSlide.editableShapes[0]
    if (!titleShape) return
    const current = titleShape.text.trim()
    const next = current ? `Overview: ${current.replace(/^Overview:\s*/i, '')}` : 'Overview'
    updateText(selectedSlide.slide_index, titleShape.shape_index, next)
  }

  const expandBullets = () => {
    if (!selectedSlide) return
    selectedSlide.editableShapes
      .filter((s) => s.role !== 'title')
      .forEach((shape) => {
        const expanded = `${shape.text}\n• Impact: expected business value and next action`
        updateText(selectedSlide.slide_index, shape.shape_index, expanded)
      })
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

      <main className="main-grid" style={{ gridTemplateColumns: `${leftWidth}px 8px 1fr` }}>
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
              Theme preset
              <div className="theme-cards">
                {PRESET_OPTIONS.map((option) => (
                  <button
                    type="button"
                    key={option.name}
                    data-testid={`preset-${option.name.toLowerCase()}`}
                    className={`theme-card ${form.preset === option.name ? 'active' : ''}`}
                    onClick={() => setForm((f) => ({ ...f, preset: option.name }))}
                  >
                    <div className="theme-swatches">
                      <span style={{ background: option.colors[0] }} />
                      <span style={{ background: option.colors[1] }} />
                    </div>
                    <strong>{option.name}</strong>
                    <small>{option.desc}</small>
                  </button>
                ))}
              </div>
            </label>
            <label>
              Audience
              <input value={form.audience} onChange={(e) => setForm((f) => ({ ...f, audience: e.target.value }))} />
            </label>
            <label>
              Tone
              <input value={form.tone} onChange={(e) => setForm((f) => ({ ...f, tone: e.target.value }))} />
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
              <select data-testid="language-select" value={form.language} onChange={(e) => setForm((f) => ({ ...f, language: e.target.value }))}>
                {LANGUAGE_OPTIONS.map((option) => (
                  <option key={option.value} value={option.value}>{option.label}</option>
                ))}
              </select>
            </label>
          </div>

          <button className="primary-btn" onClick={generatePPT} disabled={loading || !form.topic.trim()}>
            {loading ? 'Generating...' : 'Generate PPT'}
          </button>
          <p className="hint">The generated deck will load in the editor automatically.</p>
        </aside>

        <div className="resize-handle" onMouseDown={() => setIsResizing(true)} />

        <section className="panel editor-panel">
          <div className="panel-head">
            <h2>Preview and Text Editing</h2>
            <div className="row-actions">
              <button className="ghost-btn" onClick={undo} disabled={!undoStack.length}>Undo</button>
              <button className="ghost-btn" onClick={redo} disabled={!redoStack.length}>Redo</button>
              <button className="ghost-btn" onClick={() => saveManualEdits()} disabled={loading || !doc}>Save</button>
              {pptPath && (
                <a
                  className="ghost-btn link-btn"
                  href={`${API_BASE}/api/ppt/download?path=${encodeURIComponent(pptPath)}`}
                  target="_blank"
                  rel="noreferrer"
                >
                  Download .pptx
                </a>
              )}
            </div>
          </div>

          {!doc && <div className="skeleton">Start by entering a topic and generating a PPT from the left panel.</div>}
          {doc && (
            <div className="editor-layout">
              <div className="preview-column">
                {previewImages.map((img, idx) => (
                  <button
                    type="button"
                    key={img}
                    className={`preview-thumb ${selectedSlideIndex === idx ? 'active' : ''}`}
                    onClick={() => setSelectedSlideIndex(idx)}
                  >
                    <img src={toAbsolute(img)} alt={`Slide preview ${idx + 1}`} />
                    <span>
                      Slide {idx + 1}
                      {dirtySlides.has(idx) && <em className="dirty-tag"> • modified</em>}
                    </span>
                  </button>
                ))}
              </div>

              <div className="slides-list">
                {selectedSlide && (
                  <article className="slide-card" key={selectedSlide.slide_index}>
                    <h3>Slide {selectedSlide.slide_index + 1}</h3>
                    <div className="quick-actions">
                      <button className="ghost-btn" onClick={copySlideText}>Copy Slide Text</button>
                      <button className="ghost-btn" onClick={rewriteTitle}>Rewrite Title</button>
                      <button className="ghost-btn" onClick={expandBullets}>Expand Bullets</button>
                    </div>
                    {previewImages[selectedSlide.slide_index] && (
                      <div className="selected-preview">
                        <img src={toAbsolute(previewImages[selectedSlide.slide_index])} alt={`Selected slide ${selectedSlide.slide_index + 1}`} />
                      </div>
                    )}
                    {selectedSlide.editableShapes.map((sh) => (
                      <label key={sh.shape_index}>
                        <span>{sh.role} · shape #{sh.shape_index}</span>
                        <textarea value={sh.text} onChange={(e) => updateText(selectedSlide.slide_index, sh.shape_index, e.target.value)} />
                      </label>
                    ))}
                  </article>
                )}
              </div>
            </div>
          )}
        </section>
      </main>

      <button className="chat-fab" onClick={() => setChatOpen((v) => !v)}>
        {chatOpen ? 'Hide Chat' : 'Chat Edit'}
      </button>

      {chatOpen && (
        <section className="panel chat-floating">
          <h2>Chat Edits</h2>
          <div className="chat-box">
            {messages.map((m, i) => (
              <div key={i} className={`msg ${m.role}`}>{m.text}</div>
            ))}
            {loading && <div className="msg assistant">Processing...</div>}
          </div>
          <div className="chat-input-row">
            <input value={chatInput} onChange={(e) => setChatInput(e.target.value)} placeholder="Example: change slide 3 title to Growth Loop" />
            <button onClick={sendChat} disabled={loading || !doc}>Send</button>
          </div>
        </section>
      )}
    </div>
  )
}
