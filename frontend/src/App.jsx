import { useMemo, useState } from 'react'

const API_BASE =
  import.meta.env.VITE_API_BASE ||
  `${window.location.protocol}//${window.location.hostname}:8000`

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
  slide_count: 6,
  language: 'zh-CN',
}

export default function App() {
  const [form, setForm] = useState(initialForm)
  const [pptPath, setPptPath] = useState('')
  const [doc, setDoc] = useState(null)
  const [loading, setLoading] = useState(false)
  const [notice, setNotice] = useState('')
  const [chatInput, setChatInput] = useState('')
  const [messages, setMessages] = useState([
    { role: 'assistant', text: '👋 输入主题后点击「生成PPT」，随后可以继续聊天修改内容。' },
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
        language: form.language || undefined,
      }
      const data = await api('/api/generate', {
        method: 'POST',
        body: JSON.stringify(payload),
      })
      await loadPPT(data.output_path)
      setNotice(`已生成并加载：${data.output_path}（主题：${data.theme.name}）`)
      setMessages((m) => [
        ...m,
        {
          role: 'assistant',
          text: `已生成 ${data.outline.length} 页大纲，主题风格 ${data.theme.name}。你可以继续编辑或聊天改稿。`,
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
      setNotice(`文本修改已保存：${data.output_path}`)
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
          text: `已执行 ${data.plan.length} 条修改（${data.used_llm ? 'LLM' : 'fallback'}），输出：${data.output_path}`,
        },
      ])
      await loadPPT(data.output_path)
    } catch (e) {
      setMessages((m) => [...m, { role: 'assistant', text: `失败：${e.message}` }])
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div>
          <h1>ChatPPT Studio</h1>
          <p>AI 生成演示稿 · 在线预览编辑 · 对话持续改稿</p>
        </div>
        <span className="badge">FastAPI + React</span>
      </header>

      {notice && <div className="notice">{notice}</div>}

      <main className="main-grid">
        <aside className="panel generator-panel">
          <h2>生成器</h2>
          <label>
            主题 / 要求 *
            <textarea
              value={form.topic}
              placeholder="例如：给投资人汇报 AI 客服产品路线图"
              onChange={(e) => setForm((f) => ({ ...f, topic: e.target.value }))}
            />
          </label>

          <div className="field-grid">
            <label>
              受众
              <input value={form.audience} onChange={(e) => setForm((f) => ({ ...f, audience: e.target.value }))} placeholder="如：管理层、学生" />
            </label>
            <label>
              语气
              <input value={form.tone} onChange={(e) => setForm((f) => ({ ...f, tone: e.target.value }))} placeholder="如：专业、轻松" />
            </label>
            <label>
              页数
              <input
                type="number"
                min={3}
                max={20}
                value={form.slide_count}
                onChange={(e) => setForm((f) => ({ ...f, slide_count: Number(e.target.value || 6) }))}
              />
            </label>
            <label>
              语言
              <input value={form.language} onChange={(e) => setForm((f) => ({ ...f, language: e.target.value }))} placeholder="zh-CN / en" />
            </label>
          </div>

          <button className="primary-btn" onClick={generatePPT} disabled={loading || !form.topic.trim()}>
            {loading ? '生成中...' : '生成PPT'}
          </button>
          <p className="hint">生成完成后会自动加载到右侧编辑区。</p>
        </aside>

        <section className="panel editor-panel">
          <div className="panel-head">
            <h2>预览与文本编辑</h2>
            <button className="ghost-btn" onClick={saveManualEdits} disabled={loading || !doc}>保存文本修改</button>
          </div>

          {!doc && <div className="skeleton">先在左侧输入主题并生成 PPT。</div>}
          {doc && (
            <div className="slides-list">
              {titleAndBody.map((slide) => (
                <article className="slide-card" key={slide.slide_index}>
                  <h3>第 {slide.slide_index + 1} 页</h3>
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
          <h2>聊天改稿</h2>
          <div className="chat-box">
            {messages.map((m, i) => (
              <div key={i} className={`msg ${m.role}`}>{m.text}</div>
            ))}
            {loading && <div className="msg assistant">处理中...</div>}
          </div>
          <div className="chat-input-row">
            <input
              value={chatInput}
              onChange={(e) => setChatInput(e.target.value)}
              placeholder="例如：把第3页标题改为增长飞轮"
            />
            <button onClick={sendChat} disabled={loading || !doc}>发送</button>
          </div>
        </section>
      </main>
    </div>
  )
}
