import { useMemo, useState } from 'react'

const API_BASE = import.meta.env.VITE_API_BASE || 'http://127.0.0.1:8000'

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

export default function App() {
  const [pptPath, setPptPath] = useState('')
  const [doc, setDoc] = useState(null)
  const [loading, setLoading] = useState(false)
  const [chatInput, setChatInput] = useState('')
  const [messages, setMessages] = useState([])
  const [notice, setNotice] = useState('')

  const titleAndBody = useMemo(() => {
    if (!doc) return []
    return doc.slides.map((s) => ({
      ...s,
      editableShapes: s.shapes.filter((x) => ['title', 'body', 'text'].includes(x.role)),
    }))
  }, [doc])

  const loadPPT = async () => {
    setLoading(true)
    setNotice('')
    try {
      const data = await api(`/api/ppt?path=${encodeURIComponent(pptPath)}`)
      setDoc(data)
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
    if (!doc) return
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
      setNotice(`已保存到: ${data.output_path}`)
      setPptPath(data.output_path)
      await loadPPT()
    } catch (e) {
      setNotice(e.message)
    } finally {
      setLoading(false)
    }
  }

  const sendChat = async () => {
    if (!chatInput.trim()) return
    setLoading(true)
    setMessages((m) => [...m, { role: 'user', text: chatInput }])
    try {
      const data = await api('/api/chat', {
        method: 'POST',
        body: JSON.stringify({ path: pptPath, message: chatInput }),
      })
      setMessages((m) => [
        ...m,
        {
          role: 'assistant',
          text: `已执行 ${data.plan.length} 条修改（${data.used_llm ? 'LLM' : 'fallback'}）。输出：${data.output_path}`,
        },
      ])
      setPptPath(data.output_path)
      setChatInput('')
      await loadPPT()
    } catch (e) {
      setMessages((m) => [...m, { role: 'assistant', text: `失败: ${e.message}` }])
    } finally {
      setLoading(false)
    }
  }

  const applyTheme = async () => {
    setLoading(true)
    try {
      const data = await api('/api/theme/apply', {
        method: 'POST',
        body: JSON.stringify({ path: pptPath }),
      })
      setNotice(`主题 ${data.theme.name} 已应用，输出: ${data.output_path}`)
      setPptPath(data.output_path)
      await loadPPT()
    } catch (e) {
      setNotice(e.message)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="container">
      <h1>ChatPPT Web Editor</h1>
      <div className="row">
        <input value={pptPath} onChange={(e) => setPptPath(e.target.value)} placeholder="输入本机 PPT 路径，如 /tmp/demo.pptx" />
        <button onClick={loadPPT} disabled={loading || !pptPath}>加载</button>
        <button onClick={saveManualEdits} disabled={loading || !doc}>保存文本修改</button>
        <button onClick={applyTheme} disabled={loading || !doc}>自动应用主题</button>
      </div>
      {notice && <p className="notice">{notice}</p>}

      <div className="layout">
        <section>
          <h2>PPT 页面编辑</h2>
          {!doc && <p>先加载一个 PPT 文件。</p>}
          {titleAndBody.map((slide) => (
            <div className="slide" key={slide.slide_index}>
              <h3>第 {slide.slide_index + 1} 页</h3>
              {slide.editableShapes.map((sh) => (
                <label key={sh.shape_index}>
                  <span>{sh.role} (shape #{sh.shape_index})</span>
                  <textarea value={sh.text} onChange={(e) => updateText(slide.slide_index, sh.shape_index, e.target.value)} />
                </label>
              ))}
            </div>
          ))}
        </section>

        <section>
          <h2>聊天改PPT</h2>
          <div className="chat-box">
            {messages.map((m, i) => (
              <div key={i} className={`msg ${m.role}`}>{m.text}</div>
            ))}
          </div>
          <div className="row">
            <input value={chatInput} onChange={(e) => setChatInput(e.target.value)} placeholder="例如：把第3页标题改为年度规划" />
            <button onClick={sendChat} disabled={loading || !doc}>发送</button>
          </div>
        </section>
      </div>
    </div>
  )
}
