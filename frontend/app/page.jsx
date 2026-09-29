
"use client"

import { useState, useRef, useEffect, useCallback, useMemo } from "react"
import { Canvas, useFrame } from "@react-three/fiber"
import * as THREE from "three"
import { motion, AnimatePresence } from "framer-motion"

const API_BASE = "/api"

// ─── Volumetric 3D Backdrop ───────────────────────────────────────
// A lit, physically-shaded scene: a machined core inside a wireframe
// containment shell, suspended in a slow-drifting particle field.
// The whole rig parallaxes gently against the pointer for real depth.

function BackdropScene() {
  const rig = useRef()
  const core = useRef()
  const shell = useRef()
  const halo = useRef()
  const dust = useRef()

  const dustGeometry = useMemo(() => {
    const count = 460
    const positions = new Float32Array(count * 3)
    for (let i = 0; i < count; i++) {
      const radius = 5.5 + Math.random() * 11
      const theta = Math.random() * Math.PI * 2
      positions[i * 3] = Math.cos(theta) * radius
      positions[i * 3 + 1] = (Math.random() - 0.5) * 14
      positions[i * 3 + 2] = Math.sin(theta) * radius
    }
    const geometry = new THREE.BufferGeometry()
    geometry.setAttribute("position", new THREE.BufferAttribute(positions, 3))
    return geometry
  }, [])

  useEffect(() => () => dustGeometry.dispose(), [dustGeometry])

  useFrame((state, delta) => {
    const time = state.clock.elapsedTime
    const pointer = state.pointer || state.mouse

    if (core.current) {
      core.current.rotation.y += delta * 0.14
      core.current.rotation.x = Math.sin(time * 0.2) * 0.18
      core.current.position.y = Math.sin(time * 0.5) * 0.22
    }
    if (shell.current) {
      shell.current.rotation.y -= delta * 0.07
      shell.current.rotation.z = Math.cos(time * 0.16) * 0.12
      shell.current.position.y = Math.sin(time * 0.5) * 0.22
    }
    if (halo.current) {
      halo.current.rotation.z += delta * 0.05
      halo.current.rotation.x = 1.15 + Math.sin(time * 0.24) * 0.1
    }
    if (dust.current) {
      dust.current.rotation.y += delta * 0.012
    }
    if (rig.current && pointer) {
      rig.current.rotation.y += (pointer.x * 0.22 - rig.current.rotation.y) * 0.035
      rig.current.rotation.x += (-pointer.y * 0.14 - rig.current.rotation.x) * 0.035
    }
  })

  return (
    <group ref={rig}>
      <ambientLight intensity={0.35} />
      <directionalLight position={[6, 8, 6]} intensity={1.1} color="#c7d2ff" />
      <pointLight position={[-7, -3, 4]} intensity={45} distance={26} color="#35e0ff" />
      <pointLight position={[6, 5, -4]} intensity={40} distance={26} color="#7c8cff" />

      {/* Machined inner core */}
      <mesh ref={core}>
        <icosahedronGeometry args={[1.35, 1]} />
        <meshStandardMaterial
          color="#2b3566"
          metalness={0.92}
          roughness={0.22}
          flatShading
        />
      </mesh>

      {/* Wireframe containment shell */}
      <mesh ref={shell}>
        <icosahedronGeometry args={[2.35, 2]} />
        <meshBasicMaterial color="#7c8cff" wireframe transparent opacity={0.13} />
      </mesh>

      {/* Orbital halo ring */}
      <mesh ref={halo} rotation={[1.15, 0, 0]}>
        <torusGeometry args={[3.5, 0.012, 8, 160]} />
        <meshBasicMaterial color="#35e0ff" transparent opacity={0.3} />
      </mesh>

      {/* Suspended particle field */}
      <points ref={dust} geometry={dustGeometry}>
        <pointsMaterial
          size={0.035}
          color="#a6b2ff"
          transparent
          opacity={0.55}
          sizeAttenuation
        />
      </points>
    </group>
  )
}

// ─── Simple Markdown Renderer ─────────────────────────────────────

function renderMarkdown(text) {
  if (!text) return null
  if (typeof text !== "string") {
    if (Array.isArray(text)) {
      text = text.map(item => (typeof item === "object" ? item?.text || JSON.stringify(item) : String(item))).join("\n")
    } else if (typeof text === "object") {
      text = text?.text || JSON.stringify(text, null, 2)
    } else {
      text = String(text)
    }
  }

  const lines = text.split("\n")
  const elements = []
  let inCodeBlock = false
  let codeContent = []
  let listItems = []

  const flushList = () => {
    if (listItems.length > 0) {
      elements.push(
        <ul key={`ul-${elements.length}`}>
          {listItems.map((item, i) => (
            <li key={i} dangerouslySetInnerHTML={{ __html: inlineFormat(item) }} />
          ))}
        </ul>
      )
      listItems = []
    }
  }

  const inlineFormat = (str) => {
    return str
      .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
      .replace(/\*(.+?)\*/g, "<em>$1</em>")
      .replace(/`([^`]+)`/g, "<code>$1</code>")
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]

    if (line.startsWith("```")) {
      if (inCodeBlock) {
        elements.push(
          <pre key={`code-${elements.length}`}>
            <code>{codeContent.join("\n")}</code>
          </pre>
        )
        codeContent = []
        inCodeBlock = false
      } else {
        flushList()
        inCodeBlock = true
      }
      continue
    }

    if (inCodeBlock) {
      codeContent.push(line)
      continue
    }

    if (line.match(/^[-*•]\s/)) {
      listItems.push(line.replace(/^[-*•]\s/, ""))
      continue
    } else {
      flushList()
    }

    if (line.match(/^#{1,3}\s/)) {
      const level = line.match(/^(#{1,3})\s/)[1].length
      const text = line.replace(/^#{1,3}\s/, "")
      const Tag = `h${level + 1}`
      elements.push(
        <Tag
          key={`h-${elements.length}`}
          dangerouslySetInnerHTML={{ __html: inlineFormat(text) }}
          style={{ marginTop: "12px", marginBottom: "6px" }}
        />
      )
    } else if (line.trim() === "") {
      // skip
    } else {
      elements.push(
        <p
          key={`p-${elements.length}`}
          dangerouslySetInnerHTML={{ __html: inlineFormat(line) }}
        />
      )
    }
  }

  flushList()
  return elements
}

// ─── Suggestion Prompts ───────────────────────────────────────────

const SUGGESTIONS = [
  "Give me a summary of this spreadsheet",
  "What are the key statistics?",
  "Show me a bar chart of the data",
  "Which column has the highest values?",
]

// ─── Main Page Component ──────────────────────────────────────────

export default function Home() {
  const [files, setFiles] = useState([])
  const [activeFile, setActiveFile] = useState(null)
  const [messages, setMessages] = useState([])
  const [input, setInput] = useState("")
  const [isLoading, setIsLoading] = useState(false)
  const [isDragOver, setIsDragOver] = useState(false)
  const [uploadProgress, setUploadProgress] = useState(null)
  const [deletingFile, setDeletingFile] = useState(null)
  const [threadId] = useState(() => "session_" + Math.random().toString(36).substring(2, 11))
  const [previewModal, setPreviewModal] = useState({
    isOpen: false,
    filename: "",
    activeSheet: "",
    sheets: [],
    columns: [],
    data: [],
    totalRows: 0,
    offset: 0,
    limit: 100,
    loading: false,
    error: null,
    range: "",
  })

  const openPreview = useCallback(async (filename, sheetName = "", offset = 0, limit = 100, range = "") => {
    setPreviewModal((prev) => ({
      ...prev,
      isOpen: true,
      filename,
      activeSheet: sheetName || prev.activeSheet,
      offset,
      limit,
      loading: true,
      error: null,
      range,
    }))

    try {
      const url = new URL(`${API_BASE}/preview/${encodeURIComponent(filename)}`, window.location.origin)
      if (sheetName) url.searchParams.set("sheet_name", sheetName)
      if (range) url.searchParams.set("cell_range", range)
      url.searchParams.set("offset", offset)
      url.searchParams.set("limit", limit)
      
      const res = await fetch(url.toString())
      if (!res.ok) {
        const err = await res.json()
        throw new Error(err.detail || "Failed to load preview")
      }
      const data = await res.json()
      setPreviewModal({
        isOpen: true,
        filename,
        activeSheet: data.active_sheet,
        sheets: data.sheets || [],
        columns: data.columns || [],
        data: data.data || [],
        totalRows: data.total_rows || 0,
        offset: data.offset ?? offset,
        limit: data.limit ?? limit,
        loading: false,
        error: null,
        range,
      })
    } catch (err) {
      setPreviewModal((prev) => ({
        ...prev,
        loading: false,
        error: err.message,
      }))
    }
  }, [])

  const closePreview = () => {
    setPreviewModal((prev) => ({ ...prev, isOpen: false }))
  }

  const messagesEndRef = useRef(null)
  const fileInputRef = useRef(null)

  const refreshFiles = useCallback(async () => {
    const response = await fetch(`${API_BASE}/files`)
    if (!response.ok) throw new Error("Could not refresh files")
    const data = await response.json()
    setFiles((previous) => (data.files || []).map((name) => ({
      name,
      size: previous.find((file) => file.name === name)?.size || null,
    })))
    return data.files || []
  }, [])

  const handleDeleteFile = useCallback(async (filename) => {
    if (!window.confirm(`Delete ${filename}? This cannot be undone.`)) return

    setDeletingFile(filename)
    try {
      const response = await fetch(`${API_BASE}/files/${encodeURIComponent(filename)}`, {
        method: "DELETE",
      })
      if (!response.ok) {
        const error = await response.json().catch(() => ({}))
        throw new Error(error.detail || "Could not delete file")
      }

      setFiles((previous) => previous.filter((file) => file.name !== filename))
      setActiveFile((previous) => previous === filename ? null : previous)
      setPreviewModal((previous) => previous.filename === filename
        ? { ...previous, isOpen: false }
        : previous)
    } catch (error) {
      setMessages((previous) => [
        ...previous,
        { role: "assistant", content: `Unable to delete **${filename}**: ${error.message}` },
      ])
    } finally {
      setDeletingFile(null)
    }
  }, [])

  useEffect(() => {
    refreshFiles().catch(() => {})
  }, [refreshFiles])

  // Auto-scroll to bottom on new messages
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" })
  }, [messages, isLoading])

  // ─── File Upload ──────────────────────────────────────────────

  const handleUpload = useCallback(async (file) => {
    if (!file) return
    setUploadProgress(0)

    const formData = new FormData()
    formData.append("file", file)

    try {
      // Simulate progress
      const progressInterval = setInterval(() => {
        setUploadProgress((prev) => Math.min((prev || 0) + 15, 90))
      }, 200)

      const res = await fetch(`${API_BASE}/upload`, {
        method: "POST",
        body: formData,
      })

      clearInterval(progressInterval)
      setUploadProgress(100)

      if (!res.ok) {
        const err = await res.json()
        throw new Error(err.detail || "Upload failed")
      }

      const data = await res.json()
      setFiles((prev) => [...prev.filter((item) => item.name !== data.file), {
        name: data.file,
        size: data.size_bytes,
      }])
      setActiveFile(data.file)

      // Add system message
      setMessages((prev) => [
        ...prev,
        {
          role: "assistant",
          content: `📊 **${data.file}** uploaded successfully! You can now ask me questions about your data.`,
        },
      ])

      setTimeout(() => setUploadProgress(null), 1000)
    } catch (error) {
      setUploadProgress(null)
      setMessages((prev) => [
        ...prev,
        {
          role: "assistant",
          content: `❌ Upload failed: ${error.message}`,
        },
      ])
    }
  }, [])

  const onDrop = useCallback((e) => {
    e.preventDefault()
    setIsDragOver(false)
    const file = e.dataTransfer?.files?.[0]
    if (file) handleUpload(file)
  }, [handleUpload])

  const onDragOver = useCallback((e) => {
    e.preventDefault()
    setIsDragOver(true)
  }, [])

  const onDragLeave = useCallback(() => {
    setIsDragOver(false)
  }, [])

  // ─── Chat ─────────────────────────────────────────────────────

  const sendMessage = useCallback(async (text) => {
    const question = text || input.trim()
    if (!question || isLoading) return

    setInput("")
    setMessages((prev) => [...prev, { role: "user", content: question }])
    setIsLoading(true)

    // The agent can legitimately take several minutes (multiple LLM + MCP tool
    // calls, plus provider-overload retries). This client-side ceiling is kept
    // slightly above the Next.js proxy's own timeout (see next.config.js) so a
    // real proxy/backend timeout is reported before this abort ever fires.
    const controller = new AbortController()
    const timeoutId = setTimeout(() => controller.abort(), 630_000) // 10.5 minutes

    try {
      const targetFile = activeFile || (files.length > 0 ? files[files.length - 1].name : null)
      const res = await fetch(`${API_BASE}/chat`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          question,
          thread_id: threadId,
          file_name: targetFile,
        }),
        signal: controller.signal,
      })

      const data = await res.json()
      if (Array.isArray(data.files)) {
        setFiles((previous) => data.files.map((name) => ({
          name,
          size: previous.find((file) => file.name === name)?.size || null,
        })))
      } else {
        refreshFiles().catch((error) => setMessages((previous) => [
          ...previous, { role: "assistant", content: `Workbook list could not refresh: ${error.message}` },
        ]))
      }
      if (data.created_files?.length) setActiveFile(data.created_files[0])

      setMessages((prev) => [
        ...prev,
        {
          role: "assistant",
          content: data.answer || "No response received.",
          chart: data.chart_image || null,
          updatedTables: data.updated_tables || [],
          syncError: data.sync_error || null,
        },
      ])
    } catch (error) {
      // A workbook write may have succeeded before the response was lost.
      // Refresh files only; never resend the Excel request automatically.
      await refreshFiles().catch(() => {})
      const message = error.name === "AbortError"
        ? "⚠️ The request took too long and was stopped on this side. The backend may still finish the workbook change — check the file list."
        : `⚠️ Connection error: Could not reach the backend. Make sure the server is running at \`${API_BASE}\`.`
      setMessages((prev) => [
        ...prev,
        {
          role: "assistant",
          content: message,
        },
      ])
    } finally {
      clearTimeout(timeoutId)
      setIsLoading(false)
    }
  }, [input, isLoading, activeFile, files, threadId, refreshFiles])

  const handleKeyDown = (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault()
      sendMessage()
    }
  }

  const formatFileSize = (bytes) => {
    if (!bytes) return ""
    if (bytes < 1024) return `${bytes} B`
    if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`
    return `${(bytes / 1048576).toFixed(1)} MB`
  }

  // ─── Render ───────────────────────────────────────────────────

  return (
    <div className="app">
      {/* ── Volumetric environment ── */}
      <div className="scene-backdrop" />
      <div className="canvas-container">
        <Canvas
          camera={{ position: [0, 0, 9], fov: 45 }}
          dpr={[1, 1.8]}
          gl={{ antialias: true, alpha: true }}
        >
          <BackdropScene />
        </Canvas>
      </div>
      <div className="backdrop-grid" />
      <div className="backdrop-vignette" />

      <div className="workspace">
        {/* ── Instrument rail ── */}
        <aside className="sidebar panel">
          <div className="logo">
            <div className="logo-icon">📊</div>
            <div>
              <h1>AI Excel Analyst</h1>
              <p>OpenRouter · LangGraph · MCP</p>
            </div>
          </div>

          <div className="sidebar-scroll">
            {/* Upload Zone */}
            <div
              className={`upload-zone ${isDragOver ? "dragover" : ""}`}
              onClick={() => fileInputRef.current?.click()}
              onDrop={onDrop}
              onDragOver={onDragOver}
              onDragLeave={onDragLeave}
            >
              <span className="upload-icon">📁</span>
              <h3>Upload Excel File</h3>
              <p>.xlsx or .xls</p>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => handleUpload(e.target.files?.[0])}
              />
              {uploadProgress !== null && (
                <div className="upload-progress">
                  <div
                    className="upload-progress-bar"
                    style={{ width: `${uploadProgress}%` }}
                  />
                </div>
              )}
            </div>

            {/* File List */}
            {files.length > 0 && (
              <div className="file-section">
                <h4>Workbooks</h4>
                {files.map((file, i) => (
                  <motion.div
                    key={i}
                    className={`file-item ${activeFile === file.name ? "active-file" : ""}`}
                    onClick={() => setActiveFile(file.name)}
                    style={{ cursor: "pointer" }}
                    initial={{ opacity: 0, y: 8 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: i * 0.1 }}
                  >
                    <span className="file-icon">📄</span>
                    <div className="file-info">
                      <div className="file-name" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%" }}>
                        <span style={{ overflow: "hidden", textOverflow: "ellipsis" }}>{file.name}</span>
                        <div className="file-actions">
                          <button
                            className="action-btn"
                            onClick={(e) => {
                              e.stopPropagation()
                              setActiveFile(file.name)
                              openPreview(file.name)
                            }}
                            title="Preview spreadsheet"
                          >
                            👁️
                          </button>
                          <button
                            className="action-btn delete-file-btn"
                            onClick={(e) => {
                              e.stopPropagation()
                              handleDeleteFile(file.name)
                            }}
                            disabled={deletingFile === file.name}
                            title={`Delete ${file.name}`}
                            aria-label={`Delete ${file.name}`}
                          >
                            {deletingFile === file.name ? "…" : "🗑️"}
                          </button>
                          <a
                            href={`${API_BASE}/download/${encodeURIComponent(file.name)}`}
                            download
                            className="action-btn"
                            onClick={(e) => e.stopPropagation()}
                            title="Download file"
                          >
                            ⬇️
                          </a>
                        </div>
                      </div>
                      <div className="file-size">{formatFileSize(file.size)}</div>
                    </div>
                    <div className="file-status" />
                  </motion.div>
                ))}
              </div>
            )}
          </div>

          {/* Rail telemetry */}
          <div className="rail-footer">
            <div className="rail-stat">
              <span>Workbooks</span>
              <strong>{files.length}</strong>
            </div>
            <div className="rail-stat">
              <span>Active</span>
              <strong>{activeFile || "none"}</strong>
            </div>
            <div className="rail-stat">
              <span>Session</span>
              <strong>{threadId.replace("session_", "")}</strong>
            </div>
          </div>
        </aside>

        {/* ── Chat deck ── */}
        <main className="chat-area panel">
          <header className="chat-header">
            <div className="chat-title">
              <div>
                <h2>Analysis Console</h2>
                <div className="chat-subtitle">
                  {activeFile ? `Bound to ${activeFile}` : "No workbook bound · agent can create one"}
                </div>
              </div>
            </div>
            <div className="header-chips">
              {files.length > 0 && (
                <span className="chip">
                  <span>◆</span>
                  <strong>{files.length}</strong> loaded
                </span>
              )}
              <div className="status">
                <div className="status-dot" />
                OpenRouter Connected
              </div>
            </div>
          </header>

          {/* Messages or Welcome Screen */}
          {messages.length === 0 ? (
            <div className="welcome">
              <div className="welcome-icon">✨</div>
              <h2>Welcome to AI Excel Analyst</h2>
              <p>
                Upload a spreadsheet and ask questions in natural language.
                I'll analyze your data, compute statistics, and generate
                visualizations.
              </p>
              <div className="suggestions">
                {SUGGESTIONS.map((text, i) => (
                  <motion.button
                    key={i}
                    className="suggestion"
                    onClick={() => sendMessage(text)}
                    initial={{ opacity: 0, y: 12 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: 0.1 + i * 0.08 }}
                  >
                    {text}
                  </motion.button>
                ))}
              </div>
            </div>
          ) : (
            <div className="messages">
              <AnimatePresence>
                {messages.map((msg, i) => (
                  <motion.div
                    key={i}
                    className={`message ${msg.role}`}
                    initial={{ opacity: 0, y: 14, scale: 0.98 }}
                    animate={{ opacity: 1, y: 0, scale: 1 }}
                    transition={{ duration: 0.36, ease: [0.22, 1, 0.36, 1] }}
                  >
                    <div className="message-avatar">
                      {msg.role === "assistant" ? "🤖" : "👤"}
                    </div>
                    <div className="message-content">
                      {renderMarkdown(msg.content)}
                      {msg.updatedTables?.map((table, tableIndex) => (
                        <div className="inline-table" key={`${table.workbook}-${table.sheet}-${tableIndex}`}>
                          <div className="inline-table-heading">
                            <strong>{table.workbook} · {table.sheet} · {table.range}</strong>
                            <button onClick={() => openPreview(table.workbook, table.sheet, 0, 100, table.range)}>Open preview</button>
                          </div>
                          <div className="inline-table-scroll">
                            <table className="preview-table">
                              <thead><tr>{table.headers.map((header, i) => <th key={i}>{header}</th>)}</tr></thead>
                              <tbody>{table.rows.map((row, i) => (
                                <tr key={i}>{row.map((value, j) => <td key={j}>{value}</td>)}</tr>
                              ))}</tbody>
                            </table>
                          </div>
                          <small>Showing {table.rows.length} of {table.total_row_count} rows</small>
                        </div>
                      ))}
                      {msg.syncError && <p>{msg.syncError}</p>}
                      {msg.chart && (
                        <img
                          className="chart-image"
                          src={`data:image/png;base64,${msg.chart}`}
                          alt="Generated chart"
                        />
                      )}
                    </div>
                  </motion.div>
                ))}
              </AnimatePresence>

              {/* Thinking indicator */}
              {isLoading && (
                <div className="thinking">
                  <div className="message-avatar">🤖</div>
                  <div className="thinking-dots">
                    <span />
                    <span />
                    <span />
                  </div>
                </div>
              )}

              <div ref={messagesEndRef} />
            </div>
          )}

          {/* Input */}
          <div className="input-area">
            <div className="input-wrapper">
              <input
                id="chat-input"
                type="text"
                placeholder={
                  files.length === 0
                    ? "Ask me to create a workbook, or upload one..."
                    : "Ask a question about your data..."
                }
                value={input}
                onChange={(e) => setInput(e.target.value)}
                onKeyDown={handleKeyDown}
                disabled={isLoading}
              />
              <button
                className="send-button"
                onClick={() => sendMessage()}
                disabled={!input.trim() || isLoading}
                aria-label="Send message"
              >
                ➤
              </button>
            </div>
          </div>
        </main>
      </div>

      {/* Spreadsheet Preview Modal */}
      {previewModal.isOpen && (
        <div className="preview-overlay" onClick={closePreview}>
          <div className="preview-modal" onClick={(e) => e.stopPropagation()}>
            <div className="preview-header">
              <div className="preview-header-left">
                <span>📊</span>
                <h3>{previewModal.filename}</h3>
                {previewModal.loading && <span className="preview-loading-tag">Loading…</span>}
              </div>
              <div className="preview-header-actions">
                <button
                  className="action-btn"
                  onClick={() => openPreview(previewModal.filename, previewModal.activeSheet, previewModal.offset, previewModal.limit, previewModal.range)}
                  title="Refresh data"
                >
                  🔄 Refresh
                </button>
                <button
                  className="preview-close-btn"
                  onClick={closePreview}
                  title="Close preview"
                >
                  ✕
                </button>
              </div>
            </div>

            {/* Sheet Tabs */}
            {previewModal.sheets.length > 1 && (
              <div className="preview-tabs">
                {previewModal.sheets.map((sheet) => (
                  <button
                    key={sheet}
                    className={`preview-tab ${previewModal.activeSheet === sheet ? "active" : ""}`}
                    onClick={() => openPreview(previewModal.filename, sheet, 0, previewModal.limit)}
                  >
                    {sheet}
                  </button>
                ))}
              </div>
            )}

            {/* Table Content */}
            <div className="preview-body">
              {previewModal.error ? (
                <div className="preview-state error">
                  ❌ {previewModal.error}
                </div>
              ) : previewModal.loading ? (
                <div className="preview-state">
                  Loading spreadsheet data...
                </div>
              ) : previewModal.columns.length === 0 ? (
                <div className="preview-state">
                  This sheet appears to be empty.
                </div>
              ) : (
                <table className="preview-table">
                  <thead>
                    <tr>
                      <th className="row-index">#</th>
                      {previewModal.columns.map((col, idx) => (
                        <th key={idx}>{String(col)}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {previewModal.data.map((row, rIdx) => (
                      <tr key={rIdx}>
                        <td className="row-index">
                          {previewModal.offset + rIdx + 1}
                        </td>
                        {previewModal.columns.map((col, cIdx) => (
                          <td key={cIdx}>{row[col] !== undefined && row[col] !== null ? String(row[col]) : ""}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>

            <div className="preview-footer">
              <div className="preview-footer-left">
                <span>Sheet: <strong>{previewModal.activeSheet || "N/A"}</strong></span>
                <span>
                  Showing rows {previewModal.totalRows > 0 ? previewModal.offset + 1 : 0}–{Math.min(previewModal.offset + previewModal.data.length, previewModal.totalRows)} of {previewModal.totalRows}
                </span>
              </div>
              <div className="preview-pagination">
                <button
                  className="preview-page-btn"
                  disabled={previewModal.offset === 0 || previewModal.loading}
                  onClick={() => openPreview(previewModal.filename, previewModal.activeSheet, Math.max(0, previewModal.offset - previewModal.limit), previewModal.limit, previewModal.range)}
                  title="Previous page"
                >
                  ◀ Prev
                </button>
                <span>
                  Page {Math.floor(previewModal.offset / previewModal.limit) + 1} of {Math.max(1, Math.ceil(previewModal.totalRows / previewModal.limit))}
                </span>
                <button
                  className="preview-page-btn"
                  disabled={previewModal.offset + previewModal.limit >= previewModal.totalRows || previewModal.loading}
                  onClick={() => openPreview(previewModal.filename, previewModal.activeSheet, previewModal.offset + previewModal.limit, previewModal.limit, previewModal.range)}
                  title="Next page"
                >
                  Next ▶
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
