
import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

/* ═══════════════════════════════════════════════════════════
   SECURITY LAYER  — enterprise-grade validation & sanitisation
   ═══════════════════════════════════════════════════════════ */
const SECURITY = {
  MAX_FILE_SIZE_MB: 10,
  MAX_QUESTIONS: 200,
  ALLOWED_EXTENSIONS: [".xlsx", ".xls"],
  ALLOWED_MIME: [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel",
  ],
  // Strip HTML / script injection from any cell value before display
  sanitize: (v) =>
    String(v ?? "")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/javascript:/gi, "")
      .trim()
      .slice(0, 2000),

  validateFile(file) {
    const ext = "." + file.name.split(".").pop().toLowerCase();
    if (!this.ALLOWED_EXTENSIONS.includes(ext))
      return { ok: false, msg: `File type not allowed. Use: ${this.ALLOWED_EXTENSIONS.join(", ")}` };
    if (!this.ALLOWED_MIME.includes(file.type) && file.type !== "")
      return { ok: false, msg: "MIME type not allowed." };
    if (file.size > this.MAX_FILE_SIZE_MB * 1024 * 1024)
      return { ok: false, msg: `File exceeds ${this.MAX_FILE_SIZE_MB} MB limit.` };
    return { ok: true };
  },
};

/* ═══════════════════════════════════════════════════════════
   EXCEL PARSER  — multi-sheet, auto-detect question columns
   ═══════════════════════════════════════════════════════════ */
function parseExcel(buffer) {
  const wb = XLSX.read(buffer, { type: "array" });
  const questions = [];

  wb.SheetNames.forEach((sheetName) => {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (rows.length < 2) return;

    const headers = rows[0].map((h) => String(h).trim());
    // Auto-detect: prefer a column whose header contains "question|query|q"
    let qIdx = headers.findIndex((h) => /question|query|\bq\b/i.test(h));
    if (qIdx < 0) qIdx = 0; // fallback: first column

    // Also collect a "context" column if present
    const ctxIdx = headers.findIndex((h) => /context|background|description/i.test(h));

    rows.slice(1).forEach((row, i) => {
      const raw = String(row[qIdx] ?? "").trim();
      if (!raw) return;
      questions.push({
        id: `${sheetName}::R${i + 2}`,
        sheet: sheetName,
        excelRow: i + 2,
        questionHeader: headers[qIdx] || "Column A",
        question: SECURITY.sanitize(raw),
        context: ctxIdx >= 0 ? SECURITY.sanitize(row[ctxIdx]) : "",
        status: "pending",   // pending | processing | done | error | skipped
        result: null,
        processedAt: null,
        durationMs: null,
      });
    });
  });

  return questions.slice(0, SECURITY.MAX_QUESTIONS);
}

/* ═══════════════════════════════════════════════════════════
   LOCAL SLM LAYER  — calls Anthropic API (swap for Ollama /
   LM Studio by changing ENDPOINT + headers in production)
   ═══════════════════════════════════════════════════════════ */
const SLM_CONFIG = {
  // In production: "http://localhost:11434/api/chat"  (Ollama)
  //                "http://localhost:1234/v1/chat/completions"  (LM Studio)
  ENDPOINT: "https://api.anthropic.com/v1/messages",
  MODEL: "claude-sonnet-4-20250514",
  MAX_TOKENS: 800,
  TIMEOUT_MS: 30_000,
  RETRY_ATTEMPTS: 2,
  RETRY_DELAY_MS: 1000,
};

const SYSTEM_PROMPT = `You are a precise enterprise Q&A engine connected to a local SLM.
For every question, respond ONLY with a valid JSON object — no markdown fences, no preamble.

Schema (all fields required):
{
  "answer": "<concise, factual answer — max 3 sentences>",
  "confidence": <float 0.0–1.0>,
  "category": "<one of: factual | procedural | analytical | opinion | unknown>",
  "reasoning": "<1–2 sentences explaining how you arrived at the answer>",
  "flags": {
    "needs_verification": <boolean>,
    "sensitive_content": <boolean>,
    "out_of_scope": <boolean>
  }
}`;

async function callSLM(question, context, signal) {
  const body = JSON.stringify({
    model: SLM_CONFIG.MODEL,
    max_tokens: SLM_CONFIG.MAX_TOKENS,
    system: SYSTEM_PROMPT,
    messages: [
      {
        role: "user",
        content: context
          ? `Context: ${context}\n\nQuestion: ${question}`
          : `Question: ${question}`,
      },
    ],
  });

  for (let attempt = 0; attempt <= SLM_CONFIG.RETRY_ATTEMPTS; attempt++) {
    if (signal?.aborted) throw new Error("Cancelled");
    try {
      const res = await fetch(SLM_CONFIG.ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body,
        signal,
      });
      if (!res.ok) {
        const err = await res.text().catch(() => res.statusText);
        throw new Error(`HTTP ${res.status}: ${err}`);
      }
      const data = await res.json();
      const raw = (data.content ?? []).find((b) => b.type === "text")?.text ?? "{}";
      const cleaned = raw.replace(/```json|```/g, "").trim();
      return JSON.parse(cleaned);
    } catch (e) {
      if (e.message === "Cancelled" || attempt === SLM_CONFIG.RETRY_ATTEMPTS) throw e;
      await new Promise((r) => setTimeout(r, SLM_CONFIG.RETRY_DELAY_MS * (attempt + 1)));
    }
  }
}

/* ═══════════════════════════════════════════════════════════
   EXCEL EXPORT  — structured output with metadata sheet
   ═══════════════════════════════════════════════════════════ */
function exportResults(items, fileInfo) {
  const answered = items.filter((i) => i.status === "done");
  const failed   = items.filter((i) => i.status === "error");

  // Sheet 1 — Answers
  const answerRows = items.map((item) => ({
    "Sheet":            item.sheet,
    "Excel Row":        item.excelRow,
    "Question Column":  item.questionHeader,
    "Question":         item.question,
    "Answer":           item.result?.answer ?? "",
    "Confidence (0-1)": item.result?.confidence ?? "",
    "Category":         item.result?.category ?? "",
    "Reasoning":        item.result?.reasoning ?? "",
    "Needs Verification": item.result?.flags?.needs_verification ? "Yes" : item.result ? "No" : "",
    "Sensitive Content":  item.result?.flags?.sensitive_content  ? "Yes" : item.result ? "No" : "",
    "Out of Scope":       item.result?.flags?.out_of_scope       ? "Yes" : item.result ? "No" : "",
    "Status":           item.status,
    "Duration (ms)":    item.durationMs ?? "",
    "Processed At":     item.processedAt ?? "",
  }));

  const ws1 = XLSX.utils.json_to_sheet(answerRows);
  ws1["!cols"] = [10, 10, 18, 45, 45, 18, 15, 45, 18, 16, 14, 12, 14, 22].map((w) => ({ wch: w }));

  // Sheet 2 — Raw JSON
  const jsonRows = items
    .filter((i) => i.result)
    .map((i) => ({
      "ID": i.id,
      "Question": i.question,
      "Raw JSON": JSON.stringify(i.result, null, 2),
    }));
  const ws2 = XLSX.utils.json_to_sheet(jsonRows);
  ws2["!cols"] = [25, 45, 80].map((w) => ({ wch: w }));

  // Sheet 3 — Metadata / Audit trail
  const meta = [
    { Key: "Source File",        Value: fileInfo?.name ?? "unknown" },
    { Key: "File Size",          Value: fileInfo?.size ?? "" },
    { Key: "Generated At",       Value: new Date().toISOString() },
    { Key: "Model",              Value: SLM_CONFIG.MODEL },
    { Key: "Total Questions",    Value: items.length },
    { Key: "Answered",           Value: answered.length },
    { Key: "Failed",             Value: failed.length },
    { Key: "Avg Confidence",     Value: answered.length
        ? (answered.reduce((s, i) => s + (i.result?.confidence ?? 0), 0) / answered.length).toFixed(3)
        : "N/A" },
    { Key: "Security: Max File Size", Value: `${SECURITY.MAX_FILE_SIZE_MB} MB` },
    { Key: "Security: Max Questions", Value: SECURITY.MAX_QUESTIONS },
  ];
  const ws3 = XLSX.utils.json_to_sheet(meta);
  ws3["!cols"] = [28, 50].map((w) => ({ wch: w }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, "AI Answers");
  XLSX.utils.book_append_sheet(wb, ws2, "Raw JSON");
  XLSX.utils.book_append_sheet(wb, ws3, "Audit Metadata");

  XLSX.writeFile(wb, `ai_answers_${Date.now()}.xlsx`);
}

/* ═══════════════════════════════════════════════════════════
   UI COMPONENTS
   ═══════════════════════════════════════════════════════════ */
const css = {
  // colour tokens
  bg:      "#07090f",
  surface: "#0d1117",
  border:  "#1e2a3a",
  accent:  "#00c8ff",
  green:   "#22d3a5",
  yellow:  "#f5c542",
  red:     "#ff4d6d",
  muted:   "#4a6070",
  text:    "#c8dce8",
  heading: "#e8f4fc",
};

function Badge({ status }) {
  const map = {
    pending:    { bg: "#151c28", c: css.muted,  label: "PENDING"     },
    processing: { bg: "#0a1e35", c: css.accent, label: "PROCESSING…" },
    done:       { bg: "#071a14", c: css.green,  label: "DONE"        },
    error:      { bg: "#1a0810", c: css.red,    label: "ERROR"       },
    skipped:    { bg: "#151c28", c: css.muted,  label: "SKIPPED"     },
  };
  const s = map[status] || map.pending;
  return (
    <span style={{
      background: s.bg, color: s.c, border: `1px solid ${s.c}33`,
      borderRadius: 3, padding: "2px 7px", fontSize: 9,
      fontFamily: "'JetBrains Mono', monospace", fontWeight: 700, letterSpacing: "0.08em",
    }}>{s.label}</span>
  );
}

function ConfBar({ val }) {
  if (val == null || val === "") return <span style={{ color: css.muted, fontSize: 11 }}>—</span>;
  const pct = Math.round(val * 100);
  const c = pct >= 80 ? css.green : pct >= 55 ? css.yellow : css.red;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <div style={{ flex: 1, height: 3, background: css.border, borderRadius: 99 }}>
        <div style={{ width: `${pct}%`, height: "100%", background: c, borderRadius: 99, transition: "width .5s" }} />
      </div>
      <span style={{ color: c, fontSize: 11, fontWeight: 700, minWidth: 30 }}>{pct}%</span>
    </div>
  );
}

function StatCard({ label, value, sub, color }) {
  return (
    <div style={{
      background: css.surface, border: `1px solid ${css.border}`,
      borderRadius: 10, padding: "18px 22px", minWidth: 110, flex: 1,
    }}>
      <div style={{ color: color || css.accent, fontSize: 28, fontWeight: 800, fontFamily: "'JetBrains Mono', monospace", lineHeight: 1 }}>{value}</div>
      <div style={{ color: css.muted, fontSize: 11, marginTop: 4, textTransform: "uppercase", letterSpacing: "0.07em" }}>{label}</div>
      {sub && <div style={{ color: css.text, fontSize: 11, marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════
   MAIN APP
   ═══════════════════════════════════════════════════════════ */
export default function App() {
  const [phase, setPhase]       = useState("upload");   // upload | review | running | done
  const [fileInfo, setFileInfo] = useState(null);
  const [items, setItems]       = useState([]);          // the question+result records
  const [error, setError]       = useState("");
  const [progress, setProgress] = useState({ done: 0, total: 0 });
  const [expanded, setExpanded] = useState(null);
  const [concurrency]           = useState(3);
  const abortRef                = useRef(null);
  const dropZoneRef             = useRef();

  /* ── drag-and-drop ── */
  const onDragOver  = (e) => { e.preventDefault(); dropZoneRef.current?.classList.add("dz-hover"); };
  const onDragLeave = ()  => dropZoneRef.current?.classList.remove("dz-hover");
  const onDrop      = (e) => { e.preventDefault(); onDragLeave(); ingestFile(e.dataTransfer.files[0]); };

  const ingestFile = useCallback(async (file) => {
    if (!file) return;
    setError("");
    const v = SECURITY.validateFile(file);
    if (!v.ok) return setError(`⚠ ${v.msg}`);

    const buf = await file.arrayBuffer();
    let qs;
    try { qs = parseExcel(buf); }
    catch (e) { return setError(`Parse error: ${e.message}`); }

    if (!qs.length)
      return setError("No questions found. Add a column with 'question' in its header.");

    setFileInfo({
      name: file.name,
      size: (file.size / 1024).toFixed(1) + " KB",
      sheets: [...new Set(qs.map((q) => q.sheet))],
    });
    setItems(qs);
    setPhase("review");
  }, []);

  /* ── agent runner ── */
  const runAgent = async () => {
    abortRef.current = new AbortController();
    setPhase("running");
    setProgress({ done: 0, total: items.length });

    const queue  = [...items];
    let doneCount = 0;

    const worker = async () => {
      while (queue.length && !abortRef.current.signal.aborted) {
        const item = queue.shift();

        setItems((prev) =>
          prev.map((i) => i.id === item.id ? { ...i, status: "processing" } : i)
        );

        const t0 = Date.now();
        try {
          const result = await callSLM(item.question, item.context, abortRef.current.signal);
          const ms = Date.now() - t0;
          setItems((prev) =>
            prev.map((i) =>
              i.id === item.id
                ? { ...i, status: "done", result, processedAt: new Date().toISOString(), durationMs: ms }
                : i
            )
          );
        } catch (e) {
          if (e.message === "Cancelled") break;
          const ms = Date.now() - t0;
          setItems((prev) =>
            prev.map((i) =>
              i.id === item.id
                ? { ...i, status: "error",
                    result: { answer: e.message, confidence: 0, category: "error",
                              reasoning: "API call failed", flags: {} },
                    processedAt: new Date().toISOString(), durationMs: ms }
                : i
            )
          );
        }

        doneCount++;
        setProgress({ done: doneCount, total: items.length });
      }
    };

    await Promise.all(Array.from({ length: concurrency }, worker));
    setPhase("done");
  };

  const cancel = () => abortRef.current?.abort();
  const reset  = () => { setPhase("upload"); setFileInfo(null); setItems([]); setError(""); setExpanded(null); };

  /* ── derived stats ── */
  const done       = items.filter((i) => i.status === "done");
  const failed     = items.filter((i) => i.status === "error");
  const pct        = progress.total ? Math.round((progress.done / progress.total) * 100) : 0;
  const avgConf    = done.length
    ? (done.reduce((s, i) => s + (i.result?.confidence ?? 0), 0) / done.length * 100).toFixed(0) + "%"
    : "—";

  /* ── inject font ── */
  useEffect(() => {
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;700&family=Syne:wght@700;800&family=DM+Sans:wght@400;500&display=swap";
    document.head.appendChild(link);

    const style = document.createElement("style");
    style.textContent = `
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body { background: ${css.bg}; }
      ::-webkit-scrollbar { width: 6px; } ::-webkit-scrollbar-track { background: ${css.bg}; }
      ::-webkit-scrollbar-thumb { background: ${css.border}; border-radius: 3px; }
      .dz { border: 2px dashed ${css.border}; border-radius: 14px; transition: all .25s; cursor: pointer; }
      .dz:hover, .dz-hover { border-color: ${css.accent}; background: #0a1a2a; }
      .btn { cursor: pointer; border: none; border-radius: 8px; font-family: 'DM Sans', sans-serif;
             font-weight: 500; transition: opacity .15s, transform .1s; }
      .btn:hover { opacity: .88; }  .btn:active { transform: scale(.97); }
      .row-card { border-bottom: 1px solid ${css.border}; transition: background .15s; }
      .row-card:hover { background: #0d1a26; }
      .tag { background: #0a1a2a; border: 1px solid ${css.border}; border-radius: 4px;
             color: ${css.muted}; font-size: 11px; padding: 2px 8px; }
    `;
    document.head.appendChild(style);
  }, []);

  /* ════════════════════════════════════════════════════
     RENDER
     ════════════════════════════════════════════════════ */
  return (
    <div style={{ minHeight: "100vh", background: css.bg, color: css.text,
                  fontFamily: "'DM Sans', sans-serif", fontSize: 14 }}>

      {/* ── Header ── */}
      <div style={{ borderBottom: `1px solid ${css.border}`, padding: "18px 32px",
                    display: "flex", alignItems: "center", gap: 14,
                    background: `linear-gradient(90deg, #0d1117 0%, #070d18 100%)` }}>
        <div style={{ width: 32, height: 32, borderRadius: 8,
                      background: `linear-gradient(135deg, ${css.accent}, #0066ff)`,
                      display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>⚡</div>
        <div>
          <div style={{ color: css.heading, fontFamily: "'Syne', sans-serif",
                        fontWeight: 800, fontSize: 17, letterSpacing: "-0.02em" }}>
            Excel Q&A Agent
          </div>
          <div style={{ color: css.muted, fontSize: 11 }}>
            Local SLM · Structured JSON · Enterprise MVP
          </div>
        </div>

        {/* Pipeline steps */}
        <div style={{ marginLeft: "auto", display: "flex", gap: 6 }}>
          {["Upload", "Review", "Process", "Export"].map((s, i) => {
            const phaseIdx = { upload: 0, review: 1, running: 2, done: 3 }[phase];
            const active   = i === phaseIdx;
            const past     = i < phaseIdx;
            return (
              <div key={s} style={{ display: "flex", alignItems: "center", gap: 4 }}>
                {i > 0 && <div style={{ width: 16, height: 1, background: past ? css.accent : css.border }} />}
                <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                  <div style={{ width: 20, height: 20, borderRadius: "50%",
                                background: active ? css.accent : past ? "#0a3a55" : css.border,
                                color: active ? css.bg : past ? css.accent : css.muted,
                                fontSize: 10, fontWeight: 700, display: "flex",
                                alignItems: "center", justifyContent: "center" }}>
                    {past ? "✓" : i + 1}
                  </div>
                  <span style={{ color: active ? css.heading : css.muted, fontSize: 11,
                                 fontWeight: active ? 600 : 400 }}>{s}</span>
                </div>
              </div>
            );
          })}
        </div>
      </div>

      <div style={{ maxWidth: 980, margin: "0 auto", padding: "32px 24px" }}>

        {/* ── ERROR BANNER ── */}
        {error && (
          <div style={{ background: "#180810", border: `1px solid ${css.red}44`,
                        borderRadius: 8, padding: "12px 16px", color: css.red,
                        marginBottom: 24, fontSize: 13 }}>
            {error}
          </div>
        )}

        {/* ══════════════════════════════════════════
            PHASE: UPLOAD
            ══════════════════════════════════════════ */}
        {phase === "upload" && (
          <div>
            {/* drop zone */}
            <div
              ref={dropZoneRef}
              className="dz"
              onDragOver={onDragOver}
              onDragLeave={onDragLeave}
              onDrop={onDrop}
              onClick={() => document.getElementById("_fi").click()}
              style={{ padding: "56px 24px", textAlign: "center", marginBottom: 32 }}
            >
              <input id="_fi" type="file" accept=".xlsx,.xls" hidden
                onChange={(e) => ingestFile(e.target.files[0])} />
              <div style={{ fontSize: 40, marginBottom: 12 }}>📊</div>
              <div style={{ color: css.heading, fontFamily: "'Syne', sans-serif",
                            fontWeight: 700, fontSize: 18, marginBottom: 6 }}>
                Drop your Excel file here
              </div>
              <div style={{ color: css.muted, fontSize: 13 }}>
                or click to browse · .xlsx / .xls · max {SECURITY.MAX_FILE_SIZE_MB} MB
              </div>
            </div>

            {/* Architecture overview */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 16 }}>
              {[
                { icon: "🔒", title: "Security Layer", desc: "File-type & MIME validation, size limits, cell sanitisation, injection prevention" },
                { icon: "🧠", title: "Local SLM Engine", desc: "Pluggable model endpoint (Ollama / LM Studio / Anthropic). Structured JSON schema, retry logic, concurrency control" },
                { icon: "📤", title: "Structured Export", desc: "Multi-sheet Excel: Answers + Raw JSON + Audit metadata. Ready for BI / downstream pipelines" },
              ].map((c) => (
                <div key={c.title} style={{ background: css.surface, border: `1px solid ${css.border}`,
                                            borderRadius: 12, padding: "20px 18px" }}>
                  <div style={{ fontSize: 26, marginBottom: 10 }}>{c.icon}</div>
                  <div style={{ color: css.heading, fontWeight: 600, fontSize: 13, marginBottom: 6 }}>{c.title}</div>
                  <div style={{ color: css.muted, fontSize: 12, lineHeight: 1.6 }}>{c.desc}</div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ══════════════════════════════════════════
            PHASE: REVIEW
            ══════════════════════════════════════════ */}
        {phase === "review" && (
          <div>
            {/* file card */}
            <div style={{ background: css.surface, border: `1px solid ${css.border}`,
                          borderRadius: 10, padding: "16px 20px", marginBottom: 24,
                          display: "flex", alignItems: "center", gap: 16 }}>
              <div style={{ fontSize: 28 }}>📁</div>
              <div style={{ flex: 1 }}>
                <div style={{ color: css.heading, fontWeight: 600 }}>{fileInfo?.name}</div>
                <div style={{ color: css.muted, fontSize: 12, marginTop: 2 }}>
                  {fileInfo?.size} · Sheets: {fileInfo?.sheets?.join(", ")} · {items.length} questions detected
                </div>
              </div>
              <button className="btn" onClick={reset}
                style={{ background: css.border, color: css.muted, padding: "8px 16px", fontSize: 12 }}>
                Replace file
              </button>
            </div>

            {/* question preview */}
            <div style={{ color: css.heading, fontWeight: 600, marginBottom: 12,
                          display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span>Questions preview</span>
              <span className="tag">{items.length} / {SECURITY.MAX_QUESTIONS} max</span>
            </div>

            <div style={{ background: css.surface, border: `1px solid ${css.border}`,
                          borderRadius: 10, overflow: "hidden", marginBottom: 24 }}>
              {items.slice(0, 8).map((item, i) => (
                <div key={item.id} style={{
                  padding: "12px 18px", borderBottom: i < Math.min(items.length, 8) - 1 ? `1px solid ${css.border}` : "none",
                  display: "grid", gridTemplateColumns: "auto 1fr auto", gap: 12, alignItems: "center",
                }}>
                  <span style={{ color: css.muted, fontSize: 11, fontFamily: "monospace",
                                 minWidth: 50 }}>
                    {item.sheet} R{item.excelRow}
                  </span>
                  <span style={{ color: css.text, fontSize: 13 }}>{item.question}</span>
                  <Badge status="pending" />
                </div>
              ))}
              {items.length > 8 && (
                <div style={{ padding: "10px 18px", color: css.muted, fontSize: 12, textAlign: "center" }}>
                  + {items.length - 8} more questions…
                </div>
              )}
            </div>

            <div style={{ display: "flex", gap: 12 }}>
              <button className="btn" onClick={runAgent}
                style={{ flex: 1, background: `linear-gradient(90deg, ${css.accent}, #0066ff)`,
                         color: "#fff", padding: "14px 24px", fontSize: 14, fontWeight: 600 }}>
                ⚡ Run Agent  ({items.length} questions · {concurrency} concurrent)
              </button>
            </div>
          </div>
        )}

        {/* ══════════════════════════════════════════
            PHASE: RUNNING
            ══════════════════════════════════════════ */}
        {phase === "running" && (
          <div>
            {/* progress header */}
            <div style={{ background: css.surface, border: `1px solid ${css.border}`,
                          borderRadius: 10, padding: "20px 24px", marginBottom: 24 }}>
              <div style={{ display: "flex", justifyContent: "space-between",
                            alignItems: "center", marginBottom: 10 }}>
                <span style={{ color: css.heading, fontWeight: 600 }}>Processing…</span>
                <span style={{ color: css.accent, fontFamily: "monospace",
                               fontSize: 20, fontWeight: 700 }}>{pct}%</span>
              </div>
              <div style={{ height: 6, background: css.border, borderRadius: 99, overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${pct}%`, borderRadius: 99, transition: "width .4s",
                              background: `linear-gradient(90deg, ${css.accent}, #0066ff)` }} />
              </div>
              <div style={{ color: css.muted, fontSize: 12, marginTop: 8 }}>
                {progress.done} of {progress.total} · {concurrency} workers · retries enabled
              </div>
              <button className="btn" onClick={cancel}
                style={{ marginTop: 12, background: css.border, color: css.red,
                         padding: "8px 18px", fontSize: 12 }}>
                ✕ Cancel
              </button>
            </div>

            {/* live list */}
            <div style={{ background: css.surface, border: `1px solid ${css.border}`,
                          borderRadius: 10, overflow: "hidden" }}>
              {items.map((item, i) => (
                <div key={item.id} className="row-card"
                  style={{ padding: "12px 18px",
                           display: "grid", gridTemplateColumns: "auto 1fr auto auto", gap: 12, alignItems: "center" }}>
                  <span style={{ color: css.muted, fontSize: 11, fontFamily: "monospace", minWidth: 50 }}>
                    R{item.excelRow}
                  </span>
                  <div>
                    <div style={{ color: css.text, fontSize: 13, marginBottom: 2 }}>{item.question}</div>
                    {item.result?.answer && (
                      <div style={{ color: css.muted, fontSize: 11 }}>
                        {item.result.answer.slice(0, 90)}{item.result.answer.length > 90 ? "…" : ""}
                      </div>
                    )}
                  </div>
                  <ConfBar val={item.result?.confidence} />
                  <Badge status={item.status} />
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ══════════════════════════════════════════
            PHASE: DONE
            ══════════════════════════════════════════ */}
        {phase === "done" && (
          <div>
            {/* stats row */}
            <div style={{ display: "flex", gap: 14, marginBottom: 24, flexWrap: "wrap" }}>
              <StatCard label="Total"     value={items.length}  color={css.accent} />
              <StatCard label="Answered"  value={done.length}   color={css.green}  />
              <StatCard label="Failed"    value={failed.length} color={failed.length ? css.red : css.muted} />
              <StatCard label="Avg Conf"  value={avgConf}       color={css.yellow} />
            </div>

            {/* actions */}
            <div style={{ display: "flex", gap: 12, marginBottom: 24 }}>
              <button className="btn" onClick={() => exportResults(items, fileInfo)}
                style={{ flex: 1, background: `linear-gradient(90deg, ${css.green}, #00997a)`,
                         color: "#fff", padding: "14px 24px", fontSize: 14, fontWeight: 600 }}>
                ⬇ Download Excel Report
              </button>
              <button className="btn" onClick={reset}
                style={{ background: css.border, color: css.text, padding: "14px 20px", fontSize: 14 }}>
                ↩ Start Over
              </button>
            </div>

            {/* results table */}
            <div style={{ background: css.surface, border: `1px solid ${css.border}`,
                          borderRadius: 10, overflow: "hidden" }}>
              {items.map((item) => (
                <div key={item.id} className="row-card">
                  {/* row header */}
                  <div
                    style={{ padding: "14px 18px", cursor: "pointer",
                             display: "grid", gridTemplateColumns: "auto 1fr 140px auto auto", gap: 12, alignItems: "center" }}
                    onClick={() => setExpanded(expanded === item.id ? null : item.id)}
                  >
                    <span style={{ color: css.muted, fontSize: 11, fontFamily: "monospace", minWidth: 50 }}>
                      {item.sheet} R{item.excelRow}
                    </span>
                    <span style={{ color: css.text, fontSize: 13 }}>{item.question}</span>
                    <ConfBar val={item.result?.confidence} />
                    <span className="tag">{item.result?.category ?? "—"}</span>
                    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                      <Badge status={item.status} />
                      <span style={{ color: css.muted, fontSize: 14 }}>
                        {expanded === item.id ? "▲" : "▼"}
                      </span>
                    </div>
                  </div>

                  {/* expanded detail */}
                  {expanded === item.id && item.result && (
                    <div style={{ borderTop: `1px solid ${css.border}`,
                                  background: "#090e18", padding: "16px 20px" }}>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                        <div>
                          <div style={{ color: css.muted, fontSize: 11, marginBottom: 4, textTransform: "uppercase" }}>Answer</div>
                          <div style={{ color: css.heading, fontSize: 13, lineHeight: 1.6 }}>
                            {item.result.answer}
                          </div>
                        </div>
                        <div>
                          <div style={{ color: css.muted, fontSize: 11, marginBottom: 4, textTransform: "uppercase" }}>Reasoning</div>
                          <div style={{ color: css.text, fontSize: 13, lineHeight: 1.6 }}>
                            {item.result.reasoning}
                          </div>
                        </div>
                      </div>

                      <div style={{ marginTop: 14, display: "flex", gap: 10, flexWrap: "wrap" }}>
                        {item.result.flags?.needs_verification && (
                          <span className="tag" style={{ color: css.yellow, borderColor: css.yellow + "55" }}>⚠ Needs Verification</span>
                        )}
                        {item.result.flags?.sensitive_content && (
                          <span className="tag" style={{ color: css.red, borderColor: css.red + "55" }}>🔒 Sensitive</span>
                        )}
                        {item.result.flags?.out_of_scope && (
                          <span className="tag" style={{ color: css.muted }}>⊘ Out of Scope</span>
                        )}
                        {item.durationMs && (
                          <span className="tag">{item.durationMs} ms</span>
                        )}
                        {item.processedAt && (
                          <span className="tag">{new Date(item.processedAt).toLocaleTimeString()}</span>
                        )}
                      </div>

                      {/* raw JSON */}
                      <details style={{ marginTop: 12 }}>
                        <summary style={{ color: css.muted, fontSize: 11, cursor: "pointer" }}>
                          Raw JSON
                        </summary>
                        <pre style={{ marginTop: 8, background: css.bg, borderRadius: 6, padding: "10px 14px",
                                      color: css.green, fontSize: 11, overflow: "auto",
                                      fontFamily: "'JetBrains Mono', monospace" }}>
                          {JSON.stringify(item.result, null, 2)}
                        </pre>
                      </details>
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}