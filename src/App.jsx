import { useState, useEffect } from "react";

const C = {
  bg:     "#07090D",
  s1:     "#0C1018",
  s2:     "#101620",
  border: "#1A2535",
  cyan:   "#00D4FF",
  purple: "#9333EA",
  green:  "#10D98A",
  yellow: "#F5A623",
  red:    "#FF4060",
  muted:  "#3D5468",
  text:   "#8BA8C0",
  head:   "#E2EEF8",
};

const MONO = "'JetBrains Mono', monospace";
const SANS = "'Outfit', sans-serif";

// ── Sample Excel text representation (what the AI actually sees) ──
const EXCEL_TEXT_SAMPLE = `============================================================
SHEET: Customer Survey Q3
Dimensions: 12 rows × 4 columns
Merged ranges: A1:D1, A2:D2
============================================================
  [A1]=Customer Satisfaction Survey — Q3 2024
  [A3]=Section  |  [B3]=Question  |  [C3]=Response  |  [D3]=Score
  [A4]=Product  |  [B4]=How satisfied are you with product quality?  |  [C4]=Very Satisfied  |  [D4]=5
  [A5]=Product  |  [B5]=Would you recommend us?  |  [C5]=Yes  |  [D5]=4
  [A6]=Support  |  [B6]=How was your support experience?  |  [C6]=Good  |  [D6]=4
  [A7]=Support  |  [B7]=Response time acceptable?  |  [C7]=No  |  [D7]=2
  [A8]=Pricing  |  [B8]=Is pricing fair for the value provided?  |  [C8]=Neutral  |  [D8]=3
  [A10]=Comments  |  [B10]=Please leave any additional feedback

PANDAS VIEW (raw rows, no header assumed)
============================================================
--- Sheet: Customer Survey Q3 ---
     0         1                                      2               3
0  NaN  Customer Satisfaction Survey — Q3 2024      NaN            NaN
2  Section   Question                               Response        Score
3  Product   How satisfied are you with quality?   Very Satisfied  5
...`;

// ── Sample AI JSON output ──
const AI_JSON_SAMPLE = {
  file_type: "survey",
  file_purpose: "Customer satisfaction survey for Q3 2024 with product, support, and pricing sections",
  sheets_summary: [{ sheet: "Customer Survey Q3", purpose: "Survey responses with section, question, response, and score columns" }],
  structure_notes: "Merged header rows A1:D1 and A2:D2. Data starts at row 3 with headers. Sections: Product, Support, Pricing.",
  items: [
    { id: "Q1", sheet: "Customer Survey Q3", cell_ref: "B4", question_or_task: "How satisfied are you with product quality?", answer: "Very Satisfied (Score: 5/5)", data_source: "file_data", confidence: 1.0, category: "lookup", flags: { needs_human_review: false, ambiguous: false, data_missing: false } },
    { id: "Q2", sheet: "Customer Survey Q3", cell_ref: "B7", question_or_task: "Response time acceptable?", answer: "No (Score: 2/5) — support response time was rated below acceptable", data_source: "file_data", confidence: 1.0, category: "lookup", flags: { needs_human_review: true, ambiguous: false, data_missing: false } },
    { id: "Q3", sheet: "Customer Survey Q3", cell_ref: "B10", question_or_task: "Please leave any additional feedback", answer: "No response recorded in cell C10", data_source: "file_data", confidence: 0.9, category: "lookup", flags: { needs_human_review: false, ambiguous: false, data_missing: true } },
  ],
  summary: { total_items: 5, answered: 5, needs_review: 1, avg_confidence: 0.96 },
};

// ── Pipeline steps ──
const N8N_STEPS = [
  { icon: "📡", color: C.cyan,   label: "Webhook",        sub: "POST multipart/form-data",        detail: "n8n receives the .xlsx binary via webhook. The binary data is stored in $binary.data and is immediately available to subsequent nodes." },
  { icon: "🔒", color: C.red,    label: "Security Gate",  sub: "Code node — ext, size, MIME",     detail: "JS Code node validates extension (.xlsx/.xls), file size (<10MB), generates jobId. Throws error and stops workflow on violation." },
  { icon: "🔄", color: C.yellow, label: "Parser Service", sub: "HTTP POST binary → text",         detail: "n8n cannot parse Excel natively in code. An HTTP Request node POSTs the binary to the Excel Parser microservice (Python/FastAPI on port 8001) which returns text_repr." },
  { icon: "🧩", color: C.purple, label: "Build Prompt",   sub: "Code node — inject text into LLM prompt", detail: "Code node takes text_repr from the parser and injects it into the Ollama system+user prompt. Truncates to 12,000 chars if needed." },
  { icon: "🧠", color: C.cyan,   label: "Ollama Call",    sub: "HTTP POST → http://ollama:11434", detail: "HTTP Request node POSTs the full prompt payload to Ollama's /api/chat. timeout: 120s, retry: 3×. Model reads the raw text — no column assumptions." },
  { icon: "✅", color: C.green,  label: "Parse & Validate", sub: "Code node — JSON + schema check", detail: "Code node strips markdown fences, JSON.parse(), validates items array, normalises confidence to 0–1, flags missing fields." },
  { icon: "📤", color: C.green,  label: "Export Excel",   sub: "HTTP → Excel Exporter service",   detail: "POSTs the validated analysis JSON to the Excel Exporter microservice which creates a multi-sheet .xlsx with answers, raw JSON, and audit trail." },
];

const LG_STEPS = [
  { icon: "📡", color: C.cyan,   label: "FastAPI /analyse", sub: "POST multipart upload",        detail: "FastAPI endpoint receives bytes, validates file, writes to /tmp, queues async background job, returns job_id immediately (202 Accepted)." },
  { icon: "🔒", color: C.red,    label: "validate_file()",  sub: "Python — ext + size check",    detail: "Checks extension, MIME, file size. HTTPException raised on violation — no LangGraph graph is entered." },
  { icon: "🔵", color: C.purple, label: "node_extract_text", sub: "LangGraph node 1",            detail: "excel_to_text_full() uses openpyxl to iterate every non-empty cell with its reference, plus a pandas .to_string() view. Both views merged into one text blob." },
  { icon: "🧠", color: C.cyan,   label: "node_analyse_with_ai", sub: "LangGraph node 2",        detail: "langchain-ollama sends the full text + system prompt to Ollama. The LLM reads raw content — no pandas headers, no column guessing. Returns raw text." },
  { icon: "✅", color: C.green,  label: "node_parse_validate", sub: "LangGraph node 3",         detail: "JSON parse, strip fences, find first '{'. Pydantic-validates schema. Clamps confidence 0–1. Returns SLMAnswer or sets error for retry." },
  { icon: "🔄", color: C.yellow, label: "Conditional Edge", sub: "should_retry() — up to 2×",   detail: "If parse fails: retry_count < 2 → back to node_analyse_with_ai. On 3rd failure or success → END. State carries retry_count." },
  { icon: "📤", color: C.green,  label: "export_to_excel()", sub: "openpyxl — 5 sheets",        detail: "Writes AI Analysis, File Understanding, Sheet Summaries, Raw JSON, Audit tabs. FileResponse streams to client." },
];

function PipelineRow({ steps, color, label }) {
  const [active, setActive] = useState(null);
  return (
    <div style={{ marginBottom: 8 }}>
      <div style={{ color, fontSize: 12, fontWeight: 700, fontFamily: MONO,
                    marginBottom: 10, letterSpacing: "0.05em" }}>{label}</div>
      <div style={{ display: "flex", alignItems: "flex-start", gap: 0, overflowX: "auto", paddingBottom: 4 }}>
        {steps.map((s, i) => (
          <div key={i} style={{ display: "flex", alignItems: "center" }}>
            <div
              onClick={() => setActive(active === i ? null : i)}
              style={{
                minWidth: 110, cursor: "pointer", textAlign: "center",
                background: active === i ? s.color + "18" : C.s2,
                border: `1.5px solid ${active === i ? s.color : C.border}`,
                borderRadius: 10, padding: "10px 8px",
                transition: "all .2s",
              }}
            >
              <div style={{ fontSize: 20, marginBottom: 5 }}>{s.icon}</div>
              <div style={{ color: active === i ? s.color : C.head, fontSize: 11,
                            fontWeight: 600, lineHeight: 1.3, marginBottom: 3 }}>{s.label}</div>
              <div style={{ color: C.muted, fontSize: 9, fontFamily: MONO, lineHeight: 1.3 }}>{s.sub}</div>
            </div>
            {i < steps.length - 1 && (
              <div style={{ color: C.muted, fontSize: 18, padding: "0 4px", marginTop: -8 }}>→</div>
            )}
          </div>
        ))}
      </div>
      {active !== null && (
        <div style={{
          marginTop: 10, background: steps[active].color + "0C",
          border: `1px solid ${steps[active].color}33`,
          borderRadius: 8, padding: "12px 16px",
          color: C.text, fontSize: 12, lineHeight: 1.7,
        }}>
          <span style={{ color: steps[active].color, fontWeight: 700 }}>
            {steps[active].label}:{" "}
          </span>
          {steps[active].detail}
        </div>
      )}
    </div>
  );
}

function CodeBlock({ children, label, color }) {
  return (
    <div style={{ background: "#050810", border: `1px solid ${color || C.border}33`,
                  borderRadius: 8, overflow: "hidden" }}>
      {label && (
        <div style={{ background: (color || C.cyan) + "14", padding: "6px 14px",
                      color: color || C.cyan, fontSize: 10, fontFamily: MONO,
                      fontWeight: 700, letterSpacing: "0.06em", borderBottom: `1px solid ${color || C.border}22` }}>
          {label}
        </div>
      )}
      <pre style={{ padding: "14px", fontFamily: MONO, fontSize: 11, color: C.green,
                    overflowX: "auto", margin: 0, lineHeight: 1.65, whiteSpace: "pre-wrap",
                    wordBreak: "break-word" }}>
        {children}
      </pre>
    </div>
  );
}

export default function App() {
  const [tab, setTab] = useState("how");

  useEffect(() => {
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;700;800&family=JetBrains+Mono:wght@400;700&display=swap";
    document.head.appendChild(link);
    const s = document.createElement("style");
    s.textContent = `* { box-sizing: border-box; margin: 0; padding: 0; }
    body { background: ${C.bg}; }
    ::-webkit-scrollbar { width: 5px; height: 5px; }
    ::-webkit-scrollbar-track { background: ${C.bg}; }
    ::-webkit-scrollbar-thumb { background: ${C.border}; border-radius: 3px; }
    .tab-btn { cursor: pointer; border: none; font-family: '${SANS}'; font-size: 13px;
               font-weight: 600; transition: all .2s; border-radius: 7px; padding: 9px 20px; }
    .tab-btn:hover { opacity: .9; }`;
    document.head.appendChild(s);
  }, []);

  const tabs = [
    { id: "how",      label: "💡 How It Works"    },
    { id: "pipeline", label: "🔀 Pipelines"       },
    { id: "text",     label: "📄 Excel → Text"    },
    { id: "output",   label: "✅ AI JSON Output"  },
    { id: "guide",    label: "🚀 Setup Guide"     },
  ];

  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.text, fontFamily: SANS, fontSize: 14 }}>

      {/* Header */}
      <div style={{ borderBottom: `1px solid ${C.border}`, padding: "18px 28px",
                    background: "linear-gradient(180deg, #0E1520 0%, #07090D 100%)" }}>
        <div style={{ maxWidth: 1000, margin: "0 auto" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 18 }}>
            <div style={{ width: 38, height: 38, borderRadius: 10, fontSize: 20,
                          background: `linear-gradient(135deg, ${C.cyan}33, ${C.purple}33)`,
                          border: `1px solid ${C.cyan}44`,
                          display: "flex", alignItems: "center", justifyContent: "center" }}>⚡</div>
            <div>
              <div style={{ color: C.head, fontWeight: 800, fontSize: 20, letterSpacing: "-0.025em" }}>
                Unstructured Excel → AI Analysis
              </div>
              <div style={{ color: C.muted, fontSize: 12 }}>
                How n8n and LangGraph read ANY Excel file — no column assumptions · Ollama local SLM
              </div>
            </div>
          </div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {tabs.map(t => (
              <button key={t.id} className="tab-btn"
                style={{ background: tab === t.id ? C.cyan + "18" : "transparent",
                         color: tab === t.id ? C.cyan : C.muted,
                         border: `1.5px solid ${tab === t.id ? C.cyan + "55" : C.border}` }}
                onClick={() => setTab(t.id)}>
                {t.label}
              </button>
            ))}
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1000, margin: "0 auto", padding: "28px 20px" }}>

        {/* ── HOW IT WORKS ── */}
        {tab === "how" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
            {/* The core idea */}
            <div style={{ background: `linear-gradient(135deg, ${C.cyan}0A, ${C.purple}0A)`,
                          border: `1px solid ${C.cyan}22`, borderRadius: 14, padding: "22px 26px" }}>
              <div style={{ color: C.head, fontWeight: 800, fontSize: 18, marginBottom: 12 }}>
                The Core Idea
              </div>
              <div style={{ color: C.text, fontSize: 14, lineHeight: 1.8 }}>
                Instead of trying to <em style={{ color: C.yellow }}>parse the structure first</em> and
                then extract questions, we <strong style={{ color: C.cyan }}>dump the entire Excel file as text</strong> —
                including every cell reference, value, merged range, and sheet — and let the AI figure out
                the structure itself. This works on forms, surveys, audits, reports, and totally irregular layouts.
              </div>
            </div>

            {/* 3 steps */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 }}>
              {[
                { n: "01", icon: "📊", color: C.cyan,   title: "Excel → Text",
                  body: "openpyxl reads every non-empty cell with its exact reference ([B4]=text). Merged ranges are listed. Pandas adds a second raw-row view. Both merged into one text blob." },
                { n: "02", icon: "🧠", color: C.purple, title: "Text → Ollama",
                  body: "The full text is sent to Ollama with a system prompt that says: find every question/task/item, answer it using data in the file or your knowledge, return structured JSON." },
                { n: "03", icon: "📤", color: C.green,  title: "JSON → Excel",
                  body: "The validated JSON response is written back as a clean Excel report: Answers sheet + Raw JSON + Audit metadata. No data is lost." },
              ].map(c => (
                <div key={c.n} style={{ background: C.s1, border: `1px solid ${C.border}`,
                                        borderRadius: 12, padding: "18px 16px" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                    <div style={{ color: c.color, fontFamily: MONO, fontSize: 11, fontWeight: 700,
                                  background: c.color + "18", border: `1px solid ${c.color}44`,
                                  borderRadius: 4, padding: "2px 8px" }}>{c.n}</div>
                    <div style={{ fontSize: 20 }}>{c.icon}</div>
                  </div>
                  <div style={{ color: C.head, fontWeight: 700, fontSize: 14, marginBottom: 8 }}>{c.title}</div>
                  <div style={{ color: C.text, fontSize: 12, lineHeight: 1.65 }}>{c.body}</div>
                </div>
              ))}
            </div>

            {/* Why this works */}
            <div style={{ background: C.s1, border: `1px solid ${C.border}`, borderRadius: 12, padding: "18px 22px" }}>
              <div style={{ color: C.head, fontWeight: 700, marginBottom: 12 }}>
                Why this works for unstructured files
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                {[
                  ["✓ No header assumptions", "AI sees [A4]=Question  [B4]=Answer — understands context from content, not column position"],
                  ["✓ Handles merged cells", "openpyxl exposes merge ranges (A1:D1) which tells AI this is likely a title or section header"],
                  ["✓ Multi-sheet aware", "Each sheet is labelled and separated — AI understands which sheet holds which data"],
                  ["✓ Mixed layouts", "A file with a table on Sheet1 and a form on Sheet2 is handled in one pass — AI sees both"],
                ].map(([t, d]) => (
                  <div key={t} style={{ display: "flex", gap: 10 }}>
                    <span style={{ color: C.green, flexShrink: 0 }}>✓</span>
                    <div>
                      <div style={{ color: C.head, fontSize: 12, fontWeight: 600 }}>{t.replace("✓ ", "")}</div>
                      <div style={{ color: C.text, fontSize: 12, lineHeight: 1.55 }}>{d}</div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* ── PIPELINES ── */}
        {tab === "pipeline" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>
            <div style={{ color: C.muted, fontSize: 12 }}>
              Click any node to see what it does in detail.
            </div>
            <div style={{ background: C.s1, border: `1px solid ${C.border}`, borderRadius: 12, padding: "20px" }}>
              <PipelineRow steps={N8N_STEPS} color={C.cyan} label="n8n WORKFLOW — NODES" />
            </div>
            <div style={{ background: C.s1, border: `1px solid ${C.border}`, borderRadius: 12, padding: "20px" }}>
              <PipelineRow steps={LG_STEPS} color={C.purple} label="LANGGRAPH — STATE GRAPH NODES" />
            </div>

            {/* Key difference */}
            <div style={{ background: C.s2, border: `1px solid ${C.yellow}22`, borderRadius: 10, padding: "16px 20px" }}>
              <div style={{ color: C.yellow, fontWeight: 700, marginBottom: 8 }}>
                ⚡ Key Difference for Unstructured Files
              </div>
              <div style={{ color: C.text, fontSize: 13, lineHeight: 1.7 }}>
                <strong style={{ color: C.head }}>n8n</strong> cannot parse .xlsx binary natively in a Code node.
                You need the <code style={{ color: C.cyan, fontFamily: MONO, background: C.s1, padding: "1px 5px", borderRadius: 3 }}>excel_parser_service.py</code> microservice running on port 8001.
                The HTTP Request node POSTs the binary and receives <code style={{ color: C.cyan, fontFamily: MONO }}>text_repr</code>.
                <br /><br />
                <strong style={{ color: C.head }}>LangGraph</strong> handles it in pure Python —
                openpyxl and pandas are called directly inside the <code style={{ color: C.purple, fontFamily: MONO }}>node_extract_text</code> LangGraph node.
                No extra microservice needed.
              </div>
            </div>
          </div>
        )}

        {/* ── EXCEL → TEXT ── */}
        {tab === "text" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <div style={{ color: C.text, fontSize: 13, lineHeight: 1.7 }}>
              This is what the AI actually receives — a plain text dump of the entire Excel file.
              Cell references, sheet names, merged areas, and all values are preserved.
              <strong style={{ color: C.cyan }}> The AI has no idea what "column C" means</strong> — it reads content like a human would read printed text.
            </div>
            <CodeBlock label="WHAT THE AI SEES (text_repr)" color={C.cyan}>
              {EXCEL_TEXT_SAMPLE}
            </CodeBlock>
            <div style={{ background: C.s1, border: `1px solid ${C.border}`, borderRadius: 10, padding: "16px 20px" }}>
              <div style={{ color: C.head, fontWeight: 700, marginBottom: 10 }}>
                Conversion code (Python — works in both n8n service and LangGraph node)
              </div>
              <CodeBlock color={C.purple}>{`wb = openpyxl.load_workbook(path, data_only=True)

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    
    # Merged cells → structural hint for AI
    if ws.merged_cells.ranges:
        merged = [str(r) for r in ws.merged_cells.ranges]
        output.append(f"Merged ranges: {', '.join(merged)}")
    
    # Every non-empty cell with its reference
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
        cells = []
        for col_idx, value in enumerate(row, 1):
            if value is not None and str(value).strip():
                col = openpyxl.utils.get_column_letter(col_idx)
                cells.append(f"[{col}{row_idx}]={value}")
        if cells:
            output.append("  " + "  |  ".join(cells))

# Also add pandas view (different formatting catches more)
df = pd.read_excel(path, header=None, dtype=str)
output.append(df.fillna("").to_string(index=True))`}
              </CodeBlock>
            </div>
          </div>
        )}

        {/* ── AI JSON OUTPUT ── */}
        {tab === "output" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <div style={{ color: C.text, fontSize: 13, lineHeight: 1.7 }}>
              The AI returns a structured JSON object with file understanding, discovered items,
              answers, confidence scores, data sources (file vs general knowledge), and flags for human review.
            </div>

            {/* Items preview */}
            <div style={{ background: C.s1, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden" }}>
              {[["Key", "Value", "bg"], ["file_type", AI_JSON_SAMPLE.file_type, false],
                ["file_purpose", AI_JSON_SAMPLE.file_purpose, false],
                ["structure_notes", AI_JSON_SAMPLE.structure_notes, false]].map(([k, v, hdr], i) => (
                <div key={i} style={{ display: "grid", gridTemplateColumns: "160px 1fr",
                                      padding: "10px 18px", borderBottom: `1px solid ${C.border}`,
                                      background: hdr === "bg" ? C.s2 : "transparent" }}>
                  <span style={{ color: C.muted, fontFamily: MONO, fontSize: 11, fontWeight: 700 }}>{k}</span>
                  <span style={{ color: C.head, fontSize: 12 }}>{String(v)}</span>
                </div>
              ))}
            </div>

            <div style={{ color: C.head, fontWeight: 700, marginBottom: 4 }}>Discovered Items</div>
            {AI_JSON_SAMPLE.items.map((item) => (
              <div key={item.id} style={{ background: C.s1, border: `1px solid ${C.border}`,
                                          borderRadius: 10, padding: "14px 18px" }}>
                <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 8 }}>
                  <span style={{ color: C.cyan, fontFamily: MONO, fontSize: 10, fontWeight: 700,
                                 background: C.cyan + "18", border: `1px solid ${C.cyan}33`,
                                 borderRadius: 4, padding: "2px 7px" }}>{item.id}</span>
                  <span style={{ color: C.muted, fontFamily: MONO, fontSize: 10 }}>{item.cell_ref}</span>
                  <span style={{ color: C.muted, fontFamily: MONO, fontSize: 10 }}>{item.category}</span>
                  <span style={{ color: item.confidence >= 0.9 ? C.green : C.yellow,
                                 fontFamily: MONO, fontSize: 10, marginLeft: "auto" }}>
                    {Math.round(item.confidence * 100)}% conf
                  </span>
                </div>
                <div style={{ color: C.text, fontSize: 12, marginBottom: 6 }}>
                  <strong style={{ color: C.yellow }}>Q: </strong>{item.question_or_task}
                </div>
                <div style={{ color: C.head, fontSize: 13 }}>
                  <strong style={{ color: C.green }}>A: </strong>{item.answer}
                </div>
                {(item.flags.needs_human_review || item.flags.data_missing) && (
                  <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
                    {item.flags.needs_human_review && (
                      <span style={{ color: C.yellow, background: C.yellow + "18",
                                     border: `1px solid ${C.yellow}33`, borderRadius: 4,
                                     fontSize: 10, padding: "2px 8px", fontFamily: MONO }}>⚠ Needs Review</span>
                    )}
                    {item.flags.data_missing && (
                      <span style={{ color: C.red, background: C.red + "18",
                                     border: `1px solid ${C.red}33`, borderRadius: 4,
                                     fontSize: 10, padding: "2px 8px", fontFamily: MONO }}>⊘ Data Missing</span>
                    )}
                  </div>
                )}
              </div>
            ))}
          </div>
        )}

        {/* ── SETUP GUIDE ── */}
        {tab === "guide" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            {[
              {
                step: "1", color: C.cyan, title: "Start Ollama + pull model",
                code: `# docker-compose.yml already handles this — or manually:
docker run -d --name ollama -p 11434:11434 ollama/ollama
docker exec ollama ollama pull mistral:7b
# Verify:
curl http://localhost:11434/api/tags`,
              },
              {
                step: "2", color: C.purple, title: "Run the Excel Parser microservice",
                code: `# Required by n8n (LangGraph uses it as a library)
pip install fastapi uvicorn openpyxl pandas python-multipart
uvicorn excel_parser_service:app --port 8001
# Verify:
curl http://localhost:8001/health`,
              },
              {
                step: "3a", color: C.cyan, title: "n8n — Import workflow",
                code: `# Start n8n
docker run -d -p 5678:5678 n8nio/n8n
# Open http://localhost:5678
# Import: unstructured_excel_workflow.json
# Set Ollama URL in HTTP node → http://ollama:11434
# Set Parser URL → http://excel-parser:8001
# Activate workflow → test with:
curl -X POST http://localhost:5678/webhook/analyse-excel \\
  -F "data=@your_file.xlsx"`,
              },
              {
                step: "3b", color: C.purple, title: "LangGraph — Run the API",
                code: `pip install -r requirements.txt
uvicorn unstructured_agent:app --port 8000

# Submit a job:
curl -X POST http://localhost:8000/analyse \\
  -F "file=@your_file.xlsx"
# → { "job_id": "abc-123", "status": "queued" }

# Poll status:
curl http://localhost:8000/jobs/abc-123

# Download result:
curl http://localhost:8000/jobs/abc-123/download -o result.xlsx`,
              },
              {
                step: "4", color: C.green, title: "Full stack with Docker Compose",
                code: `docker compose up -d
# Starts: Ollama, ollama-init (model pull), 
#         excel-parser:8001, langgraph-agent:8000, n8n:5678
# All on same Docker network — service names as hostnames`,
              },
            ].map(s => (
              <div key={s.step} style={{ background: C.s1, border: `1px solid ${C.border}`,
                                          borderRadius: 12, overflow: "hidden" }}>
                <div style={{ padding: "12px 18px", borderBottom: `1px solid ${C.border}`,
                              display: "flex", gap: 10, alignItems: "center",
                              background: s.color + "0A" }}>
                  <span style={{ color: s.color, fontFamily: MONO, fontSize: 11, fontWeight: 700,
                                 background: s.color + "18", border: `1px solid ${s.color}44`,
                                 borderRadius: 4, padding: "2px 8px" }}>STEP {s.step}</span>
                  <span style={{ color: C.head, fontWeight: 600 }}>{s.title}</span>
                </div>
                <CodeBlock color={s.color}>{s.code}</CodeBlock>
              </div>
            ))}
          </div>
        )}

      </div>
    </div>
  );
}
