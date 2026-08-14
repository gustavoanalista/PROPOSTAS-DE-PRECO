import { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ── IndexedDB helpers ─────────────────────────────────────
const DB_NAME = "propostaFarmaDB";
const DB_VERSION = 1;
const STORE = "planilha";

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => e.target.result.createObjectStore(STORE);
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror = () => reject(req.error);
  });
}
async function dbSave(key, value) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, "readwrite");
    tx.objectStore(STORE).put(value, key);
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
}
async function dbLoad(key) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, "readonly");
    const req = tx.objectStore(STORE).get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function dbDelete(key) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, "readwrite");
    tx.objectStore(STORE).delete(key);
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
}

// ── Formatters ────────────────────────────────────────────
const formatBRL = (v) =>
  typeof v === "number"
    ? v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })
    : "—";

const formatDate = (v) => {
  if (!v) return "—";
  if (typeof v === "number") {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    if (d.getFullYear() > 2400) return "S/V";
    return d.toLocaleDateString("pt-BR");
  }
  return String(v).substring(0, 10);
};

// ── Message parser ────────────────────────────────────────
function parseMessage(text) {
  const header = {};
  const products = [];
  const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);
  for (const line of lines) {
    const m = (pat) => { const r = line.match(pat); return r ? r[1].trim() : null; };
    if (/^COORD[:\s]/i.test(line)) header.coord = m(/^COORD[:\s]+(.+)/i);
    else if (/^SUP[:\s]/i.test(line)) header.sup = m(/^SUP[:\s]+(.+)/i);
    else if (/^RCA[:\s]/i.test(line)) {
      const raw = m(/^RCA[:\s]+(.+)/i);
      if (raw) { const parts = raw.match(/(\d+)\s+(.*)/); header.rcaCod = parts ? parts[1] : raw; header.rcaNome = parts ? parts[2] : ""; }
    } else if (/^RAZ[ÃA]O[:\s]/i.test(line)) header.razao = m(/^RAZ[ÃA]O[:\s]+(.+)/i);
    else if (/^CNPJ[:\s]/i.test(line)) header.cnpj = m(/^CNPJ[:\s]+(.+)/i);
    else if (/^C[OÓ]DIGO[:\s]/i.test(line)) header.codigoCliente = m(/^C[OÓ]DIGO[:\s]+(.+)/i);
    else if (/^DESCONTO[:\s]/i.test(line)) { const d = m(/^DESCONTO[:\s]+(.+)/i); header.desconto = d ? parseFloat(d.replace(",", ".").replace("%", "")) : null; }
  }
  const prodRegex = /^(\d{5,7})\s+(\d+)\s+([\d,\.]+)$/;
  for (const line of lines) {
    const match = line.match(prodRegex);
    if (match) products.push({ cod: parseInt(match[1], 10), qty: parseInt(match[2], 10), price: parseFloat(match[3].replace(",", ".")) });
  }
  if (products.length === 0) {
    for (const line of lines) {
      const parts = line.split(/\s{2,}|\t/).map((s) => s.trim()).filter(Boolean);
      if (parts.length === 3 && /^\d{5,7}$/.test(parts[0]) && /^\d+$/.test(parts[1]) && /^[\d,\.]+$/.test(parts[2]))
        products.push({ cod: parseInt(parts[0], 10), qty: parseInt(parts[1], 10), price: parseFloat(parts[2].replace(",", ".")) });
    }
  }
  return { header, products };
}

function buildLookup(workbook) {
  const sheet = workbook.Sheets["BASE PROPOSTA"];
  if (!sheet) return null;
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const cod = r[2];
    if (!cod) continue;
    map[parseInt(cod, 10)] = { descricao: r[7] || r[1] || "", estoque: r[3], custo: r[4], markup: r[5], multiplo: r[6] ? parseInt(r[6], 10) : 1, validade: r[9], categoria: r[0] || "" };
  }
  return map;
}

function calcRow(p, data) {
  const custo = data.custo * 1.025;
  const mkp = p.markup !== undefined ? p.markup : data.markup;
  const precoVenda = custo * (1 + mkp);
  const difUnid = p.price - precoVenda;
  const pctDesc = precoVenda !== 0 ? difUnid / precoVenda : 0;
  const valorVerba = difUnid * p.qty;
  const valorPedido = p.price * p.qty;
  const investimento = valorPedido !== 0 ? valorVerba / valorPedido : 0;
  return { ...p, descricao: data.descricao, estoque: data.estoque, custo, markup: mkp, precoVenda, difUnid, pctDesc, valorVerba, valorPedido, investimento, multiplo: data.multiplo, validade: data.validade, notFound: false };
}

function totals(rows) {
  const totalPedido = rows.reduce((s, r) => s + (r.valorPedido || 0), 0);
  const totalVerba = rows.reduce((s, r) => s + (r.valorVerba || 0), 0);
  const totalInvest = totalPedido !== 0 ? totalVerba / totalPedido : 0;
  return { totalPedido, totalVerba, totalInvest };
}

// ── EditField ─────────────────────────────────────────────
function EditField({ label, value, displayValue, color, onCommit }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState("");
  const inputRef = useRef();
  const start = () => { setDraft(String(value).replace(".", ",")); setEditing(true); setTimeout(() => inputRef.current?.focus(), 50); };
  const commit = () => { const parsed = parseFloat(draft.replace(",", ".")); if (!isNaN(parsed) && parsed > 0) onCommit(parsed); setEditing(false); };
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 0", borderBottom: "1px solid #1e2236" }}>
      <span style={{ fontSize: 12, color: "#7c84a0", flex: 1 }}>{label}</span>
      {editing ? (
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          <input ref={inputRef} value={draft} onChange={(e) => setDraft(e.target.value)} onBlur={commit}
            onKeyDown={(e) => { if (e.key === "Enter") commit(); if (e.key === "Escape") setEditing(false); }}
            style={{ width: 90, background: "#0f1117", border: "2px solid #2563eb", borderRadius: 8, color: "#fff", fontSize: 15, fontWeight: 700, padding: "6px 10px", outline: "none", textAlign: "right" }} />
          <button onClick={commit} style={{ background: "#2563eb", border: "none", borderRadius: 6, color: "#fff", padding: "6px 10px", fontSize: 13, cursor: "pointer" }}>✓</button>
        </div>
      ) : (
        <div onClick={start} style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
          <span style={{ fontSize: 15, fontWeight: 700, color: color || "#60a5fa" }}>{displayValue}</span>
          <span style={{ fontSize: 11, color: "#2563eb", background: "#1e2743", borderRadius: 4, padding: "2px 6px" }}>✏️</span>
        </div>
      )}
    </div>
  );
}

// ── ProductCard ───────────────────────────────────────────
function ProductCard({ row, index, onUpdatePrice, onUpdateMarkup, onUpdateQty, onRemove }) {
  const [expanded, setExpanded] = useState(false);
  if (row.notFound) return (
    <div style={{ background: "#1f1015", borderRadius: 14, padding: 16, marginBottom: 10, border: "1px solid #ef4444", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
      <div style={{ color: "#f87171", fontWeight: 700, fontSize: 14 }}>⚠️ Cód. {row.cod} — não encontrado</div>
      {onRemove && <button onClick={() => onRemove(index)} style={{ background: "none", border: "none", color: "#f87171", fontSize: 20, cursor: "pointer" }}>✕</button>}
    </div>
  );
  const isNeg = row.difUnid < 0;
  const descPct = Math.abs(row.pctDesc) * 100;
  const descColor = isNeg ? (descPct > 15 ? "#ef4444" : descPct > 8 ? "#f59e0b" : "#4ade80") : "#4ade80";
  const investColor = row.investimento < 0 ? "#ef4444" : "#4ade80";
  const statusBg = isNeg ? (descPct > 15 ? "#3b0a0a" : descPct > 8 ? "#2d1f00" : "#0a2010") : "#0a2010";
  const statusBorder = isNeg ? (descPct > 15 ? "#ef4444" : descPct > 8 ? "#f59e0b" : "#22c55e") : "#22c55e";
  return (
    <div style={{ background: "#1a1d27", borderRadius: 16, marginBottom: 12, border: `1px solid ${statusBorder}`, overflow: "hidden" }}>
      <div style={{ padding: "14px 16px", cursor: "pointer", background: statusBg }} onClick={() => setExpanded(!expanded)}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 6 }}>
          <div style={{ flex: 1, paddingRight: 8 }}>
            <div style={{ fontSize: 11, color: "#7c84a0", marginBottom: 2 }}>Cód. {row.cod}</div>
            <div style={{ fontSize: 14, fontWeight: 700, color: "#e8eaf0", lineHeight: 1.3 }}>{row.descricao || "—"}</div>
          </div>
          <div style={{ textAlign: "right", flexShrink: 0, display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 4 }}>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#60a5fa" }}>{row.price.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })}</div>
            <div style={{ fontSize: 12, color: "#7c84a0" }}>× {row.qty} un.</div>
            {onRemove && (
              <button onClick={(e) => { e.stopPropagation(); onRemove(index); }}
                style={{ background: "#2d0f0f", border: "1px solid #f87171", borderRadius: 6, color: "#f87171", fontSize: 11, padding: "2px 8px", cursor: "pointer", marginTop: 2 }}>
                remover
              </button>
            )}
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 4 }}>
          <span style={{ fontSize: 12, background: "#13151f", borderRadius: 6, padding: "3px 8px", color: descColor, fontWeight: 700 }}>{isNeg ? "" : "+"}{descPct.toFixed(1)}% desc.</span>
          <span style={{ fontSize: 12, background: "#13151f", borderRadius: 6, padding: "3px 8px", color: investColor, fontWeight: 600 }}>Invest. {(Math.abs(row.investimento) * 100).toFixed(1)}%</span>
          <span style={{ fontSize: 12, background: "#13151f", borderRadius: 6, padding: "3px 8px", color: row.estoque > 0 ? "#4ade80" : "#ef4444" }}>Estoque: {row.estoque ?? "—"}</span>
          <span style={{ fontSize: 12, background: "#13151f", borderRadius: 6, padding: "3px 8px", color: "#94a3b8", marginLeft: "auto" }}>{expanded ? "▲" : "▼"}</span>
        </div>
      </div>
      {expanded && (
        <div style={{ padding: "12px 16px" }}>
          <div style={{ background: "#13151f", borderRadius: 10, padding: "4px 14px", marginBottom: 12 }}>
            <EditField label="Preço promocional" value={row.price} displayValue={row.price.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })} color="#60a5fa" onCommit={(v) => onUpdatePrice(index, v)} />
            <EditField label="Quantidade" value={row.qty} displayValue={`${row.qty} un.`} color="#34d399" onCommit={(v) => onUpdateQty(index, Math.round(v))} />
            <EditField label="Markup %" value={(row.markup * 100).toFixed(2)} displayValue={(row.markup * 100).toFixed(2) + "%"} color="#c084fc" onCommit={(v) => onUpdateMarkup(index, v)} />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
            {[
              { label: "Custo", value: formatBRL(row.custo), color: "#94a3b8" },
              { label: "Preço de venda", value: formatBRL(row.precoVenda), color: "#94a3b8" },
              { label: "Dif. por unidade", value: formatBRL(row.difUnid), color: isNeg ? "#f87171" : "#4ade80" },
              { label: "Verba total", value: formatBRL(row.valorVerba), color: isNeg ? "#f87171" : "#4ade80" },
              { label: "Valor do pedido", value: formatBRL(row.valorPedido), color: "#e8eaf0" },
              { label: "Investimento", value: (Math.abs(row.investimento) * 100).toFixed(2) + "%", color: investColor },
              { label: "Múltiplo", value: row.multiplo ?? "—", color: "#94a3b8" },
              { label: "Validade", value: formatDate(row.validade), color: "#94a3b8" },
            ].map(({ label, value, color }) => (
              <div key={label} style={{ background: "#13151f", borderRadius: 10, padding: "10px 12px" }}>
                <div style={{ fontSize: 10, color: "#7c84a0", marginBottom: 3 }}>{label.toUpperCase()}</div>
                <div style={{ fontSize: 14, fontWeight: 700, color }}>{value}</div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ── TotalsBar ─────────────────────────────────────────────
function TotalsBar({ rows }) {
  const { totalPedido, totalVerba, totalInvest } = totals(rows);
  return (
    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 16 }}>
      {[
        { label: "PEDIDO", value: formatBRL(totalPedido), color: "#60a5fa", icon: "💰" },
        { label: "VERBA", value: formatBRL(totalVerba), color: totalVerba < 0 ? "#f87171" : "#4ade80", icon: "📉" },
        { label: "INVEST.", value: (Math.abs(totalInvest) * 100).toFixed(1) + "%", color: totalInvest < 0 ? "#f87171" : "#4ade80", icon: "📊" },
      ].map((item) => (
        <div key={item.label} style={{ background: "#1a1d27", borderRadius: 12, padding: "12px 8px", border: "1px solid #252836", textAlign: "center" }}>
          <div style={{ fontSize: 16, marginBottom: 2 }}>{item.icon}</div>
          <div style={{ fontSize: 9, color: "#7c84a0", letterSpacing: 1, marginBottom: 4 }}>{item.label}</div>
          <div style={{ fontSize: 13, fontWeight: 800, color: item.color }}>{item.value}</div>
        </div>
      ))}
    </div>
  );
}

// ── PropostaManualTab ─────────────────────────────────────
function PropostaManualTab({ lookup }) {
  const [rows, setRows] = useState([]);
  const [codInput, setCodInput] = useState("");
  const [qtyInput, setQtyInput] = useState("1");
  const [priceInput, setPriceInput] = useState("");
  const [searchResult, setSearchResult] = useState(null); // null | false | data
  const [addError, setAddError] = useState("");
  const codRef = useRef();

  if (!lookup) return (
    <div style={{ padding: 24, textAlign: "center" }}>
      <div style={{ fontSize: 40, marginBottom: 12 }}>📂</div>
      <div style={{ color: "#7c84a0", fontSize: 14 }}>Carregue a planilha de estoque na aba Proposta para usar esta funcionalidade.</div>
    </div>
  );

  const buscarProduto = () => {
    setAddError("");
    const cod = parseInt(codInput.trim(), 10);
    if (isNaN(cod)) { setSearchResult(false); return; }
    const data = lookup[cod];
    setSearchResult(data ? { cod, ...data } : false);
    if (data) {
      const custo = data.custo * 1.025;
      const precoVenda = custo * (1 + data.markup);
      setPriceInput(precoVenda.toFixed(2).replace(".", ","));
    }
  };

  const adicionarProduto = () => {
    setAddError("");
    if (!searchResult) { setAddError("Busque um produto primeiro."); return; }
    const qty = parseInt(qtyInput, 10);
    const price = parseFloat(priceInput.replace(",", "."));
    if (isNaN(qty) || qty <= 0) { setAddError("Quantidade inválida."); return; }
    if (isNaN(price) || price <= 0) { setAddError("Preço inválido."); return; }
    const data = lookup[searchResult.cod];
    const newRow = calcRow({ cod: searchResult.cod, qty, price }, data);
    setRows((prev) => [...prev, newRow]);
    // reset
    setCodInput(""); setQtyInput("1"); setPriceInput(""); setSearchResult(null);
    codRef.current?.focus();
  };

  const removeRow = (i) => setRows((prev) => prev.filter((_, idx) => idx !== i));

  const updateRowPrice = (i, v) => setRows((prev) => prev.map((row, idx) => idx !== i ? row : calcRow({ ...row, price: v }, lookup[row.cod])));
  const updateRowQty = (i, v) => setRows((prev) => prev.map((row, idx) => idx !== i ? row : calcRow({ ...row, qty: v }, lookup[row.cod])));
  const updateRowMarkup = (i, v) => setRows((prev) => prev.map((row, idx) => idx !== i ? row : calcRow({ ...row, markup: v / 100 }, lookup[row.cod])));

  const limparTudo = () => { setRows([]); setCodInput(""); setQtyInput("1"); setPriceInput(""); setSearchResult(null); setAddError(""); };

  return (
    <div style={{ padding: 16 }}>
      <div style={{ fontSize: 11, fontWeight: 700, color: "#7c84a0", letterSpacing: 1, marginBottom: 12 }}>PROPOSTA MANUAL</div>

      {/* Add product form */}
      <div style={{ background: "#1a1d27", borderRadius: 16, padding: 16, marginBottom: 16, border: "1px solid #252836" }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: "#7c84a0", marginBottom: 12 }}>ADICIONAR PRODUTO</div>

        {/* Código + busca */}
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <input
            ref={codRef}
            value={codInput}
            onChange={(e) => { setCodInput(e.target.value); setSearchResult(null); setAddError(""); }}
            onKeyDown={(e) => e.key === "Enter" && buscarProduto()}
            placeholder="Código do produto"
            inputMode="numeric"
            style={{ flex: 1, background: "#0f1117", border: "2px solid #252836", borderRadius: 10, color: "#e8eaf0", fontSize: 15, padding: "12px 14px", outline: "none", fontFamily: "monospace" }}
          />
          <button onClick={buscarProduto} style={{ background: "#2563eb", border: "none", borderRadius: 10, color: "#fff", fontWeight: 700, fontSize: 15, padding: "0 18px", cursor: "pointer" }}>
            🔍
          </button>
        </div>

        {/* Produto encontrado */}
        {searchResult === false && (
          <div style={{ background: "#2d0f0f", border: "1px solid #ef4444", borderRadius: 8, padding: "10px 14px", color: "#f87171", fontSize: 13, marginBottom: 10 }}>
            ⚠️ Produto não encontrado
          </div>
        )}
        {searchResult && searchResult.descricao && (
          <div style={{ background: "#0a2010", border: "1px solid #22c55e", borderRadius: 10, padding: "10px 14px", marginBottom: 12 }}>
            <div style={{ fontSize: 11, color: "#7c84a0", marginBottom: 2 }}>Cód. {searchResult.cod}</div>
            <div style={{ fontSize: 14, fontWeight: 700, color: "#e8eaf0", marginBottom: 4 }}>{searchResult.descricao}</div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <span style={{ fontSize: 11, color: "#94a3b8" }}>Custo: {formatBRL(searchResult.custo * 1.025)}</span>
              <span style={{ fontSize: 11, color: "#94a3b8" }}>Markup: {((searchResult.markup || 0) * 100).toFixed(1)}%</span>
              <span style={{ fontSize: 11, color: searchResult.estoque > 0 ? "#4ade80" : "#ef4444" }}>Estoque: {searchResult.estoque ?? "—"}</span>
            </div>
          </div>
        )}

        {/* Qty + Price */}
        {searchResult && (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 12 }}>
              <div>
                <div style={{ fontSize: 10, color: "#7c84a0", marginBottom: 4 }}>QUANTIDADE</div>
                <input
                  value={qtyInput}
                  onChange={(e) => setQtyInput(e.target.value)}
                  inputMode="numeric"
                  style={{ width: "100%", background: "#0f1117", border: "2px solid #252836", borderRadius: 10, color: "#34d399", fontSize: 16, fontWeight: 700, padding: "12px 14px", outline: "none", boxSizing: "border-box", textAlign: "center" }}
                />
              </div>
              <div>
                <div style={{ fontSize: 10, color: "#7c84a0", marginBottom: 4 }}>PREÇO PROMO (R$)</div>
                <input
                  value={priceInput}
                  onChange={(e) => setPriceInput(e.target.value)}
                  inputMode="decimal"
                  style={{ width: "100%", background: "#0f1117", border: "2px solid #2563eb", borderRadius: 10, color: "#60a5fa", fontSize: 16, fontWeight: 700, padding: "12px 14px", outline: "none", boxSizing: "border-box", textAlign: "center" }}
                />
              </div>
            </div>
            {addError && <div style={{ color: "#f87171", fontSize: 12, marginBottom: 8 }}>⚠️ {addError}</div>}
            <button onClick={adicionarProduto} style={{ width: "100%", background: "#16a34a", border: "none", borderRadius: 12, color: "#fff", fontWeight: 700, fontSize: 15, padding: "14px 0", cursor: "pointer" }}>
              ＋ Adicionar à proposta
            </button>
          </>
        )}
      </div>

      {/* Rows */}
      {rows.length > 0 && (
        <>
          <TotalsBar rows={rows} />
          <div style={{ fontSize: 11, fontWeight: 700, color: "#7c84a0", letterSpacing: 1, marginBottom: 10 }}>PRODUTOS ({rows.length})</div>
          {rows.map((row, i) => (
            <ProductCard key={i} row={row} index={i}
              onUpdatePrice={updateRowPrice} onUpdateQty={updateRowQty} onUpdateMarkup={updateRowMarkup}
              onRemove={removeRow}
            />
          ))}
          <button onClick={limparTudo} style={{ width: "100%", marginTop: 4, background: "#2d0f0f", color: "#f87171", border: "1px solid #f87171", borderRadius: 14, padding: "14px 0", fontWeight: 700, fontSize: 14, cursor: "pointer" }}>
            🗑️ Limpar proposta
          </button>
        </>
      )}
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("proposta");
  const [step, setStep] = useState("input");
  const [msgText, setMsgText] = useState("");
  const [lookup, setLookup] = useState(null);
  const [fileName, setFileName] = useState("");
  const [result, setResult] = useState(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [dbLoading, setDbLoading] = useState(true);

  useEffect(() => {
    (async () => {
      try {
        const saved = await dbLoad("lookup");
        const savedName = await dbLoad("fileName");
        if (saved) { setLookup(saved); setFileName(savedName || "planilha salva"); }
      } catch (e) {}
      finally { setDbLoading(false); }
    })();
  }, []);

  const handleFile = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true);
    const reader = new FileReader();
    reader.onload = async (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const map = buildLookup(wb);
        if (!map) throw new Error("Aba 'BASE PROPOSTA' não encontrada.");
        setLookup(map); setFileName(file.name); setError("");
        await dbSave("lookup", map); await dbSave("fileName", file.name);
      } catch (err) { setError("Erro ao ler o arquivo: " + err.message); }
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleRemovePlanilha = async () => {
    await dbDelete("lookup"); await dbDelete("fileName");
    setLookup(null); setFileName("");
  };

  const buildTotals = (header, rows) => {
    const { totalPedido, totalVerba, totalInvest } = totals(rows);
    setResult({ header, rows, totalPedido, totalVerba, totalInvest });
  };

  const handleAnalyse = () => {
    setError("");
    if (!msgText.trim()) { setError("Cole a mensagem do vendedor acima."); return; }
    if (!lookup) { setError("Faça o upload da planilha primeiro."); return; }
    const { header, products } = parseMessage(msgText);
    if (products.length === 0) { setError("Nenhum produto encontrado. Verifique o formato COD | QTD | PREÇO."); return; }
    const rows = products.map((p) => { const data = lookup[p.cod]; if (!data) return { ...p, notFound: true }; return calcRow(p, data); });
    buildTotals(header, rows);
    setStep("result");
  };

  const handleClear = () => { setMsgText(""); setResult(null); setError(""); setStep("input"); };
  const updatePrice = (i, v) => { const rows = result.rows.map((row, idx) => idx !== i || row.notFound ? row : calcRow({ ...row, price: v }, lookup[row.cod])); buildTotals(result.header, rows); };
  const updateQty = (i, v) => { const rows = result.rows.map((row, idx) => idx !== i || row.notFound ? row : calcRow({ ...row, qty: v }, lookup[row.cod])); buildTotals(result.header, rows); };
  const updateMarkup = (i, v) => { const rows = result.rows.map((row, idx) => idx !== i || row.notFound ? row : calcRow({ ...row, markup: v / 100 }, lookup[row.cod])); buildTotals(result.header, rows); };

  if (dbLoading) return (
    <div style={{ fontFamily: "'Inter','Segoe UI',sans-serif", background: "#0f1117", minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 12 }}>
      <div style={{ fontSize: 32 }}>⏳</div>
      <div style={{ color: "#7c84a0", fontSize: 14 }}>Carregando...</div>
    </div>
  );

  const tabBtn = (id, label, icon) => (
    <button onClick={() => setTab(id)} style={{ flex: 1, background: "none", border: "none", cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 3, padding: "10px 0", color: tab === id ? "#2563eb" : "#7c84a0" }}>
      <span style={{ fontSize: 22 }}>{icon}</span>
      <span style={{ fontSize: 10, fontWeight: 700, letterSpacing: 0.5 }}>{label}</span>
      {tab === id && <div style={{ width: 24, height: 3, background: "#2563eb", borderRadius: 2 }} />}
    </button>
  );

  return (
    <div style={{ fontFamily: "'Inter','Segoe UI',sans-serif", background: "#0f1117", minHeight: "100vh", color: "#e8eaf0", maxWidth: 480, margin: "0 auto", display: "flex", flexDirection: "column" }}>

      {/* Top bar */}
      <div style={{ background: "#1a1d27", borderBottom: "2px solid #2563eb", padding: "14px 16px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ background: "#2563eb", borderRadius: 7, padding: "5px 10px", fontWeight: 700, fontSize: 12, color: "#fff" }}>CD 331</div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 15 }}>Proposta Farma</div>
            <div style={{ fontSize: 11, color: "#7c84a0" }}>Grupo Mateus</div>
          </div>
        </div>
        {tab === "proposta" && step === "result" && (
          <button onClick={handleClear} style={{ background: "#2d0f0f", color: "#f87171", border: "1px solid #f87171", borderRadius: 8, padding: "8px 14px", fontWeight: 600, fontSize: 13, cursor: "pointer" }}>
            🗑️ Nova
          </button>
        )}
      </div>

      {/* Content */}
      <div style={{ flex: 1, overflowY: "auto", paddingBottom: 80 }}>

        {/* ══ PROPOSTA POR MENSAGEM ══ */}
        {tab === "proposta" && (
          <div style={{ padding: 16 }}>
            {step === "input" && (
              <div>
                <div style={{ marginBottom: 14 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#7c84a0", letterSpacing: 1, marginBottom: 8 }}>1. PLANILHA DE ESTOQUE</div>
                  {lookup ? (
                    <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "14px 16px", background: "#0a2010", borderRadius: 14, border: "2px solid #22c55e" }}>
                      <span style={{ fontSize: 28 }}>✅</span>
                      <div style={{ flex: 1 }}>
                        <div style={{ fontSize: 14, fontWeight: 600, color: "#22c55e" }}>Planilha carregada</div>
                        <div style={{ fontSize: 11, color: "#7c84a0" }}>{Object.keys(lookup).length.toLocaleString()} produtos · {fileName}</div>
                        <div style={{ fontSize: 10, color: "#4ade80", marginTop: 2 }}>💾 Salva no dispositivo</div>
                      </div>
                      <button onClick={handleRemovePlanilha} style={{ background: "transparent", border: "none", color: "#f87171", fontSize: 20, cursor: "pointer", padding: 4 }}>🗑️</button>
                    </div>
                  ) : (
                    <label style={{ display: "flex", alignItems: "center", gap: 12, padding: "14px 16px", background: "#1a1d27", borderRadius: 14, border: "2px dashed #252836", cursor: "pointer" }}>
                      <span style={{ fontSize: 28 }}>{loading ? "⏳" : "📂"}</span>
                      <div>
                        <div style={{ fontSize: 14, fontWeight: 600, color: "#e8eaf0" }}>{loading ? "Carregando..." : "Selecionar planilha .xlsx"}</div>
                        <div style={{ fontSize: 11, color: "#7c84a0" }}>ESTOQUE_CD_331.xlsx</div>
                      </div>
                      <input type="file" accept=".xlsx,.xlsm" onChange={handleFile} style={{ display: "none" }} />
                    </label>
                  )}
                </div>

                <div style={{ marginBottom: 14 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#7c84a0", letterSpacing: 1, marginBottom: 8 }}>2. MENSAGEM DO VENDEDOR</div>
                  <textarea value={msgText} onChange={(e) => setMsgText(e.target.value)}
                    placeholder={"Cole a mensagem do WhatsApp aqui...\n\nALTERAÇÃO MATEUS MAIS\nCOORD: ...\nRCA: 39595  Nome\n...\nCOD:    QTD    PREÇO\n148101  8      18,98"}
                    style={{ width: "100%", height: 220, background: "#1a1d27", border: "1px solid #252836", borderRadius: 14, color: "#e8eaf0", fontSize: 14, padding: 16, resize: "none", outline: "none", boxSizing: "border-box", fontFamily: "monospace", lineHeight: 1.7 }} />
                </div>

                {error && <div style={{ background: "#2d0f0f", border: "1px solid #ef4444", borderRadius: 10, padding: "12px 16px", color: "#f87171", marginBottom: 14, fontSize: 13 }}>⚠️ {error}</div>}

                <button onClick={handleAnalyse} style={{ width: "100%", background: "#2563eb", color: "#fff", border: "none", borderRadius: 14, padding: "18px 0", fontWeight: 700, fontSize: 17, cursor: "pointer", boxShadow: "0 4px 20px #2563eb44" }}>
                  🔍 Analisar Proposta
                </button>
              </div>
            )}

            {step === "result" && result && (
              <div>
                <div style={{ background: "#1a1d27", borderRadius: 14, padding: 16, marginBottom: 14, border: "1px solid #252836" }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#7c84a0", letterSpacing: 1, marginBottom: 10 }}>DADOS DA PROPOSTA</div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    {[["RCA", result.header.rcaCod ? `${result.header.rcaCod} · ${result.header.rcaNome}` : "—"], ["Coordenador", result.header.coord || "—"], ["Supervisor", result.header.sup || "—"], ["Desconto", result.header.desconto != null ? result.header.desconto + "%" : "—"]].map(([label, value]) => (
                      <div key={label} style={{ background: "#13151f", borderRadius: 10, padding: "8px 12px" }}>
                        <div style={{ fontSize: 10, color: "#7c84a0" }}>{label.toUpperCase()}</div>
                        <div style={{ fontSize: 13, fontWeight: 600, color: "#e8eaf0", marginTop: 2 }}>{value}</div>
                      </div>
                    ))}
                  </div>
                  {result.header.razao && (
                    <div style={{ background: "#13151f", borderRadius: 10, padding: "8px 12px", marginTop: 8 }}>
                      <div style={{ fontSize: 10, color: "#7c84a0" }}>CLIENTE</div>
                      <div style={{ fontSize: 13, fontWeight: 600, color: "#e8eaf0", marginTop: 2 }}>{result.header.razao}</div>
                      {result.header.cnpj && <div style={{ fontSize: 11, color: "#7c84a0" }}>{result.header.cnpj}</div>}
                    </div>
                  )}
                </div>

                <TotalsBar rows={result.rows} />

                <div style={{ fontSize: 11, fontWeight: 700, color: "#7c84a0", letterSpacing: 1, marginBottom: 10 }}>PRODUTOS ({result.rows.length})</div>
                {result.rows.map((row, i) => (
                  <ProductCard key={i} row={row} index={i} onUpdatePrice={updatePrice} onUpdateQty={updateQty} onUpdateMarkup={updateMarkup} />
                ))}

                <button onClick={handleClear} style={{ width: "100%", marginTop: 8, background: "#2d0f0f", color: "#f87171", border: "1px solid #f87171", borderRadius: 14, padding: "16px 0", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>
                  🗑️ Limpar e nova proposta
                </button>
              </div>
            )}
          </div>
        )}

        {/* ══ PROPOSTA MANUAL ══ */}
        {tab === "manual" && <PropostaManualTab lookup={lookup} />}
      </div>

      {/* Bottom nav */}
      <div style={{ position: "fixed", bottom: 0, left: "50%", transform: "translateX(-50%)", width: "100%", maxWidth: 480, background: "#1a1d27", borderTop: "1px solid #252836", display: "flex", zIndex: 10 }}>
        {tabBtn("proposta", "MENSAGEM", "💬")}
        {tabBtn("manual", "MANUAL", "✍️")}
      </div>
    </div>
  );
}
