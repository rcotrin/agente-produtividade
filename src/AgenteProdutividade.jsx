import { useState, useCallback, useMemo } from 'react';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, ScatterChart, Scatter, ZAxis,
  RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis,
} from 'recharts';

// ─── palette ────────────────────────────────────────────────────────────────
const C = {
  bg: '#0d1117', surface: '#161b22', border: '#30363d',
  primary: '#58a6ff', success: '#3fb950', warn: '#d29922',
  danger: '#f85149', muted: '#8b949e', text: '#e6edf3',
  purple: '#bc8cff', cyan: '#39d353', orange: '#f0883e',
};
const CHART_COLORS = [C.primary, C.success, C.warn, C.danger, C.purple, C.cyan, C.orange, '#a5d6ff'];

// ─── helpers ─────────────────────────────────────────────────────────────────
function isWorkDay(d) {
  const dow = d.getDay();
  return dow !== 0 && dow !== 6;
}

function calcHorasUteis(inicio, fim, pausas = []) {
  if (!inicio || !fim) return 0;
  const start = new Date(inicio);
  const end = new Date(fim);
  if (end <= start) return 0;

  let totalMs = 0;
  const cur = new Date(start);

  while (cur < end) {
    if (isWorkDay(cur)) {
      const dayStart = new Date(cur);
      dayStart.setHours(9, 0, 0, 0);
      const dayEnd = new Date(cur);
      dayEnd.setHours(18, 0, 0, 0);

      const segStart = cur < dayStart ? dayStart : new Date(cur);
      const segEnd = end < dayEnd ? new Date(end) : dayEnd;

      if (segEnd > segStart) {
        let ms = segEnd - segStart;

        // subtract pausa periods that overlap this segment
        for (const p of pausas) {
          const ps = new Date(p.inicio);
          const pe = new Date(p.fim);
          const os = ps < segStart ? segStart : ps;
          const oe = pe > segEnd ? segEnd : pe;
          if (oe > os) ms -= (oe - os);
        }
        if (ms > 0) totalMs += ms;
      }
    }
    cur.setDate(cur.getDate() + 1);
    cur.setHours(9, 0, 0, 0);
  }
  return totalMs / 3_600_000;
}

function extrairPausas(updates) {
  const pausas = [];
  let pausaInicio = null;
  for (const u of updates) {
    const label = (u.fields?.['System.BoardColumn'] || '').toLowerCase();
    const ts = u.fields?.['System.ChangedDate'];
    if (!ts) continue;
    if ((label.includes('pausa') || label.includes('bloqueio')) && !pausaInicio) {
      pausaInicio = ts;
    } else if (pausaInicio && !label.includes('pausa') && !label.includes('bloqueio')) {
      pausas.push({ inicio: pausaInicio, fim: ts });
      pausaInicio = null;
    }
  }
  return pausas;
}

function horasToStr(h) {
  const hrs = Math.floor(h);
  const min = Math.round((h - hrs) * 60);
  return `${hrs}h${min > 0 ? ` ${min}m` : ''}`;
}

function fmtDate(iso) {
  if (!iso) return '—';
  return new Date(iso).toLocaleDateString('pt-BR');
}

function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

// ─── Azure DevOps API ────────────────────────────────────────────────────────
function azHeaders(pat) {
  return {
    Authorization: `Basic ${btoa(`:${pat}`)}`,
    'Content-Type': 'application/json',
  };
}

async function azFetch(pat, url) {
  const res = await fetch(url, { headers: azHeaders(pat) });
  if (!res.ok) throw new Error(`Azure DevOps ${res.status}: ${await res.text()}`);
  return res.json();
}

async function fetchProjects(org, pat) {
  const data = await azFetch(pat, `/devops/${org}/_apis/projects?api-version=7.1`);
  return data.value || [];
}

async function fetchIterations(org, project, pat) {
  const data = await azFetch(pat, `/devops/${org}/${project}/_apis/work/teamsettings/iterations?api-version=7.1`);
  return data.value || [];
}

async function fetchTeamMembers(org, project, pat) {
  try {
    const data = await azFetch(pat, `/devops/${org}/_apis/projects/${project}/teams?api-version=7.1`);
    const teams = data.value || [];
    if (!teams.length) return [];
    const first = teams[0].id;
    const mem = await azFetch(pat, `/devops/${org}/_apis/projects/${project}/teams/${first}/members?api-version=7.1`);
    return (mem.value || []).map(m => m.identity?.displayName).filter(Boolean);
  } catch {
    return [];
  }
}

async function fetchWorkItems(org, project, pat, filters) {
  const conditions = [`[System.TeamProject] = '${project}'`];
  if (filters.iterationPath) conditions.push(`[System.IterationPath] = '${filters.iterationPath}'`);
  if (filters.assignedTo) conditions.push(`[System.AssignedTo] = '${filters.assignedTo}'`);
  if (filters.dateFrom) conditions.push(`[System.ChangedDate] >= '${filters.dateFrom}'`);
  if (filters.dateTo) conditions.push(`[System.ChangedDate] <= '${filters.dateTo}'`);

  const wiql = {
    query: `SELECT [System.Id] FROM WorkItems WHERE ${conditions.join(' AND ')} ORDER BY [System.ChangedDate] DESC`,
  };

  const res = await fetch(`/devops/${org}/${project}/_apis/wit/wiql?api-version=7.1`, {
    method: 'POST',
    headers: azHeaders(pat),
    body: JSON.stringify(wiql),
  });
  if (!res.ok) throw new Error(`WIQL ${res.status}: ${await res.text()}`);
  const data = await res.json();
  const ids = (data.workItems || []).map(w => w.id);
  if (!ids.length) return [];

  const fields = [
    'System.Id', 'System.Title', 'System.WorkItemType', 'System.State',
    'System.AssignedTo', 'System.BoardColumn', 'System.CreatedDate',
    'System.ChangedDate', 'Microsoft.VSTS.Common.ActivatedDate',
    'Microsoft.VSTS.Common.ResolvedDate', 'Microsoft.VSTS.Common.ClosedDate',
    'Microsoft.VSTS.Scheduling.OriginalEstimate',
    'Microsoft.VSTS.Scheduling.CompletedWork',
    'Microsoft.VSTS.Scheduling.RemainingWork',
  ].join(',');

  const chunks = chunk(ids, 200);
  const items = [];
  for (const ch of chunks) {
    const data2 = await azFetch(pat, `/devops/${org}/_apis/wit/workitems?ids=${ch.join(',')}&fields=${fields}&api-version=7.1`);
    items.push(...(data2.value || []));
  }
  return items;
}

async function fetchUpdates(org, project, pat, id) {
  const data = await azFetch(pat, `/devops/${org}/${project}/_apis/wit/workitems/${id}/updates?api-version=7.1`);
  return data.value || [];
}

// ─── sub-components ───────────────────────────────────────────────────────────
function KpiCard({ label, value, sub, color = C.primary }) {
  return (
    <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: '16px 20px', minWidth: 140 }}>
      <div style={{ fontSize: 12, color: C.muted, marginBottom: 4 }}>{label}</div>
      <div style={{ fontSize: 26, fontWeight: 700, color }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

function Section({ title, children }) {
  return (
    <div style={{ marginBottom: 32 }}>
      <h2 style={{ fontSize: 15, fontWeight: 600, color: C.muted, marginBottom: 12, textTransform: 'uppercase', letterSpacing: 1 }}>{title}</h2>
      {children}
    </div>
  );
}

function Tag({ children, color = C.muted }) {
  return (
    <span style={{ background: `${color}22`, color, border: `1px solid ${color}55`, borderRadius: 4, padding: '2px 6px', fontSize: 11, fontWeight: 600 }}>
      {children}
    </span>
  );
}

const TYPE_COLOR = {
  'Epic': C.purple, 'Feature': C.primary, 'User Story': C.success,
  'Task': C.cyan, 'Bug': C.danger, 'Test Case': C.orange,
};

function stateColor(s) {
  if (!s) return C.muted;
  const l = s.toLowerCase();
  if (l.includes('done') || l.includes('closed') || l.includes('resolved')) return C.success;
  if (l.includes('active') || l.includes('progress')) return C.primary;
  if (l.includes('blocked') || l.includes('pausa')) return C.danger;
  return C.warn;
}

const CustomTooltip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: '8px 12px', fontSize: 12 }}>
      {label && <div style={{ color: C.muted, marginBottom: 4 }}>{label}</div>}
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color || C.text }}>{p.name}: <strong>{typeof p.value === 'number' ? p.value.toFixed(1) : p.value}</strong></div>
      ))}
    </div>
  );
};

// ─── main component ───────────────────────────────────────────────────────────
export default function AgenteProdutividade() {
  // credentials / config
  const [azOrg, setAzOrg] = useState('');
  const [azProject, setAzProject] = useState('');
  const [azPat, setAzPat] = useState('');
  const [anthropicKey, setAnthropicKey] = useState('');

  // remote data
  const [projects, setProjects] = useState([]);
  const [iterations, setIterations] = useState([]);
  const [members, setMembers] = useState([]);

  // filters
  const [filterDateFrom, setFilterDateFrom] = useState('');
  const [filterDateTo, setFilterDateTo] = useState('');
  const [filterIteration, setFilterIteration] = useState('');
  const [filterMember, setFilterMember] = useState('');

  // work data
  const [workItems, setWorkItems] = useState([]);
  const [updatesMap, setUpdatesMap] = useState({}); // id → updates[]
  const [metricsMap, setMetricsMap] = useState({}); // id → {leadTime, cycleTime, horasUteis, pausas, interacoes, boardTimes}

  // ui
  const [loading, setLoading] = useState(false);
  const [loadingMsg, setLoadingMsg] = useState('');
  const [error, setError] = useState('');
  const [tab, setTab] = useState('dashboard');
  const [aiReport, setAiReport] = useState('');
  const [aiLoading, setAiLoading] = useState(false);
  const [configOpen, setConfigOpen] = useState(true);

  // ── connect: load projects/iters/members ───────────────────────────────────
  const handleConnect = useCallback(async () => {
    if (!azOrg || !azPat) return setError('Informe organização e PAT.');
    setError(''); setLoading(true); setLoadingMsg('Conectando ao Azure DevOps…');
    try {
      const projs = await fetchProjects(azOrg, azPat);
      setProjects(projs);
      if (azProject) {
        const [iters, mems] = await Promise.all([
          fetchIterations(azOrg, azProject, azPat),
          fetchTeamMembers(azOrg, azProject, azPat),
        ]);
        setIterations(iters);
        setMembers(mems);
      }
      setConfigOpen(false);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, [azOrg, azProject, azPat]);

  const handleProjectChange = useCallback(async (proj) => {
    setAzProject(proj);
    if (!azOrg || !azPat || !proj) return;
    try {
      const [iters, mems] = await Promise.all([
        fetchIterations(azOrg, proj, azPat),
        fetchTeamMembers(azOrg, proj, azPat),
      ]);
      setIterations(iters);
      setMembers(mems);
    } catch { /* ignore */ }
  }, [azOrg, azPat]);

  // ── fetch & compute ────────────────────────────────────────────────────────
  const handleFetch = useCallback(async () => {
    if (!azOrg || !azProject || !azPat) return setError('Configure a conexão antes de buscar dados.');
    setError(''); setLoading(true); setWorkItems([]); setUpdatesMap({}); setMetricsMap({}); setAiReport('');

    try {
      setLoadingMsg('Buscando work items…');
      const filters = { iterationPath: filterIteration, assignedTo: filterMember, dateFrom: filterDateFrom, dateTo: filterDateTo };
      const items = await fetchWorkItems(azOrg, azProject, azPat, filters);
      setWorkItems(items);

      const uMap = {};
      const mMap = {};
      const total = items.length;

      for (let i = 0; i < items.length; i++) {
        const wi = items[i];
        const id = wi.id;
        setLoadingMsg(`Carregando histórico ${i + 1}/${total}…`);

        let updates = [];
        try { updates = await fetchUpdates(azOrg, azProject, azPat, id); } catch { /* skip */ }
        uMap[id] = updates;

        // board column times
        const boardTimes = {};
        let prevCol = null, prevTs = null;
        for (const u of updates) {
          const col = u.fields?.['System.BoardColumn']?.newValue;
          const ts = u.revisedDate || u.fields?.['System.ChangedDate']?.newValue;
          if (col && ts) {
            if (prevCol && prevTs) {
              boardTimes[prevCol] = (boardTimes[prevCol] || 0) + calcHorasUteis(prevTs, ts);
            }
            prevCol = col; prevTs = ts;
          }
        }
        if (prevCol && prevTs) {
          boardTimes[prevCol] = (boardTimes[prevCol] || 0) + calcHorasUteis(prevTs, new Date().toISOString());
        }

        // pausa periods
        const pausas = extrairPausas(updates);

        // dates
        const createdDate = wi.fields?.['System.CreatedDate'];
        const activatedDate = wi.fields?.['Microsoft.VSTS.Common.ActivatedDate'];
        const closedDate = wi.fields?.['Microsoft.VSTS.Common.ClosedDate'] || wi.fields?.['Microsoft.VSTS.Common.ResolvedDate'];

        // lead time = created → closed (business hours)
        const leadTime = calcHorasUteis(createdDate, closedDate, pausas);
        // cycle time = activated → closed (business hours)
        const cycleTime = calcHorasUteis(activatedDate, closedDate, pausas);
        // horas úteis total = created → now or closed
        const horasUteis = calcHorasUteis(createdDate, closedDate || new Date().toISOString(), pausas);

        // interaction count = number of meaningful updates (comments, state changes, assignments)
        const interacoes = updates.filter(u =>
          u.fields?.['System.State'] || u.fields?.['System.AssignedTo'] || u.commentVersionRef
        ).length;

        mMap[id] = { leadTime, cycleTime, horasUteis, pausas, interacoes, boardTimes };
      }

      setUpdatesMap(uMap);
      setMetricsMap(mMap);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false); setLoadingMsg('');
    }
  }, [azOrg, azProject, azPat, filterIteration, filterMember, filterDateFrom, filterDateTo]);

  // ── derived metrics ────────────────────────────────────────────────────────
  const metrics = useMemo(() => {
    if (!workItems.length) return null;

    const totalItems = workItems.length;
    const completedItems = workItems.filter(w => {
      const s = (w.fields?.['System.State'] || '').toLowerCase();
      return s === 'done' || s === 'closed' || s === 'resolved';
    }).length;

    // hours per person
    const hoursByPerson = {};
    for (const wi of workItems) {
      const name = wi.fields?.['System.AssignedTo']?.displayName || 'Não atribuído';
      const h = metricsMap[wi.id]?.horasUteis || 0;
      hoursByPerson[name] = (hoursByPerson[name] || 0) + h;
    }

    // hours by type
    const hoursByType = {};
    for (const wi of workItems) {
      const type = wi.fields?.['System.WorkItemType'] || 'Outro';
      const h = metricsMap[wi.id]?.horasUteis || 0;
      hoursByType[type] = (hoursByType[type] || 0) + h;
    }

    // board column aggregates
    const boardTotals = {};
    for (const m of Object.values(metricsMap)) {
      for (const [col, h] of Object.entries(m.boardTimes || {})) {
        boardTotals[col] = (boardTotals[col] || 0) + h;
      }
    }

    // lead/cycle scatter
    const scatter = workItems
      .filter(w => metricsMap[w.id]?.leadTime > 0)
      .map(w => ({
        id: w.id,
        title: w.fields?.['System.Title']?.slice(0, 30),
        leadTime: +(metricsMap[w.id]?.leadTime || 0).toFixed(1),
        cycleTime: +(metricsMap[w.id]?.cycleTime || 0).toFixed(1),
        interacoes: metricsMap[w.id]?.interacoes || 0,
      }));

    // radar: interactions per person
    const interByPerson = {};
    for (const wi of workItems) {
      const name = wi.fields?.['System.AssignedTo']?.displayName || 'Não atribuído';
      interByPerson[name] = (interByPerson[name] || 0) + (metricsMap[wi.id]?.interacoes || 0);
    }

    const totalHours = Object.values(hoursByPerson).reduce((a, b) => a + b, 0);
    const avgLeadTime = scatter.length ? scatter.reduce((a, s) => a + s.leadTime, 0) / scatter.length : 0;
    const avgCycleTime = scatter.length ? scatter.reduce((a, s) => a + s.cycleTime, 0) / scatter.length : 0;

    return {
      totalItems, completedItems, totalHours, avgLeadTime, avgCycleTime,
      hoursByPerson, hoursByType, boardTotals, scatter, interByPerson,
    };
  }, [workItems, metricsMap]);

  // ── AI analysis ────────────────────────────────────────────────────────────
  const handleAiAnalysis = useCallback(async () => {
    if (!anthropicKey || !metrics) return;
    setAiLoading(true); setAiReport('');
    try {
      const payload = {
        totalItems: metrics.totalItems,
        completedItems: metrics.completedItems,
        totalHorasUteis: +metrics.totalHours.toFixed(1),
        avgLeadTimeHoras: +metrics.avgLeadTime.toFixed(1),
        avgCycleTimeHoras: +metrics.avgCycleTime.toFixed(1),
        horasPorPessoa: Object.fromEntries(Object.entries(metrics.hoursByPerson).map(([k, v]) => [k, +v.toFixed(1)])),
        horasPorTipo: Object.fromEntries(Object.entries(metrics.hoursByType).map(([k, v]) => [k, +v.toFixed(1)])),
        tempoPorColuna: Object.fromEntries(Object.entries(metrics.boardTotals).map(([k, v]) => [k, +v.toFixed(1)])),
        interacoesPorPessoa: metrics.interByPerson,
        project: azProject,
        filtros: { iteracao: filterIteration, membro: filterMember, de: filterDateFrom, ate: filterDateTo },
      };

      const res = await fetch('/anthropic/v1/messages', {
        method: 'POST',
        headers: { 'x-api-key': anthropicKey, 'anthropic-version': '2023-06-01', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: 'claude-haiku-4-5-20251001',
          max_tokens: 1024,
          system: `Você é um consultor sênior de produtividade de equipes de software. Analise os dados abaixo e gere um relatório executivo em português com: (1) principais achados, (2) pontos de atenção, (3) recomendações acionáveis. Seja objetivo e direto. Use markdown.`,
          messages: [{ role: 'user', content: `Dados do projeto ${azProject}:\n\n${JSON.stringify(payload, null, 2)}` }],
        }),
      });
      const data = await res.json();
      setAiReport(data.content?.[0]?.text || 'Sem resposta.');
    } catch (e) {
      setAiReport(`Erro: ${e.message}`);
    } finally {
      setAiLoading(false);
    }
  }, [anthropicKey, metrics, azProject, filterIteration, filterMember, filterDateFrom, filterDateTo]);

  // ── PDF export ─────────────────────────────────────────────────────────────
  const handleExportPdf = useCallback(async () => {
    const { default: html2canvas } = await import('html2canvas');
    const { jsPDF } = await import('jspdf');
    const el = document.getElementById('relatorio-pdf');
    if (!el) return;
    const canvas = await html2canvas(el, { scale: 1.5, backgroundColor: '#0d1117', useCORS: true });
    const img = canvas.toDataURL('image/png');
    const pdf = new jsPDF({ orientation: 'landscape', unit: 'px', format: [canvas.width / 1.5, canvas.height / 1.5] });
    pdf.addImage(img, 'PNG', 0, 0, canvas.width / 1.5, canvas.height / 1.5);
    pdf.save(`produtividade-${azProject}-${new Date().toISOString().slice(0, 10)}.pdf`);
  }, [azProject]);

  // ── render ────────────────────────────────────────────────────────────────
  const sx = {
    wrap: { minHeight: '100vh', background: C.bg, color: C.text, fontFamily: 'inherit' },
    header: { background: C.surface, borderBottom: `1px solid ${C.border}`, padding: '12px 24px', display: 'flex', alignItems: 'center', gap: 12 },
    main: { maxWidth: 1400, margin: '0 auto', padding: '24px 24px' },
    card: { background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: 20, marginBottom: 20 },
    input: { background: '#010409', border: `1px solid ${C.border}`, borderRadius: 6, color: C.text, padding: '6px 10px', fontSize: 13, width: '100%' },
    btn: (col = C.primary) => ({ background: col, color: '#fff', border: 'none', borderRadius: 6, padding: '8px 16px', cursor: 'pointer', fontWeight: 600, fontSize: 13 }),
    btnOutline: { background: 'transparent', border: `1px solid ${C.border}`, color: C.muted, borderRadius: 6, padding: '8px 16px', cursor: 'pointer', fontSize: 13 },
    label: { fontSize: 12, color: C.muted, display: 'block', marginBottom: 4 },
    row: { display: 'flex', gap: 12, flexWrap: 'wrap' },
    tabs: { display: 'flex', gap: 2, borderBottom: `1px solid ${C.border}`, marginBottom: 24 },
  };

  const tabBtn = (key, label) => (
    <button
      onClick={() => setTab(key)}
      style={{ ...sx.btnOutline, borderBottom: tab === key ? `2px solid ${C.primary}` : '2px solid transparent', color: tab === key ? C.primary : C.muted, borderRadius: '6px 6px 0 0', borderLeft: 'none', borderRight: 'none', borderTop: 'none' }}
    >
      {label}
    </button>
  );

  return (
    <div style={sx.wrap}>
      {/* header */}
      <div style={sx.header}>
        <div style={{ fontSize: 18, fontWeight: 700, color: C.text }}>⚡ Agente de Produtividade</div>
        <div style={{ flex: 1 }} />
        <button onClick={() => setConfigOpen(o => !o)} style={sx.btnOutline}>⚙ Configuração</button>
        {metrics && <button onClick={handleExportPdf} style={sx.btn(C.warn)}>↓ PDF</button>}
      </div>

      <div style={sx.main}>
        {/* config panel */}
        {configOpen && (
          <div style={sx.card}>
            <h3 style={{ marginBottom: 16, fontWeight: 600 }}>Configuração</h3>
            <div style={{ ...sx.row, marginBottom: 12 }}>
              <div style={{ flex: 1 }}>
                <label style={sx.label}>Organização Azure DevOps</label>
                <input style={sx.input} placeholder="ex: minha-org" value={azOrg} onChange={e => setAzOrg(e.target.value)} />
              </div>
              <div style={{ flex: 1 }}>
                <label style={sx.label}>PAT (Personal Access Token)</label>
                <input style={sx.input} type="password" placeholder="Token com permissão de leitura" value={azPat} onChange={e => setAzPat(e.target.value)} />
              </div>
              <div style={{ flex: 1 }}>
                <label style={sx.label}>Chave Anthropic (Claude Haiku)</label>
                <input style={sx.input} type="password" placeholder="sk-ant-..." value={anthropicKey} onChange={e => setAnthropicKey(e.target.value)} />
              </div>
              <div style={{ display: 'flex', alignItems: 'flex-end' }}>
                <button onClick={handleConnect} style={sx.btn()} disabled={loading}>
                  {loading && loadingMsg.startsWith('Conectando') ? 'Conectando…' : 'Conectar'}
                </button>
              </div>
            </div>
            <div style={sx.row}>
              <div style={{ flex: 1 }}>
                <label style={sx.label}>Projeto</label>
                <select style={sx.input} value={azProject} onChange={e => handleProjectChange(e.target.value)}>
                  <option value="">Selecione o projeto</option>
                  {projects.map(p => <option key={p.id} value={p.name}>{p.name}</option>)}
                </select>
              </div>
              <div style={{ flex: 1 }}>
                <label style={sx.label}>Sprint / Iteração</label>
                <select style={sx.input} value={filterIteration} onChange={e => setFilterIteration(e.target.value)}>
                  <option value="">Todas as iterações</option>
                  {iterations.map(it => <option key={it.id} value={it.path}>{it.name}</option>)}
                </select>
              </div>
              <div style={{ flex: 1 }}>
                <label style={sx.label}>Profissional</label>
                <select style={sx.input} value={filterMember} onChange={e => setFilterMember(e.target.value)}>
                  <option value="">Todos</option>
                  {members.map(m => <option key={m} value={m}>{m}</option>)}
                </select>
              </div>
              <div style={{ flex: 1 }}>
                <label style={sx.label}>Data de</label>
                <input type="date" style={sx.input} value={filterDateFrom} onChange={e => setFilterDateFrom(e.target.value)} />
              </div>
              <div style={{ flex: 1 }}>
                <label style={sx.label}>Data até</label>
                <input type="date" style={sx.input} value={filterDateTo} onChange={e => setFilterDateTo(e.target.value)} />
              </div>
              <div style={{ display: 'flex', alignItems: 'flex-end' }}>
                <button onClick={handleFetch} style={sx.btn(C.success)} disabled={loading || !azProject}>
                  {loading ? loadingMsg || 'Carregando…' : '⟳ Buscar Dados'}
                </button>
              </div>
            </div>
          </div>
        )}

        {error && <div style={{ background: `${C.danger}22`, border: `1px solid ${C.danger}`, borderRadius: 6, padding: '10px 14px', color: C.danger, marginBottom: 16, fontSize: 13 }}>{error}</div>}

        {loading && (
          <div style={{ textAlign: 'center', color: C.muted, padding: 40 }}>
            <div style={{ fontSize: 13 }}>{loadingMsg}</div>
            <div style={{ marginTop: 8, height: 2, background: C.border, borderRadius: 1, overflow: 'hidden' }}>
              <div style={{ height: '100%', background: C.primary, width: '60%', animation: 'pulse 1.5s ease-in-out infinite' }} />
            </div>
          </div>
        )}

        {metrics && !loading && (
          <div id="relatorio-pdf">
            {/* KPI strip */}
            <div style={{ ...sx.row, marginBottom: 24 }}>
              <KpiCard label="Work Items" value={metrics.totalItems} sub={`${metrics.completedItems} concluídos`} />
              <KpiCard label="Horas Úteis Totais" value={horasToStr(metrics.totalHours)} color={C.success} />
              <KpiCard label="Lead Time Médio" value={horasToStr(metrics.avgLeadTime)} sub="criação → conclusão" color={C.warn} />
              <KpiCard label="Cycle Time Médio" value={horasToStr(metrics.avgCycleTime)} sub="início → conclusão" color={C.primary} />
              <KpiCard label="Itens em Aberto" value={metrics.totalItems - metrics.completedItems} color={C.danger} />
            </div>

            {/* tabs */}
            <div style={sx.tabs}>
              {tabBtn('dashboard', 'Dashboard')}
              {tabBtn('board', 'Board & Fluxo')}
              {tabBtn('items', 'Work Items')}
              {tabBtn('ai', 'Análise IA')}
            </div>

            {/* dashboard tab */}
            {tab === 'dashboard' && (
              <div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 }}>
                  {/* horas por pessoa */}
                  <div style={sx.card}>
                    <h4 style={{ marginBottom: 12, fontSize: 13, color: C.muted }}>Horas Úteis por Profissional</h4>
                    <ResponsiveContainer width="100%" height={240}>
                      <BarChart data={Object.entries(metrics.hoursByPerson).map(([name, h]) => ({ name: name.split(' ')[0], horas: +h.toFixed(1) }))}>
                        <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                        <XAxis dataKey="name" tick={{ fill: C.muted, fontSize: 11 }} />
                        <YAxis tick={{ fill: C.muted, fontSize: 11 }} />
                        <Tooltip content={<CustomTooltip />} />
                        <Bar dataKey="horas" fill={C.primary} radius={[4, 4, 0, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>

                  {/* horas por tipo */}
                  <div style={sx.card}>
                    <h4 style={{ marginBottom: 12, fontSize: 13, color: C.muted }}>Distribuição por Tipo</h4>
                    <ResponsiveContainer width="100%" height={240}>
                      <PieChart>
                        <Pie
                          data={Object.entries(metrics.hoursByType).map(([name, value]) => ({ name, value: +value.toFixed(1) }))}
                          cx="50%" cy="50%" innerRadius={60} outerRadius={100}
                          dataKey="value" nameKey="name"
                        >
                          {Object.keys(metrics.hoursByType).map((k, i) => (
                            <Cell key={k} fill={TYPE_COLOR[k] || CHART_COLORS[i % CHART_COLORS.length]} />
                          ))}
                        </Pie>
                        <Tooltip content={<CustomTooltip />} />
                        <Legend wrapperStyle={{ fontSize: 11, color: C.muted }} />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>

                  {/* lead time vs cycle time scatter */}
                  <div style={sx.card}>
                    <h4 style={{ marginBottom: 12, fontSize: 13, color: C.muted }}>Lead Time vs. Cycle Time (h)</h4>
                    <ResponsiveContainer width="100%" height={240}>
                      <ScatterChart>
                        <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                        <XAxis dataKey="leadTime" name="Lead Time" type="number" tick={{ fill: C.muted, fontSize: 11 }} label={{ value: 'Lead Time (h)', fill: C.muted, fontSize: 11, position: 'insideBottom', offset: -4 }} />
                        <YAxis dataKey="cycleTime" name="Cycle Time" type="number" tick={{ fill: C.muted, fontSize: 11 }} label={{ value: 'Cycle Time (h)', fill: C.muted, fontSize: 11, angle: -90, position: 'insideLeft' }} />
                        <ZAxis dataKey="interacoes" range={[40, 200]} name="Interações" />
                        <Tooltip cursor={{ strokeDasharray: '3 3' }} content={<CustomTooltip />} />
                        <Scatter data={metrics.scatter} fill={C.primary} opacity={0.7} />
                      </ScatterChart>
                    </ResponsiveContainer>
                  </div>

                  {/* interactions radar */}
                  <div style={sx.card}>
                    <h4 style={{ marginBottom: 12, fontSize: 13, color: C.muted }}>Índice de Interação por Profissional</h4>
                    <ResponsiveContainer width="100%" height={240}>
                      <RadarChart data={Object.entries(metrics.interByPerson).map(([name, v]) => ({ name: name.split(' ')[0], interacoes: v }))}>
                        <PolarGrid stroke={C.border} />
                        <PolarAngleAxis dataKey="name" tick={{ fill: C.muted, fontSize: 11 }} />
                        <PolarRadiusAxis tick={{ fill: C.muted, fontSize: 9 }} />
                        <Radar name="Interações" dataKey="interacoes" stroke={C.purple} fill={C.purple} fillOpacity={0.3} />
                        <Tooltip content={<CustomTooltip />} />
                      </RadarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </div>
            )}

            {/* board tab */}
            {tab === 'board' && (
              <div>
                <div style={sx.card}>
                  <h4 style={{ marginBottom: 12, fontSize: 13, color: C.muted }}>Tempo Médio por Coluna do Board (h)</h4>
                  <ResponsiveContainer width="100%" height={320}>
                    <BarChart
                      data={Object.entries(metrics.boardTotals)
                        .sort((a, b) => b[1] - a[1])
                        .map(([col, h]) => ({ col, horas: +h.toFixed(1) }))}
                      layout="vertical"
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                      <XAxis type="number" tick={{ fill: C.muted, fontSize: 11 }} />
                      <YAxis dataKey="col" type="category" width={140} tick={{ fill: C.muted, fontSize: 11 }} />
                      <Tooltip content={<CustomTooltip />} />
                      <Bar dataKey="horas" fill={C.cyan} radius={[0, 4, 4, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div style={sx.card}>
                  <h4 style={{ marginBottom: 12, fontSize: 13, color: C.muted }}>Indicador de Interação por Card</h4>
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                      <thead>
                        <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                          {['ID', 'Título', 'Tipo', 'Responsável', 'Interações', 'Pausas', 'Lead Time'].map(h => (
                            <th key={h} style={{ padding: '8px 10px', textAlign: 'left', color: C.muted, fontWeight: 500 }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {workItems
                          .sort((a, b) => (metricsMap[b.id]?.interacoes || 0) - (metricsMap[a.id]?.interacoes || 0))
                          .slice(0, 50)
                          .map(wi => {
                            const m = metricsMap[wi.id] || {};
                            return (
                              <tr key={wi.id} style={{ borderBottom: `1px solid ${C.border}22` }}>
                                <td style={{ padding: '7px 10px', color: C.muted }}>{wi.id}</td>
                                <td style={{ padding: '7px 10px' }}>{(wi.fields?.['System.Title'] || '').slice(0, 50)}</td>
                                <td style={{ padding: '7px 10px' }}>
                                  <Tag color={TYPE_COLOR[wi.fields?.['System.WorkItemType']] || C.muted}>
                                    {wi.fields?.['System.WorkItemType']}
                                  </Tag>
                                </td>
                                <td style={{ padding: '7px 10px' }}>{wi.fields?.['System.AssignedTo']?.displayName?.split(' ')[0] || '—'}</td>
                                <td style={{ padding: '7px 10px', color: m.interacoes > 10 ? C.warn : C.text }}>{m.interacoes || 0}</td>
                                <td style={{ padding: '7px 10px', color: m.pausas?.length > 0 ? C.danger : C.muted }}>{m.pausas?.length || 0}</td>
                                <td style={{ padding: '7px 10px' }}>{m.leadTime > 0 ? horasToStr(m.leadTime) : '—'}</td>
                              </tr>
                            );
                          })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* items tab */}
            {tab === 'items' && (
              <div style={sx.card}>
                <h4 style={{ marginBottom: 12, fontSize: 13, color: C.muted }}>Work Items ({workItems.length})</h4>
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                    <thead>
                      <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                        {['ID', 'Título', 'Tipo', 'Estado', 'Responsável', 'Criado', 'Fechado', 'Lead Time', 'Cycle Time', 'Horas Úteis'].map(h => (
                          <th key={h} style={{ padding: '8px 10px', textAlign: 'left', color: C.muted, fontWeight: 500, whiteSpace: 'nowrap' }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {workItems.map(wi => {
                        const m = metricsMap[wi.id] || {};
                        const state = wi.fields?.['System.State'] || '';
                        return (
                          <tr key={wi.id} style={{ borderBottom: `1px solid ${C.border}22` }}>
                            <td style={{ padding: '7px 10px', color: C.muted }}>{wi.id}</td>
                            <td style={{ padding: '7px 10px', maxWidth: 240 }}>{(wi.fields?.['System.Title'] || '').slice(0, 50)}</td>
                            <td style={{ padding: '7px 10px' }}>
                              <Tag color={TYPE_COLOR[wi.fields?.['System.WorkItemType']] || C.muted}>
                                {wi.fields?.['System.WorkItemType']}
                              </Tag>
                            </td>
                            <td style={{ padding: '7px 10px' }}>
                              <Tag color={stateColor(state)}>{state}</Tag>
                            </td>
                            <td style={{ padding: '7px 10px' }}>{wi.fields?.['System.AssignedTo']?.displayName || '—'}</td>
                            <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>{fmtDate(wi.fields?.['System.CreatedDate'])}</td>
                            <td style={{ padding: '7px 10px', whiteSpace: 'nowrap' }}>{fmtDate(wi.fields?.['Microsoft.VSTS.Common.ClosedDate'])}</td>
                            <td style={{ padding: '7px 10px' }}>{m.leadTime > 0 ? horasToStr(m.leadTime) : '—'}</td>
                            <td style={{ padding: '7px 10px' }}>{m.cycleTime > 0 ? horasToStr(m.cycleTime) : '—'}</td>
                            <td style={{ padding: '7px 10px' }}>{m.horasUteis > 0 ? horasToStr(m.horasUteis) : '—'}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* ai tab */}
            {tab === 'ai' && (
              <div>
                <div style={{ ...sx.row, marginBottom: 16 }}>
                  <button onClick={handleAiAnalysis} style={sx.btn(C.purple)} disabled={aiLoading || !anthropicKey}>
                    {aiLoading ? 'Analisando…' : '✦ Gerar Análise com Claude Haiku'}
                  </button>
                  {!anthropicKey && <span style={{ color: C.warn, fontSize: 12, alignSelf: 'center' }}>Configure a chave Anthropic para usar a análise IA.</span>}
                </div>
                {aiReport && (
                  <div style={{ ...sx.card, whiteSpace: 'pre-wrap', fontSize: 13, lineHeight: 1.7 }}>
                    {aiReport}
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {!metrics && !loading && (
          <div style={{ textAlign: 'center', color: C.muted, padding: 80, fontSize: 14 }}>
            Configure a conexão com o Azure DevOps e clique em <strong>Buscar Dados</strong> para iniciar a análise.
          </div>
        )}
      </div>

      <style>{`
        @keyframes pulse { 0%,100%{opacity:.4} 50%{opacity:1} }
        button:disabled { opacity: .5; cursor: default; }
        select option { background: #161b22; }
      `}</style>
    </div>
  );
}
