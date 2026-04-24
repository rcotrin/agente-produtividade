import { useState, useCallback, useMemo } from 'react';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, ScatterChart, Scatter, ZAxis,
  RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis,
} from 'recharts';

// ─── paleta Y Tecnologia ─────────────────────────────────────────────────────
const C = {
  bg: '#F9F9F9',       surface: '#FFFFFF',    border: '#E8ECF5',
  accent: '#39ADE3',   navy: '#00366C',       navyMed: '#07447A',
  cyanLight: '#87D5F6', sectionBg: '#E9F5FA',
  green: '#166534',    amber: '#92400e',      coral: '#9a3412',
  red: '#991b1b',      muted: '#94a3b8',
  text: '#627C89',     textDim: '#74768B',    textBright: '#444762',
};
const CHART_COLORS = [C.accent, C.navy, C.navyMed, '#0284c7', C.green, C.amber, C.coral, C.cyanLight];

const TYPE_COLOR = {
  'Epic': C.navy, 'Feature': C.navyMed, 'User Story': C.accent,
  'Task': '#0284c7', 'Bug': C.coral, 'Test Case': '#6366f1',
};

// ─── feriados nacionais brasileiros ──────────────────────────────────────────
// Fixos: MM-DD
const FERIADOS_FIXOS = new Set([
  '01-01', // Confraternização Universal
  '04-21', // Tiradentes
  '05-01', // Dia do Trabalho
  '09-07', // Independência do Brasil
  '10-12', // Nossa Senhora Aparecida
  '11-02', // Finados
  '11-15', // Proclamação da República
  '11-20', // Consciência Negra (lei federal desde 2024)
  '12-25', // Natal
]);

// Algoritmo de Butcher para calcular a Páscoa
function calcPascoa(year) {
  const a = year % 19, b = Math.floor(year / 100), c = year % 100;
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month - 1, day);
}

function toKey(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

const _feriadosMoveis = {};
function getFeriadosMoveis(year) {
  if (_feriadosMoveis[year]) return _feriadosMoveis[year];
  const pascoa = calcPascoa(year);
  const offsets = { carnaval2: -47, carnaval3: -46, sextaSanta: -2, pascoa: 0, corpusChristi: 60 };
  const set = new Set(
    Object.values(offsets).map(o => {
      const d = new Date(pascoa);
      d.setDate(pascoa.getDate() + o);
      return toKey(d);
    })
  );
  _feriadosMoveis[year] = set;
  return set;
}

function isFeriado(d) {
  const mmdd = `${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  if (FERIADOS_FIXOS.has(mmdd)) return true;
  return getFeriadosMoveis(d.getFullYear()).has(toKey(d));
}

// ─── helpers ─────────────────────────────────────────────────────────────────
function isWorkDay(d) {
  const dow = d.getDay();
  return dow !== 0 && dow !== 6 && !isFeriado(d);
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

// Nas updates, cada campo vem como { oldValue, newValue } — não uma string direta.
function colFromUpdate(u) {
  return u.fields?.['System.BoardColumn']?.newValue || '';
}
function tsFromUpdate(u) {
  return u.fields?.['System.ChangedDate']?.newValue || u.revisedDate || '';
}

function extrairPausas(updates) {
  const pausas = [];
  let pausaInicio = null;
  for (const u of updates) {
    const label = colFromUpdate(u).toLowerCase();
    const ts = tsFromUpdate(u);
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

function isIdleColumn(col) {
  if (!col) return true;
  const l = col.toLowerCase();
  return (
    l.includes('new') || l.includes('to do') || l.includes('todo') ||
    l.includes('backlog') || l.includes('pausa') || l.includes('bloqueio') ||
    l.includes('done') || l.includes('closed') || l.includes('resolved') ||
    l.includes('cancelled') || l.includes('cancelad') || l.includes('aguardando') ||
    l.includes('a fazer') || l.includes('novo') || l.includes('fila')
  );
}

// Horas úteis baseadas no tempo real em colunas ativas do board
function calcHorasUteisFromBoard(updates) {
  const timeline = [];
  for (const u of updates) {
    const col = colFromUpdate(u);
    const ts = tsFromUpdate(u);
    if (col && ts) timeline.push({ col, ts });
  }
  if (!timeline.length) return 0;
  timeline.sort((a, b) => new Date(a.ts) - new Date(b.ts));

  let total = 0;
  for (let i = 0; i < timeline.length - 1; i++) {
    if (!isIdleColumn(timeline[i].col)) {
      total += calcHorasUteis(timeline[i].ts, timeline[i + 1].ts);
    }
  }
  const last = timeline[timeline.length - 1];
  if (!isIdleColumn(last.col)) {
    total += calcHorasUteis(last.ts, new Date().toISOString());
  }
  return total;
}

// Primeira vez que o card entrou em coluna ativa (substitui ActivatedDate quando ausente)
function findActivatedDate(updates) {
  const timeline = [];
  for (const u of updates) {
    const col = colFromUpdate(u);
    const ts = tsFromUpdate(u);
    if (col && ts) timeline.push({ col, ts });
  }
  timeline.sort((a, b) => new Date(a.ts) - new Date(b.ts));
  return timeline.find(t => !isIdleColumn(t.col))?.ts || null;
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
  if (!res.ok) {
    const body = await res.text();
    const clean = body.startsWith('<') ? `recurso não encontrado (${url.split('?')[0]})` : body.slice(0, 200);
    throw new Error(`Azure DevOps ${res.status}: ${clean}`);
  }
  return res.json();
}

async function fetchProjects(org, pat) {
  const data = await azFetch(pat, `/devops/${org}/_apis/projects?api-version=7.1`);
  return data.value || [];
}

async function fetchIterations(org, project, pat) {
  const data = await azFetch(pat, `/devops/${org}/${project}/_apis/wit/classificationnodes/Iterations?$depth=10&api-version=7.1`);
  const nodes = [];
  function flatten(node) {
    if (node.path) nodes.push({ id: node.id, name: node.name, path: node.path.replace(/^\\/, '') });
    (node.children || []).forEach(flatten);
  }
  flatten(data);
  return nodes;
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
  // [System.TeamProject] omitido — já está na URL; incluí-lo causa 400 em alguns tenants
  // Épicos e Features são itens macro de backlog, não contam para produtividade
  const conditions = [`[System.WorkItemType] NOT IN ('Epic', 'Feature')`];
  if (filters.iterationPath) conditions.push(`[System.IterationPath] = '${filters.iterationPath}'`);
  if (filters.assignedTo) conditions.push(`[System.AssignedTo] = '${filters.assignedTo}'`);
  if (filters.dateFrom) conditions.push(`[System.ChangedDate] >= '${filters.dateFrom}'`);
  if (filters.dateTo) conditions.push(`[System.ChangedDate] <= '${filters.dateTo}'`);

  const where = conditions.length ? `WHERE ${conditions.join(' AND ')} ` : '';
  const wiql = {
    query: `SELECT [System.Id] FROM WorkItems ${where}ORDER BY [System.ChangedDate] DESC`,
  };

  const res = await fetch(`/devops/${org}/${project}/_apis/wit/wiql?$top=500&api-version=7.1`, {
    method: 'POST',
    headers: azHeaders(pat),
    body: JSON.stringify(wiql),
  });
  if (!res.ok) {
    const body = await res.text();
    const clean = body.startsWith('<') ? `WIQL falhou — verifique projeto e permissões (${res.status})` : body.slice(0, 200);
    throw new Error(clean);
  }
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
function KpiCard({ label, value, sub, color = C.accent }) {
  return (
    <div className="kpi-card">
      <div style={{ fontSize: 11, color: C.textDim, marginBottom: 4, fontWeight: 600, letterSpacing: '0.06em', textTransform: 'uppercase' }}>{label}</div>
      <div style={{ fontSize: 28, fontWeight: 800, color, fontFamily: "'Manrope',sans-serif", letterSpacing: '-0.01em' }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

function Tag({ children, color = C.muted }) {
  return (
    <span className="badge" style={{ background: `${color}18`, color, border: `1px solid ${color}40` }}>
      {children}
    </span>
  );
}

function stateColor(s) {
  if (!s) return C.muted;
  const l = s.toLowerCase();
  if (l.includes('done') || l.includes('closed') || l.includes('resolved')) return C.green;
  if (l.includes('active') || l.includes('progress')) return C.accent;
  if (l.includes('blocked') || l.includes('pausa')) return C.coral;
  return C.amber;
}

const CustomTooltip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 7, padding: '8px 14px', fontSize: 12, boxShadow: '0 2px 8px #00366C0e' }}>
      {label && <div style={{ color: C.textDim, marginBottom: 4, fontWeight: 600 }}>{label}</div>}
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color || C.textBright }}>
          {p.name}: <strong>{typeof p.value === 'number' ? p.value.toFixed(1) : p.value}</strong>
        </div>
      ))}
    </div>
  );
};

// ─── main component ───────────────────────────────────────────────────────────
export default function AgenteProdutividade() {
  const [azOrg, setAzOrg] = useState('');
  const [azProject, setAzProject] = useState('');
  const [azPat, setAzPat] = useState('');
  const [anthropicKey, setAnthropicKey] = useState('');

  const [projects, setProjects] = useState([]);
  const [iterations, setIterations] = useState([]);
  const [members, setMembers] = useState([]);

  const [filterDateFrom, setFilterDateFrom] = useState('');
  const [filterDateTo, setFilterDateTo] = useState('');
  const [filterIteration, setFilterIteration] = useState('');
  const [filterMember, setFilterMember] = useState('');

  const [workItems, setWorkItems] = useState([]);
  const [updatesMap, setUpdatesMap] = useState({});
  const [metricsMap, setMetricsMap] = useState({});
  const [activityHours, setActivityHours] = useState({}); // { name: horas únicas detectadas }

  const [loading, setLoading] = useState(false);
  const [loadingMsg, setLoadingMsg] = useState('');
  const [error, setError] = useState('');
  const [tab, setTab] = useState('dashboard');
  const [aiReport, setAiReport] = useState('');
  const [aiLoading, setAiLoading] = useState(false);
  const [configOpen, setConfigOpen] = useState(true);

  // ── connect ────────────────────────────────────────────────────────────────
  const handleConnect = useCallback(async () => {
    if (!azOrg || !azPat) return setError('Informe organização e PAT.');
    setError(''); setLoading(true); setLoadingMsg('Conectando ao Azure DevOps…');
    try {
      const projs = await fetchProjects(azOrg, azPat);
      setProjects(projs);
      if (azProject) {
        const [iters, mems] = await Promise.allSettled([
          fetchIterations(azOrg, azProject, azPat),
          fetchTeamMembers(azOrg, azProject, azPat),
        ]);
        if (iters.status === 'fulfilled') setIterations(iters.value);
        if (mems.status === 'fulfilled') setMembers(mems.value);
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
      const [iters, mems] = await Promise.allSettled([
        fetchIterations(azOrg, proj, azPat),
        fetchTeamMembers(azOrg, proj, azPat),
      ]);
      if (iters.status === 'fulfilled') setIterations(iters.value);
      if (mems.status === 'fulfilled') setMembers(mems.value);
    } catch { /* ignore */ }
  }, [azOrg, azPat]);

  // ── fetch & compute ────────────────────────────────────────────────────────
  const handleFetch = useCallback(async () => {
    if (!azOrg || !azProject || !azPat) return setError('Configure a conexão antes de buscar dados.');
    setError(''); setLoading(true); setWorkItems([]); setUpdatesMap({}); setMetricsMap({}); setActivityHours({}); setAiReport('');

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

        // boardTimes: tempo acumulado por coluna (usa helpers corrigidos)
        const boardTimes = {};
        let prevCol = null, prevTs = null;
        for (const u of updates) {
          const col = colFromUpdate(u);
          const ts = tsFromUpdate(u);
          if (col && ts) {
            if (prevCol && prevTs) boardTimes[prevCol] = (boardTimes[prevCol] || 0) + calcHorasUteis(prevTs, ts);
            prevCol = col; prevTs = ts;
          }
        }
        if (prevCol && prevTs) boardTimes[prevCol] = (boardTimes[prevCol] || 0) + calcHorasUteis(prevTs, new Date().toISOString());

        const pausas = extrairPausas(updates);
        const createdDate = wi.fields?.['System.CreatedDate'];
        const closedDate = wi.fields?.['Microsoft.VSTS.Common.ClosedDate'] || wi.fields?.['Microsoft.VSTS.Common.ResolvedDate'];

        // Activated: campo nativo ou primeira entrada em coluna ativa
        const activatedDate = wi.fields?.['Microsoft.VSTS.Common.ActivatedDate'] || findActivatedDate(updates);

        const leadTime = calcHorasUteis(createdDate, closedDate, pausas);
        const cycleTime = calcHorasUteis(activatedDate, closedDate, pausas);

        // Esforço registrado manualmente no Azure DevOps (campo CompletedWork)
        const completedWork = wi.fields?.['Microsoft.VSTS.Scheduling.CompletedWork'] || 0;

        // Tempo em colunas ativas — útil para Lead/Cycle Time de cards individuais
        const horasUteisBoard = calcHorasUteisFromBoard(updates);

        const interacoes = updates.filter(u => u.fields?.['System.State'] || u.fields?.['System.AssignedTo'] || u.commentVersionRef).length;

        mMap[id] = { leadTime, cycleTime, completedWork, horasUteisBoard, pausas, interacoes, boardTimes };
      }

      // ── Janelas de Atividade ──────────────────────────────────────────────
      // Para cada update, registra a hora-slot (pessoa+data+hora) como ativa.
      // Horas únicas = limite inferior conservador de horas realmente trabalhadas.
      const slots = {}; // { name: Set<"YYYY-MM-DD-HH"> }
      for (const updates of Object.values(uMap)) {
        for (const u of updates) {
          const who = u.revisedBy?.displayName;
          const ts = u.revisedDate || u.fields?.['System.ChangedDate']?.newValue;
          if (!who || !ts) continue;
          const d = new Date(ts);
          if (!isWorkDay(d)) continue;
          const h = d.getHours();
          if (h < 9 || h >= 18) continue;
          const key = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}-${h}`;
          if (!slots[who]) slots[who] = new Set();
          slots[who].add(key);
        }
      }
      const aHours = {};
      for (const [name, s] of Object.entries(slots)) aHours[name] = s.size;
      setActivityHours(aHours);

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

    // completedWorkByPerson: soma de CompletedWork (esforço manual registrado)
    // itemsByPerson: contagem de cards por pessoa (para quem não usa CompletedWork)
    const completedWorkByPerson = {};
    const itemsByPerson = {};
    const completedWorkByType = {};
    const boardTotals = {};
    const interByPerson = {};
    let totalCompletedWork = 0;
    let anyCompletedWork = false;

    for (const wi of workItems) {
      const name = wi.fields?.['System.AssignedTo']?.displayName || 'Não atribuído';
      const type = wi.fields?.['System.WorkItemType'] || 'Outro';
      const cw = metricsMap[wi.id]?.completedWork || 0;
      if (cw > 0) anyCompletedWork = true;
      completedWorkByPerson[name] = (completedWorkByPerson[name] || 0) + cw;
      completedWorkByType[type] = (completedWorkByType[type] || 0) + cw;
      itemsByPerson[name] = (itemsByPerson[name] || 0) + 1;
      totalCompletedWork += cw;
      interByPerson[name] = (interByPerson[name] || 0) + (metricsMap[wi.id]?.interacoes || 0);
    }

    for (const m of Object.values(metricsMap)) {
      for (const [col, h] of Object.entries(m.boardTimes || {})) {
        boardTotals[col] = (boardTotals[col] || 0) + h;
      }
    }

    // Se nenhum item tem CompletedWork preenchido, usar contagem de cards como proxy
    const hoursByPerson = anyCompletedWork ? completedWorkByPerson : itemsByPerson;
    const hoursByType = anyCompletedWork ? completedWorkByType : {};
    const hoursLabel = anyCompletedWork ? 'Horas (CompletedWork)' : 'Qtd. Cards';

    const scatter = workItems
      .filter(w => metricsMap[w.id]?.leadTime > 0)
      .map(w => ({
        id: w.id,
        title: w.fields?.['System.Title']?.slice(0, 30),
        leadTime: +(metricsMap[w.id]?.leadTime || 0).toFixed(1),
        cycleTime: +(metricsMap[w.id]?.cycleTime || 0).toFixed(1),
        interacoes: metricsMap[w.id]?.interacoes || 0,
      }));

    const avgLeadTime = scatter.length ? scatter.reduce((a, s) => a + s.leadTime, 0) / scatter.length : 0;
    const avgCycleTime = scatter.length ? scatter.reduce((a, s) => a + s.cycleTime, 0) / scatter.length : 0;

    // ── Cobertura CompletedWork por pessoa ───────────────────────────────────
    const cwCoverage = {}; // { name: { filled, total } }
    for (const wi of workItems) {
      const name = wi.fields?.['System.AssignedTo']?.displayName || 'Não atribuído';
      if (!cwCoverage[name]) cwCoverage[name] = { filled: 0, total: 0 };
      cwCoverage[name].total++;
      if ((metricsMap[wi.id]?.completedWork || 0) > 0) cwCoverage[name].filled++;
    }
    const globalCwCoverage = workItems.length
      ? Math.round(100 * workItems.filter(w => (metricsMap[w.id]?.completedWork || 0) > 0).length / workItems.length)
      : 0;

    // ── Gráfico cruzado: Atividade vs CompletedWork ──────────────────────────
    const allNames = new Set([...Object.keys(activityHours), ...Object.keys(completedWorkByPerson)]);
    const activityVsCw = [...allNames].map(name => ({
      name: name.split(' ')[0],
      fullName: name,
      atividade: activityHours[name] || 0,
      completedWork: completedWorkByPerson[name] || 0,
      cobertura: cwCoverage[name] ? Math.round(100 * cwCoverage[name].filled / cwCoverage[name].total) : 0,
    })).sort((a, b) => b.atividade - a.atividade);

    const totalActivityHours = Object.values(activityHours).reduce((a, b) => a + b, 0);

    return {
      totalItems, completedItems, totalCompletedWork, anyCompletedWork,
      avgLeadTime, avgCycleTime, hoursByPerson, hoursByType, hoursLabel,
      boardTotals, scatter, interByPerson,
      cwCoverage, globalCwCoverage, activityVsCw, totalActivityHours,
    };
  }, [workItems, metricsMap, activityHours]);

  // ── AI analysis ────────────────────────────────────────────────────────────
  const handleAiAnalysis = useCallback(async () => {
    if (!anthropicKey || !metrics) return;
    setAiLoading(true); setAiReport('');
    try {
      const payload = {
        totalItems: metrics.totalItems,
        completedItems: metrics.completedItems,
        totalHorasUteis: +metrics.totalCompletedWork.toFixed(1),
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
    const canvas = await html2canvas(el, { scale: 1.5, backgroundColor: '#F9F9F9', useCORS: true });
    const img = canvas.toDataURL('image/png');
    const pdf = new jsPDF({ orientation: 'landscape', unit: 'px', format: [canvas.width / 1.5, canvas.height / 1.5] });
    pdf.addImage(img, 'PNG', 0, 0, canvas.width / 1.5, canvas.height / 1.5);
    pdf.save(`produtividade-${azProject}-${new Date().toISOString().slice(0, 10)}.pdf`);
  }, [azProject]);

  // ── render ─────────────────────────────────────────────────────────────────
  const tabActive = key => tab === key;

  return (
    <div style={{ fontFamily: "'Roboto','system-ui',sans-serif", background: C.bg, minHeight: '100vh', color: C.text }}>
      {/* Google Fonts */}
      <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&family=Manrope:wght@600;700;800&display=swap" rel="stylesheet" />

      <style>{`
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #F9F9F9; }
        ::-webkit-scrollbar-thumb { background: #E8ECF5; border-radius: 3px; }
        ::-webkit-scrollbar-thumb:hover { background: #87D5F6; }
        .kpi-card { background: #FFFFFF; border: 1px solid #E8ECF5; border-radius: 8px; padding: 16px 22px; min-width: 150px; flex: 1; }
        .section-card { background: #FFFFFF; border: 1px solid #E8ECF5; border-radius: 8px; padding: 20px; margin-bottom: 20px; }
        .section-card h4 { font-size: 12px; font-weight: 700; color: #74768B; margin-bottom: 14px; text-transform: uppercase; letter-spacing: 0.08em; }
        .badge { font-size: 10px; padding: 2px 8px; border-radius: 20px; font-weight: 600; letter-spacing: .03em; flex-shrink: 0; }
        .ap-input { background: #FFFFFF; border: 1px solid #E8ECF5; border-radius: 6px; color: #444762; padding: 7px 10px; font-size: 13px; width: 100%; font-family: inherit; transition: border-color .15s, box-shadow .15s; }
        .ap-input:focus { outline: none; border-color: #39ADE3; box-shadow: 0 0 0 3px #39ADE314; }
        .ap-btn { border: none; border-radius: 6px; padding: 8px 18px; cursor: pointer; font-weight: 600; font-size: 13px; font-family: inherit; transition: opacity .15s; }
        .ap-btn:disabled { opacity: .45; cursor: default; }
        .ap-btn:hover:not(:disabled) { opacity: .88; }
        .ap-btn-outline { background: transparent; border: 1px solid #E8ECF5; border-radius: 6px; color: #74768B; padding: 7px 14px; cursor: pointer; font-size: 13px; font-family: inherit; transition: border-color .15s, color .15s; }
        .ap-btn-outline:hover { border-color: #39ADE3; color: #39ADE3; }
        .tab-btn { background: none; border: none; border-bottom: 2px solid transparent; cursor: pointer; font-family: inherit; padding: 8px 14px; margin-bottom: -1px; font-size: 13px; font-weight: 600; color: #74768B; transition: all .15s; }
        .tab-btn.active { border-bottom-color: #39ADE3; color: #39ADE3; }
        .tab-btn:hover:not(.active) { color: #444762; border-bottom-color: #E8ECF5; }
        .wi-row { transition: background .12s; }
        .wi-row:hover { background: #E9F5FA; }
        @keyframes shimmer { 0%,100%{opacity:.4} 50%{opacity:1} }
      `}</style>

      {/* ── header ── */}
      <div style={{ background: '#FFFFFF', borderBottom: '1px solid #E8ECF5', padding: '14px 28px', display: 'flex', alignItems: 'center', gap: 12, boxShadow: '0 2px 8px #00366C0e' }}>
        <div style={{ width: 10, height: 10, borderRadius: '50%', background: 'linear-gradient(135deg, #00366C, #39ADE3)', flexShrink: 0 }} />
        <span style={{ fontFamily: "'Manrope',sans-serif", fontWeight: 800, fontSize: 18, background: 'linear-gradient(30deg, #00366C 0%, #39ADE3 100%)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent', backgroundClip: 'text', letterSpacing: '-0.01em' }}>
          Agente de Produtividade
        </span>
        {azProject && (
          <>
            <span style={{ color: '#E8ECF5', fontSize: 18, margin: '0 2px' }}>│</span>
            <span style={{ fontFamily: "'Manrope',sans-serif", fontWeight: 700, fontSize: 14, color: C.accent }}>{azProject}</span>
          </>
        )}
        <div style={{ flex: 1 }} />
        {metrics && (
          <button className="ap-btn" style={{ background: C.amber, color: '#fff' }} onClick={handleExportPdf}>
            ↓ Exportar PDF
          </button>
        )}
        <button className="ap-btn-outline" onClick={() => setConfigOpen(o => !o)}>
          ⚙ Configuração
        </button>
      </div>

      <div style={{ maxWidth: 1440, margin: '0 auto', padding: '24px 28px' }}>

        {/* ── config panel ── */}
        {configOpen && (
          <div className="section-card" style={{ marginBottom: 24 }}>
            <h4 style={{ fontSize: 12, fontWeight: 700, color: C.textDim, marginBottom: 16, textTransform: 'uppercase', letterSpacing: '0.08em' }}>Configuração</h4>

            <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginBottom: 12 }}>
              <div style={{ flex: 1, minWidth: 180 }}>
                <label style={{ fontSize: 11, color: C.textDim, display: 'block', marginBottom: 4, fontWeight: 600 }}>Organização Azure DevOps</label>
                <input className="ap-input" placeholder="ex: minha-org" value={azOrg} onChange={e => setAzOrg(e.target.value)} />
              </div>
              <div style={{ flex: 1, minWidth: 180 }}>
                <label style={{ fontSize: 11, color: C.textDim, display: 'block', marginBottom: 4, fontWeight: 600 }}>PAT (Personal Access Token)</label>
                <input className="ap-input" type="password" placeholder="Token com permissão de leitura" value={azPat} onChange={e => setAzPat(e.target.value)} />
              </div>
              <div style={{ flex: 1, minWidth: 180 }}>
                <label style={{ fontSize: 11, color: C.textDim, display: 'block', marginBottom: 4, fontWeight: 600 }}>Chave Anthropic (Claude Haiku)</label>
                <input className="ap-input" type="password" placeholder="sk-ant-..." value={anthropicKey} onChange={e => setAnthropicKey(e.target.value)} />
              </div>
              <div style={{ display: 'flex', alignItems: 'flex-end' }}>
                <button className="ap-btn" style={{ background: C.navy, color: '#fff' }} onClick={handleConnect} disabled={loading}>
                  {loading && loadingMsg.startsWith('Conectando') ? 'Conectando…' : 'Conectar'}
                </button>
              </div>
            </div>

            <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
              <div style={{ flex: 1, minWidth: 160 }}>
                <label style={{ fontSize: 11, color: C.textDim, display: 'block', marginBottom: 4, fontWeight: 600 }}>Projeto</label>
                <select className="ap-input" value={azProject} onChange={e => handleProjectChange(e.target.value)}>
                  <option value="">Selecione o projeto</option>
                  {projects.map(p => <option key={p.id} value={p.name}>{p.name}</option>)}
                </select>
              </div>
              <div style={{ flex: 1, minWidth: 160 }}>
                <label style={{ fontSize: 11, color: C.textDim, display: 'block', marginBottom: 4, fontWeight: 600 }}>Sprint / Iteração</label>
                <select className="ap-input" value={filterIteration} onChange={e => setFilterIteration(e.target.value)}>
                  <option value="">Todas as iterações</option>
                  {iterations.map(it => <option key={it.id} value={it.path}>{it.name}</option>)}
                </select>
              </div>
              <div style={{ flex: 1, minWidth: 160 }}>
                <label style={{ fontSize: 11, color: C.textDim, display: 'block', marginBottom: 4, fontWeight: 600 }}>Profissional</label>
                <select className="ap-input" value={filterMember} onChange={e => setFilterMember(e.target.value)}>
                  <option value="">Todos</option>
                  {members.map(m => <option key={m} value={m}>{m}</option>)}
                </select>
              </div>
              <div style={{ flex: 1, minWidth: 130 }}>
                <label style={{ fontSize: 11, color: C.textDim, display: 'block', marginBottom: 4, fontWeight: 600 }}>Data de</label>
                <input type="date" className="ap-input" value={filterDateFrom} onChange={e => setFilterDateFrom(e.target.value)} />
              </div>
              <div style={{ flex: 1, minWidth: 130 }}>
                <label style={{ fontSize: 11, color: C.textDim, display: 'block', marginBottom: 4, fontWeight: 600 }}>Data até</label>
                <input type="date" className="ap-input" value={filterDateTo} onChange={e => setFilterDateTo(e.target.value)} />
              </div>
              <div style={{ display: 'flex', alignItems: 'flex-end' }}>
                <button className="ap-btn" style={{ background: C.accent, color: '#fff' }} onClick={handleFetch} disabled={loading || !azProject}>
                  {loading ? loadingMsg || 'Carregando…' : '⟳ Buscar Dados'}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* error */}
        {error && (
          <div style={{ background: '#FEF2F2', border: '1px solid #FECACA', borderRadius: 6, padding: '10px 14px', color: C.red, marginBottom: 16, fontSize: 13 }}>
            {error}
          </div>
        )}

        {/* loading bar */}
        {loading && (
          <div style={{ background: '#FFFFFF', border: '1px solid #E8ECF5', borderRadius: 8, padding: '24px 28px', textAlign: 'center', marginBottom: 20 }}>
            <div style={{ fontSize: 13, color: C.textDim, marginBottom: 10 }}>{loadingMsg}</div>
            <div style={{ height: 3, background: '#E8ECF5', borderRadius: 2, overflow: 'hidden' }}>
              <div style={{ height: '100%', background: `linear-gradient(90deg, ${C.navy}aa, ${C.accent})`, width: '60%', animation: 'shimmer 1.5s ease-in-out infinite', borderRadius: 2 }} />
            </div>
          </div>
        )}

        {/* ── data area ── */}
        {metrics && !loading && (
          <div id="relatorio-pdf">

            {/* KPI strip */}
            <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginBottom: 24 }}>
              <KpiCard label="Work Items" value={metrics.totalItems} sub={`${metrics.completedItems} concluídos`} color={C.navy} />
              <KpiCard
                label={metrics.anyCompletedWork ? 'Esforço Total (h)' : 'Total de Cards'}
                value={metrics.anyCompletedWork ? horasToStr(metrics.totalCompletedWork) : metrics.totalItems}
                sub={metrics.anyCompletedWork ? 'CompletedWork registrado' : 'CompletedWork não preenchido'}
                color={C.accent}
              />
              <KpiCard label="Lead Time Médio" value={horasToStr(metrics.avgLeadTime)} sub="criação → conclusão" color={C.amber} />
              <KpiCard label="Cycle Time Médio" value={horasToStr(metrics.avgCycleTime)} sub="início → conclusão" color={C.navyMed} />
              <KpiCard label="Itens em Aberto" value={metrics.totalItems - metrics.completedItems} color={C.coral} />
              <KpiCard
                label="Atividade Detectada"
                value={`${metrics.totalActivityHours}h`}
                sub="horas únicas com interação"
                color={C.navyMed}
              />
              <KpiCard
                label="Cobertura CompletedWork"
                value={`${metrics.globalCwCoverage}%`}
                sub="cards com horas declaradas"
                color={metrics.globalCwCoverage >= 70 ? C.green : metrics.globalCwCoverage >= 40 ? C.amber : C.coral}
              />
            </div>

            {/* tabs */}
            <div style={{ display: 'flex', borderBottom: '1px solid #E8ECF5', marginBottom: 24 }}>
              {['dashboard', 'board', 'items', 'ai'].map((key, _, arr) => {
                const labels = { dashboard: 'Dashboard', board: 'Board & Fluxo', items: 'Work Items', ai: 'Análise IA' };
                return (
                  <button key={key} className={`tab-btn${tabActive(key) ? ' active' : ''}`} onClick={() => setTab(key)}>
                    {labels[key]}
                  </button>
                );
              })}
            </div>

            {/* ── Dashboard ── */}
            {tab === 'dashboard' && (
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 }}>

                <div className="section-card">
                  <h4>{metrics.hoursLabel} por Profissional</h4>
                  {!metrics.anyCompletedWork && (
                    <div style={{ fontSize: 11, color: C.amber, marginBottom: 8 }}>
                      Campo "CompletedWork" não preenchido — exibindo contagem de cards. Preencha as horas concluídas nos work items para ver esforço real.
                    </div>
                  )}
                  <ResponsiveContainer width="100%" height={240}>
                    <BarChart data={Object.entries(metrics.hoursByPerson).map(([name, v]) => ({ name: name.split(' ')[0], valor: +Number(v).toFixed(1) }))}>
                      <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                      <XAxis dataKey="name" tick={{ fill: C.textDim, fontSize: 11 }} />
                      <YAxis tick={{ fill: C.textDim, fontSize: 11 }} />
                      <Tooltip content={<CustomTooltip />} />
                      <Bar dataKey="valor" name={metrics.hoursLabel} fill={C.accent} radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div className="section-card">
                  <h4>Distribuição por Tipo {metrics.anyCompletedWork ? '(h)' : '(cards)'}</h4>
                  <ResponsiveContainer width="100%" height={240}>
                    <PieChart>
                      <Pie
                        data={
                          metrics.anyCompletedWork
                            ? Object.entries(metrics.hoursByType).filter(([, v]) => v > 0).map(([name, value]) => ({ name, value: +value.toFixed(1) }))
                            : Object.entries(
                                workItems.reduce((acc, wi) => {
                                  const t = wi.fields?.['System.WorkItemType'] || 'Outro';
                                  acc[t] = (acc[t] || 0) + 1;
                                  return acc;
                                }, {})
                              ).map(([name, value]) => ({ name, value }))
                        }
                        cx="50%" cy="50%" innerRadius={60} outerRadius={100}
                        dataKey="value" nameKey="name"
                      >
                        {CHART_COLORS.map((col, i) => <Cell key={i} fill={col} />)}
                      </Pie>
                      <Tooltip content={<CustomTooltip />} />
                      <Legend wrapperStyle={{ fontSize: 11, color: C.textDim }} />
                    </PieChart>
                  </ResponsiveContainer>
                </div>

                <div className="section-card">
                  <h4>Lead Time vs. Cycle Time (h)</h4>
                  <ResponsiveContainer width="100%" height={240}>
                    <ScatterChart>
                      <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                      <XAxis dataKey="leadTime" name="Lead Time" type="number" tick={{ fill: C.textDim, fontSize: 11 }}
                        label={{ value: 'Lead Time (h)', fill: C.textDim, fontSize: 11, position: 'insideBottom', offset: -4 }} />
                      <YAxis dataKey="cycleTime" name="Cycle Time" type="number" tick={{ fill: C.textDim, fontSize: 11 }}
                        label={{ value: 'Cycle Time (h)', fill: C.textDim, fontSize: 11, angle: -90, position: 'insideLeft' }} />
                      <ZAxis dataKey="interacoes" range={[40, 200]} name="Interações" />
                      <Tooltip cursor={{ strokeDasharray: '3 3' }} content={<CustomTooltip />} />
                      <Scatter data={metrics.scatter} fill={C.navy} opacity={0.65} />
                    </ScatterChart>
                  </ResponsiveContainer>
                </div>

                <div className="section-card">
                  <h4>Índice de Interação por Profissional</h4>
                  <ResponsiveContainer width="100%" height={240}>
                    <RadarChart data={Object.entries(metrics.interByPerson).map(([name, v]) => ({ name: name.split(' ')[0], interacoes: v }))}>
                      <PolarGrid stroke={C.border} />
                      <PolarAngleAxis dataKey="name" tick={{ fill: C.textDim, fontSize: 11 }} />
                      <PolarRadiusAxis tick={{ fill: C.muted, fontSize: 9 }} />
                      <Radar name="Interações" dataKey="interacoes" stroke={C.accent} fill={C.accent} fillOpacity={0.2} />
                      <Tooltip content={<CustomTooltip />} />
                    </RadarChart>
                  </ResponsiveContainer>
                </div>

                {/* ── Atividade vs CompletedWork ── */}
                <div className="section-card" style={{ gridColumn: '1 / -1' }}>
                  <h4>Atividade Detectada vs. CompletedWork por Profissional</h4>
                  <div style={{ fontSize: 11, color: C.textDim, marginBottom: 12, lineHeight: 1.6 }}>
                    <strong style={{ color: C.navyMed }}>Atividade Detectada</strong> = horas únicas em que a pessoa fez ao menos uma interação (update, comentário, mudança de estado) dentro do horário comercial — limite inferior conservador de horas trabalhadas.&nbsp;
                    <strong style={{ color: C.accent }}>CompletedWork</strong> = horas declaradas manualmente no Azure DevOps.
                    Se Atividade {'>'} CompletedWork, a pessoa trabalhou mais do que declarou.
                  </div>
                  <ResponsiveContainer width="100%" height={260}>
                    <BarChart data={metrics.activityVsCw} barGap={4}>
                      <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                      <XAxis dataKey="name" tick={{ fill: C.textDim, fontSize: 11 }} />
                      <YAxis tick={{ fill: C.textDim, fontSize: 11 }} unit="h" />
                      <Tooltip
                        content={({ active, payload, label }) => {
                          if (!active || !payload?.length) return null;
                          const d = metrics.activityVsCw.find(x => x.name === label);
                          return (
                            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 7, padding: '10px 14px', fontSize: 12, boxShadow: '0 2px 8px #00366C0e' }}>
                              <div style={{ fontWeight: 700, color: C.textBright, marginBottom: 6 }}>{d?.fullName || label}</div>
                              {payload.map((p, i) => (
                                <div key={i} style={{ color: p.color }}>{p.name}: <strong>{p.value}h</strong></div>
                              ))}
                              {d && <div style={{ color: C.muted, marginTop: 4 }}>Cobertura CW: {d.cobertura}%</div>}
                            </div>
                          );
                        }}
                      />
                      <Legend wrapperStyle={{ fontSize: 11, color: C.textDim }} />
                      <Bar dataKey="atividade" name="Atividade Detectada (h)" fill={C.navyMed} radius={[4, 4, 0, 0]} />
                      <Bar dataKey="completedWork" name="CompletedWork (h)" fill={C.accent} radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>

                  {/* tabela de cobertura */}
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12, marginTop: 16 }}>
                    <thead>
                      <tr style={{ background: C.sectionBg }}>
                        {['Profissional', 'Atividade (h)', 'CompletedWork (h)', 'Cards c/ CW', 'Cobertura', 'Divergência'].map(h => (
                          <th key={h} style={{ padding: '7px 12px', textAlign: 'left', color: C.textDim, fontWeight: 700, fontSize: 11, textTransform: 'uppercase', letterSpacing: '0.05em' }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {metrics.activityVsCw.map(row => {
                        const div = row.atividade - row.completedWork;
                        const cov = metrics.cwCoverage[row.fullName] || { filled: 0, total: 0 };
                        return (
                          <tr key={row.fullName} className="wi-row" style={{ borderBottom: `1px solid ${C.border}` }}>
                            <td style={{ padding: '7px 12px', color: C.textBright, fontWeight: 600 }}>{row.fullName}</td>
                            <td style={{ padding: '7px 12px', color: C.navyMed, fontWeight: 700 }}>{row.atividade}h</td>
                            <td style={{ padding: '7px 12px', color: C.accent }}>{row.completedWork > 0 ? `${row.completedWork}h` : '—'}</td>
                            <td style={{ padding: '7px 12px', color: C.text }}>{cov.filled}/{cov.total}</td>
                            <td style={{ padding: '7px 12px' }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                                <div style={{ flex: 1, height: 6, background: C.border, borderRadius: 3, overflow: 'hidden', minWidth: 60 }}>
                                  <div style={{ width: `${row.cobertura}%`, height: '100%', background: row.cobertura >= 70 ? C.green : row.cobertura >= 40 ? C.amber : C.coral, borderRadius: 3 }} />
                                </div>
                                <span style={{ color: row.cobertura >= 70 ? C.green : row.cobertura >= 40 ? C.amber : C.coral, fontWeight: 700, fontSize: 11 }}>{row.cobertura}%</span>
                              </div>
                            </td>
                            <td style={{ padding: '7px 12px', color: div > 8 ? C.amber : C.muted, fontWeight: div > 8 ? 700 : 400 }}>
                              {row.completedWork > 0 ? (div > 0 ? `+${div}h não declaradas` : div < 0 ? `${Math.abs(div)}h acima da atividade` : '✓ coerente') : '—'}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* ── Board & Fluxo ── */}
            {tab === 'board' && (
              <div>
                <div className="section-card">
                  <h4>Tempo Total por Coluna do Board (h)</h4>
                  <ResponsiveContainer width="100%" height={320}>
                    <BarChart
                      data={Object.entries(metrics.boardTotals).sort((a, b) => b[1] - a[1]).map(([col, h]) => ({ col, horas: +h.toFixed(1) }))}
                      layout="vertical"
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                      <XAxis type="number" tick={{ fill: C.textDim, fontSize: 11 }} />
                      <YAxis dataKey="col" type="category" width={150} tick={{ fill: C.textDim, fontSize: 11 }} />
                      <Tooltip content={<CustomTooltip />} />
                      <Bar dataKey="horas" name="Horas" fill={C.navyMed} radius={[0, 4, 4, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div className="section-card">
                  <h4>Indicador de Interação por Card</h4>
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                      <thead>
                        <tr style={{ background: C.sectionBg }}>
                          {['ID', 'Título', 'Tipo', 'Responsável', 'Interações', 'Pausas', 'Lead Time'].map(h => (
                            <th key={h} style={{ padding: '8px 12px', textAlign: 'left', color: C.textDim, fontWeight: 700, fontSize: 11, textTransform: 'uppercase', letterSpacing: '0.05em', whiteSpace: 'nowrap' }}>{h}</th>
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
                              <tr key={wi.id} className="wi-row" style={{ borderBottom: `1px solid ${C.border}` }}>
                                <td style={{ padding: '7px 12px', color: C.muted, fontWeight: 600 }}>{wi.id}</td>
                                <td style={{ padding: '7px 12px', color: C.textBright }}>{(wi.fields?.['System.Title'] || '').slice(0, 50)}</td>
                                <td style={{ padding: '7px 12px' }}>
                                  <Tag color={TYPE_COLOR[wi.fields?.['System.WorkItemType']] || C.muted}>{wi.fields?.['System.WorkItemType']}</Tag>
                                </td>
                                <td style={{ padding: '7px 12px', color: C.text }}>{wi.fields?.['System.AssignedTo']?.displayName?.split(' ')[0] || '—'}</td>
                                <td style={{ padding: '7px 12px', color: m.interacoes > 10 ? C.amber : C.textBright, fontWeight: m.interacoes > 10 ? 700 : 400 }}>{m.interacoes || 0}</td>
                                <td style={{ padding: '7px 12px', color: m.pausas?.length > 0 ? C.coral : C.muted }}>{m.pausas?.length || 0}</td>
                                <td style={{ padding: '7px 12px', color: C.text }}>{m.leadTime > 0 ? horasToStr(m.leadTime) : '—'}</td>
                              </tr>
                            );
                          })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* ── Work Items ── */}
            {tab === 'items' && (
              <div className="section-card">
                <h4>Work Items ({workItems.length})</h4>
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: C.sectionBg }}>
                        {['ID', 'Título', 'Tipo', 'Estado', 'Responsável', 'Criado', 'Fechado', 'Lead Time', 'Cycle Time', 'Horas Úteis'].map(h => (
                          <th key={h} style={{ padding: '8px 12px', textAlign: 'left', color: C.textDim, fontWeight: 700, fontSize: 11, textTransform: 'uppercase', letterSpacing: '0.05em', whiteSpace: 'nowrap' }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {workItems.map(wi => {
                        const m = metricsMap[wi.id] || {};
                        const state = wi.fields?.['System.State'] || '';
                        return (
                          <tr key={wi.id} className="wi-row" style={{ borderBottom: `1px solid ${C.border}` }}>
                            <td style={{ padding: '7px 12px', color: C.muted, fontWeight: 600 }}>{wi.id}</td>
                            <td style={{ padding: '7px 12px', maxWidth: 260, color: C.textBright }}>{(wi.fields?.['System.Title'] || '').slice(0, 55)}</td>
                            <td style={{ padding: '7px 12px' }}>
                              <Tag color={TYPE_COLOR[wi.fields?.['System.WorkItemType']] || C.muted}>{wi.fields?.['System.WorkItemType']}</Tag>
                            </td>
                            <td style={{ padding: '7px 12px' }}>
                              <Tag color={stateColor(state)}>{state}</Tag>
                            </td>
                            <td style={{ padding: '7px 12px', color: C.text }}>{wi.fields?.['System.AssignedTo']?.displayName || '—'}</td>
                            <td style={{ padding: '7px 12px', color: C.muted, whiteSpace: 'nowrap' }}>{fmtDate(wi.fields?.['System.CreatedDate'])}</td>
                            <td style={{ padding: '7px 12px', color: C.muted, whiteSpace: 'nowrap' }}>{fmtDate(wi.fields?.['Microsoft.VSTS.Common.ClosedDate'])}</td>
                            <td style={{ padding: '7px 12px', color: C.text }}>{m.leadTime > 0 ? horasToStr(m.leadTime) : '—'}</td>
                            <td style={{ padding: '7px 12px', color: C.text }}>{m.cycleTime > 0 ? horasToStr(m.cycleTime) : '—'}</td>
                            <td style={{ padding: '7px 12px', color: C.text }}>{m.completedWork > 0 ? horasToStr(m.completedWork) : '—'}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* ── Análise IA ── */}
            {tab === 'ai' && (
              <div>
                <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 16 }}>
                  <button className="ap-btn" style={{ background: C.navy, color: '#fff' }} onClick={handleAiAnalysis} disabled={aiLoading || !anthropicKey}>
                    {aiLoading ? 'Analisando…' : '✦ Gerar Análise com Claude Haiku'}
                  </button>
                  {!anthropicKey && <span style={{ color: C.amber, fontSize: 12 }}>Configure a chave Anthropic para usar a análise IA.</span>}
                </div>
                {aiReport && (
                  <div className="section-card" style={{ whiteSpace: 'pre-wrap', fontSize: 13, lineHeight: 1.8, color: C.textBright }}>
                    {aiReport}
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* empty state */}
        {!metrics && !loading && (
          <div style={{ background: '#FFFFFF', border: '1px solid #E8ECF5', borderRadius: 8, padding: '60px 28px', textAlign: 'center' }}>
            <div style={{ width: 40, height: 40, borderRadius: '50%', background: 'linear-gradient(135deg, #00366C, #39ADE3)', margin: '0 auto 16px' }} />
            <div style={{ fontFamily: "'Manrope',sans-serif", fontWeight: 700, fontSize: 16, color: C.textBright, marginBottom: 6 }}>Agente de Produtividade</div>
            <div style={{ fontSize: 13, color: C.textDim }}>Configure a conexão com o Azure DevOps e clique em <strong>Buscar Dados</strong> para iniciar a análise.</div>
          </div>
        )}
      </div>
    </div>
  );
}
