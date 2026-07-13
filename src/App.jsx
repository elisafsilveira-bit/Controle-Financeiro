import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell, XAxis, YAxis,
  CartesianGrid, Tooltip, Legend, ResponsiveContainer, AreaChart, Area
} from "recharts";
import {
  LayoutDashboard, FileText, Scale, Wallet, Settings, Building2,
  TrendingUp, TrendingDown, Plus, Trash2, Lock, User, Eye, EyeOff,
  LogOut, Calendar, AlertCircle, Check, X, Landmark
} from "lucide-react";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";

/* ============================== CONSTANTES ============================== */

const CATEGORIAS = [
  { id: "receita_vendas", label: "Receita de Vendas", grupo: "receitaBruta", tipo: "entrada", dre: true },
  { id: "impostos_vendas", label: "Impostos sobre Vendas", grupo: "deducoes", tipo: "saida", dre: true },
  { id: "cmv", label: "Custo de Mercadorias/Serviços (CMV)", grupo: "cmv", tipo: "saida", dre: true },
  { id: "pessoal", label: "Despesas com Pessoal", grupo: "opPessoal", tipo: "saida", dre: true },
  { id: "administrativas", label: "Despesas Administrativas", grupo: "opAdmin", tipo: "saida", dre: true },
  { id: "comerciais", label: "Despesas Comerciais/Marketing", grupo: "opComercial", tipo: "saida", dre: true },
  { id: "despesas_financeiras", label: "Despesas Financeiras/Juros", grupo: "despFin", tipo: "saida", dre: true },
  { id: "receitas_financeiras", label: "Receitas Financeiras", grupo: "recFin", tipo: "entrada", dre: true },
  { id: "ircsll", label: "IR / CSLL", grupo: "ircsll", tipo: "saida", dre: true },
  { id: "aporte_socios", label: "Aporte de Sócios", grupo: null, tipo: "entrada", dre: false },
  { id: "emprestimo_entrada", label: "Empréstimo Recebido", grupo: null, tipo: "entrada", dre: false },
  { id: "emprestimo_pagamento", label: "Pagamento de Empréstimo", grupo: null, tipo: "saida", dre: false },
  { id: "investimento", label: "Investimento / Imobilizado", grupo: null, tipo: "saida", dre: false },
  { id: "distribuicao_lucros", label: "Distribuição de Lucros", grupo: null, tipo: "saida", dre: false },
  { id: "outras_entradas", label: "Outras Entradas", grupo: null, tipo: "entrada", dre: false },
  { id: "outras_saidas", label: "Outras Saídas", grupo: null, tipo: "saida", dre: false },
];

const MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
const MESES_LONGOS = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];

const CAMPOS_BALANCO = [
  { grupo: "Ativo Circulante", key: "AC", campos: [
    { id: "caixa", label: "Caixa e Equivalentes" },
    { id: "contasReceber", label: "Contas a Receber" },
    { id: "estoques", label: "Estoques" },
    { id: "outrosAC", label: "Outros Ativos Circulantes" },
  ]},
  { grupo: "Ativo Não Circulante", key: "ANC", campos: [
    { id: "imobilizado", label: "Imobilizado" },
    { id: "intangivel", label: "Intangível" },
    { id: "outrosANC", label: "Outros Ativos Não Circulantes" },
  ]},
  { grupo: "Passivo Circulante", key: "PC", campos: [
    { id: "fornecedores", label: "Fornecedores" },
    { id: "emprestimosCP", label: "Empréstimos (Curto Prazo)" },
    { id: "obrigacoesTrabalhistas", label: "Obrigações Trabalhistas" },
    { id: "impostosPagar", label: "Impostos a Pagar" },
    { id: "outrosPC", label: "Outros Passivos Circulantes" },
  ]},
  { grupo: "Passivo Não Circulante", key: "PNC", campos: [
    { id: "emprestimosLP", label: "Empréstimos (Longo Prazo)" },
    { id: "outrosPNC", label: "Outros Passivos Não Circulantes" },
  ]},
  { grupo: "Patrimônio Líquido", key: "PL", campos: [
    { id: "capitalSocial", label: "Capital Social" },
    { id: "lucrosAcumulados", label: "Lucros Acumulados" },
    { id: "reservas", label: "Reservas" },
  ]},
];

const BALANCO_VAZIO = Object.fromEntries(
  CAMPOS_BALANCO.flatMap(g => g.campos.map(c => [c.id, 0]))
);

const COLORS = { teal: "#1F6F5C", gold: "#B8863B", red: "#A8433A", ink: "#17241F", blue: "#3C6E8F" };
const PIE_COLORS = ["#1F6F5C", "#B8863B", "#3C6E8F", "#A8433A", "#7A8B5C", "#8C6D46", "#5C7A8B", "#9C5C6D"];

const EMPRESAS_PADRAO = {
  a: { nome: "Empresa A", cnpj: "00.000.000/0001-00" },
  b: { nome: "Empresa B", cnpj: "00.000.000/0002-00" },
};

/* ============================== HELPERS ============================== */

const fmtBRL = (v) => new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(v || 0);
const fmtBRLCompact = (v) => {
  const abs = Math.abs(v || 0);
  if (abs >= 1000000) return (v / 1000000).toFixed(1).replace(".", ",") + "M";
  if (abs >= 1000) return (v / 1000).toFixed(0) + "k";
  return String(Math.round(v || 0));
};
function catInfo(id) { return CATEGORIAS.find((c) => c.id === id) || {}; }

/* ---------- Importação de planilha Excel ---------- */

function normalizeStr(s) {
  return String(s || "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().trim();
}

function findCategoriaByLabel(raw) {
  const alvo = normalizeStr(raw);
  if (!alvo) return null;
  return CATEGORIAS.find((c) => normalizeStr(c.label) === alvo) ||
    CATEGORIAS.find((c) => normalizeStr(c.label).includes(alvo) || alvo.includes(normalizeStr(c.label))) ||
    null;
}

function excelSerialToDate(serial) {
  // Excel conta dias a partir de 1899-12-30
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  return new Date(utcValue * 1000);
}

function parseDataCell(v) {
  if (v instanceof Date && !isNaN(v)) {
    const y = v.getFullYear(), m = String(v.getMonth() + 1).padStart(2, "0"), d = String(v.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (typeof v === "number") {
    const d = excelSerialToDate(v);
    if (!isNaN(d)) return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
  }
  const s = String(v || "").trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return `${m[3]}-${String(m[2]).padStart(2, "0")}-${String(m[1]).padStart(2, "0")}`;
  return null;
}

function parseValorCell(v) {
  if (typeof v === "number") return v;
  let s = String(v || "").trim().replace(/[R$\s]/g, "");
  if (!s) return null;
  // trata formato brasileiro 1.234,56
  if (/,\d{1,2}$/.test(s)) s = s.replace(/\./g, "").replace(",", ".");
  else s = s.replace(/,/g, "");
  const n = Number(s);
  return isNaN(n) ? null : n;
}

function parseEmpresaCell(v, empresas) {
  const alvo = normalizeStr(v);
  if (alvo === "a" || alvo === normalizeStr(empresas.a.nome)) return "a";
  if (alvo === "b" || alvo === normalizeStr(empresas.b.nome)) return "b";
  return null;
}

function getCell(row, ...nomes) {
  const keys = Object.keys(row);
  for (const nome of nomes) {
    const alvo = normalizeStr(nome);
    const achou = keys.find((k) => normalizeStr(k) === alvo);
    if (achou !== undefined) return row[achou];
  }
  return undefined;
}

function baixarModeloExcel(empresas) {
  const linhas = [
    ["Data", "Empresa", "Categoria", "Valor", "Descrição"],
    ["2026-07-05", "A", "Receita de Vendas", 15000, "Faturamento do mês"],
    ["2026-07-06", "A", "Impostos sobre Vendas", 900, "Impostos sobre vendas"],
    ["2026-07-10", "B", "Despesas com Pessoal", 8000, "Folha de pagamento"],
  ];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(linhas);
  ws["!cols"] = [{ wch: 12 }, { wch: 10 }, { wch: 32 }, { wch: 12 }, { wch: 34 }];
  XLSX.utils.book_append_sheet(wb, ws, "Lançamentos");
  const wsCat = XLSX.utils.aoa_to_sheet([
    ["Categorias aceitas (copie e cole exatamente como está aqui)"],
    ...CATEGORIAS.map((c) => [c.label]),
  ]);
  wsCat["!cols"] = [{ wch: 42 }];
  XLSX.utils.book_append_sheet(wb, wsCat, "Categorias válidas");
  XLSX.writeFile(wb, "modelo-lancamentos.xlsx");
}

function parseLancamentosExcel(arrayBuffer, empresas) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const validos = [];
  const erros = [];
  rows.forEach((row, idx) => {
    const linhaNum = idx + 2; // +2 = considerando cabeçalho na linha 1
    const dataRaw = getCell(row, "Data");
    const empresaRaw = getCell(row, "Empresa", "CNPJ");
    const categoriaRaw = getCell(row, "Categoria");
    const valorRaw = getCell(row, "Valor");
    const descricaoRaw = getCell(row, "Descrição", "Descricao", "Obs", "Observação");

    if (!dataRaw && !empresaRaw && !categoriaRaw && !valorRaw) return; // linha em branco

    const data = parseDataCell(dataRaw);
    const cnpj = parseEmpresaCell(empresaRaw, empresas);
    const categoria = findCategoriaByLabel(categoriaRaw);
    const valor = parseValorCell(valorRaw);

    const problemas = [];
    if (!data) problemas.push("data inválida");
    if (!cnpj) problemas.push("empresa não reconhecida (use A ou B)");
    if (!categoria) problemas.push(`categoria não reconhecida ("${categoriaRaw}")`);
    if (valor === null || valor <= 0) problemas.push("valor inválido");

    if (problemas.length) {
      erros.push({ linha: linhaNum, motivo: problemas.join(", ") });
    } else {
      validos.push({
        id: Date.now() + idx,
        cnpj, data, categoriaId: categoria.id, valor,
        descricao: String(descricaoRaw || "").trim() || categoria.label,
      });
    }
  });
  return { validos, erros };
}

/* ============================== BANCO DE DADOS (Supabase) ==============================
   Usa uma tabela simples chave/valor "app_storage" (chave text PK, valor jsonb).
   Isso guarda 3 registros: "lancamentos", "balancos" e "empresas".
   Veja o arquivo supabase.sql para criar a tabela e as políticas de acesso. */

async function loadKey(key, fallback) {
  const { data, error } = await supabase.from("app_storage").select("valor").eq("chave", key).maybeSingle();
  if (error) { console.error("Erro ao carregar", key, error); return fallback; }
  return data ? data.valor : fallback;
}
async function saveKey(key, value) {
  const { error } = await supabase.from("app_storage").upsert({ chave: key, valor: value, atualizado_em: new Date().toISOString() });
  if (error) console.error("Erro ao salvar", key, error);
}

/* ============================== APP ============================== */

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = carregando, null = deslogado
  const [ready, setReady] = useState(false);
  const [view, setView] = useState("dashboard");
  const [cnpjSel, setCnpjSel] = useState("consolidado");
  const [entries, setEntries] = useState([]);
  const [balancos, setBalancos] = useState({ a: {}, b: {} });
  const [empresas, setEmpresas] = useState(EMPRESAS_PADRAO);
  const [selectedYear, setSelectedYear] = useState(new Date().getFullYear());
  const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth() + 1);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => listener.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) return;
    (async () => {
      setReady(false);
      const [e, b, emp] = await Promise.all([
        loadKey("lancamentos", []),
        loadKey("balancos", { a: {}, b: {} }),
        loadKey("empresas", EMPRESAS_PADRAO),
      ]);
      setEntries(e || []);
      setBalancos(b || { a: {}, b: {} });
      setEmpresas(emp || EMPRESAS_PADRAO);
      setReady(true);
    })();
  }, [session]);

  const persistEntries = useCallback(async (next) => { setEntries(next); await saveKey("lancamentos", next); }, []);
  const persistBalancos = useCallback(async (next) => { setBalancos(next); await saveKey("balancos", next); }, []);
  const persistEmpresas = useCallback(async (next) => { setEmpresas(next); await saveKey("empresas", next); }, []);

  if (session === undefined) {
    return <div className="loading-full">Carregando…</div>;
  }
  if (!session) {
    return <LoginScreen />;
  }

  const user = {
    nome: session.user.user_metadata?.nome || session.user.email,
    cargo: session.user.user_metadata?.cargo || "Usuário",
    email: session.user.email,
  };

  return (
    <div className="app-root">
      <style>{GLOBAL_CSS}</style>
      <Sidebar view={view} setView={setView} user={user}
        onLogout={() => supabase.auth.signOut()}
        mobileNavOpen={mobileNavOpen} setMobileNavOpen={setMobileNavOpen} />
      <main className="main-area">
        <TopBar empresas={empresas} cnpjSel={cnpjSel} setCnpjSel={setCnpjSel}
          user={user} setMobileNavOpen={setMobileNavOpen} />
        {!ready ? (
          <div className="loading-state">Carregando dados…</div>
        ) : (
          <div className="view-body">
            {view === "dashboard" && (
              <Dashboard entries={entries} cnpjSel={cnpjSel} empresas={empresas}
                selectedYear={selectedYear} selectedMonth={selectedMonth}
                setSelectedYear={setSelectedYear} setSelectedMonth={setSelectedMonth} />
            )}
            {view === "dre" && (
              <DREView entries={entries} cnpjSel={cnpjSel}
                selectedYear={selectedYear} selectedMonth={selectedMonth}
                setSelectedYear={setSelectedYear} setSelectedMonth={setSelectedMonth} />
            )}
            {view === "balanco" && (
              <BalancoView balancos={balancos} cnpjSel={cnpjSel} empresas={empresas}
                selectedYear={selectedYear} setSelectedYear={setSelectedYear}
                persistBalancos={persistBalancos} />
            )}
            {view === "fluxo" && (
              <FluxoCaixaView entries={entries} cnpjSel={cnpjSel}
                selectedYear={selectedYear} selectedMonth={selectedMonth}
                setSelectedYear={setSelectedYear} setSelectedMonth={setSelectedMonth} />
            )}
            {view === "lancamentos" && (
              <LancamentosView entries={entries} empresas={empresas} persistEntries={persistEntries} />
            )}
            {view === "config" && (
              <ConfiguracoesView empresas={empresas} persistEmpresas={persistEmpresas} user={user} />
            )}
          </div>
        )}
      </main>
    </div>
  );
}

/* ============================== LOGIN ============================== */

function LoginScreen() {
  const [email, setEmail] = useState("");
  const [p, setP] = useState("");
  const [showPass, setShowPass] = useState(false);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setLoading(true); setError("");
    const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password: p });
    setLoading(false);
    if (error) setError("E-mail ou senha inválidos.");
  };

  return (
    <div className="login-wrap">
      <style>{GLOBAL_CSS}</style>
      <div className="login-card">
        <div className="login-brand">
          <div className="login-mark"><Landmark size={22} strokeWidth={1.6} /></div>
          <div>
            <div className="login-title">Livro-Razão</div>
            <div className="login-sub">Controle Financeiro Multiempresa</div>
          </div>
        </div>
        <form onSubmit={submit} className="login-form">
          <label className="field-label">E-mail</label>
          <div className="input-icon-wrap">
            <User size={16} className="input-icon" />
            <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="seu@email.com" autoFocus required />
          </div>
          <label className="field-label">Senha</label>
          <div className="input-icon-wrap">
            <Lock size={16} className="input-icon" />
            <input type={showPass ? "text" : "password"} value={p} onChange={(e) => setP(e.target.value)} placeholder="sua senha" required />
            <button type="button" className="input-icon-btn" onClick={() => setShowPass((s) => !s)}>
              {showPass ? <EyeOff size={16} /> : <Eye size={16} />}
            </button>
          </div>
          {error && <div className="login-error"><AlertCircle size={14} /> {error}</div>}
          <button type="submit" className="btn-primary login-submit" disabled={loading}>{loading ? "Entrando…" : "Entrar"}</button>
        </form>
        <div className="login-footnote">
          Acesso restrito aos usuários cadastrados no painel do Supabase. Fale com quem administra
          o sistema caso não tenha um login ainda.
        </div>
      </div>
    </div>
  );
}

/* ============================== SIDEBAR / TOPBAR ============================== */

const NAV_ITEMS = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "dre", label: "DRE", icon: FileText },
  { id: "balanco", label: "Balanço Patrimonial", icon: Scale },
  { id: "fluxo", label: "Fluxo de Caixa", icon: Wallet },
  { id: "lancamentos", label: "Lançamentos", icon: Plus },
  { id: "config", label: "Configurações", icon: Settings },
];

function Sidebar({ view, setView, user, onLogout, mobileNavOpen, setMobileNavOpen }) {
  return (
    <>
      <aside className={"sidebar" + (mobileNavOpen ? " sidebar-open" : "")}>
        <div className="sidebar-brand">
          <div className="login-mark small"><Landmark size={18} strokeWidth={1.6} /></div>
          <div>
            <div className="sidebar-title">Livro-Razão</div>
            <div className="sidebar-sub">Controle Financeiro</div>
          </div>
        </div>
        <nav className="sidebar-nav">
          {NAV_ITEMS.map((item) => (
            <button key={item.id} className={"nav-item" + (view === item.id ? " nav-item-active" : "")}
              onClick={() => { setView(item.id); setMobileNavOpen(false); }}>
              <item.icon size={17} strokeWidth={1.8} /><span>{item.label}</span>
            </button>
          ))}
        </nav>
        <div className="sidebar-footer">
          <div className="sidebar-user">
            <div className="avatar">{user.nome.split(" ").map(w=>w[0]).slice(0,2).join("").toUpperCase()}</div>
            <div>
              <div className="sidebar-user-name">{user.nome}</div>
              <div className="sidebar-user-role">{user.cargo}</div>
            </div>
          </div>
          <button className="nav-item logout-item" onClick={onLogout}>
            <LogOut size={17} strokeWidth={1.8} /> <span>Sair</span>
          </button>
        </div>
      </aside>
      {mobileNavOpen && <div className="sidebar-backdrop" onClick={() => setMobileNavOpen(false)} />}
    </>
  );
}

function TopBar({ empresas, cnpjSel, setCnpjSel, user, setMobileNavOpen }) {
  return (
    <div className="topbar">
      <button className="hamburger" onClick={() => setMobileNavOpen(true)}>☰</button>
      <div className="cnpj-tabs">
        <button className={"cnpj-tab" + (cnpjSel === "a" ? " cnpj-tab-active" : "")} onClick={() => setCnpjSel("a")}>
          <Building2 size={14} /> {empresas.a.nome}
        </button>
        <button className={"cnpj-tab" + (cnpjSel === "b" ? " cnpj-tab-active" : "")} onClick={() => setCnpjSel("b")}>
          <Building2 size={14} /> {empresas.b.nome}
        </button>
        <button className={"cnpj-tab cnpj-tab-consol" + (cnpjSel === "consolidado" ? " cnpj-tab-active" : "")} onClick={() => setCnpjSel("consolidado")}>
          <Scale size={14} /> Consolidado
        </button>
      </div>
      <div className="topbar-right">
        <span className="topbar-user-tag">{user.nome} · {user.cargo}</span>
      </div>
    </div>
  );
}

/* ============================== CÁLCULOS ============================== */

function filterEntries(entries, cnpjSel) {
  if (cnpjSel === "consolidado") return entries;
  return entries.filter((e) => e.cnpj === cnpjSel);
}
function signedValue(entry) {
  const c = catInfo(entry.categoriaId);
  return c.tipo === "entrada" ? entry.valor : -entry.valor;
}
function computeDRE(entries, cnpjSel, year, month) {
  const filtered = filterEntries(entries, cnpjSel).filter((e) => {
    const [y, m] = e.data.split("-").map(Number);
    if (month) return y === year && m === month;
    return y === year;
  });
  const sums = {};
  filtered.forEach((e) => {
    const c = catInfo(e.categoriaId);
    if (!c.dre) return;
    sums[c.grupo] = (sums[c.grupo] || 0) + e.valor;
  });
  const receitaBruta = sums.receitaBruta || 0;
  const deducoes = sums.deducoes || 0;
  const receitaLiquida = receitaBruta - deducoes;
  const cmv = sums.cmv || 0;
  const lucroBruto = receitaLiquida - cmv;
  const opPessoal = sums.opPessoal || 0;
  const opAdmin = sums.opAdmin || 0;
  const opComercial = sums.opComercial || 0;
  const despesasOperacionais = opPessoal + opAdmin + opComercial;
  const resultadoOperacional = lucroBruto - despesasOperacionais;
  const recFin = sums.recFin || 0;
  const despFin = sums.despFin || 0;
  const resultadoAntesIR = resultadoOperacional + recFin - despFin;
  const ircsll = sums.ircsll || 0;
  const lucroLiquido = resultadoAntesIR - ircsll;
  return { receitaBruta, deducoes, receitaLiquida, cmv, lucroBruto, opPessoal, opAdmin, opComercial,
    despesasOperacionais, resultadoOperacional, recFin, despFin, resultadoAntesIR, ircsll, lucroLiquido };
}
function computeCaixaAcumulado(entries, cnpjSel, ateData) {
  const filtered = filterEntries(entries, cnpjSel).filter((e) => e.data <= ateData);
  return filtered.reduce((acc, e) => acc + signedValue(e), 0);
}
function anosDisponiveis(entries) {
  const set = new Set(entries.map((e) => Number(e.data.split("-")[0])));
  set.add(new Date().getFullYear());
  return Array.from(set).sort();
}

/* ============================== DASHBOARD ============================== */

function Dashboard({ entries, cnpjSel, selectedYear, selectedMonth, setSelectedYear, setSelectedMonth }) {
  const anos = anosDisponiveis(entries);
  const dreAtual = useMemo(() => computeDRE(entries, cnpjSel, selectedYear, selectedMonth), [entries, cnpjSel, selectedYear, selectedMonth]);
  const hojeStr = `${selectedYear}-${String(selectedMonth).padStart(2, "0")}-31`;
  const saldoCaixa = useMemo(() => computeCaixaAcumulado(entries, cnpjSel, hojeStr), [entries, cnpjSel, hojeStr]);

  const trend = useMemo(() => {
    const arr = [];
    for (let i = 5; i >= 0; i--) {
      let m = selectedMonth - i, y = selectedYear;
      while (m <= 0) { m += 12; y -= 1; }
      const dre = computeDRE(entries, cnpjSel, y, m);
      const despesas = dre.deducoes + dre.cmv + dre.despesasOperacionais + dre.despFin + dre.ircsll;
      arr.push({ label: `${MESES[m - 1]}/${String(y).slice(2)}`, Faturamento: dre.receitaBruta, Despesas: despesas, Lucro: dre.lucroLiquido });
    }
    return arr;
  }, [entries, cnpjSel, selectedYear, selectedMonth]);

  const breakdown = useMemo(() => {
    const filtered = filterEntries(entries, cnpjSel).filter((e) => {
      const [y, m] = e.data.split("-").map(Number);
      return y === selectedYear && m === selectedMonth && catInfo(e.categoriaId).tipo === "saida";
    });
    const map = {};
    filtered.forEach((e) => { const label = catInfo(e.categoriaId).label; map[label] = (map[label] || 0) + e.valor; });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [entries, cnpjSel, selectedYear, selectedMonth]);

  const despesasTotal = dreAtual.deducoes + dreAtual.cmv + dreAtual.despesasOperacionais + dreAtual.despFin + dreAtual.ircsll;
  const margem = dreAtual.receitaBruta ? (dreAtual.lucroLiquido / dreAtual.receitaBruta) * 100 : 0;

  return (
    <div className="stack">
      <PeriodBar anos={anos} selectedYear={selectedYear} selectedMonth={selectedMonth}
        setSelectedYear={setSelectedYear} setSelectedMonth={setSelectedMonth} showMonth />
      <div className="card-grid">
        <MetricCard label="Faturamento do mês" value={fmtBRL(dreAtual.receitaBruta)} icon={TrendingUp} tone="teal" />
        <MetricCard label="Despesas do mês" value={fmtBRL(despesasTotal)} icon={TrendingDown} tone="red" />
        <MetricCard label="Lucro líquido" value={fmtBRL(dreAtual.lucroLiquido)} icon={Scale} tone={dreAtual.lucroLiquido >= 0 ? "teal" : "red"} sub={`Margem: ${margem.toFixed(1)}%`} />
        <MetricCard label="Saldo em caixa (acumulado)" value={fmtBRL(saldoCaixa)} icon={Wallet} tone="gold" />
      </div>
      <div className="chart-grid">
        <div className="panel">
          <div className="panel-title">Faturamento x Despesas — últimos 6 meses</div>
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={trend} barGap={4}>
              <CartesianGrid strokeDasharray="3 3" stroke="#D8DED4" vertical={false} />
              <XAxis dataKey="label" tick={{ fontSize: 12, fill: "#3C4A42" }} axisLine={{ stroke: "#C9CFC3" }} tickLine={false} />
              <YAxis tickFormatter={fmtBRLCompact} tick={{ fontSize: 11, fill: "#3C4A42" }} axisLine={false} tickLine={false} width={50} />
              <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ fontFamily: "IBM Plex Mono, monospace", fontSize: 12, borderRadius: 8, border: "1px solid #C9CFC3" }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="Faturamento" fill={COLORS.teal} radius={[3, 3, 0, 0]} />
              <Bar dataKey="Despesas" fill={COLORS.red} radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="panel">
          <div className="panel-title">Composição das despesas — {MESES_LONGOS[selectedMonth - 1]}/{selectedYear}</div>
          {breakdown.length === 0 ? (
            <div className="empty-note">Sem despesas lançadas neste período.</div>
          ) : (
            <ResponsiveContainer width="100%" height={280}>
              <PieChart>
                <Pie data={breakdown} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={95} innerRadius={50} paddingAngle={2}>
                  {breakdown.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                </Pie>
                <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ fontFamily: "IBM Plex Mono, monospace", fontSize: 12, borderRadius: 8, border: "1px solid #C9CFC3" }} />
                <Legend wrapperStyle={{ fontSize: 11 }} layout="vertical" align="right" verticalAlign="middle" />
              </PieChart>
            </ResponsiveContainer>
          )}
        </div>
      </div>
      <div className="panel">
        <div className="panel-title">Lucro líquido — tendência</div>
        <ResponsiveContainer width="100%" height={220}>
          <AreaChart data={trend}>
            <defs>
              <linearGradient id="lucroGrad" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor={COLORS.teal} stopOpacity={0.35} />
                <stop offset="95%" stopColor={COLORS.teal} stopOpacity={0} />
              </linearGradient>
            </defs>
            <CartesianGrid strokeDasharray="3 3" stroke="#D8DED4" vertical={false} />
            <XAxis dataKey="label" tick={{ fontSize: 12, fill: "#3C4A42" }} axisLine={{ stroke: "#C9CFC3" }} tickLine={false} />
            <YAxis tickFormatter={fmtBRLCompact} tick={{ fontSize: 11, fill: "#3C4A42" }} axisLine={false} tickLine={false} width={50} />
            <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ fontFamily: "IBM Plex Mono, monospace", fontSize: 12, borderRadius: 8, border: "1px solid #C9CFC3" }} />
            <Area type="monotone" dataKey="Lucro" stroke={COLORS.teal} strokeWidth={2} fill="url(#lucroGrad)" />
          </AreaChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function MetricCard({ label, value, icon: Icon, tone, sub }) {
  return (
    <div className={"metric-card tone-" + tone}>
      <div className="metric-icon"><Icon size={17} strokeWidth={1.8} /></div>
      <div className="metric-label">{label}</div>
      <div className="metric-value">{value}</div>
      {sub && <div className="metric-sub">{sub}</div>}
    </div>
  );
}

function PeriodBar({ anos, selectedYear, selectedMonth, setSelectedYear, setSelectedMonth, showMonth }) {
  return (
    <div className="period-bar">
      <Calendar size={15} className="period-icon" />
      <select value={selectedYear} onChange={(e) => setSelectedYear(Number(e.target.value))}>
        {anos.map((y) => <option key={y} value={y}>{y}</option>)}
      </select>
      {showMonth && (
        <select value={selectedMonth} onChange={(e) => setSelectedMonth(Number(e.target.value))}>
          {MESES_LONGOS.map((m, i) => <option key={m} value={i + 1}>{m}</option>)}
        </select>
      )}
    </div>
  );
}

/* ============================== DRE ============================== */

function DREView({ entries, cnpjSel, selectedYear, selectedMonth, setSelectedYear, setSelectedMonth }) {
  const anos = anosDisponiveis(entries);
  const [modo, setModo] = useState("mensal");
  return (
    <div className="stack">
      <div className="view-header">
        <h2>Demonstração do Resultado do Exercício</h2>
        <div className="toggle-group">
          <button className={modo === "mensal" ? "toggle-active" : ""} onClick={() => setModo("mensal")}>Mensal</button>
          <button className={modo === "anual" ? "toggle-active" : ""} onClick={() => setModo("anual")}>Anual</button>
        </div>
      </div>
      <PeriodBar anos={anos} selectedYear={selectedYear} selectedMonth={selectedMonth}
        setSelectedYear={setSelectedYear} setSelectedMonth={setSelectedMonth} showMonth={modo === "mensal"} />
      {modo === "mensal" ? (
        <DREMensal entries={entries} cnpjSel={cnpjSel} year={selectedYear} month={selectedMonth} />
      ) : (
        <DREAnual entries={entries} cnpjSel={cnpjSel} year={selectedYear} />
      )}
    </div>
  );
}

const DRE_LINHAS = [
  { key: "receitaBruta", label: "Receita Bruta de Vendas", bold: false, indent: 0 },
  { key: "deducoes", label: "(–) Deduções / Impostos sobre Vendas", bold: false, indent: 0 },
  { key: "receitaLiquida", label: "= Receita Líquida", bold: true, indent: 0, divider: true },
  { key: "cmv", label: "(–) Custo das Mercadorias/Serviços (CMV)", bold: false, indent: 0 },
  { key: "lucroBruto", label: "= Lucro Bruto", bold: true, indent: 0, divider: true },
  { key: "opPessoal", label: "(–) Despesas com Pessoal", bold: false, indent: 1 },
  { key: "opAdmin", label: "(–) Despesas Administrativas", bold: false, indent: 1 },
  { key: "opComercial", label: "(–) Despesas Comerciais/Marketing", bold: false, indent: 1 },
  { key: "resultadoOperacional", label: "= Resultado Operacional", bold: true, indent: 0, divider: true },
  { key: "recFin", label: "(+) Receitas Financeiras", bold: false, indent: 0 },
  { key: "despFin", label: "(–) Despesas Financeiras", bold: false, indent: 0 },
  { key: "resultadoAntesIR", label: "= Resultado Antes do IR/CSLL", bold: true, indent: 0, divider: true },
  { key: "ircsll", label: "(–) IR / CSLL", bold: false, indent: 0 },
  { key: "lucroLiquido", label: "= LUCRO LÍQUIDO DO PERÍODO", bold: true, indent: 0, divider: true, highlight: true },
];

function DREMensal({ entries, cnpjSel, year, month }) {
  const dre = useMemo(() => computeDRE(entries, cnpjSel, year, month), [entries, cnpjSel, year, month]);
  return (
    <div className="panel ledger-panel">
      <div className="panel-title">{MESES_LONGOS[month - 1]} de {year}</div>
      <table className="ledger-table">
        <tbody>
          {DRE_LINHAS.map((l) => (
            <tr key={l.key} className={(l.bold ? "row-bold " : "") + (l.divider ? "row-divider " : "") + (l.highlight ? "row-highlight" : "")}>
              <td style={{ paddingLeft: 12 + l.indent * 18 }}>{l.label}</td>
              <td className="num-cell">{fmtBRL(dre[l.key])}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function DREAnual({ entries, cnpjSel, year }) {
  const dados = useMemo(() => {
    const porMes = [];
    for (let m = 1; m <= 12; m++) porMes.push(computeDRE(entries, cnpjSel, year, m));
    const anual = computeDRE(entries, cnpjSel, year, null);
    return { porMes, anual };
  }, [entries, cnpjSel, year]);
  return (
    <div className="panel ledger-panel table-scroll">
      <div className="panel-title">Ano de {year}</div>
      <table className="ledger-table ledger-table-wide">
        <thead>
          <tr>
            <th>Linha</th>
            {MESES.map((m) => <th key={m} className="num-cell">{m}</th>)}
            <th className="num-cell total-col">Total</th>
          </tr>
        </thead>
        <tbody>
          {DRE_LINHAS.map((l) => (
            <tr key={l.key} className={(l.bold ? "row-bold " : "") + (l.divider ? "row-divider " : "") + (l.highlight ? "row-highlight" : "")}>
              <td style={{ paddingLeft: 12 + l.indent * 18 }}>{l.label}</td>
              {dados.porMes.map((d, i) => <td key={i} className="num-cell">{fmtBRLCompact(d[l.key])}</td>)}
              <td className="num-cell total-col">{fmtBRL(dados.anual[l.key])}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* ============================== BALANÇO PATRIMONIAL ============================== */

function BalancoView({ balancos, cnpjSel, empresas, selectedYear, setSelectedYear, persistBalancos }) {
  const anos = useMemo(() => {
    const set = new Set([
      ...Object.keys(balancos.a || {}).map(Number),
      ...Object.keys(balancos.b || {}).map(Number),
      new Date().getFullYear(),
    ]);
    return Array.from(set).sort();
  }, [balancos]);

  const [editing, setEditing] = useState(false);
  const [editCnpj, setEditCnpj] = useState("a");
  const [editYear, setEditYear] = useState(selectedYear);
  const [form, setForm] = useState(BALANCO_VAZIO);

  const openEdit = (cnpj) => {
    setEditCnpj(cnpj); setEditYear(selectedYear);
    setForm({ ...BALANCO_VAZIO, ...(balancos[cnpj]?.[selectedYear] || {}) });
    setEditing(true);
  };
  const saveEdit = async () => {
    const next = { ...balancos, [editCnpj]: { ...(balancos[editCnpj] || {}), [editYear]: form } };
    await persistBalancos(next);
    setEditing(false);
  };

  const dadosA = balancos.a?.[selectedYear] || null;
  const dadosB = balancos.b?.[selectedYear] || null;
  const consolidar = (a, b) => {
    if (!a && !b) return null;
    const base = { ...BALANCO_VAZIO };
    Object.keys(base).forEach((k) => { base[k] = (a?.[k] || 0) + (b?.[k] || 0); });
    return base;
  };
  const dadosMostrados = cnpjSel === "a" ? dadosA : cnpjSel === "b" ? dadosB : consolidar(dadosA, dadosB);

  return (
    <div className="stack">
      <div className="view-header">
        <h2>Balanço Patrimonial</h2>
        {cnpjSel !== "consolidado" && <button className="btn-primary" onClick={() => openEdit(cnpjSel)}>Editar {empresas[cnpjSel].nome}</button>}
      </div>
      <PeriodBar anos={anos} selectedYear={selectedYear} setSelectedYear={setSelectedYear} showMonth={false} />
      {cnpjSel === "consolidado" && (
        <div className="toggle-group" style={{ marginBottom: 4 }}>
          <button className="toggle-active" style={{cursor:"default"}}>Editar dados individuais:</button>
          <button onClick={() => openEdit("a")}>{empresas.a.nome}</button>
          <button onClick={() => openEdit("b")}>{empresas.b.nome}</button>
        </div>
      )}
      {!dadosMostrados ? (
        <div className="panel"><div className="empty-note">Nenhum balanço cadastrado para {selectedYear}. Use o botão de edição para lançar os valores.</div></div>
      ) : <BalancoTabela dados={dadosMostrados} />}
      {editing && (
        <div className="modal-backdrop" onClick={() => setEditing(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div>Balanço — {empresas[editCnpj].nome}</div>
              <button className="icon-btn" onClick={() => setEditing(false)}><X size={16} /></button>
            </div>
            <div className="modal-body">
              <label className="field-label">Ano de referência (posição em 31/12)</label>
              <input type="number" value={editYear} onChange={(e) => setEditYear(Number(e.target.value))} className="modal-year-input" />
              {CAMPOS_BALANCO.map((grupo) => (
                <div key={grupo.key} className="balanco-form-group">
                  <div className="balanco-form-title">{grupo.grupo}</div>
                  {grupo.campos.map((c) => (
                    <div key={c.id} className="balanco-form-row">
                      <span>{c.label}</span>
                      <input type="number" value={form[c.id]} onChange={(e) => setForm({ ...form, [c.id]: Number(e.target.value) })} />
                    </div>
                  ))}
                </div>
              ))}
            </div>
            <div className="modal-footer">
              <button className="btn-secondary" onClick={() => setEditing(false)}>Cancelar</button>
              <button className="btn-primary" onClick={saveEdit}><Check size={14} /> Salvar Balanço</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function BalancoTabela({ dados }) {
  const ativoCirculante = dados.caixa + dados.contasReceber + dados.estoques + dados.outrosAC;
  const ativoNaoCirculante = dados.imobilizado + dados.intangivel + dados.outrosANC;
  const ativoTotal = ativoCirculante + ativoNaoCirculante;
  const passivoCirculante = dados.fornecedores + dados.emprestimosCP + dados.obrigacoesTrabalhistas + dados.impostosPagar + dados.outrosPC;
  const passivoNaoCirculante = dados.emprestimosLP + dados.outrosPNC;
  const patrimonioLiquido = dados.capitalSocial + dados.lucrosAcumulados + dados.reservas;
  const passivoTotal = passivoCirculante + passivoNaoCirculante + patrimonioLiquido;
  const diferenca = ativoTotal - passivoTotal;
  return (
    <div className="chart-grid">
      <div className="panel ledger-panel">
        <div className="panel-title">Ativo</div>
        <table className="ledger-table">
          <tbody>
            <tr className="row-subtitle"><td colSpan={2}>Ativo Circulante</td></tr>
            <BalancoLinha label="Caixa e Equivalentes" v={dados.caixa} />
            <BalancoLinha label="Contas a Receber" v={dados.contasReceber} />
            <BalancoLinha label="Estoques" v={dados.estoques} />
            <BalancoLinha label="Outros" v={dados.outrosAC} />
            <tr className="row-bold row-divider"><td>= Total Ativo Circulante</td><td className="num-cell">{fmtBRL(ativoCirculante)}</td></tr>
            <tr className="row-subtitle"><td colSpan={2}>Ativo Não Circulante</td></tr>
            <BalancoLinha label="Imobilizado" v={dados.imobilizado} />
            <BalancoLinha label="Intangível" v={dados.intangivel} />
            <BalancoLinha label="Outros" v={dados.outrosANC} />
            <tr className="row-bold row-divider"><td>= Total Ativo Não Circulante</td><td className="num-cell">{fmtBRL(ativoNaoCirculante)}</td></tr>
            <tr className="row-highlight row-bold row-divider"><td>ATIVO TOTAL</td><td className="num-cell">{fmtBRL(ativoTotal)}</td></tr>
          </tbody>
        </table>
      </div>
      <div className="panel ledger-panel">
        <div className="panel-title">Passivo e Patrimônio Líquido</div>
        <table className="ledger-table">
          <tbody>
            <tr className="row-subtitle"><td colSpan={2}>Passivo Circulante</td></tr>
            <BalancoLinha label="Fornecedores" v={dados.fornecedores} />
            <BalancoLinha label="Empréstimos (CP)" v={dados.emprestimosCP} />
            <BalancoLinha label="Obrigações Trabalhistas" v={dados.obrigacoesTrabalhistas} />
            <BalancoLinha label="Impostos a Pagar" v={dados.impostosPagar} />
            <BalancoLinha label="Outros" v={dados.outrosPC} />
            <tr className="row-bold row-divider"><td>= Total Passivo Circulante</td><td className="num-cell">{fmtBRL(passivoCirculante)}</td></tr>
            <tr className="row-subtitle"><td colSpan={2}>Passivo Não Circulante</td></tr>
            <BalancoLinha label="Empréstimos (LP)" v={dados.emprestimosLP} />
            <BalancoLinha label="Outros" v={dados.outrosPNC} />
            <tr className="row-bold row-divider"><td>= Total Passivo Não Circulante</td><td className="num-cell">{fmtBRL(passivoNaoCirculante)}</td></tr>
            <tr className="row-subtitle"><td colSpan={2}>Patrimônio Líquido</td></tr>
            <BalancoLinha label="Capital Social" v={dados.capitalSocial} />
            <BalancoLinha label="Lucros Acumulados" v={dados.lucrosAcumulados} />
            <BalancoLinha label="Reservas" v={dados.reservas} />
            <tr className="row-bold row-divider"><td>= Total Patrimônio Líquido</td><td className="num-cell">{fmtBRL(patrimonioLiquido)}</td></tr>
            <tr className="row-highlight row-bold row-divider"><td>PASSIVO + PL TOTAL</td><td className="num-cell">{fmtBRL(passivoTotal)}</td></tr>
          </tbody>
        </table>
        {Math.abs(diferenca) > 0.5 && <div className="balance-check"><AlertCircle size={13} /> Diferença Ativo − Passivo/PL: {fmtBRL(diferenca)}</div>}
      </div>
    </div>
  );
}
function BalancoLinha({ label, v }) { return <tr><td style={{ paddingLeft: 12 }}>{label}</td><td className="num-cell">{fmtBRL(v)}</td></tr>; }

/* ============================== FLUXO DE CAIXA ============================== */

function FluxoCaixaView({ entries, cnpjSel, selectedYear, selectedMonth, setSelectedYear, setSelectedMonth }) {
  const anos = anosDisponiveis(entries);
  const [modo, setModo] = useState("mensal");
  return (
    <div className="stack">
      <div className="view-header">
        <h2>Fluxo de Caixa</h2>
        <div className="toggle-group">
          <button className={modo === "diario" ? "toggle-active" : ""} onClick={() => setModo("diario")}>Diário</button>
          <button className={modo === "mensal" ? "toggle-active" : ""} onClick={() => setModo("mensal")}>Mensal</button>
          <button className={modo === "anual" ? "toggle-active" : ""} onClick={() => setModo("anual")}>Anual</button>
        </div>
      </div>
      <PeriodBar anos={anos} selectedYear={selectedYear} selectedMonth={selectedMonth}
        setSelectedYear={setSelectedYear} setSelectedMonth={setSelectedMonth} showMonth={modo === "diario"} />
      {modo === "diario" && <FluxoDiario entries={entries} cnpjSel={cnpjSel} year={selectedYear} month={selectedMonth} />}
      {modo === "mensal" && <FluxoMensal entries={entries} cnpjSel={cnpjSel} year={selectedYear} />}
      {modo === "anual" && <FluxoAnual entries={entries} cnpjSel={cnpjSel} anos={anos} />}
    </div>
  );
}

function FluxoDiario({ entries, cnpjSel, year, month }) {
  const filtered = filterEntries(entries, cnpjSel).filter((e) => {
    const [y, m] = e.data.split("-").map(Number);
    return y === year && m === month;
  });
  const inicioMes = `${year}-${String(month).padStart(2, "0")}-01`;
  const saldoInicial = useMemo(() => {
    const antes = filterEntries(entries, cnpjSel).filter((e) => e.data < inicioMes);
    return antes.reduce((acc, e) => acc + signedValue(e), 0);
  }, [entries, cnpjSel, inicioMes]);
  const porDia = {};
  filtered.forEach((e) => {
    const dia = e.data;
    if (!porDia[dia]) porDia[dia] = { entradas: 0, saidas: 0 };
    const c = catInfo(e.categoriaId);
    if (c.tipo === "entrada") porDia[dia].entradas += e.valor; else porDia[dia].saidas += e.valor;
  });
  const dias = Object.keys(porDia).sort();
  let acumulado = saldoInicial;
  const linhas = dias.map((d) => {
    const { entradas, saidas } = porDia[d];
    acumulado += entradas - saidas;
    return { data: d, entradas, saidas, saldoDia: entradas - saidas, acumulado };
  });
  const chartData = linhas.map((l) => ({ label: l.data.slice(8, 10), Saldo: l.acumulado }));
  return (
    <div className="stack">
      <div className="panel">
        <div className="panel-title">Saldo acumulado no mês</div>
        <ResponsiveContainer width="100%" height={220}>
          <LineChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" stroke="#D8DED4" vertical={false} />
            <XAxis dataKey="label" tick={{ fontSize: 11, fill: "#3C4A42" }} axisLine={{ stroke: "#C9CFC3" }} tickLine={false} />
            <YAxis tickFormatter={fmtBRLCompact} tick={{ fontSize: 11, fill: "#3C4A42" }} axisLine={false} tickLine={false} width={50} />
            <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ fontFamily: "IBM Plex Mono, monospace", fontSize: 12, borderRadius: 8, border: "1px solid #C9CFC3" }} />
            <Line type="monotone" dataKey="Saldo" stroke={COLORS.blue} strokeWidth={2} dot={{ r: 2 }} />
          </LineChart>
        </ResponsiveContainer>
      </div>
      <div className="panel ledger-panel table-scroll">
        <div className="panel-title">Saldo inicial do mês: {fmtBRL(saldoInicial)}</div>
        {linhas.length === 0 ? <div className="empty-note">Sem lançamentos neste mês.</div> : (
          <table className="ledger-table">
            <thead><tr><th>Data</th><th className="num-cell">Entradas</th><th className="num-cell">Saídas</th><th className="num-cell">Saldo do Dia</th><th className="num-cell">Acumulado</th></tr></thead>
            <tbody>
              {linhas.map((l) => (
                <tr key={l.data}>
                  <td>{l.data.split("-").reverse().join("/")}</td>
                  <td className="num-cell positive">{fmtBRL(l.entradas)}</td>
                  <td className="num-cell negative">{fmtBRL(l.saidas)}</td>
                  <td className={"num-cell " + (l.saldoDia >= 0 ? "positive" : "negative")}>{fmtBRL(l.saldoDia)}</td>
                  <td className="num-cell row-bold">{fmtBRL(l.acumulado)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function FluxoMensal({ entries, cnpjSel, year }) {
  const inicioAno = `${year}-01-01`;
  const saldoInicial = useMemo(() => {
    const antes = filterEntries(entries, cnpjSel).filter((e) => e.data < inicioAno);
    return antes.reduce((acc, e) => acc + signedValue(e), 0);
  }, [entries, cnpjSel, inicioAno]);
  let acumulado = saldoInicial;
  const linhas = [];
  for (let m = 1; m <= 12; m++) {
    const mm = String(m).padStart(2, "0");
    const filtered = filterEntries(entries, cnpjSel).filter((e) => e.data.startsWith(`${year}-${mm}`));
    let entradas = 0, saidas = 0;
    filtered.forEach((e) => { const c = catInfo(e.categoriaId); if (c.tipo === "entrada") entradas += e.valor; else saidas += e.valor; });
    acumulado += entradas - saidas;
    linhas.push({ mes: MESES[m - 1], entradas, saidas, saldoMes: entradas - saidas, acumulado });
  }
  return (
    <div className="stack">
      <div className="panel">
        <div className="panel-title">Entradas x Saídas por mês — {year}</div>
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={linhas}>
            <CartesianGrid strokeDasharray="3 3" stroke="#D8DED4" vertical={false} />
            <XAxis dataKey="mes" tick={{ fontSize: 12, fill: "#3C4A42" }} axisLine={{ stroke: "#C9CFC3" }} tickLine={false} />
            <YAxis tickFormatter={fmtBRLCompact} tick={{ fontSize: 11, fill: "#3C4A42" }} axisLine={false} tickLine={false} width={50} />
            <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ fontFamily: "IBM Plex Mono, monospace", fontSize: 12, borderRadius: 8, border: "1px solid #C9CFC3" }} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Bar dataKey="entradas" name="Entradas" fill={COLORS.teal} radius={[3, 3, 0, 0]} />
            <Bar dataKey="saidas" name="Saídas" fill={COLORS.red} radius={[3, 3, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div className="panel ledger-panel">
        <div className="panel-title">Saldo inicial do ano: {fmtBRL(saldoInicial)}</div>
        <table className="ledger-table">
          <thead><tr><th>Mês</th><th className="num-cell">Entradas</th><th className="num-cell">Saídas</th><th className="num-cell">Saldo do Mês</th><th className="num-cell">Acumulado</th></tr></thead>
          <tbody>
            {linhas.map((l) => (
              <tr key={l.mes}>
                <td>{l.mes}</td>
                <td className="num-cell positive">{fmtBRL(l.entradas)}</td>
                <td className="num-cell negative">{fmtBRL(l.saidas)}</td>
                <td className={"num-cell " + (l.saldoMes >= 0 ? "positive" : "negative")}>{fmtBRL(l.saldoMes)}</td>
                <td className="num-cell row-bold">{fmtBRL(l.acumulado)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function FluxoAnual({ entries, cnpjSel, anos }) {
  let acumulado = 0;
  const linhas = anos.map((year) => {
    const filtered = filterEntries(entries, cnpjSel).filter((e) => e.data.startsWith(`${year}`));
    let entradas = 0, saidas = 0;
    filtered.forEach((e) => { const c = catInfo(e.categoriaId); if (c.tipo === "entrada") entradas += e.valor; else saidas += e.valor; });
    acumulado += entradas - saidas;
    return { ano: year, entradas, saidas, saldoAno: entradas - saidas, acumulado };
  });
  return (
    <div className="stack">
      <div className="panel">
        <div className="panel-title">Entradas x Saídas por ano</div>
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={linhas}>
            <CartesianGrid strokeDasharray="3 3" stroke="#D8DED4" vertical={false} />
            <XAxis dataKey="ano" tick={{ fontSize: 12, fill: "#3C4A42" }} axisLine={{ stroke: "#C9CFC3" }} tickLine={false} />
            <YAxis tickFormatter={fmtBRLCompact} tick={{ fontSize: 11, fill: "#3C4A42" }} axisLine={false} tickLine={false} width={50} />
            <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ fontFamily: "IBM Plex Mono, monospace", fontSize: 12, borderRadius: 8, border: "1px solid #C9CFC3" }} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Bar dataKey="entradas" name="Entradas" fill={COLORS.teal} radius={[3, 3, 0, 0]} />
            <Bar dataKey="saidas" name="Saídas" fill={COLORS.red} radius={[3, 3, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div className="panel ledger-panel">
        <table className="ledger-table">
          <thead><tr><th>Ano</th><th className="num-cell">Entradas</th><th className="num-cell">Saídas</th><th className="num-cell">Saldo do Ano</th><th className="num-cell">Acumulado</th></tr></thead>
          <tbody>
            {linhas.map((l) => (
              <tr key={l.ano}>
                <td>{l.ano}</td>
                <td className="num-cell positive">{fmtBRL(l.entradas)}</td>
                <td className="num-cell negative">{fmtBRL(l.saidas)}</td>
                <td className={"num-cell " + (l.saldoAno >= 0 ? "positive" : "negative")}>{fmtBRL(l.saldoAno)}</td>
                <td className="num-cell row-bold">{fmtBRL(l.acumulado)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================== LANÇAMENTOS ============================== */

function LancamentosView({ entries, empresas, persistEntries }) {
  const [cnpj, setCnpj] = useState("a");
  const [data, setData] = useState(new Date().toISOString().slice(0, 10));
  const [categoriaId, setCategoriaId] = useState(CATEGORIAS[0].id);
  const [valor, setValor] = useState("");
  const [descricao, setDescricao] = useState("");
  const [filtroCnpj, setFiltroCnpj] = useState("todos");

  const addEntry = async (e) => {
    e.preventDefault();
    if (!valor || Number(valor) <= 0) return;
    const nova = { id: Date.now(), cnpj, data, categoriaId, valor: Number(valor), descricao: descricao || catInfo(categoriaId).label };
    await persistEntries([nova, ...entries]);
    setValor(""); setDescricao("");
  };
  const removeEntry = async (id) => { await persistEntries(entries.filter((e) => e.id !== id)); };
  const importarEmLote = async (novosLancamentos) => { await persistEntries([...novosLancamentos, ...entries]); };

  const listados = entries
    .filter((e) => filtroCnpj === "todos" || e.cnpj === filtroCnpj)
    .sort((a, b) => (a.data < b.data ? 1 : -1))
    .slice(0, 200);

  return (
    <div className="stack">
      <h2 className="view-title-only">Lançamentos</h2>
      <ImportarExcel empresas={empresas} onImport={importarEmLote} />
      <div className="panel">
        <div className="panel-title">Novo lançamento</div>
        <form className="entry-form" onSubmit={addEntry}>
          <div>
            <label className="field-label">Empresa</label>
            <select value={cnpj} onChange={(e) => setCnpj(e.target.value)}>
              <option value="a">{empresas.a.nome}</option>
              <option value="b">{empresas.b.nome}</option>
            </select>
          </div>
          <div>
            <label className="field-label">Data</label>
            <input type="date" value={data} onChange={(e) => setData(e.target.value)} required />
          </div>
          <div>
            <label className="field-label">Categoria</label>
            <select value={categoriaId} onChange={(e) => setCategoriaId(e.target.value)}>
              {CATEGORIAS.map((c) => <option key={c.id} value={c.id}>{c.label} {c.tipo === "entrada" ? "↑" : "↓"}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Valor (R$)</label>
            <input type="number" min="0" step="0.01" value={valor} onChange={(e) => setValor(e.target.value)} placeholder="0,00" required />
          </div>
          <div className="entry-form-desc">
            <label className="field-label">Descrição (opcional)</label>
            <input value={descricao} onChange={(e) => setDescricao(e.target.value)} placeholder="Ex.: NF 1234, pagamento fornecedor X…" />
          </div>
          <button type="submit" className="btn-primary entry-form-submit"><Plus size={15} /> Adicionar</button>
        </form>
      </div>
      <div className="panel">
        <div className="panel-title-row">
          <div className="panel-title">Últimos lançamentos</div>
          <select value={filtroCnpj} onChange={(e) => setFiltroCnpj(e.target.value)} className="filter-select">
            <option value="todos">Todas as empresas</option>
            <option value="a">{empresas.a.nome}</option>
            <option value="b">{empresas.b.nome}</option>
          </select>
        </div>
        <div className="table-scroll">
          <table className="ledger-table">
            <thead><tr><th>Data</th><th>Empresa</th><th>Categoria</th><th>Descrição</th><th className="num-cell">Valor</th><th></th></tr></thead>
            <tbody>
              {listados.map((e) => {
                const c = catInfo(e.categoriaId);
                return (
                  <tr key={e.id}>
                    <td>{e.data.split("-").reverse().join("/")}</td>
                    <td>{empresas[e.cnpj]?.nome}</td>
                    <td>{c.label}</td>
                    <td>{e.descricao}</td>
                    <td className={"num-cell " + (c.tipo === "entrada" ? "positive" : "negative")}>{c.tipo === "entrada" ? "+" : "−"}{fmtBRL(e.valor)}</td>
                    <td><button className="icon-btn" onClick={() => removeEntry(e.id)}><Trash2 size={14} /></button></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {listados.length === 0 && <div className="empty-note">Nenhum lançamento encontrado.</div>}
        </div>
      </div>
    </div>
  );
}

function ImportarExcel({ empresas, onImport }) {
  const [aberto, setAberto] = useState(false);
  const [preview, setPreview] = useState(null); // { validos, erros, nomeArquivo }
  const [processando, setProcessando] = useState(false);
  const [importando, setImportando] = useState(false);
  const [concluido, setConcluido] = useState(null);
  const inputRef = React.useRef(null);

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setProcessando(true);
    setConcluido(null);
    try {
      const buffer = await file.arrayBuffer();
      const { validos, erros } = parseLancamentosExcel(buffer, empresas);
      setPreview({ validos, erros, nomeArquivo: file.name });
    } catch (err) {
      console.error(err);
      setPreview({ validos: [], erros: [{ linha: "-", motivo: "Não foi possível ler este arquivo. Confirme que é um .xlsx ou .csv válido." }], nomeArquivo: file.name });
    }
    setProcessando(false);
    if (inputRef.current) inputRef.current.value = "";
  };

  const confirmarImportacao = async () => {
    if (!preview || preview.validos.length === 0) return;
    setImportando(true);
    await onImport(preview.validos);
    setImportando(false);
    setConcluido(preview.validos.length);
    setPreview(null);
  };

  return (
    <div className="panel">
      <div className="panel-title-row">
        <div className="panel-title">Importar lançamentos por planilha Excel</div>
        <button className="btn-secondary" onClick={() => setAberto((s) => !s)}>
          {aberto ? "Ocultar" : "Abrir"}
        </button>
      </div>
      {aberto && (
        <div className="import-box">
          <p className="import-text">
            Baixe o modelo, preencha uma linha para cada lançamento (usando exatamente os
            nomes de categoria da aba "Categorias válidas" do modelo) e depois envie o
            arquivo preenchido aqui.
          </p>
          <div className="import-actions">
            <button className="btn-secondary" onClick={() => baixarModeloExcel(empresas)}>
              <FileSpreadsheetIcon /> Baixar modelo (.xlsx)
            </button>
            <label className="btn-primary import-upload-btn">
              <Plus size={14} /> Selecionar planilha
              <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} hidden />
            </label>
          </div>

          {processando && <div className="empty-note">Lendo arquivo…</div>}

          {concluido !== null && (
            <div className="import-success"><Check size={14} /> {concluido} lançamento(s) importado(s) com sucesso.</div>
          )}

          {preview && (
            <div className="import-preview">
              <div className="import-preview-summary">
                <span><strong>{preview.nomeArquivo}</strong></span>
                <span className="positive">{preview.validos.length} linha(s) válida(s)</span>
                {preview.erros.length > 0 && <span className="negative">{preview.erros.length} linha(s) com problema</span>}
              </div>

              {preview.erros.length > 0 && (
                <div className="table-scroll">
                  <table className="ledger-table">
                    <thead><tr><th>Linha</th><th>Problema</th></tr></thead>
                    <tbody>
                      {preview.erros.map((e, i) => <tr key={i}><td>{e.linha}</td><td>{e.motivo}</td></tr>)}
                    </tbody>
                  </table>
                </div>
              )}

              {preview.validos.length > 0 && (
                <div className="table-scroll">
                  <table className="ledger-table">
                    <thead><tr><th>Data</th><th>Empresa</th><th>Categoria</th><th>Descrição</th><th className="num-cell">Valor</th></tr></thead>
                    <tbody>
                      {preview.validos.slice(0, 15).map((e) => {
                        const c = catInfo(e.categoriaId);
                        return (
                          <tr key={e.id}>
                            <td>{e.data.split("-").reverse().join("/")}</td>
                            <td>{empresas[e.cnpj]?.nome}</td>
                            <td>{c.label}</td>
                            <td>{e.descricao}</td>
                            <td className={"num-cell " + (c.tipo === "entrada" ? "positive" : "negative")}>{fmtBRL(e.valor)}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                  {preview.validos.length > 15 && <div className="empty-note">…e mais {preview.validos.length - 15} linha(s).</div>}
                </div>
              )}

              <div className="import-confirm-row">
                <button className="btn-secondary" onClick={() => setPreview(null)}>Cancelar</button>
                <button className="btn-primary" disabled={preview.validos.length === 0 || importando} onClick={confirmarImportacao}>
                  {importando ? "Importando…" : `Confirmar importação de ${preview.validos.length} lançamento(s)`}
                </button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function FileSpreadsheetIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
      <path d="M14 2v6h6" />
      <path d="M8 13h8M8 17h8M8 9h1" />
    </svg>
  );
}

/* ============================== CONFIGURAÇÕES ============================== */

function ConfiguracoesView({ empresas, persistEmpresas, user }) {
  const [form, setForm] = useState(empresas);
  const [saved, setSaved] = useState(false);
  const save = async () => { await persistEmpresas(form); setSaved(true); setTimeout(() => setSaved(false), 1800); };

  return (
    <div className="stack">
      <h2 className="view-title-only">Configurações</h2>
      <div className="panel">
        <div className="panel-title">Empresas cadastradas</div>
        {["a", "b"].map((k) => (
          <div key={k} className="config-row">
            <div className="config-row-title">Empresa {k.toUpperCase()}</div>
            <div className="config-fields">
              <div>
                <label className="field-label">Razão social / Nome</label>
                <input value={form[k].nome} onChange={(e) => setForm({ ...form, [k]: { ...form[k], nome: e.target.value } })} />
              </div>
              <div>
                <label className="field-label">CNPJ</label>
                <input value={form[k].cnpj} onChange={(e) => setForm({ ...form, [k]: { ...form[k], cnpj: e.target.value } })} placeholder="00.000.000/0001-00" />
              </div>
            </div>
          </div>
        ))}
        <button className="btn-primary" onClick={save}>{saved ? <><Check size={14} /> Salvo</> : "Salvar Empresas"}</button>
      </div>
      <div className="panel">
        <div className="panel-title">Sua conta</div>
        <div className="empty-note" style={{ fontStyle: "normal", textAlign: "left" }}>
          Logado como <strong>{user.email}</strong>. Os 4 usuários do sistema são gerenciados no
          painel do Supabase (Authentication → Users) — é lá que você cria, remove ou redefine
          senhas de acesso.
        </div>
      </div>
    </div>
  );
}

/* ============================== CSS ============================== */

const GLOBAL_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
:root { --paper:#EEF0EA; --surface:#FFFFFF; --ink:#17241F; --ink-soft:#3C4A42; --line:#C9CFC3; --teal:#1F6F5C; --gold:#B8863B; --red:#A8433A; --blue:#3C6E8F; --sidebar:#12201B; }
* { box-sizing: border-box; }
html, body, #root { height: 100%; margin: 0; }
.app-root, .login-wrap { font-family:'Inter',sans-serif; color:var(--ink); background:var(--paper); }
.loading-full { min-height:100vh; display:flex; align-items:center; justify-content:center; font-family:'IBM Plex Mono',monospace; color:var(--ink-soft); }
.app-root { display:flex; min-height:100vh; }
.sidebar { width:232px; flex-shrink:0; background:var(--sidebar); color:#E8EDE6; display:flex; flex-direction:column; padding:20px 14px; position:sticky; top:0; height:100vh; }
.sidebar-brand { display:flex; align-items:center; gap:10px; padding:4px 6px 22px; border-bottom:1px solid rgba(255,255,255,0.1); margin-bottom:14px; }
.login-mark { width:36px; height:36px; border-radius:8px; background:var(--teal); color:#fff; display:flex; align-items:center; justify-content:center; flex-shrink:0; }
.login-mark.small { width:30px; height:30px; }
.sidebar-title { font-family:'Fraunces',serif; font-size:16px; font-weight:600; line-height:1.1; }
.sidebar-sub { font-size:10.5px; color:#A9B8A6; letter-spacing:.03em; text-transform:uppercase; margin-top:2px; }
.sidebar-nav { display:flex; flex-direction:column; gap:2px; flex:1; }
.nav-item { display:flex; align-items:center; gap:10px; padding:9px 10px; border-radius:7px; border:none; background:transparent; color:#C9D4C6; font-family:'Inter',sans-serif; font-size:13.5px; cursor:pointer; text-align:left; transition:background .15s,color .15s; }
.nav-item:hover { background:rgba(255,255,255,0.07); color:#fff; }
.nav-item-active { background:var(--teal); color:#fff; }
.sidebar-footer { border-top:1px solid rgba(255,255,255,0.1); padding-top:12px; margin-top:8px; }
.sidebar-user { display:flex; align-items:center; gap:9px; padding:6px; margin-bottom:6px; }
.avatar { width:30px; height:30px; border-radius:50%; background:var(--gold); color:#1a1a1a; font-size:11px; font-weight:700; display:flex; align-items:center; justify-content:center; }
.sidebar-user-name { font-size:12.5px; font-weight:600; }
.sidebar-user-role { font-size:10.5px; color:#A9B8A6; }
.logout-item { color:#E8B4AC; }
.sidebar-backdrop { display:none; }
.main-area { flex:1; display:flex; flex-direction:column; min-width:0; }
.topbar { display:flex; align-items:center; gap:14px; padding:14px 26px; background:var(--surface); border-bottom:1px solid var(--line); position:sticky; top:0; z-index:5; flex-wrap:wrap; }
.hamburger { display:none; background:none; border:none; font-size:18px; cursor:pointer; color:var(--ink); }
.cnpj-tabs { display:flex; gap:6px; flex-wrap:wrap; }
.cnpj-tab { display:flex; align-items:center; gap:6px; padding:7px 13px; border-radius:999px; border:1px solid var(--line); background:var(--paper); color:var(--ink-soft); font-size:12.5px; font-weight:500; cursor:pointer; transition:all .15s; }
.cnpj-tab:hover { border-color:var(--teal); }
.cnpj-tab-active { background:var(--ink); color:#fff; border-color:var(--ink); }
.cnpj-tab-consol.cnpj-tab-active { background:var(--teal); border-color:var(--teal); }
.topbar-right { margin-left:auto; }
.topbar-user-tag { font-size:12px; color:var(--ink-soft); font-family:'IBM Plex Mono',monospace; }
.view-body { padding:22px 26px 40px; }
.loading-state { padding:60px; text-align:center; color:var(--ink-soft); font-family:'IBM Plex Mono',monospace; }
.stack { display:flex; flex-direction:column; gap:18px; }
.view-header { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px; }
.view-header h2, .view-title-only { font-family:'Fraunces',serif; font-size:22px; font-weight:600; margin:0; }
.period-bar { display:flex; align-items:center; gap:8px; }
.period-icon { color:var(--ink-soft); }
.period-bar select { padding:7px 10px; border-radius:6px; border:1px solid var(--line); background:var(--surface); font-family:'IBM Plex Mono',monospace; font-size:12.5px; color:var(--ink); cursor:pointer; }
.toggle-group { display:flex; gap:2px; background:var(--surface); border:1px solid var(--line); border-radius:8px; padding:3px; }
.toggle-group button { border:none; background:transparent; padding:6px 14px; border-radius:6px; font-size:12.5px; font-weight:500; color:var(--ink-soft); cursor:pointer; }
.toggle-active { background:var(--ink) !important; color:#fff !important; }
.card-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:14px; }
.metric-card { background:var(--surface); border:1px solid var(--line); border-radius:10px; padding:16px 18px; border-top:3px solid var(--teal); }
.metric-card.tone-red { border-top-color:var(--red); }
.metric-card.tone-gold { border-top-color:var(--gold); }
.metric-icon { color:var(--ink-soft); margin-bottom:8px; }
.metric-label { font-size:11.5px; color:var(--ink-soft); text-transform:uppercase; letter-spacing:.03em; margin-bottom:6px; }
.metric-value { font-family:'IBM Plex Mono',monospace; font-size:20px; font-weight:600; font-variant-numeric:tabular-nums; }
.metric-sub { font-size:11px; color:var(--ink-soft); margin-top:4px; }
.chart-grid { display:grid; grid-template-columns:1fr 1fr; gap:14px; }
.panel { background:var(--surface); border:1px solid var(--line); border-radius:10px; padding:18px 20px; }
.panel-title { font-family:'Fraunces',serif; font-size:14.5px; font-weight:600; margin-bottom:12px; color:var(--ink); }
.panel-title-row { display:flex; align-items:center; justify-content:space-between; margin-bottom:4px; flex-wrap:wrap; gap:8px; }
.empty-note { color:var(--ink-soft); font-size:13px; padding:20px 4px; text-align:center; font-style:italic; }
.ledger-panel { padding:18px 0 18px 0; }
.ledger-panel .panel-title { padding:0 20px; }
.table-scroll { overflow-x:auto; }
.ledger-table { width:100%; border-collapse:collapse; font-size:13px; }
.ledger-table th { text-align:left; font-size:10.5px; text-transform:uppercase; letter-spacing:.04em; color:var(--ink-soft); padding:8px 12px; border-bottom:1px solid var(--line); background:var(--paper); }
.ledger-table td { padding:7px 12px; border-bottom:1px solid #E4E7DF; font-family:'IBM Plex Mono',monospace; }
.ledger-table td:first-child { font-family:'Inter',sans-serif; }
.num-cell { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
.row-bold td { font-weight:700; }
.row-divider td { border-bottom:1px solid var(--ink); }
.row-highlight td { background:#E4EFE9; }
.row-subtitle td { padding-top:14px; font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:var(--ink-soft); border-bottom:none; }
.positive { color:var(--teal); }
.negative { color:var(--red); }
.total-col { border-left:2px solid var(--ink); background:#F5F3EC; }
.ledger-table-wide th, .ledger-table-wide td { padding:6px 8px; font-size:12px; }
.balance-check { display:flex; align-items:center; gap:6px; margin:12px 20px 0; padding:8px 10px; background:#F7EDE3; border:1px solid var(--gold); border-radius:6px; font-size:12px; color:#7A5A20; }
.field-label { display:block; font-size:11px; color:var(--ink-soft); margin-bottom:4px; font-weight:500; }
input, select { font-family:'Inter',sans-serif; font-size:13px; padding:8px 10px; border-radius:6px; border:1px solid var(--line); background:var(--surface); color:var(--ink); width:100%; }
input:focus, select:focus { outline:2px solid var(--teal); outline-offset:1px; border-color:var(--teal); }
.btn-primary { display:inline-flex; align-items:center; gap:6px; background:var(--ink); color:#fff; border:none; padding:9px 16px; border-radius:7px; font-size:13px; font-weight:600; cursor:pointer; }
.btn-primary:hover { background:var(--teal); }
.btn-primary:disabled { opacity:.6; cursor:default; }
.btn-secondary { background:var(--surface); border:1px solid var(--line); padding:9px 16px; border-radius:7px; font-size:13px; cursor:pointer; }
.icon-btn { background:none; border:none; cursor:pointer; color:var(--ink-soft); padding:4px; border-radius:5px; }
.icon-btn:hover { background:#F0E4E2; color:var(--red); }
.entry-form { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; align-items:end; }
.entry-form-desc { grid-column:span 2; }
.entry-form-submit { grid-column:span 1; height:38px; justify-content:center; }
.filter-select { width:auto; }
.import-box { display:flex; flex-direction:column; gap:12px; margin-top:6px; }
.import-text { font-size:12.5px; color:var(--ink-soft); margin:0; line-height:1.5; }
.import-actions { display:flex; gap:10px; flex-wrap:wrap; }
.import-upload-btn { cursor:pointer; }
.import-success { display:flex; align-items:center; gap:6px; background:#E4EFE9; border:1px solid var(--teal); color:#155444; padding:8px 12px; border-radius:6px; font-size:12.5px; }
.import-preview { display:flex; flex-direction:column; gap:10px; border-top:1px solid var(--line); padding-top:12px; }
.import-preview-summary { display:flex; gap:14px; flex-wrap:wrap; font-size:12.5px; }
.import-confirm-row { display:flex; justify-content:flex-end; gap:10px; }
.config-row { padding:12px 0; border-bottom:1px solid var(--line); }
.config-row:last-of-type { border-bottom:none; }
.config-row-title { font-weight:600; font-size:13px; margin-bottom:8px; }
.config-fields { display:grid; grid-template-columns:2fr 1fr; gap:12px; margin-bottom:6px; }
.modal-backdrop { position:fixed; inset:0; background:rgba(23,36,31,0.5); display:flex; align-items:center; justify-content:center; z-index:50; padding:20px; }
.modal-card { background:var(--surface); border-radius:12px; width:100%; max-width:560px; max-height:88vh; display:flex; flex-direction:column; }
.modal-header { display:flex; align-items:center; justify-content:space-between; padding:16px 20px; border-bottom:1px solid var(--line); font-family:'Fraunces',serif; font-weight:600; font-size:15px; }
.modal-body { padding:16px 20px; overflow-y:auto; }
.modal-footer { display:flex; justify-content:flex-end; gap:10px; padding:14px 20px; border-top:1px solid var(--line); }
.modal-year-input { max-width:140px; margin-bottom:14px; }
.balanco-form-group { margin-bottom:14px; }
.balanco-form-title { font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:var(--ink-soft); margin-bottom:6px; border-bottom:1px solid var(--line); padding-bottom:4px; }
.balanco-form-row { display:flex; align-items:center; justify-content:space-between; gap:10px; padding:4px 0; }
.balanco-form-row span { font-size:12.5px; flex:1; }
.balanco-form-row input { max-width:150px; text-align:right; font-family:'IBM Plex Mono',monospace; }
.login-wrap { min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px; background:linear-gradient(160deg, var(--sidebar) 0%, #1B2C24 55%, var(--paper) 55%); }
.login-card { background:var(--surface); border-radius:14px; padding:32px 30px; width:100%; max-width:380px; box-shadow:0 20px 60px rgba(0,0,0,0.25); }
.login-brand { display:flex; align-items:center; gap:12px; margin-bottom:24px; }
.login-title { font-family:'Fraunces',serif; font-size:19px; font-weight:600; }
.login-sub { font-size:11px; color:var(--ink-soft); letter-spacing:.02em; }
.login-form { display:flex; flex-direction:column; gap:4px; }
.input-icon-wrap { position:relative; margin-bottom:12px; }
.input-icon { position:absolute; left:10px; top:50%; transform:translateY(-50%); color:var(--ink-soft); }
.input-icon-wrap input { padding-left:32px; padding-right:34px; }
.input-icon-btn { position:absolute; right:8px; top:50%; transform:translateY(-50%); background:none; border:none; cursor:pointer; color:var(--ink-soft); }
.login-error { display:flex; align-items:center; gap:6px; color:var(--red); font-size:12px; margin:2px 0 8px; }
.login-submit { justify-content:center; margin-top:6px; padding:11px; font-size:13.5px; }
.login-footnote { margin-top:18px; font-size:10.5px; color:var(--ink-soft); line-height:1.5; border-top:1px solid var(--line); padding-top:12px; }
@media (max-width:980px) {
  .card-grid { grid-template-columns:repeat(2,1fr); }
  .chart-grid { grid-template-columns:1fr; }
  .entry-form { grid-template-columns:1fr 1fr; }
  .entry-form-desc, .entry-form-submit { grid-column:span 2; }
}
@media (max-width:760px) {
  .sidebar { position:fixed; left:-260px; top:0; bottom:0; z-index:40; transition:left .2s; box-shadow:4px 0 24px rgba(0,0,0,0.3); }
  .sidebar-open { left:0; }
  .sidebar-backdrop { display:block; position:fixed; inset:0; background:rgba(0,0,0,0.4); z-index:35; }
  .hamburger { display:block; }
  .card-grid { grid-template-columns:1fr 1fr; }
  .view-body { padding:16px 14px 30px; }
  .config-fields { grid-template-columns:1fr; }
}
`;
