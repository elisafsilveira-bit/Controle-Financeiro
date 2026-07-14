import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell, XAxis, YAxis,
  CartesianGrid, Tooltip, Legend, ResponsiveContainer, AreaChart, Area, ComposedChart
} from "recharts";
import {
  LayoutDashboard, FileText, Scale, Wallet, Settings, Building2,
  TrendingUp, TrendingDown, Plus, Trash2, Lock, User, Eye, EyeOff,
  LogOut, Calendar, AlertCircle, Check, X, Landmark, Percent
} from "lucide-react";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";

/* ============================== CONSTANTES ============================== */

const CATEGORIAS = [
  // Receita Operacional Bruta
  { id: "receita_vendas", label: "Venda de Mercadorias (Geral)", grupo: "receitaBruta", tipo: "entrada", dre: true },
  { id: "venda_oculos_grau", label: "Venda de Óculos de Grau", grupo: "receitaBruta", tipo: "entrada", dre: true },
  { id: "venda_armacoes", label: "Venda de Armações", grupo: "receitaBruta", tipo: "entrada", dre: true },
  { id: "venda_lentes", label: "Venda de Lentes", grupo: "receitaBruta", tipo: "entrada", dre: true },
  { id: "venda_oculos_sol", label: "Venda de Óculos de Sol", grupo: "receitaBruta", tipo: "entrada", dre: true },
  { id: "venda_lentes_contato", label: "Venda de Lentes de Contato", grupo: "receitaBruta", tipo: "entrada", dre: true },
  { id: "venda_acessorios", label: "Venda de Acessórios", grupo: "receitaBruta", tipo: "entrada", dre: true },
  { id: "venda_joias", label: "Venda de Joias", grupo: "receitaBruta", tipo: "entrada", dre: true },
  { id: "venda_relogios", label: "Venda de Relógios", grupo: "receitaBruta", tipo: "entrada", dre: true },

  // Deduções da Receita Bruta
  { id: "impostos_vendas", label: "DAS - Simples Nacional", grupo: "deducoes", tipo: "saida", dre: true },
  { id: "devolucoes_cancelamentos", label: "Devoluções de Vendas e Cancelamentos", grupo: "deducoes", tipo: "saida", dre: true },

  // Custo das Mercadorias Vendidas (CMV)
  { id: "cmv", label: "Custo das Mercadorias Vendidas (Geral)", grupo: "cmv", tipo: "saida", dre: true },
  { id: "custo_armacoes", label: "Custo de Armações", grupo: "cmv", tipo: "saida", dre: true },
  { id: "custo_lentes_lab", label: "Custo de Lentes (Laboratório Terceiro)", grupo: "cmv", tipo: "saida", dre: true },
  { id: "custo_outros_produtos", label: "Custo de Outros Produtos (óculos de sol, lentes de contato, acessórios)", grupo: "cmv", tipo: "saida", dre: true },
  { id: "custo_joias", label: "Custo de Joias", grupo: "cmv", tipo: "saida", dre: true },
  { id: "custo_relogios", label: "Custo de Relógios", grupo: "cmv", tipo: "saida", dre: true },

  // Despesas com Vendas
  { id: "comerciais", label: "Despesas Comerciais/Marketing (Outras)", grupo: "opVendas", tipo: "saida", dre: true },
  { id: "comissoes_vendedores", label: "Comissões de Vendedores", grupo: "opVendas", tipo: "saida", dre: true },
  { id: "taxas_cartao", label: "Taxas de Maquininha de Cartão", grupo: "opVendas", tipo: "saida", dre: true },
  { id: "embalagens", label: "Embalagens", grupo: "opVendas", tipo: "saida", dre: true },
  { id: "marketing_anuncios", label: "Marketing / Anúncios", grupo: "opVendas", tipo: "saida", dre: true },

  // Despesas Administrativas
  { id: "administrativas", label: "Despesas Administrativas (Outras)", grupo: "opAdmin", tipo: "saida", dre: true },
  { id: "pessoal", label: "Despesas com Pessoal (Outras)", grupo: "opAdmin", tipo: "saida", dre: true },
  { id: "pro_labore", label: "Pró-labore", grupo: "opAdmin", tipo: "saida", dre: true },
  { id: "salarios_equipe", label: "Salários da Equipe", grupo: "opAdmin", tipo: "saida", dre: true },
  { id: "fgts_inss", label: "FGTS / INSS", grupo: "opAdmin", tipo: "saida", dre: true },
  { id: "aluguel_loja", label: "Aluguel da Loja", grupo: "opAdmin", tipo: "saida", dre: true },
  { id: "agua_luz_internet", label: "Água, Luz e Internet", grupo: "opAdmin", tipo: "saida", dre: true },
  { id: "sistema_gestao", label: "Sistema de Gestão", grupo: "opAdmin", tipo: "saida", dre: true },
  { id: "material_expediente", label: "Material de Expediente", grupo: "opAdmin", tipo: "saida", dre: true },
  { id: "honorarios_contabeis", label: "Honorários Contábeis", grupo: "opAdmin", tipo: "saida", dre: true },

  // Resultado Financeiro
  { id: "receitas_financeiras", label: "Receitas Financeiras (Outras)", grupo: "recFin", tipo: "entrada", dre: true },
  { id: "descontos_fornecedores", label: "Descontos Obtidos de Fornecedores", grupo: "recFin", tipo: "entrada", dre: true },
  { id: "rendimentos_aplicacao", label: "Rendimentos de Aplicação", grupo: "recFin", tipo: "entrada", dre: true },
  { id: "despesas_financeiras", label: "Despesas Financeiras (Outras)", grupo: "despFin", tipo: "saida", dre: true },
  { id: "juros_boletos_atraso", label: "Juros de Boletos em Atraso", grupo: "despFin", tipo: "saida", dre: true },
  { id: "tarifas_bancarias", label: "Tarifas Bancárias PJ", grupo: "despFin", tipo: "saida", dre: true },

  // Fora do DRE (movimentações patrimoniais/financeiras que não são resultado)
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
  const direta = CATEGORIAS.find((c) => normalizeStr(c.label) === alvo) ||
    CATEGORIAS.find((c) => normalizeStr(c.label).includes(alvo) || alvo.includes(normalizeStr(c.label)));
  if (direta) return direta;
  // Se não bateu com nenhum nome de categoria exato, tenta pelo classificador
  // inteligente por palavras-chave (o mesmo usado no Plano de Contas)
  const viaClassificador = classificarConta(raw);
  return viaClassificador ? CATEGORIAS.find((c) => c.id === viaClassificador) || null : null;
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
  if (!alvo) return null;
  const nomeA = normalizeStr(empresas.a.nome), nomeB = normalizeStr(empresas.b.nome);
  if (alvo === "a" || alvo === nomeA) return "a";
  if (alvo === "b" || alvo === nomeB) return "b";
  if (nomeA && (alvo.includes(nomeA) || nomeA.includes(alvo))) return "a";
  if (nomeB && (alvo.includes(nomeB) || nomeB.includes(alvo))) return "b";
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

function parseStatusCell(v) {
  const alvo = normalizeStr(v);
  if (!alvo) return "liquidado"; // padrão: já pago/recebido (típico de extrato bancário/maquininha)
  if (alvo.includes("pendente") || alvo.includes("aberto") || alvo.includes("a pagar") || alvo.includes("a receber")) return "pendente";
  return "liquidado";
}

function baixarModeloExcel(empresas) {
  const linhas = [
    ["Data de Competência", "Data de Pagamento/Recebimento", "Empresa", "Categoria", "Valor", "Status", "Descrição"],
    ["2026-07-05", "2026-07-05", "A", "Venda de Óculos de Grau", 9500, "Pago/Recebido", "Faturamento do mês - à vista"],
    ["2026-07-05", "2026-08-05", "A", "Venda de Armações", 3200, "Pendente", "Venda parcelada - só cai no caixa em agosto"],
    ["2026-07-06", "2026-07-06", "A", "DAS - Simples Nacional", 900, "Pago/Recebido", "Guia DAS do mês"],
    ["2026-07-10", "2026-07-10", "B", "Pró-labore", 8000, "Pago/Recebido", "Retirada do sócio"],
  ];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(linhas);
  ws["!cols"] = [{ wch: 16 }, { wch: 22 }, { wch: 10 }, { wch: 32 }, { wch: 12 }, { wch: 14 }, { wch: 34 }];
  XLSX.utils.book_append_sheet(wb, ws, "Lançamentos");
  const wsCat = XLSX.utils.aoa_to_sheet([
    ["Categorias aceitas (copie e cole exatamente como está aqui)"],
    ...CATEGORIAS.map((c) => [c.label]),
  ]);
  wsCat["!cols"] = [{ wch: 42 }];
  XLSX.utils.book_append_sheet(wb, wsCat, "Categorias válidas");
  const wsAjuda = XLSX.utils.aoa_to_sheet([
    ["Coluna", "Obrigatório?", "O que é"],
    ["Data de Competência", "Opcional (usa a data de pagamento se deixar em branco)", "Data da nota/emissão — usada na DRE"],
    ["Data de Pagamento/Recebimento", "Sim (se Status = Pago/Recebido)", "Data em que o dinheiro entrou/saiu — usada no Fluxo de Caixa"],
    ["Empresa", "Sim", "A ou B"],
    ["Categoria", "Sim", "Copie da aba 'Categorias válidas'"],
    ["Valor", "Sim", "Use ponto ou vírgula decimal"],
    ["Status", "Opcional (padrão: Pago/Recebido)", "'Pago/Recebido' ou 'Pendente'"],
    ["Descrição", "Opcional", "Texto livre"],
  ]);
  wsAjuda["!cols"] = [{ wch: 30 }, { wch: 42 }, { wch: 46 }];
  XLSX.utils.book_append_sheet(wb, wsAjuda, "Como preencher");
  XLSX.writeFile(wb, "modelo-lancamentos.xlsx");
}

// Grupos de "apelidos" de coluna usados tanto para achar o cabeçalho quanto para
// mapear campo -> índice de coluna quando o arquivo tem cabeçalho.
const ALIASES_CAMPOS = {
  dataCompetencia: ["data de competência", "data competencia", "competência", "data emissão", "data de emissão", "data movimento"],
  dataCaixa: ["data de pagamento/recebimento", "data pagamento", "data recebimento", "data de pagamento", "data de recebimento", "data"],
  empresa: ["empresa", "cnpj", "loja", "filial"],
  categoria: ["categoria", "conta", "plano de contas"],
  valor: ["valor", "valor (r$)", "valor total", "montante"],
  status: ["status", "situação"],
  descricao: ["descrição", "descricao", "obs", "observação", "histórico", "lançamento"],
};

function detectarCabecalho(row) {
  const mapa = {};
  let acertos = 0;
  row.forEach((cel, idx) => {
    const alvo = normalizeStr(cel);
    if (!alvo) return;
    for (const [campo, apelidos] of Object.entries(ALIASES_CAMPOS)) {
      if (apelidos.includes(alvo) && mapa[campo] === undefined) {
        mapa[campo] = idx;
        acertos++;
      }
    }
  });
  // considera cabeçalho só se reconheceu pelo menos 3 dos campos essenciais
  const essenciais = ["empresa", "categoria", "valor"].filter((c) => mapa[c] !== undefined).length;
  return essenciais >= 3 ? mapa : null;
}

function parseLancamentosExcel(arrayBuffer, empresas) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const linhasBrutas = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
  if (linhasBrutas.length === 0) return { validos: [], erros: [] };

  const mapaCabecalho = detectarCabecalho(linhasBrutas[0]);
  // Com cabeçalho: pula a 1ª linha e usa as posições encontradas nela.
  // Sem cabeçalho (arquivo "cru" tipo extrato/relatório direto): assume a ordem
  // Data, Empresa, Categoria, Valor, Descrição (posições 0,1,2,3,4).
  const mapa = mapaCabecalho || { dataCaixa: 0, empresa: 1, categoria: 2, valor: 3, descricao: 4 };
  const linhasDados = mapaCabecalho ? linhasBrutas.slice(1) : linhasBrutas;
  const offsetLinha = mapaCabecalho ? 2 : 1;

  const validos = [];
  const erros = [];
  linhasDados.forEach((row, idx) => {
    const linhaNum = idx + offsetLinha;
    const get = (campo) => (mapa[campo] !== undefined ? row[mapa[campo]] : undefined);

    const dataCompetenciaRaw = get("dataCompetencia");
    const dataCaixaRaw = get("dataCaixa") || dataCompetenciaRaw;
    const empresaRaw = get("empresa");
    const categoriaRaw = get("categoria");
    const valorRaw = get("valor");
    const statusRaw = get("status");
    let descricaoRaw = get("descricao");
    // evita usar a descrição quando ela só repete o nome da categoria (comum em exports diretos)
    if (descricaoRaw && categoriaRaw && normalizeStr(descricaoRaw) === normalizeStr(categoriaRaw)) descricaoRaw = "";

    if (!dataCaixaRaw && !dataCompetenciaRaw && !empresaRaw && !categoriaRaw && !valorRaw) return; // linha em branco

    const dataCaixa = parseDataCell(dataCaixaRaw);
    const dataCompetencia = parseDataCell(dataCompetenciaRaw) || dataCaixa;
    const cnpj = parseEmpresaCell(empresaRaw, empresas);
    const categoria = findCategoriaByLabel(categoriaRaw);
    const valor = parseValorCell(valorRaw);
    const status = parseStatusCell(statusRaw);

    const problemas = [];
    if (!dataCompetencia) problemas.push("data de competência inválida");
    if (status === "liquidado" && !dataCaixa) problemas.push("data de pagamento/recebimento inválida");
    if (!cnpj) problemas.push(`empresa não reconhecida ("${empresaRaw}") — confira o nome cadastrado em Configurações`);
    if (!categoria) problemas.push(`categoria não reconhecida ("${categoriaRaw}")`);
    if (valor === null || valor <= 0) problemas.push("valor inválido");

    if (problemas.length) {
      erros.push({ linha: linhaNum, motivo: problemas.join(", ") });
    } else {
      validos.push({
        id: Date.now() + idx,
        cnpj, data: dataCompetencia, categoriaId: categoria.id, valor,
        descricao: String(descricaoRaw || "").trim() || categoria.label,
        status, dataCaixa: status === "liquidado" ? dataCaixa : null,
      });
    }
  });
  return { validos, erros };
}

/* ---------- Importação do Relatório de Análise do Plano de Contas (para a DRE) ----------
   Este relatório vem do sistema contábil já consolidado por conta (sem data por lançamento),
   então cada linha vira UM lançamento de competência no mês/ano escolhido, com status
   "pendente" (não afeta o Fluxo de Caixa — que é regime de caixa e já tem sua própria
   importação). O sistema tenta classificar cada conta automaticamente por palavras-chave,
   mas sempre mostra uma prévia para o usuário revisar/corrigir antes de confirmar. */

const REGRAS_CLASSIFICACAO = [
  // Financiamento / patrimônio (NÃO entram na DRE)
  [/aporte/, "aporte_socios"],
  [/bens de pequeno valor|bem de pequeno valor|ativo imobilizado|^imobilizado/, "investimento"],
  [/distribui[cç][aã]o.*lucro/, "distribuicao_lucros"],
  [/s[oó]cio/, "pro_labore"],

  // Taxas de cartão / bancárias (checadas antes das regras genéricas de aluguel/convênio)
  [/aluguel.*maquin|maquin.*aluguel|reten[cç][aã]o.*cart[aã]o|taxa.*convenio/, "taxas_cartao"],
  [/\biof\b|manuten[cç][aã]o.*conta|taxa.*pix|\bpix\b/, "tarifas_bancarias"],

  // Folha de pagamento e encargos
  [/pro.?labor/, "pro_labore"],
  [/fgts/, "fgts_inss"],
  [/\binss\b/, "fgts_inss"],
  [/13.*sal[aá]rio|d[eé]cimo terceiro/, "salarios_equipe"],
  [/f[ée]rias/, "salarios_equipe"],
  [/rescis|multa recisoria|multa rescis/, "salarios_equipe"],
  [/plano.*sa[uú]de/, "salarios_equipe"],
  [/conv[eê]nio/, "salarios_equipe"],
  [/servi[cç]o extra/, "salarios_equipe"],
  [/sindicato/, "salarios_equipe"],
  [/vale.*(alimenta|refei[cç]|transport)/, "salarios_equipe"],
  [/sal[aá]rio maternidade/, "salarios_equipe"],
  [/pr[eê]mio/, "comissoes_vendedores"],
  [/exame/, "administrativas"],
  [/uniforme/, "administrativas"],
  [/sal[aá]rio/, "salarios_equipe"],

  // Estrutura / utilidades
  [/predial/, "administrativas"],
  [/aluguel|condom[ií]nio/, "aluguel_loja"],
  [/(^|\s)[aá]gua(\s|$)|luz|energia|internet|telefone/, "agua_luz_internet"],
  [/contabilidade|honor[aá]rio.*cont[aá]bil/, "honorarios_contabeis"],
  [/sistema.*gest|software/, "sistema_gestao"],
  [/mat\.?\s*escrit[oó]rio|material.*(expediente|escrit[oó]rio)|papelaria/, "material_expediente"],

  // Vendas / comercial
  [/comiss/, "comissoes_vendedores"],
  [/(maquineta|maquin|taxa.*cart[aã]o|adquirente)/, "taxas_cartao"],
  [/embalag/, "embalagens"],
  [/evento|(marketing|propaganda|an[uú]ncio|publicidade)/, "marketing_anuncios"],

  // Financeiro
  [/juros/, "juros_boletos_atraso"],
  [/tarifa.*banc/, "tarifas_bancarias"],
  [/desconto.*fornecedor|cr[eé]dito.*fornecedor/, "descontos_fornecedores"],
  [/rendimento.*aplica|receita financeira/, "rendimentos_aplicacao"],
  [/\bicms\b|(^| )das( |$)|simples nacional/, "impostos_vendas"],
  [/devolu[cç][aã]o|cancelamento|cr[eé]dito.*(dep[oó]sito|cliente)/, "devolucoes_cancelamentos"],

  // Custo (CMV) — compras, insumos e consertos
  [/conserto.*joia/, "custo_joias"],
  [/conserto.*[oó]culos/, "custo_outros_produtos"],
  [/insumos.*lab|laborat[oó]rio/, "custo_lentes_lab"],
  [/compras.*len[st]e|(custo|cmv).*len[st]e/, "custo_lentes_lab"],
  [/compras.*(arma[cç]|solar)|(custo|cmv).*arma[cç]/, "custo_armacoes"],
  [/compras.*joia|(custo|cmv).*joia/, "custo_joias"],
  [/compras.*rel[oó]gio|(custo|cmv).*rel[oó]gio/, "custo_relogios"],
  [/acess[oó]rio|consulta|optometra|fornitura|compras|(custo|cmv)|mercadoria vendida/, "custo_outros_produtos"],

  // Receita — vendas
  [/[oó]culos de grau/, "venda_oculos_grau"],
  [/lente de contato/, "venda_lentes_contato"],
  [/[oó]culos de sol/, "venda_oculos_sol"],
  [/\bjoia/, "venda_joias"],
  [/rel[oó]gio/, "venda_relogios"],
  [/\blente/, "venda_lentes"],
  [/arma[cç][aã]o/, "venda_armacoes"],
  [/acess[oó]rio/, "venda_acessorios"],
  [/venda|receita de vendas|faturamento|^caixa$|rec cart[aã]o|cart[aã]o|pedido|ordem de servi[cç]o/, "receita_vendas"],

  // Genéricos (última linha de defesa antes do fallback por sinal)
  [/curso|treinamento|higieniza|decora[cç][aã]o|supermercado|terceiro/, "administrativas"],
  [/despesa.*administrativ|administrativ/, "administrativas"],
  [/despesa.*financeira|financeira/, "despesas_financeiras"],
  [/despesa.*comercial|comercial/, "comerciais"],
];

function classificarConta(nomeConta) {
  const alvo = normalizeStr(nomeConta);
  for (const [regex, categoriaId] of REGRAS_CLASSIFICACAO) {
    if (regex.test(alvo)) return categoriaId;
  }
  return null; // não identificado — usa um valor padrão (marcado para revisão)
}

// Linhas de estrutura do relatório que nunca são lançamentos de verdade
const LINHAS_IGNORAR = /^impresso em|^p[aá]gina \d|^filtro:|^total:?$|^acumulado$|^contas a (pagar|receber)$|^an[aá]lise do plano de contas/i;

function parsePlanoContasExcel(arrayBuffer, mesSelecionado) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });

  let colunaDoMes = null; // índice da coluna que corresponde ao mês escolhido
  let anoDetectado = null;
  const linhas = [];
  let linhaId = 0;

  rows.forEach((row) => {
    // Detecta linha de cabeçalho com nomes de mês (repete várias vezes no relatório)
    const mesesEncontrados = row.reduce((acc, cel, idx) => {
      const i = MESES_LONGOS.findIndex((m) => normalizeStr(m) === normalizeStr(cel));
      if (i !== -1) acc[i + 1] = idx;
      return acc;
    }, {});
    if (Object.keys(mesesEncontrados).length >= 6) {
      colunaDoMes = mesesEncontrados[mesSelecionado] ?? null;
      const possivelAno = row.find((c) => typeof c === "number" && c >= 2000 && c <= 2100);
      if (possivelAno) anoDetectado = possivelAno;
      return; // linha de cabeçalho, não é lançamento
    }
    if (colunaDoMes === null) return; // ainda não achamos nenhum cabeçalho de mês

    const contaRaw = row[0];
    const conta = String(contaRaw ?? "").trim();
    if (!conta || LINHAS_IGNORAR.test(conta) || /^\d{4}$/.test(conta)) return;

    const valorCel = row[colunaDoMes];
    const valorNum = typeof valorCel === "number" ? valorCel : parseValorCell(valorCel);
    if (!valorNum) return; // 0 ou vazio = sem movimento naquele mês

    let categoriaId = classificarConta(conta);
    let confiante = !!categoriaId;
    if (!categoriaId) {
      categoriaId = valorNum >= 0 ? "receita_vendas" : "outras_saidas";
    } else {
      // se o sinal do valor não bate com o tipo esperado da categoria (ex: um "crédito"
      // com valor positivo caindo numa categoria de saída), marca para revisão manual
      const categoriaInfo = CATEGORIAS.find((c) => c.id === categoriaId);
      const tipoEsperado = valorNum >= 0 ? "entrada" : "saida";
      if (categoriaInfo && categoriaInfo.tipo !== tipoEsperado) confiante = false;
    }

    linhas.push({ linhaId: linhaId++, conta, valor: Math.abs(valorNum), categoriaId, confiante });
  });

  return { linhas, anoDetectado };
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
              <DREView entries={entries} cnpjSel={cnpjSel} empresas={empresas} persistEntries={persistEntries}
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
function statusEfetivo(e) { return e.status || "liquidado"; }
function dataCaixaEfetiva(e) { return e.dataCaixa || e.data; }
function entriesCaixa(entries, cnpjSel) {
  // Regime de caixa: só entra no Fluxo de Caixa o que já foi efetivamente pago/recebido
  return filterEntries(entries, cnpjSel).filter((e) => statusEfetivo(e) === "liquidado");
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
  const opVendas = sums.opVendas || 0;
  const opAdmin = sums.opAdmin || 0;
  const despesasOperacionais = opVendas + opAdmin;
  const resultadoOperacional = lucroBruto - despesasOperacionais;
  const recFin = sums.recFin || 0;
  const despFin = sums.despFin || 0;
  const resultadoFinanceiro = recFin - despFin;
  const lucroLiquido = resultadoOperacional + resultadoFinanceiro;
  return { receitaBruta, deducoes, receitaLiquida, cmv, lucroBruto, opVendas, opAdmin,
    despesasOperacionais, resultadoOperacional, recFin, despFin, resultadoFinanceiro, lucroLiquido };
}
function computeCaixaAcumulado(entries, cnpjSel, ateData) {
  const filtered = entriesCaixa(entries, cnpjSel).filter((e) => dataCaixaEfetiva(e) <= ateData);
  return filtered.reduce((acc, e) => acc + signedValue(e), 0);
}
function anosDisponiveis(entries) {
  const set = new Set();
  entries.forEach((e) => {
    set.add(Number(e.data.split("-")[0]));
    set.add(Number(dataCaixaEfetiva(e).split("-")[0]));
  });
  set.add(new Date().getFullYear());
  return Array.from(set).sort();
}

/* ============================== DASHBOARD ============================== */

function Dashboard({ entries, cnpjSel, selectedYear, selectedMonth, setSelectedYear, setSelectedMonth }) {
  const anos = anosDisponiveis(entries);
  const dreAtual = useMemo(() => computeDRE(entries, cnpjSel, selectedYear, selectedMonth), [entries, cnpjSel, selectedYear, selectedMonth]);
  const hojeStr = `${selectedYear}-${String(selectedMonth).padStart(2, "0")}-31`;
  const saldoFimMes = useMemo(() => computeCaixaAcumulado(entries, cnpjSel, hojeStr), [entries, cnpjSel, hojeStr]);
  const saldoHoje = useMemo(() => computeCaixaAcumulado(entries, cnpjSel, new Date().toISOString().slice(0, 10)), [entries, cnpjSel]);

  const geracaoCaixaMes = useMemo(() => {
    return entriesCaixa(entries, cnpjSel)
      .filter((e) => dataCaixaEfetiva(e).startsWith(`${selectedYear}-${String(selectedMonth).padStart(2, "0")}`))
      .reduce((acc, e) => acc + signedValue(e), 0);
  }, [entries, cnpjSel, selectedYear, selectedMonth]);

  const trend = useMemo(() => {
    const arr = [];
    for (let i = 5; i >= 0; i--) {
      let m = selectedMonth - i, y = selectedYear;
      while (m <= 0) { m += 12; y -= 1; }
      const dre = computeDRE(entries, cnpjSel, y, m);
      const despesas = dre.deducoes + dre.cmv + dre.despesasOperacionais + dre.despFin;
      const mm = String(m).padStart(2, "0");
      const geracaoCaixa = entriesCaixa(entries, cnpjSel)
        .filter((e) => dataCaixaEfetiva(e).startsWith(`${y}-${mm}`))
        .reduce((acc, e) => acc + signedValue(e), 0);
      arr.push({
        label: `${MESES[m - 1]}/${String(y).slice(2)}`, Faturamento: dre.receitaBruta, Despesas: despesas,
        "Lucro Líquido (Competência)": dre.lucroLiquido, "Geração de Caixa (Caixa)": geracaoCaixa,
      });
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

  const despesasTotal = dreAtual.deducoes + dreAtual.cmv + dreAtual.despesasOperacionais + dreAtual.despFin;
  const margem = dreAtual.receitaBruta ? (dreAtual.lucroLiquido / dreAtual.receitaBruta) * 100 : 0;

  const composicaoReceita = useMemo(() => {
    if (!dreAtual.receitaBruta) return null;
    return [{
      nome: "Receita Bruta",
      "Deduções": dreAtual.deducoes,
      "CMV": dreAtual.cmv,
      "Despesas Operacionais": dreAtual.despesasOperacionais,
      "Resultado Operacional": Math.max(dreAtual.resultadoOperacional, 0),
    }];
  }, [dreAtual]);

  return (
    <div className="stack">
      <PeriodBar anos={anos} selectedYear={selectedYear} selectedMonth={selectedMonth}
        setSelectedYear={setSelectedYear} setSelectedMonth={setSelectedMonth} showMonth />
      <div className="card-grid">
        <MetricCard label="Faturamento do mês" value={fmtBRL(dreAtual.receitaBruta)} icon={TrendingUp} tone="teal" />
        <MetricCard label="Margem de lucro atual" value={`${margem.toFixed(1)}%`} icon={Percent} tone={margem >= 0 ? "teal" : "red"} sub={`Lucro líquido: ${fmtBRL(dreAtual.lucroLiquido)}`} />
        <MetricCard label="Saldo em conta disponível hoje" value={fmtBRL(saldoHoje)} icon={Wallet} tone="gold" sub={`Em ${new Date().toLocaleDateString("pt-BR")}`} />
        <MetricCard label="Despesas do mês" value={fmtBRL(despesasTotal)} icon={TrendingDown} tone="red" />
        <MetricCard label="Saldo em caixa no fim do mês" value={fmtBRL(saldoFimMes)} icon={Scale} tone={saldoFimMes >= 0 ? "teal" : "red"} />
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
        <div className="panel-title">Visão Unificada: Lucro Líquido (Competência) x Geração de Caixa (Caixa) — últimos 6 meses</div>
        <div className="import-text" style={{ marginBottom: 6 }}>
          Cruza o resultado econômico da DRE (o que foi vendido/gasto na competência) com o
          que realmente entrou/saiu do caixa no mesmo mês — útil para ver se a empresa está
          "lucrando no papel" mas sem caixa disponível, ou o contrário.
        </div>
        <ResponsiveContainer width="100%" height={280}>
          <ComposedChart data={trend}>
            <CartesianGrid strokeDasharray="3 3" stroke="#D8DED4" vertical={false} />
            <XAxis dataKey="label" tick={{ fontSize: 12, fill: "#3C4A42" }} axisLine={{ stroke: "#C9CFC3" }} tickLine={false} />
            <YAxis tickFormatter={fmtBRLCompact} tick={{ fontSize: 11, fill: "#3C4A42" }} axisLine={false} tickLine={false} width={50} />
            <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ fontFamily: "IBM Plex Mono, monospace", fontSize: 12, borderRadius: 8, border: "1px solid #C9CFC3" }} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Bar dataKey="Lucro Líquido (Competência)" fill={COLORS.teal} radius={[3, 3, 0, 0]} barSize={28} />
            <Line type="monotone" dataKey="Geração de Caixa (Caixa)" stroke={COLORS.blue} strokeWidth={2.5} dot={{ r: 3 }} />
          </ComposedChart>
        </ResponsiveContainer>
      </div>

      <div className="panel">
        <div className="panel-title">CMV e Despesas Operacionais em relação à Receita Bruta — {MESES_LONGOS[selectedMonth - 1]}/{selectedYear}</div>
        {!composicaoReceita ? (
          <div className="empty-note">Sem receita lançada neste período.</div>
        ) : (
          <>
            <ResponsiveContainer width="100%" height={130}>
              <BarChart data={composicaoReceita} layout="vertical" margin={{ left: 10, right: 20 }}>
                <XAxis type="number" tickFormatter={fmtBRLCompact} tick={{ fontSize: 11, fill: "#3C4A42" }} axisLine={false} tickLine={false} />
                <YAxis type="category" dataKey="nome" tick={false} axisLine={false} width={0} />
                <Tooltip formatter={(v) => fmtBRL(v)} contentStyle={{ fontFamily: "IBM Plex Mono, monospace", fontSize: 12, borderRadius: 8, border: "1px solid #C9CFC3" }} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                <Bar dataKey="Deduções" stackId="a" fill={COLORS.gold} />
                <Bar dataKey="CMV" stackId="a" fill={COLORS.red} />
                <Bar dataKey="Despesas Operacionais" stackId="a" fill="#8C6D46" />
                <Bar dataKey="Resultado Operacional" stackId="a" fill={COLORS.teal} radius={[0, 4, 4, 0]} />
              </BarChart>
            </ResponsiveContainer>
            <div className="formula-caption" style={{ padding: "4px 4px 0" }}>
              Receita Bruta = {fmtBRL(dreAtual.receitaBruta)} · Lucro Líquido do mês (após resultado financeiro): {fmtBRL(dreAtual.lucroLiquido)}
            </div>
          </>
        )}
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

function DREView({ entries, cnpjSel, empresas, persistEntries, selectedYear, selectedMonth, setSelectedYear, setSelectedMonth }) {
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
      <ImportarPlanoContas empresas={empresas} entries={entries} persistEntries={persistEntries}
        selectedYear={selectedYear} selectedMonth={selectedMonth} />
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
  { key: "receitaBruta", label: "(=) Receita Operacional Bruta", bold: false, indent: 0 },
  { key: "deducoes", label: "(–) Deduções da Receita Bruta", bold: false, indent: 0 },
  { key: "receitaLiquida", label: "(=) Receita Operacional Líquida", bold: true, indent: 0, divider: true },
  { key: "cmv", label: "(–) Custo das Mercadorias Vendidas (CMV)", bold: false, indent: 0 },
  { key: "lucroBruto", label: "(=) Lucro Bruto", bold: true, indent: 0, divider: true },
  { key: "opVendas", label: "Despesas com Vendas", bold: false, indent: 1 },
  { key: "opAdmin", label: "Despesas Administrativas", bold: false, indent: 1 },
  { key: "despesasOperacionais", label: "(–) Total Despesas Operacionais", bold: true, indent: 0, divider: true },
  { key: "resultadoOperacional", label: "(=) Resultado Operacional antes do Result. Financeiro", bold: true, indent: 0, divider: true },
  { key: "recFin", label: "(+) Receitas Financeiras", bold: false, indent: 1 },
  { key: "despFin", label: "(–) Despesas Financeiras", bold: false, indent: 1 },
  { key: "resultadoFinanceiro", label: "(=) Resultado Financeiro Líquido", bold: true, indent: 0, divider: true },
  { key: "lucroLiquido", label: "(=) LUCRO OU PREJUÍZO LÍQUIDO DO PERÍODO", bold: true, indent: 0, divider: true, highlight: true },
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

  const pendentes = useMemo(() => filterEntries(entries, cnpjSel).filter((e) => statusEfetivo(e) === "pendente"), [entries, cnpjSel]);
  const pendReceber = pendentes.filter((e) => catInfo(e.categoriaId).tipo === "entrada").reduce((a, e) => a + e.valor, 0);
  const pendPagar = pendentes.filter((e) => catInfo(e.categoriaId).tipo === "saida").reduce((a, e) => a + e.valor, 0);

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
      <div className="import-text" style={{ marginTop: -8 }}>
        Este módulo segue o <strong>regime de caixa</strong>: só entram aqui os lançamentos já
        marcados como <strong>pago/recebido</strong>, na data em que o dinheiro realmente
        entrou ou saiu (não na data da nota).
      </div>
      {pendentes.length > 0 && (
        <div className="panel pendencias-panel">
          <div className="panel-title">Pendências (ainda não entraram no caixa)</div>
          <div className="pendencias-row">
            <span>A receber: <strong className="positive">{fmtBRL(pendReceber)}</strong></span>
            <span>A pagar: <strong className="negative">{fmtBRL(pendPagar)}</strong></span>
            <span className="empty-note" style={{ padding: 0, fontStyle: "normal" }}>({pendentes.length} lançamento(s) — marque como pago/recebido em Lançamentos)</span>
          </div>
        </div>
      )}
      <PeriodBar anos={anos} selectedYear={selectedYear} selectedMonth={selectedMonth}
        setSelectedYear={setSelectedYear} setSelectedMonth={setSelectedMonth} showMonth={modo === "diario"} />
      {modo === "diario" && <FluxoDiario entries={entries} cnpjSel={cnpjSel} year={selectedYear} month={selectedMonth} />}
      {modo === "mensal" && <FluxoMensal entries={entries} cnpjSel={cnpjSel} year={selectedYear} />}
      {modo === "anual" && <FluxoAnual entries={entries} cnpjSel={cnpjSel} anos={anos} />}
    </div>
  );
}

function FluxoDiario({ entries, cnpjSel, year, month }) {
  const filtered = entriesCaixa(entries, cnpjSel).filter((e) => {
    const [y, m] = dataCaixaEfetiva(e).split("-").map(Number);
    return y === year && m === month;
  });
  const inicioMes = `${year}-${String(month).padStart(2, "0")}-01`;
  const saldoInicial = useMemo(() => {
    const antes = entriesCaixa(entries, cnpjSel).filter((e) => dataCaixaEfetiva(e) < inicioMes);
    return antes.reduce((acc, e) => acc + signedValue(e), 0);
  }, [entries, cnpjSel, inicioMes]);
  const porDia = {};
  filtered.forEach((e) => {
    const dia = dataCaixaEfetiva(e);
    if (!porDia[dia]) porDia[dia] = { entradas: 0, saidas: 0 };
    const c = catInfo(e.categoriaId);
    if (c.tipo === "entrada") porDia[dia].entradas += e.valor; else porDia[dia].saidas += e.valor;
  });
  const dias = Object.keys(porDia).sort();
  let acumulado = saldoInicial;
  const linhas = dias.map((d) => {
    const { entradas, saidas } = porDia[d];
    const saldoAbertura = acumulado;
    acumulado += entradas - saidas;
    return { data: d, saldoAbertura, entradas, saidas, saldoDia: entradas - saidas, acumulado };
  });
  const chartData = linhas.map((l) => ({ label: l.data.slice(8, 10), Saldo: l.acumulado }));
  return (
    <div className="stack">
      <div className="panel">
        <div className="panel-title">Saldo acumulado no mês (regime de caixa)</div>
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
        <div className="formula-caption">Saldo Inicial + Entradas − Saídas = Saldo Final do Dia</div>
        {linhas.length === 0 ? <div className="empty-note">Sem lançamentos pagos/recebidos neste mês.</div> : (
          <table className="ledger-table">
            <thead><tr><th>Data</th><th className="num-cell">Saldo Inicial</th><th className="num-cell">Entradas</th><th className="num-cell">Saídas</th><th className="num-cell">Saldo Final do Dia</th></tr></thead>
            <tbody>
              {linhas.map((l) => (
                <tr key={l.data}>
                  <td>{l.data.split("-").reverse().join("/")}</td>
                  <td className="num-cell">{fmtBRL(l.saldoAbertura)}</td>
                  <td className="num-cell positive">{fmtBRL(l.entradas)}</td>
                  <td className="num-cell negative">{fmtBRL(l.saidas)}</td>
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
    const antes = entriesCaixa(entries, cnpjSel).filter((e) => dataCaixaEfetiva(e) < inicioAno);
    return antes.reduce((acc, e) => acc + signedValue(e), 0);
  }, [entries, cnpjSel, inicioAno]);
  let acumulado = saldoInicial;
  const linhas = [];
  for (let m = 1; m <= 12; m++) {
    const mm = String(m).padStart(2, "0");
    const filtered = entriesCaixa(entries, cnpjSel).filter((e) => dataCaixaEfetiva(e).startsWith(`${year}-${mm}`));
    let entradas = 0, saidas = 0;
    filtered.forEach((e) => { const c = catInfo(e.categoriaId); if (c.tipo === "entrada") entradas += e.valor; else saidas += e.valor; });
    const saldoAbertura = acumulado;
    acumulado += entradas - saidas;
    linhas.push({ mes: MESES[m - 1], saldoAbertura, entradas, saidas, saldoMes: entradas - saidas, acumulado });
  }
  return (
    <div className="stack">
      <div className="panel">
        <div className="panel-title">Entradas x Saídas por mês (regime de caixa) — {year}</div>
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
        <div className="formula-caption">Saldo Inicial + Entradas − Saídas = Saldo Final do Mês</div>
        <table className="ledger-table">
          <thead><tr><th>Mês</th><th className="num-cell">Saldo Inicial</th><th className="num-cell">Entradas</th><th className="num-cell">Saídas</th><th className="num-cell">Saldo Final</th></tr></thead>
          <tbody>
            {linhas.map((l) => (
              <tr key={l.mes}>
                <td>{l.mes}</td>
                <td className="num-cell">{fmtBRL(l.saldoAbertura)}</td>
                <td className="num-cell positive">{fmtBRL(l.entradas)}</td>
                <td className="num-cell negative">{fmtBRL(l.saidas)}</td>
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
    const filtered = entriesCaixa(entries, cnpjSel).filter((e) => dataCaixaEfetiva(e).startsWith(`${year}`));
    let entradas = 0, saidas = 0;
    filtered.forEach((e) => { const c = catInfo(e.categoriaId); if (c.tipo === "entrada") entradas += e.valor; else saidas += e.valor; });
    const saldoAbertura = acumulado;
    acumulado += entradas - saidas;
    return { ano: year, saldoAbertura, entradas, saidas, saldoAno: entradas - saidas, acumulado };
  });
  return (
    <div className="stack">
      <div className="panel">
        <div className="panel-title">Entradas x Saídas por ano (regime de caixa)</div>
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
        <div className="formula-caption">Saldo Inicial + Entradas − Saídas = Saldo Final do Ano</div>
        <table className="ledger-table">
          <thead><tr><th>Ano</th><th className="num-cell">Saldo Inicial</th><th className="num-cell">Entradas</th><th className="num-cell">Saídas</th><th className="num-cell">Saldo Final</th></tr></thead>
          <tbody>
            {linhas.map((l) => (
              <tr key={l.ano}>
                <td>{l.ano}</td>
                <td className="num-cell">{fmtBRL(l.saldoAbertura)}</td>
                <td className="num-cell positive">{fmtBRL(l.entradas)}</td>
                <td className="num-cell negative">{fmtBRL(l.saidas)}</td>
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
  const [status, setStatus] = useState("liquidado");
  const [dataCaixa, setDataCaixa] = useState(new Date().toISOString().slice(0, 10));
  const [filtroCnpj, setFiltroCnpj] = useState("todos");
  const [filtroStatus, setFiltroStatus] = useState("todos");
  const [marcandoPago, setMarcandoPago] = useState(null); // entry sendo marcada como paga

  const addEntry = async (e) => {
    e.preventDefault();
    if (!valor || Number(valor) <= 0) return;
    const nova = {
      id: Date.now(), cnpj, data, categoriaId, valor: Number(valor),
      descricao: descricao || catInfo(categoriaId).label,
      status, dataCaixa: status === "liquidado" ? (dataCaixa || data) : null,
    };
    await persistEntries([nova, ...entries]);
    setValor(""); setDescricao("");
  };
  const removeEntry = async (id) => { await persistEntries(entries.filter((e) => e.id !== id)); };
  const importarEmLote = async (novosLancamentos) => { await persistEntries([...novosLancamentos, ...entries]); };
  const confirmarPagamento = async (id, novaDataCaixa) => {
    await persistEntries(entries.map((e) => e.id === id ? { ...e, status: "liquidado", dataCaixa: novaDataCaixa } : e));
    setMarcandoPago(null);
  };

  const listados = entries
    .filter((e) => filtroCnpj === "todos" || e.cnpj === filtroCnpj)
    .filter((e) => filtroStatus === "todos" || statusEfetivo(e) === filtroStatus)
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
            <label className="field-label">Data de Competência (nota/emissão)</label>
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
          <div>
            <label className="field-label">Status (regime de caixa)</label>
            <select value={status} onChange={(e) => setStatus(e.target.value)}>
              <option value="liquidado">Pago / Recebido</option>
              <option value="pendente">Pendente (a pagar/receber)</option>
            </select>
          </div>
          {status === "liquidado" && (
            <div>
              <label className="field-label">Data do Pagamento/Recebimento</label>
              <input type="date" value={dataCaixa} onChange={(e) => setDataCaixa(e.target.value)} />
            </div>
          )}
          <div className="entry-form-desc">
            <label className="field-label">Descrição (opcional)</label>
            <input value={descricao} onChange={(e) => setDescricao(e.target.value)} placeholder="Ex.: NF 1234, pagamento fornecedor X…" />
          </div>
          <button type="submit" className="btn-primary entry-form-submit"><Plus size={15} /> Adicionar</button>
        </form>
        <div className="import-text" style={{ marginTop: 10 }}>
          A <strong>Data de Competência</strong> é usada na DRE. A <strong>Data de Pagamento/Recebimento</strong>
          é usada no Fluxo de Caixa. Se marcar como "Pendente", o lançamento entra na DRE mas só aparece
          no Fluxo de Caixa quando você marcá-lo como pago/recebido depois.
        </div>
      </div>
      <div className="panel">
        <div className="panel-title-row">
          <div className="panel-title">Últimos lançamentos</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <select value={filtroStatus} onChange={(e) => setFiltroStatus(e.target.value)} className="filter-select">
              <option value="todos">Todos os status</option>
              <option value="liquidado">Pago / Recebido</option>
              <option value="pendente">Pendente</option>
            </select>
            <select value={filtroCnpj} onChange={(e) => setFiltroCnpj(e.target.value)} className="filter-select">
              <option value="todos">Todas as empresas</option>
              <option value="a">{empresas.a.nome}</option>
              <option value="b">{empresas.b.nome}</option>
            </select>
          </div>
        </div>
        <div className="table-scroll">
          <table className="ledger-table">
            <thead><tr><th>Data Competência</th><th>Data Caixa</th><th>Empresa</th><th>Categoria</th><th>Descrição</th><th className="num-cell">Valor</th><th>Status</th><th></th></tr></thead>
            <tbody>
              {listados.map((e) => {
                const c = catInfo(e.categoriaId);
                const st = statusEfetivo(e);
                return (
                  <tr key={e.id}>
                    <td>{e.data.split("-").reverse().join("/")}</td>
                    <td>{st === "liquidado" ? dataCaixaEfetiva(e).split("-").reverse().join("/") : "—"}</td>
                    <td>{empresas[e.cnpj]?.nome}</td>
                    <td>{c.label}</td>
                    <td>{e.descricao}</td>
                    <td className={"num-cell " + (c.tipo === "entrada" ? "positive" : "negative")}>{c.tipo === "entrada" ? "+" : "−"}{fmtBRL(e.valor)}</td>
                    <td>
                      {st === "liquidado"
                        ? <span className="status-badge status-pago">Pago/Recebido</span>
                        : <span className="status-badge status-pendente">Pendente</span>}
                    </td>
                    <td style={{ display: "flex", gap: 4 }}>
                      {st === "pendente" && (
                        <button className="icon-btn" title="Marcar como pago/recebido" onClick={() => setMarcandoPago(e)}>
                          <Check size={14} />
                        </button>
                      )}
                      <button className="icon-btn" onClick={() => removeEntry(e.id)}><Trash2 size={14} /></button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {listados.length === 0 && <div className="empty-note">Nenhum lançamento encontrado.</div>}
        </div>
      </div>

      {marcandoPago && (
        <MarcarPagoModal
          entry={marcandoPago}
          onCancel={() => setMarcandoPago(null)}
          onConfirm={confirmarPagamento}
        />
      )}
    </div>
  );
}

function MarcarPagoModal({ entry, onCancel, onConfirm }) {
  const [dataCaixa, setDataCaixa] = useState(new Date().toISOString().slice(0, 10));
  const c = catInfo(entry.categoriaId);
  return (
    <div className="modal-backdrop" onClick={onCancel}>
      <div className="modal-card" onClick={(e) => e.stopPropagation()} style={{ maxWidth: 420 }}>
        <div className="modal-header">
          <div>Marcar como {c.tipo === "entrada" ? "Recebido" : "Pago"}</div>
          <button className="icon-btn" onClick={onCancel}><X size={16} /></button>
        </div>
        <div className="modal-body">
          <p className="import-text">{c.label} — {fmtBRL(entry.valor)} — {entry.descricao}</p>
          <label className="field-label">Data do {c.tipo === "entrada" ? "recebimento" : "pagamento"}</label>
          <input type="date" value={dataCaixa} onChange={(e) => setDataCaixa(e.target.value)} />
        </div>
        <div className="modal-footer">
          <button className="btn-secondary" onClick={onCancel}>Cancelar</button>
          <button className="btn-primary" onClick={() => onConfirm(entry.id, dataCaixa)}><Check size={14} /> Confirmar</button>
        </div>
      </div>
    </div>
  );
}

function ImportarPlanoContas({ empresas, entries, persistEntries, selectedYear, selectedMonth }) {
  const [aberto, setAberto] = useState(false);
  const [cnpj, setCnpj] = useState("a");
  const [ano, setAno] = useState(selectedYear);
  const [mes, setMes] = useState(selectedMonth);
  const [linhas, setLinhas] = useState(null); // preview editável
  const [anoDetectado, setAnoDetectado] = useState(null);
  const [nomeArquivo, setNomeArquivo] = useState("");
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
      const { linhas: contas, anoDetectado: anoArq } = parsePlanoContasExcel(buffer, mes);
      setLinhas(contas);
      setAnoDetectado(anoArq);
      setNomeArquivo(file.name);
    } catch (err) {
      console.error(err);
      setLinhas([]);
      setAnoDetectado(null);
      setNomeArquivo(file.name);
    }
    setProcessando(false);
    if (inputRef.current) inputRef.current.value = "";
  };

  const atualizarCategoria = (linhaId, categoriaId) => {
    setLinhas((prev) => prev.map((l) => l.linhaId === linhaId ? { ...l, categoriaId, confiante: true } : l));
  };

  const todasLinhas = linhas || [];
  const confiantes = todasLinhas.filter((l) => l.confiante);
  const paraRevisar = todasLinhas.filter((l) => !l.confiante);

  const confirmarImportacao = async () => {
    if (todasLinhas.length === 0) return;
    setImportando(true);
    const dataCompetencia = `${ano}-${String(mes).padStart(2, "0")}-${String(new Date(ano, mes, 0).getDate()).padStart(2, "0")}`;
    const novos = todasLinhas.map((l, i) => ({
      id: Date.now() + i,
      cnpj, data: dataCompetencia, categoriaId: l.categoriaId, valor: l.valor,
      descricao: `Plano de Contas: ${l.conta}`,
      status: "pendente", dataCaixa: null, // só afeta a DRE (competência), não o Fluxo de Caixa
    }));
    await persistEntries([...novos, ...entries]);
    setImportando(false);
    setConcluido(novos.length);
    setLinhas(null);
  };

  return (
    <div className="panel">
      <div className="panel-title-row">
        <div className="panel-title">Importar Relatório de Análise do Plano de Contas (Excel)</div>
        <button className="btn-secondary" onClick={() => setAberto((s) => !s)}>{aberto ? "Ocultar" : "Abrir"}</button>
      </div>
      {aberto && (
        <div className="import-box">
          <p className="import-text">
            Suba aqui o relatório "Análise do plano de contas (realizado)" exportado do seu
            sistema contábil. O app localiza a coluna do mês escolhido abaixo e tenta reconhecer
            automaticamente cada conta (receita, custo, despesa, ou até aporte/empréstimo — que
            não entram na DRE). Linhas com fundo amarelo são sugestões genéricas: confira antes
            de confirmar. Esses lançamentos entram só na DRE (competência), não no Fluxo de Caixa.
          </p>
          <div className="import-actions" style={{ alignItems: "flex-end" }}>
            <div>
              <label className="field-label">Empresa</label>
              <select value={cnpj} onChange={(e) => setCnpj(e.target.value)}>
                <option value="a">{empresas.a.nome}</option>
                <option value="b">{empresas.b.nome}</option>
              </select>
            </div>
            <div>
              <label className="field-label">Mês de referência</label>
              <select value={mes} onChange={(e) => setMes(Number(e.target.value))}>
                {MESES_LONGOS.map((m, i) => <option key={m} value={i + 1}>{m}</option>)}
              </select>
            </div>
            <div>
              <label className="field-label">Ano</label>
              <input type="number" value={ano} onChange={(e) => setAno(Number(e.target.value))} style={{ width: 90 }} />
            </div>
            <label className="btn-primary import-upload-btn">
              <Plus size={14} /> Selecionar relatório
              <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} hidden />
            </label>
          </div>
          <div className="import-text" style={{ marginTop: -4 }}>
            Escolha o mês/ano <strong>antes</strong> de selecionar o arquivo — o app lê a coluna
            correspondente a esse mês dentro do relatório.
          </div>

          {processando && <div className="empty-note">Lendo arquivo…</div>}
          {concluido !== null && (
            <div className="import-success"><Check size={14} /> {concluido} conta(s) importada(s) para a DRE de {MESES_LONGOS[mes - 1]}/{ano}.</div>
          )}

          {linhas && (
            <div className="import-preview">
              {anoDetectado && anoDetectado !== ano && (
                <div className="balance-check"><AlertCircle size={13} /> Este arquivo parece ser do ano {anoDetectado}, mas você selecionou {ano}. Confira o campo "Ano" acima se não for intencional.</div>
              )}
              <div className="import-preview-summary">
                <span><strong>{nomeArquivo}</strong></span>
                <span className="positive">{confiantes.length} conta(s) reconhecida(s) automaticamente</span>
                {paraRevisar.length > 0 && <span className="negative">{paraRevisar.length} conta(s) com sugestão genérica — revise antes de confirmar</span>}
              </div>
              {todasLinhas.length === 0 ? (
                <div className="empty-note">
                  Nenhum valor encontrado para {MESES_LONGOS[mes - 1]}/{ano} neste arquivo. Confira se
                  escolheu o mês certo, ou se o relatório realmente tem movimento nesse período.
                </div>
              ) : (
                <div className="table-scroll">
                  <table className="ledger-table">
                    <thead><tr><th>Conta (original)</th><th className="num-cell">Valor</th><th>Classificação na DRE</th></tr></thead>
                    <tbody>
                      {todasLinhas.map((l) => (
                        <tr key={l.linhaId} className={!l.confiante ? "row-atencao" : ""}>
                          <td>{l.conta}</td>
                          <td className="num-cell">{fmtBRL(l.valor)}</td>
                          <td>
                            <select value={l.categoriaId} onChange={(e) => atualizarCategoria(l.linhaId, e.target.value)}>
                              <optgroup label="Contas de resultado (entram na DRE)">
                                {CATEGORIAS.filter((c) => c.dre).map((c) => (
                                  <option key={c.id} value={c.id}>{c.label} {c.tipo === "entrada" ? "↑" : "↓"}</option>
                                ))}
                              </optgroup>
                              <optgroup label="Fora da DRE (financiamento/patrimônio)">
                                {CATEGORIAS.filter((c) => !c.dre).map((c) => (
                                  <option key={c.id} value={c.id}>{c.label} {c.tipo === "entrada" ? "↑" : "↓"}</option>
                                ))}
                              </optgroup>
                            </select>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
              <div className="import-confirm-row">
                <button className="btn-secondary" onClick={() => setLinhas(null)}>Cancelar</button>
                <button className="btn-primary" disabled={todasLinhas.length === 0 || importando} onClick={confirmarImportacao}>
                  {importando ? "Importando…" : `Confirmar importação de ${todasLinhas.length} conta(s) para ${MESES_LONGOS[mes - 1]}/${ano}`}
                </button>
              </div>
            </div>
          )}
        </div>
      )}
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
            nomes de categoria da aba "Categorias válidas" e as instruções da aba "Como
            preencher") e depois envie o arquivo preenchido aqui. Serve tanto para lançar
            manualmente quanto para importar o extrato do banco ou o relatório da maquininha
            de cartão — nesse caso, deixe o Status como "Pago/Recebido" para cada linha.
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
                    <thead><tr><th>Competência</th><th>Pagto/Receb.</th><th>Empresa</th><th>Categoria</th><th>Descrição</th><th className="num-cell">Valor</th><th>Status</th></tr></thead>
                    <tbody>
                      {preview.validos.slice(0, 15).map((e) => {
                        const c = catInfo(e.categoriaId);
                        return (
                          <tr key={e.id}>
                            <td>{e.data.split("-").reverse().join("/")}</td>
                            <td>{e.dataCaixa ? e.dataCaixa.split("-").reverse().join("/") : "—"}</td>
                            <td>{empresas[e.cnpj]?.nome}</td>
                            <td>{c.label}</td>
                            <td>{e.descricao}</td>
                            <td className={"num-cell " + (c.tipo === "entrada" ? "positive" : "negative")}>{fmtBRL(e.valor)}</td>
                            <td>{e.status === "liquidado" ? <span className="status-badge status-pago">Pago/Receb.</span> : <span className="status-badge status-pendente">Pendente</span>}</td>
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
.card-grid { display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:14px; }
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
.status-badge { display:inline-block; padding:2px 9px; border-radius:999px; font-size:10.5px; font-weight:600; white-space:nowrap; }
.status-pago { background:#E4EFE9; color:#155444; }
.status-pendente { background:#FBEFD8; color:#8A5A16; }
.formula-caption { font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--ink-soft); padding:0 20px 10px; }
.pendencias-panel { border-left:3px solid var(--gold); }
.pendencias-row { display:flex; gap:20px; flex-wrap:wrap; align-items:center; font-size:13px; }
.row-atencao td { background:#FBF3E3; }
.row-atencao select { border-color:var(--gold); }
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
