// ═══════════════════════════════════════════════════════════════
//  SMK · Indicadores Mensais — Google Apps Script
//  Versão com histórico completo de lançamentos
//
//  CONFIGURAÇÃO:
//  1. Abra o Google Sheets onde quer salvar os dados
//  2. Extensões → Apps Script → cole este código
//  3. Implante como Web App:
//       Executar como: "Eu (seu e-mail)"
//       Quem tem acesso: "Qualquer pessoa"  ← necessário para o HTML enviar dados
//  4. Copie a URL gerada e cole no campo "URL do Apps Script" no sistema
// ═══════════════════════════════════════════════════════════════

// Nome das abas na planilha
const ABA_HISTORICO = 'Histórico';   // Todos os lançamentos (com duplicatas por mês)
const ABA_ATUAL     = 'Atual';       // Apenas o mais recente por vendedor+mês
const ABA_LOG       = 'Log';         // Registro de sincronizações

// Colunas na ordem exata enviada pelo sistema
const COLUNAS = [
  'id',
  'mes',
  'vendedor',
  'filial',
  'contratos',
  'valorContratos',
  'faturamento',
  'recebimentos',
  'vencimentos',
  'entradas',
  'dataLancamento',
  'savedAt'
];

const CABECALHO_LEGIVEL = [
  'ID',
  'Mês/Ano',
  'Vendedor',
  'Filial',
  'Contratos',
  'Valor Contratos',
  'Faturamento',
  'Recebimentos',
  'Vencimentos',
  'Entradas do Mês',
  'Data do Lançamento',
  'Salvo Em'
];

// ───────────────────────────────────────────────────────────────
//  doPost — recebe dados do sistema HTML
// ───────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const acao = payload.acao || 'sincronizar';

    if (acao === 'sincronizar') {
      return sincronizarDados(payload);
    }

    return resposta({ ok: false, erro: 'Ação desconhecida: ' + acao });

  } catch (err) {
    registrarLog('ERRO doPost', err.message);
    return resposta({ ok: false, erro: err.message });
  }
}

// ───────────────────────────────────────────────────────────────
//  doGet — permite carregar os dados de volta para o sistema
// ───────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = obterOuCriarAba(ss, ABA_HISTORICO);
    const dados = lerAba(aba);
    return resposta({ ok: true, dados: dados });
  } catch (err) {
    return resposta({ ok: false, erro: err.message });
  }
}

// ───────────────────────────────────────────────────────────────
//  Sincronização principal
// ───────────────────────────────────────────────────────────────
function sincronizarDados(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registros = payload.dados || [];

  if (!registros.length) {
    return resposta({ ok: false, erro: 'Nenhum dado recebido.' });
  }

  // ── 1. Aba Histórico — todos os registros enviados pelo sistema ──
  const abaHist = obterOuCriarAba(ss, ABA_HISTORICO);
  escreverAba(abaHist, registros);
  formatarAba(abaHist);

  // ── 2. Aba Atual — apenas o mais recente por vendedor+mês ──
  const abaAtual = obterOuCriarAba(ss, ABA_ATUAL);
  const registrosAtuais = filtrarMaisRecentes(registros);
  escreverAba(abaAtual, registrosAtuais);
  formatarAba(abaAtual);

  // ── 3. Log de sincronização ──
  registrarLog(
    'SYNC OK',
    `${registros.length} registros no histórico | ${registrosAtuais.length} registros atuais`
  );

  return resposta({
    ok: true,
    mensagem: 'Sincronizado com sucesso!',
    totalHistorico: registros.length,
    totalAtual: registrosAtuais.length,
    timestamp: new Date().toISOString()
  });
}

// ───────────────────────────────────────────────────────────────
//  Filtra apenas o registro mais recente por vendedor+mês
// ───────────────────────────────────────────────────────────────
function filtrarMaisRecentes(registros) {
  const mapa = {};
  registros.forEach(r => {
    const chave = r.mes + '|' + r.vendedor;
    const atual = mapa[chave];
    if (!atual || r.savedAt > atual.savedAt) {
      mapa[chave] = r;
    }
  });
  return Object.values(mapa).sort((a, b) => {
    if (a.mes !== b.mes) return a.mes < b.mes ? 1 : -1; // mais recente primeiro
    return a.vendedor < b.vendedor ? -1 : 1;
  });
}

// ───────────────────────────────────────────────────────────────
//  Escreve dados em uma aba (limpa e reescreve tudo)
// ───────────────────────────────────────────────────────────────
function escreverAba(aba, registros) {
  aba.clearContents();

  const linhas = [CABECALHO_LEGIVEL];

  registros.forEach(r => {
    const linha = COLUNAS.map(col => {
      const val = r[col];
      if (val === undefined || val === null) return '';
      // Formata valores monetários como número
      if (['valorContratos','faturamento','recebimentos','vencimentos','entradas'].includes(col)) {
        return typeof val === 'number' ? val : parseFloat(val) || 0;
      }
      // Formata contratos como número inteiro
      if (col === 'contratos') {
        return typeof val === 'number' ? val : parseInt(val) || 0;
      }
      return val;
    });
    linhas.push(linha);
  });

  aba.getRange(1, 1, linhas.length, CABECALHO_LEGIVEL.length).setValues(linhas);
}

// ───────────────────────────────────────────────────────────────
//  Formata visual da aba (cabeçalho, larguras, formato R$)
// ───────────────────────────────────────────────────────────────
function formatarAba(aba) {
  const numLinhas = aba.getLastRow();
  if (numLinhas < 1) return;

  const numCols = CABECALHO_LEGIVEL.length;

  // Cabeçalho
  const cabRange = aba.getRange(1, 1, 1, numCols);
  cabRange
    .setBackground('#1a1a2e')
    .setFontColor('#ff8c00')
    .setFontWeight('bold')
    .setFontSize(9);

  // Congela cabeçalho
  aba.setFrozenRows(1);

  if (numLinhas > 1) {
    // Colunas monetárias: índices 5,6,7,8,9 (base 1 = 6,7,8,9,10)
    const colsMon = [6, 7, 8, 9, 10];
    colsMon.forEach(col => {
      aba.getRange(2, col, numLinhas - 1, 1)
        .setNumberFormat('R$ #,##0.00');
    });

    // Coluna de contratos (índice 4 = col 5)
    aba.getRange(2, 5, numLinhas - 1, 1)
      .setNumberFormat('#,##0');

    // Zebra nas linhas de dados
    for (let i = 2; i <= numLinhas; i++) {
      aba.getRange(i, 1, 1, numCols)
        .setBackground(i % 2 === 0 ? '#f9f9f9' : '#ffffff')
        .setFontColor('#222222')
        .setFontSize(9);
    }
  }

  // Larguras das colunas
  const larguras = [160, 90, 120, 130, 80, 120, 120, 120, 120, 120, 130, 180];
  larguras.forEach((w, i) => aba.setColumnWidth(i + 1, w));

  // Bordas
  if (numLinhas > 1) {
    aba.getRange(1, 1, numLinhas, numCols)
      .setBorder(true, true, true, true, true, true, '#dddddd', SpreadsheetApp.BorderStyle.SOLID);
  }
}

// ───────────────────────────────────────────────────────────────
//  Utilitários
// ───────────────────────────────────────────────────────────────
function obterOuCriarAba(ss, nome) {
  let aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
  }
  return aba;
}

function lerAba(aba) {
  const valores = aba.getDataRange().getValues();
  if (valores.length < 2) return [];
  const cab = valores[0];
  return valores.slice(1).map(linha => {
    const obj = {};
    cab.forEach((col, i) => { obj[col] = linha[i]; });
    return obj;
  });
}

function registrarLog(tipo, mensagem) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaLog = obterOuCriarAba(ss, ABA_LOG);
    if (abaLog.getLastRow() === 0) {
      abaLog.appendRow(['Timestamp', 'Tipo', 'Mensagem']);
      abaLog.getRange(1, 1, 1, 3)
        .setBackground('#1a1a2e')
        .setFontColor('#ff8c00')
        .setFontWeight('bold');
    }
    abaLog.appendRow([new Date().toLocaleString('pt-BR'), tipo, mensagem]);
  } catch (e) {
    // silencia erros de log para não atrapalhar a resposta principal
  }
}

function resposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
