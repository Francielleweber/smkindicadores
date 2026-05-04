// ═══════════════════════════════════════════════════════════════
//  SMK · Indicadores — Google Apps Script
//  Cole este código em https://script.google.com
//  e publique como Web App (acesso: Qualquer pessoa)
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME_HIST = 'Histórico';
const COLS = ['id','mes','vendedor','filial','contratos','valorContratos',
              'faturamento','recebimentos','vencimentos','entradas',
              'dataLancamento','savedAt'];

// ── GET: buscar dados do Sheets → retorna JSON ─────────────────
function doGet(e) {
  const acao = (e && e.parameter && e.parameter.acao) || 'buscar';

  if (acao === 'buscar') {
    try {
      const ss    = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEET_NAME_HIST);

      if (!sheet || sheet.getLastRow() < 2) {
        return jsonResp({ ok: true, dados: [], totalHistorico: 0 });
      }

      const range  = sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), COLS.length));
      const values = range.getValues();
      const numCols = ['contratos','valorContratos','faturamento',
                       'recebimentos','vencimentos','entradas'];

      // Usa COLS como chaves (independente do cabeçalho real da planilha)
      // pois os dados estão na mesma ordem das colunas de COLS
      const dados = values.slice(1)
        .filter(row => row.some(c => c !== '' && c !== null))
        .map(row => {
          const obj = {};
          COLS.forEach((h, i) => obj[h] = row[i] ?? '');
          numCols.forEach(k => {
            if (obj[k] !== undefined)
              obj[k] = parseFloat(String(obj[k]).replace(',', '.')) || 0;
          });
          return obj;
        });

      return jsonResp({ ok: true, dados: dados, totalHistorico: dados.length });
    } catch (err) {
      return jsonResp({ ok: false, erro: err.message });
    }
  }

  return jsonResp({ ok: false, erro: 'Ação GET desconhecida: ' + acao });
}

// ── POST: receber dados do app → gravar no Sheets ─────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const acao    = payload.acao || 'sincronizar';

    if (acao === 'sincronizar') {
      const dados = payload.dados || [];
      const ss    = SpreadsheetApp.getActiveSpreadsheet();

      let sheet = ss.getSheetByName(SHEET_NAME_HIST);
      if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME_HIST);
      }

      // Limpa e reescreve
      sheet.clearContents();

      if (dados.length === 0) {
        sheet.getRange(1, 1, 1, COLS.length).setValues([COLS]);
        return jsonResp({ ok: true, totalHistorico: 0 });
      }

      const rows = [COLS, ...dados.map(r => COLS.map(c => r[c] ?? ''))];
      sheet.getRange(1, 1, rows.length, COLS.length).setValues(rows);

      // Formata cabeçalho
      sheet.getRange(1, 1, 1, COLS.length)
           .setFontWeight('bold')
           .setBackground('#222222')
           .setFontColor('#ffffff');

      return jsonResp({ ok: true, totalHistorico: dados.length });
    }

    return jsonResp({ ok: false, erro: 'Ação POST desconhecida: ' + acao });

  } catch (err) {
    return jsonResp({ ok: false, erro: err.message });
  }
}

// ── Helper ─────────────────────────────────────────────────────
function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
