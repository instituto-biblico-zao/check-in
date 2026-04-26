// ══════════════════════════════════════════════════════════════
//  INSTITUTO BÍBLICO ZAO — App de Presença
//  Google Apps Script v4.0 — Banco de Dados Central
//
//  INSTRUÇÕES:
//  1. Abra o Google Sheets → Extensões → Apps Script
//  2. Apague o código existente e cole este arquivo completo
//  3. Salve (Ctrl+S)
//  4. Implantar → Nova implantação → App da Web
//     - Executar como: Eu
//     - Quem tem acesso: Qualquer pessoa
//  5. Copie a URL e cole no app em Configurações → Sistema
// ══════════════════════════════════════════════════════════════

var ABA_DADOS = '_dados';  // Aba oculta que armazena o JSON completo

// ──────────────────────────────────────────────────────────────
// ENTRY POINTS
// ──────────────────────────────────────────────────────────────

function doGet(e) {
  var acao = e && e.parameter ? e.parameter.acao : 'ping';

  if (acao === 'carregar') {
    return responder({ ok: true, dados: carregarDados() });
  }

  return responder({ ok: true, mensagem: 'ZAO App de Presença v4.0 — ativo.' });
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var acao = payload.acao;

    if (acao === 'salvar') {
      salvarDados(payload.dados);
      atualizarPlanilhasVisuais(payload.dados);
      return responder({ ok: true, mensagem: 'Dados salvos.' });
    }

    return responder({ ok: false, erro: 'Ação desconhecida: ' + acao });
  } catch (err) {
    return responder({ ok: false, erro: err.message });
  }
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ──────────────────────────────────────────────────────────────
// PERSISTÊNCIA — JSON na aba _dados
// ──────────────────────────────────────────────────────────────

function getAbaDados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(ABA_DADOS);
  if (!aba) {
    aba = ss.insertSheet(ABA_DADOS);
    aba.hideSheet();
    aba.getRange(1, 1).setValue('{}');
  }
  return aba;
}

function carregarDados() {
  var aba = getAbaDados();
  var valor = aba.getRange(1, 1).getValue();
  if (!valor || valor === '') return {};
  try { return JSON.parse(valor); } catch(e) { return {}; }
}

function salvarDados(dados) {
  var aba = getAbaDados();
  aba.getRange(1, 1).setValue(JSON.stringify(dados));
}

// ──────────────────────────────────────────────────────────────
// PLANILHAS VISUAIS — geradas automaticamente para consulta
// ──────────────────────────────────────────────────────────────

function atualizarPlanilhasVisuais(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var turmas = dados.turmas || [];

  // Aba Resumo Geral
  atualizarResumoGeral(ss, turmas);

  // Uma aba por turma+aula
  turmas.forEach(function(turma) {
    (turma.aulas || []).forEach(function(aula) {
      atualizarAbaPresenca(ss, turma, aula);
    });
  });
}

function atualizarResumoGeral(ss, turmas) {
  var nomeAba = 'Resumo Geral';
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) {
    aba = ss.insertSheet(nomeAba, 0);
  }
  aba.clearContents();

  // Cabeçalho
  var cab = ['Turma', 'Aula', 'Data', 'Total Alunos', 'Presentes', 'Ausentes', 'Frequência (%)'];
  var cabRange = aba.getRange(1, 1, 1, cab.length);
  cabRange.setValues([cab]);
  cabRange.setFontWeight('bold').setBackground('#1A3A5C').setFontColor('#FFFFFF');

  var linha = 2;
  turmas.forEach(function(turma) {
    (turma.aulas || []).forEach(function(aula) {
      var presenca = aula.presenca || {};
      var alunos = turma.alunos || [];
      var presentes = alunos.filter(function(a) { return presenca[a] > 0; }).length;
      var total = alunos.length;
      var pct = total ? Math.round(presentes / total * 100) : 0;
      var ausentes = total - presentes;

      aba.getRange(linha, 1, 1, 7).setValues([[
        turma.nome, aula.nome, fmtData(aula.data),
        total, presentes, ausentes, pct + '%'
      ]]);

      // Cor da frequência
      var pctCell = aba.getRange(linha, 7);
      if (pct >= 75) pctCell.setBackground('#EAF3DE').setFontColor('#27500A');
      else if (pct >= 50) pctCell.setBackground('#FAEEDA').setFontColor('#854F0B');
      else pctCell.setBackground('#FCEBEB').setFontColor('#A32D2D');

      linha++;
    });
  });

  aba.setColumnWidths(1, 7, 150);
  aba.setFrozenRows(1);
}

function atualizarAbaPresenca(ss, turma, aula) {
  var nomeAba = sanitizar(turma.nome + ' — ' + aula.nome);
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) aba = ss.insertSheet(nomeAba);
  aba.clearContents();

  // Título
  var t1 = aba.getRange(1, 1, 1, 5); t1.merge();
  t1.setValue('INSTITUTO BÍBLICO ZAO — ' + turma.nome);
  t1.setFontSize(13).setFontWeight('bold').setBackground('#1A3A5C').setFontColor('#FFFFFF').setHorizontalAlignment('center');

  var t2 = aba.getRange(2, 1, 1, 5); t2.merge();
  t2.setValue(aula.nome + ' — ' + fmtData(aula.data));
  t2.setBackground('#2E5F8A').setFontColor('#FFFFFF').setHorizontalAlignment('center');

  // Cabeçalho colunas
  var cab = ['Nº', 'Nome do Aluno', 'Status', 'Horário', 'Última atualização'];
  var cabRange = aba.getRange(3, 1, 1, 5);
  cabRange.setValues([cab]).setFontWeight('bold').setBackground('#D0E4F7').setFontColor('#1A3A5C');

  // Alunos
  var alunos = turma.alunos || [];
  var presenca = aula.presenca || {};
  var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');

  var presentes = alunos.filter(function(a) { return presenca[a] > 0; })
    .sort(function(a, b) { return presenca[a] - presenca[b]; });
  var ausentes = alunos.filter(function(a) { return !presenca[a] || presenca[a] <= 0; });
  var ordenados = presentes.concat(ausentes);

  ordenados.forEach(function(nome, i) {
    var ts = presenca[nome];
    var isP = ts > 0;
    var linha = i + 4;
    var status = isP ? 'Presente' : 'Ausente';
    var horario = isP ? fmtHora(ts) : '';

    aba.getRange(linha, 1, 1, 5).setValues([[i + 1, nome, status, horario, agora]]);

    var statusCell = aba.getRange(linha, 3);
    if (isP) statusCell.setBackground('#EAF3DE').setFontColor('#27500A').setFontWeight('bold');
    else statusCell.setBackground('#FCEBEB').setFontColor('#A32D2D').setFontWeight('bold');

    var rowBg = (linha % 2 === 0) ? '#F7F6F2' : '#FFFFFF';
    aba.getRange(linha, 1).setBackground(rowBg);
    aba.getRange(linha, 2).setBackground(rowBg);
    aba.getRange(linha, 4).setBackground(rowBg);
    aba.getRange(linha, 5).setBackground(rowBg).setFontColor('#6B6860');
  });

  // Totais
  var ultLinha = ordenados.length + 4;
  var totCell = aba.getRange(ultLinha + 1, 1, 1, 5); totCell.merge();
  var pres = presentes.length;
  var tot = alunos.length;
  var pct = tot ? Math.round(pres / tot * 100) : 0;
  totCell.setValue('Total: ' + pres + '/' + tot + ' presentes (' + pct + '%)')
    .setFontWeight('bold').setBackground('#1A3A5C').setFontColor('#FFFFFF').setHorizontalAlignment('center');

  aba.setColumnWidth(1, 40);
  aba.setColumnWidth(2, 220);
  aba.setColumnWidth(3, 90);
  aba.setColumnWidth(4, 80);
  aba.setColumnWidth(5, 160);
  aba.setFrozenRows(3);
}

// ──────────────────────────────────────────────────────────────
// UTILITÁRIOS
// ──────────────────────────────────────────────────────────────

function fmtData(s) {
  if (!s) return '';
  var p = s.split('-');
  return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : s;
}

function fmtHora(ts) {
  if (!ts) return '';
  var d = new Date(ts);
  return Utilities.formatDate(d, 'America/Sao_Paulo', 'HH:mm');
}

function sanitizar(nome) {
  return nome.replace(/[\\\/\?\*\[\]]/g, '-').substring(0, 100);
}
