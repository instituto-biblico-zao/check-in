// ══════════════════════════════════════════════════════════════
//  INSTITUTO BÍBLICO ZAO — App de Presença
//  Google Apps Script v5.0 — Operações Atômicas
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

var ABA_DADOS = '_dados';

// ──────────────────────────────────────────────────────────────
// ENTRY POINTS
// ──────────────────────────────────────────────────────────────

function doGet(e) {
  var acao = e && e.parameter ? e.parameter.acao : 'ping';
  if (acao === 'carregar') {
    return responder({ ok: true, dados: carregarDados() });
  }
  return responder({ ok: true, mensagem: 'ZAO App de Presença v5.0 — ativo.' });
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var acao = payload.acao;

    if (acao === 'salvar_estrutura') {
      // Salva turmas/aulas/alunos preservando presenças existentes
      salvarEstrutura(payload.dados);
      atualizarPlanilhasVisuais(carregarDados());
      return responder({ ok: true, mensagem: 'Estrutura salva.' });
    }

    if (acao === 'registrar_presenca') {
      // Operação atômica — registra só um aluno
      var resultado = registrarPresenca(
        payload.turmaId,
        payload.aulaId,
        payload.aluno,
        payload.timestamp
      );
      if (resultado) {
        atualizarPlanilhasVisuais(carregarDados());
        return responder({ ok: true, mensagem: 'Presença registrada.' });
      } else {
        return responder({ ok: false, erro: 'Turma ou aula não encontrada.' });
      }
    }

    if (acao === 'cancelar_presenca') {
      // Operação atômica — remove presença de um aluno
      cancelarPresenca(payload.turmaId, payload.aulaId, payload.aluno);
      atualizarPlanilhasVisuais(carregarDados());
      return responder({ ok: true, mensagem: 'Presença cancelada.' });
    }

    if (acao === 'salvar') {
      // Compatibilidade retroativa — salva tudo (usado pelo botão Sincronizar do tablet)
      salvarDadosCompleto(payload.dados);
      atualizarPlanilhasVisuais(carregarDados());
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
// PERSISTÊNCIA
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
  if (!valor || valor === '') return { turmas: [] };
  try { return JSON.parse(valor); } catch(e) { return { turmas: [] }; }
}

function gravarDados(dados) {
  var aba = getAbaDados();
  aba.getRange(1, 1).setValue(JSON.stringify(dados));
}

// ──────────────────────────────────────────────────────────────
// OPERAÇÕES
// ──────────────────────────────────────────────────────────────

function salvarEstrutura(dadosNovos) {
  // Salva turmas/aulas/alunos MAS preserva presenças que já existem no Sheets
  var dadosAtuais = carregarDados();
  var turmasAtuais = dadosAtuais.turmas || [];

  var turmasNovas = (dadosNovos.turmas || []).map(function(turmaNova) {
    var turmaAtual = turmasAtuais.filter(function(t){ return t.id === turmaNova.id; })[0];

    var aulasNovas = (turmaNova.aulas || []).map(function(aulaNova) {
      // Preserva presencas existentes no Sheets
      var aulaAtual = turmaAtual
        ? (turmaAtual.aulas || []).filter(function(a){ return a.id === aulaNova.id; })[0]
        : null;
      var presencasExistentes = aulaAtual ? (aulaAtual.presenca || {}) : {};
      var presencasLocais = aulaNova.presenca || {};

      // Mescla: usa timestamp mais recente para cada aluno
      var merged = {};
      var todosAlunos = Object.keys(presencasExistentes).concat(Object.keys(presencasLocais));
      todosAlunos.forEach(function(aluno) {
        var tsExistente = presencasExistentes[aluno] || 0;
        var tsLocal = presencasLocais[aluno] || 0;
        if (tsExistente > 0 || tsLocal > 0) {
          merged[aluno] = Math.max(tsExistente, tsLocal);
        } else {
          merged[aluno] = Math.min(tsExistente, tsLocal);
        }
      });

      return {
        id: aulaNova.id,
        nome: aulaNova.nome,
        data: aulaNova.data,
        presenca: merged
      };
    });

    return {
      id: turmaNova.id,
      nome: turmaNova.nome,
      cor: turmaNova.cor,
      icone: turmaNova.icone,
      alunos: turmaNova.alunos,
      aulas: aulasNovas,
      aulaAtiva: turmaNova.aulaAtiva
    };
  });

  var dadosFinal = {
    turmas: turmasNovas,
    turmaAtiva: dadosNovos.turmaAtiva
  };

  gravarDados(dadosFinal);
}

function registrarPresenca(turmaId, aulaId, aluno, timestamp) {
  // Lê → modifica só aquele aluno → grava
  var dados = carregarDados();
  var turma = (dados.turmas || []).filter(function(t){ return t.id === turmaId; })[0];
  if (!turma) return false;
  var aula = (turma.aulas || []).filter(function(a){ return a.id === aulaId; })[0];
  if (!aula) return false;
  if (!aula.presenca) aula.presenca = {};
  aula.presenca[aluno] = timestamp;
  gravarDados(dados);
  return true;
}

function cancelarPresenca(turmaId, aulaId, aluno) {
  var dados = carregarDados();
  var turma = (dados.turmas || []).filter(function(t){ return t.id === turmaId; })[0];
  if (!turma) return;
  var aula = (turma.aulas || []).filter(function(a){ return a.id === aulaId; })[0];
  if (!aula || !aula.presenca) return;
  delete aula.presenca[aluno];
  gravarDados(dados);
}

function salvarDadosCompleto(dados) {
  // Usado pelo botão Sincronizar do tablet — mescla antes de salvar
  salvarEstrutura(dados);
}

// ──────────────────────────────────────────────────────────────
// PLANILHAS VISUAIS
// ──────────────────────────────────────────────────────────────

function atualizarPlanilhasVisuais(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var turmas = dados.turmas || [];
  atualizarResumoGeral(ss, turmas);
  turmas.forEach(function(turma) {
    (turma.aulas || []).forEach(function(aula) {
      atualizarAbaPresenca(ss, turma, aula);
    });
  });
}

function atualizarResumoGeral(ss, turmas) {
  var nomeAba = 'Resumo Geral';
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) aba = ss.insertSheet(nomeAba, 0);
  aba.clearContents();
  var cab = ['Turma', 'Aula', 'Data', 'Total Alunos', 'Presentes', 'Ausentes', 'Frequência (%)'];
  var cabRange = aba.getRange(1, 1, 1, cab.length);
  cabRange.setValues([cab]).setFontWeight('bold').setBackground('#1A3A5C').setFontColor('#FFFFFF');
  var linha = 2;
  turmas.forEach(function(turma) {
    (turma.aulas || []).forEach(function(aula) {
      var presenca = aula.presenca || {};
      var alunos = turma.alunos || [];
      var presentes = alunos.filter(function(a){ return presenca[a] > 0; }).length;
      var total = alunos.length;
      var pct = total ? Math.round(presentes / total * 100) : 0;
      aba.getRange(linha, 1, 1, 7).setValues([[
        turma.nome, aula.nome, fmtData(aula.data),
        total, presentes, total - presentes, pct + '%'
      ]]);
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
  var t1 = aba.getRange(1, 1, 1, 5); t1.merge();
  t1.setValue('INSTITUTO BÍBLICO ZAO — ' + turma.nome)
    .setFontSize(13).setFontWeight('bold').setBackground('#1A3A5C').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  var t2 = aba.getRange(2, 1, 1, 5); t2.merge();
  t2.setValue(aula.nome + ' — ' + fmtData(aula.data))
    .setBackground('#2E5F8A').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  var cab = ['Nº', 'Nome do Aluno', 'Status', 'Horário', 'Última atualização'];
  aba.getRange(3, 1, 1, 5).setValues([cab]).setFontWeight('bold').setBackground('#D0E4F7').setFontColor('#1A3A5C');
  var alunos = turma.alunos || [];
  var presenca = aula.presenca || {};
  var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  var presentes = alunos.filter(function(a){ return presenca[a] > 0; })
    .sort(function(a,b){ return presenca[a] - presenca[b]; });
  var ausentes = alunos.filter(function(a){ return !presenca[a] || presenca[a] <= 0; });
  presentes.concat(ausentes).forEach(function(nome, i) {
    var ts = presenca[nome]; var isP = ts > 0; var linha = i + 4;
    aba.getRange(linha, 1, 1, 5).setValues([[i+1, nome, isP?'Presente':'Ausente', isP?fmtHora(ts):'', agora]]);
    var sc = aba.getRange(linha, 3);
    if (isP) sc.setBackground('#EAF3DE').setFontColor('#27500A').setFontWeight('bold');
    else sc.setBackground('#FCEBEB').setFontColor('#A32D2D').setFontWeight('bold');
    var rowBg = (linha % 2 === 0) ? '#F7F6F2' : '#FFFFFF';
    aba.getRange(linha, 1).setBackground(rowBg);
    aba.getRange(linha, 2).setBackground(rowBg);
    aba.getRange(linha, 4).setBackground(rowBg);
    aba.getRange(linha, 5).setBackground(rowBg).setFontColor('#6B6860');
  });
  var ultLinha = alunos.length + 4;
  var totCell = aba.getRange(ultLinha + 1, 1, 1, 5); totCell.merge();
  var pres = presentes.length; var tot = alunos.length;
  var pct = tot ? Math.round(pres/tot*100) : 0;
  totCell.setValue('Total: '+pres+'/'+tot+' presentes ('+pct+'%)')
    .setFontWeight('bold').setBackground('#1A3A5C').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  aba.setColumnWidth(1,40); aba.setColumnWidth(2,220); aba.setColumnWidth(3,90);
  aba.setColumnWidth(4,80); aba.setColumnWidth(5,160); aba.setFrozenRows(3);
}

// ──────────────────────────────────────────────────────────────
// UTILITÁRIOS
// ──────────────────────────────────────────────────────────────

function fmtData(s) {
  if (!s) return '';
  var p = s.split('-');
  return p.length === 3 ? p[2]+'/'+p[1]+'/'+p[0] : s;
}

function fmtHora(ts) {
  if (!ts) return '';
  return Utilities.formatDate(new Date(ts), 'America/Sao_Paulo', 'HH:mm');
}

function sanitizar(nome) {
  return nome.replace(/[\\\/\?\*\[\]]/g, '-').substring(0, 100);
}
