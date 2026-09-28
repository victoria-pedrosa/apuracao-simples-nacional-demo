function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Automações')
      .addItem('▶️ Extrair Dados (Aba 062026 <- Extratos 05/2026)', 'extrairSelecionados')
      .addItem('🕵️ Detetive: Checar Pastas e Arquivos', 'detetivePastas')
      .addToUi();
}

function parseMoney(str) {
  if (!str) return 0;
  var limpo = str.toString().replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.');
  return parseFloat(limpo) || 0;
}

function formatMoney(num) {
  return "R$ " + num.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function obterPastaPrincipal() {
  var id = "ID_EXEMPLO";
  try { return DriveApp.getFolderById(id); } catch (e) {
    var busca = DriveApp.getFoldersByName("1. Impostos");
    if (busca.hasNext()) return busca.next();
  }
  return null;
}

function encontrarPastaMes(pastaEmpresa) {
  var variacoes = ["052026", "202605", "05-2026", "2026-05", "MAIO", "05.2026"];
  var subPastas = pastaEmpresa.getFolders();
  while (subPastas.hasNext()) {
    var sub = subPastas.next();
    var nome = sub.getName().toUpperCase().replace(/\s+/g, '');
    for (var i = 0; i < variacoes.length; i++) {
      if (nome.indexOf(variacoes[i]) !== -1) return sub;
    }
  }
  return null;
}

function buscarPastaEmpresa(pastaPai, nomeEmpresa) {
  var nomeLimpo = nomeEmpresa.toString().replace(/[^\w\sÀ-ÿ]/gi, ' ').replace(/\s+/g, ' ').trim();
  var palavras = nomeLimpo.split(" ").slice(0, 3);
  if (palavras.length > 0) {
    var termoJunto = palavras.join(" ");
    var buscaJunta = pastaPai.searchFolders("title contains '" + termoJunto + "'");
    if (buscaJunta.hasNext()) return buscaJunta.next();
    var querySeparada = palavras.map(function(p) { return "title contains '" + p + "'"; }).join(" and ");
    var buscaSeparada = pastaPai.searchFolders(querySeparada);
    if (buscaSeparada.hasNext()) return buscaSeparada.next();
  }
  return null;
}

function buscarExtrato(pastaPai, pastaEmpresa, codigoDominio) {
  var filtrosNome = ["EXTRATO", "GUIA DAS", "052026", "DECLARACAO", "DECLARAÇÃO"];
  if (pastaEmpresa) {
    var pastaMes = encontrarPastaMes(pastaEmpresa);
    var alvo = pastaMes || pastaEmpresa;
    for (var i = 0; i < filtrosNome.length; i++) {
      var arq = alvo.searchFiles("title contains '" + filtrosNome[i] + "' and mimeType = 'application/pdf'");
      if (arq.hasNext()) return arq.next();
    }
  }
  if (codigoDominio) {
    var arqGlobal = pastaPai.searchFiles("title contains 'EXTRATO' and title contains '" + codigoDominio + "'");
    if (arqGlobal.hasNext()) return arqGlobal.next();
  }
  return null;
}

function detetivePastas() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var pastaPai = obterPastaPrincipal();
  if (!pastaPai) { ui.alert("Erro de acesso ao Drive."); return; }
  var range = sheet.getActiveRange();
  var data = sheet.getRange(range.getRow(), 1, range.getNumRows(), sheet.getLastColumn()).getValues();
  var alertas = [];
  for (var i = 0; i < data.length; i++) {
    var nome = data[i][0], cod = data[i][3];
    if (!nome) continue;
    var pasta = buscarPastaEmpresa(pastaPai, nome);
    var pdf = buscarExtrato(pastaPai, pasta, cod);
    if (!pasta) alertas.push("❌ " + nome + ": Pasta não encontrada.");
    else if (!pdf) alertas.push("⚠️ " + nome + ": Arquivo não achado.");
  }
  ui.alert("Resultado", alertas.length > 0 ? alertas.join("\n") : "✅ Tudo OK nas linhas selecionadas!", ui.ButtonSet.OK);
}

// ─────────────────────────────────────────────────────────────
//  HELPERS DE STATUS VISUAL (gravam imediatamente na célula)
// ─────────────────────────────────────────────────────────────

/**
 * Grava status de progresso na linha imediatamente,
 * sem aguardar o fim do loop.
 */
function gravarStatus(sheet, row, txtAnexo, valorRBT, status) {
  // Grava as três colunas de uma vez (J, K, L) em batch de 1 linha — muito mais rápido que 3 setValue separados
  sheet.getRange(row, 10, 1, 3).setValues([[txtAnexo, valorRBT, status]]);
  // Força o flush imediato para o usuário ver na hora
  SpreadsheetApp.flush();
}

/**
 * Marca a linha como "em processamento" logo no início,
 * assim o usuário vê que ela está sendo trabalhada.
 */
function marcarEmProcessamento(sheet, row) {
  sheet.getRange(row, 10, 1, 3).setValues([["⏳ Processando...", "", ""]]);
  SpreadsheetApp.flush();
}

// ─────────────────────────────────────────────────────────────
//  FUNÇÃO PRINCIPAL
// ─────────────────────────────────────────────────────────────

function extrairSelecionados() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var pastaPai = obterPastaPrincipal();
  if (!pastaPai) return;

  var range = sheet.getActiveRange();
  var startRow = range.getRow();
  // Lê todos os dados de uma vez só (evita múltiplas chamadas ao Sheets API)
  var data = sheet.getRange(startRow, 1, range.getNumRows(), sheet.getLastColumn()).getValues();

  for (var i = 0; i < data.length; i++) {
    var nome = data[i][0], cod = data[i][3];
    if (!nome) continue;

    var row = startRow + i;

    // ✅ MUDANÇA PRINCIPAL: avisa o usuário que essa linha está sendo processada AGORA
    marcarEmProcessamento(sheet, row);

    var tempFileId = null;

    try {
      var pasta = buscarPastaEmpresa(pastaPai, nome);
      var pdf   = buscarExtrato(pastaPai, pasta, cod);

      if (!pdf) {
        // Grava resultado negativo imediatamente e passa para a próxima
        gravarStatus(sheet, row, "❌ PDF não encontrado", "", "ERRO");
        continue;
      }

      // Converte PDF → Google Doc para extração de texto
      var tempFile = Drive.Files.copy(
        { title: "Temp_" + nome, mimeType: MimeType.GOOGLE_DOCS },
        pdf.getId(),
        { convert: true, supportsAllDrives: true }
      );
      tempFileId = tempFile.id;

      var doc   = DocumentApp.openById(tempFileId);
      var texto = doc.getBody().getText();

      // ── 1. RPA e RBT12 ──────────────────────────────────────
      var rpa = 0, rbt12 = 0;
      var mRPA = texto.match(/Receita Bruta do PA \(RPA\)[\s\S]{0,50}?(Competência|Caixa)[\s\S]{0,30}?([\d\.]+(?:,\d{2}))/i);
      if (mRPA) rpa = parseMoney(mRPA[2]);

      var mRBT = texto.match(/RBT12\)[\s\S]{0,40}?([\d\.]+(?:,\d{2}))/i);
      if (mRBT) rbt12 = parseMoney(mRBT[1]);

      // ── 2. Meses Abertos via Tabela 2.2 ─────────────────────
      var mesesAberto = 12;
      var blocoReceitas = texto.match(/2\.2.*?Receitas Brutas Anteriores[\s\S]*?(?=2\.3)/i);

      if (blocoReceitas) {
        var matchesMeses = blocoReceitas[0].match(/\d{2}\/\d{4}/g);
        if (matchesMeses) {
          var mesesUnicos = matchesMeses.filter(function(v, idx, arr) { return arr.indexOf(v) === idx; });
          if (mesesUnicos.length < 12) mesesAberto = mesesUnicos.length + 1;
        }
      }

      // ── 3. Cálculo Final RBT12 ───────────────────────────────
      var finalRBTValue = 0;
      var statusConf    = "OK";

      if (mesesAberto < 12) {
        finalRBTValue = ((rbt12 + rpa) / mesesAberto) * 12;
        statusConf    = "OK-Prop";
      } else {
        var antigo = 0;
        if (blocoReceitas) {
          var rAnt = /04\/2025[\s\S]{0,20}?([\d\.]+(?:,\d{2}))/g, m;
          while ((m = rAnt.exec(blocoReceitas[0])) !== null) antigo += parseMoney(m[1]);
        }
        finalRBTValue = (rbt12 + rpa) - antigo;
      }

      // ── 4. Anexos ────────────────────────────────────────────
      var anexos = [];
      var mAnx, rAnx = /Anexo\s+(VI|V|IV|III|II|I)\b/gi;
      while ((mAnx = rAnx.exec(texto)) !== null) {
        var nomeAnx = "ANEXO " + mAnx[1].toUpperCase();
        if (anexos.indexOf(nomeAnx) === -1) anexos.push(nomeAnx);
      }

      var txtAnexo = anexos.length > 0
        ? anexos.join(", ")
        : (pdf.getName().indexOf("DECLARA") !== -1 ? "SEM FATURAMENTO" : "NÃO ENCONTRADO");

      // ── 5. ✅ Gravação IMEDIATA assim que termina essa empresa ─
      gravarStatus(sheet, row, txtAnexo, formatMoney(finalRBTValue), statusConf);

    } catch (e) {
      // Mesmo em erro, grava imediatamente para o usuário saber
      gravarStatus(sheet, row, "❌ ERRO: " + e.message, "", "ERRO");
      Logger.log(nome + ": " + e.message);
    } finally {
      // Limpa o arquivo temporário sempre, com ou sem erro
      if (tempFileId) {
        try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (e2) {}
      }
    }
  }

  ui.alert("✅ Processamento Concluído!");
}